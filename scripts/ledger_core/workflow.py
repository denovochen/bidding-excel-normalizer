"""有界结构语义交接；由 Python 编译、验证并持久化解析计划。 @author denovochen"""
from __future__ import annotations

import copy
import json
import re
import shlex
import sys
from pathlib import Path

from .contract import VERSION, ROLES, BUSINESS_ROLES, LedgerError, MappingRevisionRequired, clean, column_label, column_number, stable_id, text
from .normalize import build_outputs, collect_table
from .relations import infer_table_kind
from .review import create_work_directory, save_state, create_state, load_state, pending_questions, review_result
from .structure import source_layout, _block_summary, simple_lot
from .workbook import (HEADER_NAMES, _row_header_matches, _header_text, _effective_columns, _has_company_list,
                       inspect_workbooks, read_workbook, describe_source_failure, header_extent)
from .contract import RecoverableWorkbookError
from .artifacts import publish

BATCH_SIZE = 3
MAX_RESPONSE_CHARS = 24000
MAX_ATTEMPTS = 2


def command(name: str, **paths) -> str:
    entry = Path(__file__).resolve().parents[1] / "excel_ledger.py"
    args = [sys.executable, str(entry), name]
    for key, value in paths.items():
        args.extend(["--" + key.replace("_", "-"), str(value)])
    return shlex.join(args)


def _regions(book, sheet) -> list[dict]:
    rows = sheet.row_numbers
    if not rows:
        return []
    candidates = []
    for row in rows:
        matches = _row_header_matches(sheet, row)
        if BUSINESS_ROLES.intersection(matches) and (len(matches) >= 2 or
                {"bidder_name", "award_name"}.intersection(matches)):
            candidates.append(row)
    runs = []
    for row in candidates:
        if runs and row == runs[-1][-1] + 1 and len(runs[-1]) < 3:
            runs[-1].append(row)
        else:
            runs.append([row])
    if not runs:
        runs = [[next((row for row in rows[:10] if len(sheet.populated_columns_by_row.get(row, [])) >= 2), rows[0])]]
    regions = []
    for index, headers in enumerate(runs):
        detected = {}
        for row in headers:
            detected.update(_row_header_matches(sheet, row))
        end = runs[index + 1][0] - 1 if index + 1 < len(runs) else sheet.max_row
        headers = [row for row in header_extent(sheet, headers, detected) if row <= end]
        labels = [column_label(col) for col in _effective_columns(sheet, max(headers) + 1, end, headers)]
        # 同一表头重复出现某角色时不采用“最后一列胜出”。
        ambiguous = []
        for role, label in detected.items():
            values = {column_label(col) for row in headers for col in sheet.populated_columns_by_row.get(row, [])
                      if clean(sheet.raw(row, col).value).replace(" ", "") in HEADER_NAMES[role]}
            if len(values) > 1:
                ambiguous.append(role)
        for role in ambiguous:
            detected.pop(role)
        profiles = []
        for label in labels:
            col = column_number(label)
            populated = [row for row in rows if max(headers) < row <= end and text(sheet.raw(row, col).value)]
            profiles.append({"column": label, "header": _header_text(sheet, headers, col)[:100],
                             "samples": [{"cell": f"{label}{row}", "text": text(sheet.raw(row, col).value)[:72]}
                                         for row in populated[:2]], "nonempty": len(populated),
                             "formula_count": sum(sheet.raw(row, col).formula for row in populated),
                             "error_count": sum(sheet.raw(row, col).error for row in populated),
                             "suggested_role": next((role for role, value in detected.items() if value == label), "unresolved")})
        signature = [(profile["header"], profile["suggested_role"]) for profile in profiles]
        regions.append({"id": stable_id("region", book.sha256, sheet.index, index),
                        "source_id": book.source_id, "sheet_index": sheet.index, "sheet": sheet.name,
                        "source_file": book.path.name, "start": headers[0], "end": end,
                        "headers": headers, "columns": profiles, "detected": detected,
                        "cohort": stable_id("structure", signature), "ambiguous_roles": ambiguous})
    return regions


def _book_sheet(books, region):
    book = next(book for book in books if book.source_id == region["source_id"])
    return book, book.sheets[region["sheet_index"]]


def _table(region, answer, books, verify=True) -> dict:
    allowed = {"action", "columns", "record_layout", "project_context_columns", "lot_context_columns",
               "award_completeness", "basis", "header_rows"}
    if not isinstance(answer, dict) or set(answer) - allowed:
        raise LedgerError("结构答案只接受列含义、记录布局、项目上下文来源、完整性和依据；不能设置分组参数")
    if answer.get("action") != "interpret" or not isinstance(answer.get("basis"), str) or not answer["basis"].strip():
        raise LedgerError("结构判断必须包含 action=interpret 和非空 basis")
    mapping = answer.get("columns", {})
    labels = {item["column"] for item in region["columns"]}
    if not isinstance(mapping, dict) or set(mapping) != labels:
        raise LedgerError("必须为本区域每一列指定角色或 evidence，不能漏列或增加来源外的列")
    if any(not isinstance(role, str) or role not in ROLES | {"evidence", "group_context"} for role in mapping.values()):
        raise LedgerError("列角色无效；尚不能解释的区域请选择 review")
    roles = [role for role in mapping.values() if role not in {"evidence", "group_context"}]
    if len(roles) != len(set(roles)) or not BUSINESS_ROLES.intersection(roles):
        raise LedgerError("业务角色不能重复，且必须包含业务字段")
    book, sheet = _book_sheet(books, region)
    headers = answer.get("header_rows", region["headers"])
    if (not isinstance(headers, list) or not headers or any(type(row) is not int for row in headers) or
            min(headers) < region["start"] or max(headers) >= min(region["end"] + 1, region["start"] + 10)):
        raise LedgerError("表头行必须来自区域开头十行以内的已展示来源")
    if headers != list(range(min(headers), max(headers) + 1)) or max(headers) == region["end"]:
        raise LedgerError("表头必须连续且后面存在数据行")
    profiles = [{**item, "header": _header_text(sheet, headers, column_number(item["column"]))}
                for item in region["columns"]]
    if headers == region["headers"] and BUSINESS_ROLES.intersection(region["ambiguous_roles"]):
        raise LedgerError("区域含并排重复业务列，不能选择一列后静默丢弃其他列；需要区域复核")
    layout = answer.get("record_layout")
    if not isinstance(layout, str) or layout not in {"bidder_rows", "lists_in_cells", "award_rows", "company_list", "project_rows"}:
        raise LedgerError("record_layout 必须为 bidder_rows/lists_in_cells/award_rows/company_list/project_rows")
    columns = {role: label for label, role in mapping.items() if role not in {"evidence", "group_context"}}
    for profile in profiles:
        if mapping[profile["column"]] not in BUSINESS_ROLES:
            continue
        duplicates = [item for item in profiles if item["column"] != profile["column"] and
                      item["header"] and item["header"] == profile["header"]]
        if duplicates:
            raise LedgerError("重复业务表头需要先区分区域，不能将另一份业务列降为 evidence")
    serial = region["detected"].get("bidder_serial")
    if serial and mapping[serial] not in {"bidder_serial", "evidence"}:
        raise LedgerError("投标企业序号不能解释为项目序号或排名")
    if layout == "award_rows" and ("bidder_name" in columns or "award_name" not in columns):
        raise LedgerError("award_rows 只能用于仅列中标企业的区域")
    if layout == "project_rows" and ({"bidder_name", "award_name"}.intersection(columns) or
                                     not {"project_name", "project_code", "lot_name", "lot_code"}.intersection(columns)):
        raise LedgerError("project_rows 只用于没有企业字段的项目/标段上下文区域")
    if layout in {"bidder_rows", "company_list"} and "bidder_name" not in columns:
        raise LedgerError("投标行/企业名单区域必须映射 bidder_name")
    for role in ("bidder_name", "bidder_serial", "lot_name", "lot_code"):
        label = region["detected"].get(role)
        if label and mapping[label] == "evidence" and (role != "bidder_serial" or "bidder_name" in columns):
            raise LedgerError(f"{label} 已有 {role} 表头依据，不能直接降为 evidence；请确认语义或转复核")
    completeness = answer.get("award_completeness", "unknown")
    if not isinstance(completeness, str) or completeness not in {"complete", "partial", "unknown"}:
        raise LedgerError("award_completeness 必须为 complete/partial/unknown")
    context = answer.get("project_context_columns", [])
    if not isinstance(context, list) or any(not isinstance(label, str) or mapping.get(label) not in {"lot_name", "lot_code", "notes"} for label in context):
        raise LedgerError("项目上下文来源只能引用已映射的标段/备注列")
    project_role = next((role for role in ("project_name", "project_code") if role in columns), None)
    if context and not project_role:
        raise LedgerError("备用项目上下文需要同时映射主要项目字段")
    lot_context = answer.get("lot_context_columns", [])
    if (not isinstance(lot_context, list) or any(not isinstance(label, str) or mapping.get(label) not in {"lot_name", "lot_code"}
                                               for label in lot_context) or set(lot_context) & set(context)):
        raise LedgerError("lot_context_columns 只接受已映射且未兼作项目的标段列")
    project_mode = "none"
    if project_role:
        project_col = column_number(columns[project_role])
        populated = [row for row in sheet.row_numbers if max(headers) < row <= region["end"]]
        sparse = any(not text(sheet.raw(row, project_col).value) for row in populated)
        project_mode = "blocks" if "bidder_name" in columns and (sparse or context) else "repeated"
    group_mode = "source_blocks" if "bidder_name" in columns else "row"
    if layout == "company_list":
        if project_role or {"lot_name", "lot_code", "award_name", "award_status", "bidder_count"}.intersection(columns):
            raise LedgerError("纯企业名单不能同时声明项目、标段或中标业务角色")
        group_mode = "source"
    table = {"table_id": region["id"], "table_kind": infer_table_kind(columns),
             "header_rows": headers, "data_start_row": max(headers) + 1, "data_end_row": region["end"],
             "columns": columns, "column_dispositions": [
                 {"column": item["column"], "header": item["header"], "disposition": mapping[item["column"]],
                  **({"mode": "blocks"} if mapping[item["column"]] == "group_context" else {}),
                  "reason": "结构语义判断确认的辅助来源列，区域依据见 award_completeness.basis"}
                 for item in profiles if mapping[item["column"]] in {"evidence", "group_context"}],
             "project_mode": project_mode, "group_mode": group_mode,
             "bidder_separator": "delimited" if layout == "lists_in_cells" else "single",
             "award_mode": "auto", "award_completeness": {
                 "status": completeness, "basis_type": "structural" if completeness != "unknown" else "none",
                 "basis": answer["basis"]}, "structure_warnings": [],
             "summary_markers": ["项目汇总", "合计", "小计", "总计"], "non_tender_markers": ["未招投标", "未招标"]}
    if context:
        table["project_context_fields"] = [mapping[label] for label in context]
    if lot_context:
        table["lot_context_fields"] = [mapping[label] for label in lot_context]
    # 语义答案通过真实来源约束才可以进入运行计划。
    if verify:
        collect_table(book, sheet, table, 0, {})
    return table


def _question(cohort, regions, books, state) -> dict:
    representative = regions[0]
    detected = representative["detected"]
    company = detected.get("bidder_name") or detected.get("award_name")
    book, sheet = _book_sheet(books, representative)
    layout = "award_rows" if "award_name" in detected and "bidder_name" not in detected else "bidder_rows"
    if not company and {"project_name", "project_code", "lot_name", "lot_code"}.intersection(detected):
        layout = "project_rows"
    if company and any(_has_company_list(text(sheet.raw(row, column_number(company)).value))
                       for row in sheet.row_numbers if representative["start"] < row <= representative["end"]):
        layout = "lists_in_cells"
    elif "bidder_name" in detected and not {"project_name", "project_code", "lot_name", "lot_code", "award_name"}.intersection(detected):
        layout = "company_list"
    alternatives = []
    for region in regions:
        _, source_sheet = _book_sheet(books, region)
        project = region["detected"].get("project_name")
        for role in ("lot_name", "lot_code", "notes"):
            col = region["detected"].get(role)
            if not project or not col:
                continue
            for row in source_sheet.row_numbers:
                if not max(region["headers"]) < row <= region["end"]:
                    continue
                value = source_sheet.raw(row, column_number(col)).value
                if value and not simple_lot(value) and not text(source_sheet.resolved(row, column_number(project))[0].value):
                    alternatives.append({"region": region["id"], "cell": f"{col}{row}", "text": text(value)[:160]})
                    break
    mapping = {item["column"]: item["suggested_role"] for item in representative["columns"]}
    answer = {"action": "interpret", "columns": mapping, "record_layout": layout,
              "project_context_columns": [], "lot_context_columns": [],
              "award_completeness": "unknown", "basis": "填写区域语义判断依据"}
    candidate = copy.deepcopy(answer)
    candidate["columns"] = {key: "evidence" if role == "unresolved" else role for key, role in mapping.items()}
    boundaries = []
    try:
        table = _table(representative, candidate, books, verify=False)
        boundaries = [_block_summary(block) for block in source_layout(book, sheet, table)["blocks"][:3]]
    except LedgerError:
        pass
    question = {"question_id": cohort, "task_type": "structure", "owner": "model",
            "regions": [{key: region[key] for key in ("id", "source_file", "sheet", "start", "end", "headers")}
                        for region in regions[:12]], "region_count": len(regions),
            "regions_truncated": len(regions) > 12, "columns": representative["columns"],
            "header_preview": [sheet.row_view(row) for row in sheet.row_numbers
                               if representative["start"] <= row <= min(representative["end"], representative["start"] + 3)],
            "source_block_examples": boundaries, "project_context_candidates": alternatives[:5],
            "ambiguous_roles": representative["ambiguous_roles"], "answer_template": answer,
            "review_answer": {"action": "review", "basis": "填写当前区域无法可靠解释的具体原因"},
            "previous_error": state["errors"].get(cohort),
            "inspect_command": command("inspect", state=state["state_path"], region=representative["id"],
                                       rows=f"{representative['start']}:{min(representative['start'] + 19, representative['end'])}")}
    if len(json.dumps(question, ensure_ascii=False)) > MAX_RESPONSE_CHARS - 3000:
        question["columns"] = [{"column": item["column"], "header": item["header"][:16]}
                               for item in representative["columns"]]
        question["column_profiles_compacted"] = True
        question["header_preview"] = []
        question["source_block_examples"] = []
        question["inspect_columns_hint"] = "inspect_command 可增加 --columns A:P 分页查看完整列头和样例，每页最多16列"
        question["regions"] = question["regions"][:1]
        question["regions_truncated"] = len(regions) > 1
    return question


def _plan(state, books) -> dict:
    sources = []
    for book in books:
        sheets = []
        for sheet in book.sheets:
            regions = [region for region in state["regions"] if region["source_id"] == book.source_id and region["sheet_index"] == sheet.index]
            if not regions:
                sheets.append({"name": sheet.name, "action": "skip", "reason": "空白工作表"})
                continue
            tables, reviews, ignored = [], [], []
            first = min(region["start"] for region in regions)
            if first > 1:
                ignored.append({"start": 1, "end": first - 1, "reason": "已确认表头之前的标题区域"})
            for region in regions:
                decision = state["interpretations"][region["cohort"]]
                if decision["action"] == "review":
                    reviews.append({"start": region["start"], "end": region["end"], "reason": decision["basis"]})
                else:
                    compiled = copy.deepcopy(state["compiled_tables"][region["id"]])
                    compiled = compiled if isinstance(compiled, list) else [compiled]
                    tables.extend(compiled)
                    reviews.extend(state.get("region_reviews", {}).get(region["id"], []))
                    if compiled and min(compiled[0]["header_rows"]) > region["start"]:
                        ignored.append({"start": region["start"], "end": min(compiled[0]["header_rows"]) - 1,
                                        "reason": "已确认表头前的区域标题"})
            if tables:
                sheets.append({"name": sheet.name, "action": "parse", "tables": tables,
                               "ignored_rows": ignored, "review_regions": reviews})
            else:
                sheets.append({"name": sheet.name, "action": "review", "reason": "; ".join(x["reason"] for x in reviews)})
        sources.append({"file_name": book.path.name, "sha256": book.sha256, "sheets": sheets})
    return {"schema_version": 1, "sources": sources}


def _relationship_questions(state, plan) -> list[dict]:
    tables = {table["table_id"]: table for source in plan["sources"] for sheet in source["sheets"]
              for table in sheet.get("tables", [])}
    cohorts = {}
    for region in state["regions"]:
        compiled = state["compiled_tables"].get(region["id"], [])
        compiled = compiled if isinstance(compiled, list) else [compiled]
        ids = [item["table_id"] for item in compiled if item["table_id"] in tables]
        table = tables[ids[0]] if ids else None
        if table and {"project_name", "project_code"}.intersection(table["columns"]):
            cohorts.setdefault(region["cohort"], []).append({**region, "compiled_ids": ids})
    bidder_sets = [{"set_id": key, "tables": [table_id for region in regions for table_id in region["compiled_ids"]],
                   "sources": [region["sheet"] for region in regions]}
                  for key, regions in cohorts.items()
                  if tables[regions[0]["compiled_ids"][0]]["table_kind"] in {"bidder_roster", "complete_results"}]
    if not bidder_sets:
        return []
    return [{"question_id": "connect_" + key, "task_type": "table_relationship", "owner": "model",
             "award_tables": [table_id for region in regions for table_id in region["compiled_ids"]],
             "sources": [region["sheet"] for region in regions], "candidate_bidder_sets": bidder_sets,
             "answer_template": {"bidder_sets": [], "basis": "按区域业务含义选择同类招标名册；无对应名册保持空数组"}}
            for key, regions in cohorts.items() if tables[regions[0]["compiled_ids"][0]]["table_kind"] == "award_summary"]


def handoff(state: dict, books) -> dict:
    cohorts = {}
    for region in state["regions"]:
        if region["cohort"] not in state["interpretations"]:
            cohorts.setdefault(region["cohort"], []).append(region)
    questions = [_question(key, regions, books, state) for key, regions in list(cohorts.items())[:BATCH_SIZE]]
    if not cohorts:
        plan = _plan(state, books)
        questions = [question for question in _relationship_questions(state, plan)
                     if question["question_id"] not in state["connections"]]
    remaining = len(cohorts) if cohorts else len(questions)
    selected = []
    for question in questions[:BATCH_SIZE]:
        if selected and len(json.dumps(selected + [question], ensure_ascii=False)) > MAX_RESPONSE_CHARS - 3000:
            break
        selected.append(question)
    result = {"kind": "mapping_required", "stage": "structure", "state": state["state_path"],
              "remaining_task_count": remaining, "questions": selected,
              "message": "只返回 question_id 到结构语义答案的 JSON；保留有依据的建议，未知列必须解释。无需读取源码或编写 plan。",
              "answer_roles": sorted(ROLES | {"evidence", "group_context"}),
              "next_command": command("resolve", state=state["state_path"], answers=Path(state["state_path"]).with_name("answers.json"))}
    state["issued_questions"] = [question["question_id"] for question in selected]
    save_state(Path(state["state_path"]), state)
    return result


def start(inputs, books, failures, output) -> dict:
    directory = create_work_directory(output)
    inspection = inspect_workbooks(books, failures)
    inspection_path = directory / "inspection.json"
    inspection_path.write_text(json.dumps(inspection, ensure_ascii=False), encoding="utf-8")
    regions = [region for book in books for sheet in book.sheets for region in _regions(book, sheet)]
    state = {"stage": "structure", "schema_version": 1, "parser_version": VERSION,
             "inputs": [str(path.expanduser().resolve(strict=True)) for path in inputs],
             "source_hashes": {book.source_id: book.sha256 for book in books},
             "failures": failures, "output": str(output.absolute()), "regions": regions,
             "interpretations": {}, "compiled_tables": {}, "connections": {}, "errors": {}, "attempts": {},
             "state_path": str(directory / "structure.json")}
    reply = handoff(state, books)
    # 旧 plan 客户端仍可读取完整 inspection，正常流程使用 questions。
    reply["inspection"] = {"inspection_path": str(inspection_path),
                           "relationship_candidates": inspection["relationship_candidates"],
                           "source_failures": failures}
    return reply


def read_inputs(state):
    books, failures = [], []
    for path in state["inputs"]:
        try:
            books.append(read_workbook(Path(path)))
        except RecoverableWorkbookError as exc:
            failures.append(describe_source_failure(Path(path), str(exc)))
    if ({book.source_id: book.sha256 for book in books} != state["source_hashes"] or
            failures != state["failures"]):
        raise LedgerError("输入文件已变化，不能继续使用旧结构判断")
    return books, failures


def _quarantine_failed_regions(state, regions, answer, books, message):
    """同结构区域中仅隔离失败成员；有明确块边界时保留前后可解析部分。"""
    state.setdefault("region_reviews", {})
    for region in regions:
        try:
            state["compiled_tables"][region["id"]] = _table(region, answer, books)
            continue
        except LedgerError as exc:
            evidence = getattr(exc, "evidence", {})
        reviews = [{"start": region["start"], "end": region["end"], "reason": message}]
        kept = []
        try:
            base = _table(region, answer, books, verify=False)
            book, sheet = _book_sheet(books, region)
            blocks = source_layout(book, sheet, base, strict=False)["blocks"]
            failed = next((index for index, block in enumerate(blocks) if evidence.get("row") in block["rows"]), None)
            if failed is not None:
                def beginning(block):
                    return min([block["anchor_row"]] + [int(re.search(r"\d+$", item["cell"])[0])
                                                         for item in block["evidence"]])
                start = max(base["data_start_row"], beginning(blocks[failed]))
                end = beginning(blocks[failed + 1]) - 1 if failed + 1 < len(blocks) else base["data_end_row"]
                for left, right in ((base["data_start_row"], start - 1), (end + 1, base["data_end_row"])):
                    if left > right:
                        continue
                    table = copy.deepcopy(base)
                    table.update(table_id=stable_id("table", region["id"], left, right),
                                 data_start_row=left, data_end_row=right)
                    collect_table(book, sheet, table, len(kept), {})
                    kept.append(table)
                if kept:
                    reviews = [{"start": start, "end": end, "reason": message}]
        except LedgerError:
            kept = []
        state["compiled_tables"][region["id"]] = kept
        state["region_reviews"][region["id"]] = reviews


def resume(path: Path, state: dict, answers: dict) -> dict:
    if state.get("parser_version") != VERSION or state.get("schema_version") != 1:
        raise LedgerError("结构状态版本无效")
    state["state_path"] = str(path.resolve(strict=True))
    books, failures = read_inputs(state)
    if not isinstance(answers, dict) or set(answers) - set(state["issued_questions"]):
        raise LedgerError("结构答案只能引用本批已返回的问题")
    for key, answer in answers.items():
        regions = [region for region in state["regions"] if region["cohort"] == key]
        try:
            if regions:
                if isinstance(answer, dict) and answer.get("action") == "review":
                    if set(answer) != {"action", "basis"} or not isinstance(answer["basis"], str) or not answer["basis"].strip():
                        raise LedgerError("review 必须提供具体 basis")
                else:
                    for region in regions:
                        state["compiled_tables"][region["id"]] = _table(region, answer, books)
                state["interpretations"][key] = answer
            else:
                question = next(question for question in _relationship_questions(state, _plan(state, books))
                                if question["question_id"] == key)
                candidates = {item["set_id"] for item in question["candidate_bidder_sets"]}
                if (not isinstance(answer, dict) or set(answer) != {"bidder_sets", "basis"} or
                        not isinstance(answer["bidder_sets"], list) or
                        any(not isinstance(value, str) for value in answer["bidder_sets"]) or
                        len(set(answer["bidder_sets"])) != len(answer["bidder_sets"]) or
                        not set(answer["bidder_sets"]) <= candidates or
                        not isinstance(answer["basis"], str) or not answer["basis"].strip()):
                    raise LedgerError("跨表语义答案只能选择已给出的 bidder_sets 并提供 basis")
                state["connections"][key] = answer
            state["errors"].pop(key, None)
        except LedgerError as exc:
            state["attempts"][key] = state["attempts"].get(key, 0) + 1
            evidence = getattr(exc, "evidence", {})
            if len(json.dumps(evidence, ensure_ascii=False)) > 1600:
                evidence = {"excerpt": json.dumps(evidence, ensure_ascii=False)[:1600], "truncated": True}
            state["errors"][key] = {"message": str(exc)[:1000], "evidence": evidence}
            if regions and state["attempts"][key] >= MAX_ATTEMPTS:
                reason = "两次结构判断未通过来源约束: " + str(exc)
                state["interpretations"][key] = {"action": "interpret", "basis": reason}
                _quarantine_failed_regions(state, regions, answer, books, reason)
            elif not regions and state["attempts"][key] >= MAX_ATTEMPTS:
                save_state(path, state)
                raise LedgerError("表关系答案连续两次不符合候选契约；可保留原 state 修正，无产物发布")
    reply = handoff(state, books)
    if reply["remaining_task_count"]:
        return reply
    plan = _plan(state, books)
    relationships = []
    for question in _relationship_questions(state, plan):
        answer = state["connections"][question["question_id"]]
        targets = [table for candidate in question["candidate_bidder_sets"]
                   if candidate["set_id"] in answer["bidder_sets"] for table in candidate["tables"]]
        if targets:
            relationships.append({"id": question["question_id"], "kind": "award_to_bidder_roster",
                                  "award_tables": question["award_tables"], "bidder_tables": targets,
                                  "project_keys": ["project_code", "project_name"]})
    if relationships:
        plan["relationships"] = relationships
    plan_path = path.with_name("plan.json")
    plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
    prepared = build_outputs(books, plan, source_failures=failures)
    _, count, remaining = pending_questions(prepared[2], {})
    if remaining:
        state_path = create_state([Path(value) for value in state["inputs"]], plan, Path(state["output"]),
                                  prepared[2]["generated_at"], count, work_directory=path.parent)
        return review_result(state_path, prepared[2], load_state(state_path))
    return publish(Path(state["output"]), *prepared)


def inspect_region(path: Path, state: dict, region_id: str, span: str, columns: str | None = None) -> dict:
    books, _ = read_inputs(state)
    region = next((region for region in state["regions"] if region["id"] == region_id), None)
    if not region:
        raise LedgerError("区域不存在")
    start, end = [int(value) for value in span.split(":")]
    if start < region["start"] or end > region["end"] or end < start or end - start >= 20:
        raise LedgerError("补证据范围必须位于当前区域内，每次最多20行")
    _, sheet = _book_sheet(books, region)
    left, right = [column_number(value) for value in columns.split(":")] if columns else [1, min(16, sheet.max_col)]
    if right < left or right > sheet.max_col or right - left >= 16:
        raise LedgerError("列范围须位于当前 Sheet 内，每页最多16列")
    return {"kind": "structure_evidence", "region": region_id,
            "columns": [{"column": column_label(col), "header": _header_text(sheet, region["headers"], col)[:160]}
                        for col in range(left, right + 1)],
            "rows": [{f"{column_label(col)}{row}": text(sheet.raw(row, col).value)[:48]
                      for col in range(left, right + 1) if text(sheet.raw(row, col).value)} for row in range(start, end + 1)],
            "values_may_be_truncated": True,
            "next_column_page": (f"{column_label(right + 1)}:{column_label(min(right + 16, sheet.max_col))}"
                                 if right < sheet.max_col else None),
            "merges": [bounds for bounds in sheet.merges if bounds[0] <= end and bounds[2] >= start][:20],
            "next_command": command("resolve", state=path, answers=path.with_name("answers.json"))}
