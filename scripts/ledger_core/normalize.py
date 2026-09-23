"""按验证后的映射清洗企业参与记录，保留原始值与纠错依据。 @author denovochen"""
from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from typing import Any

from .contract import (FIELDS, VERSION, LedgerError, MappingRevisionRequired, clean, column_number, coordinate,
                       name_key, stable_id, text)
from .workbook import HEADER_NAMES, Workbook, Sheet, validate_plan
from .names import POLICY, match_group
from .relations import apply_relationships, infer_table_kind, table_identity

EMPTY_NAMES = {"", "/", "-", "—", "无", "暂无", "未招标", "未招投标", "未确定", "待定", "未中标", "否", "不适用", "待招标", "未开标"}
COMPANY_END = re.compile(r"(?:公司|工程队|工程处|合作社|事务所|中心|研究院|设计院|厂|经营部)$")


def procurement_signal(notes: str, markers: list[str]) -> str | None:
    """只把明确陈述作为未招投标依据；否定、疑问与矛盾分别保留。"""
    signals = set()
    for clause in re.split(r"[，,；;。\n]", clean(notes)):
        for marker in markers:
            if marker not in clause:
                continue
            for match in re.finditer(re.escape(marker), clause):
                before, after = clause[:match.start()].strip(), clause[match.end():].strip()
                if re.search(r"是否|可能|疑似|待核|不确定|不清楚|不能认定|未核实|未能核实|没有证据|尚待", clause) or "?" in clause:
                    signals.add("uncertain")
                elif re.search(r"(?:不存在|没有|并非|不是|并无|未发现|未发生|未出现)(?:任何)?$", before) or re.match(r"(?:情形|情况|问题)?(?:不存在|未发生|未出现)", after):
                    signals.add("not_non_tender")
                else:
                    signals.add("non_tender")
    if "uncertain" in signals or {"not_non_tender", "non_tender"}.issubset(signals):
        return "uncertain"
    return next(iter(signals)) if signals else None


def aliases_from_json(payload: dict[str, Any]) -> dict[str, dict[str, str]]:
    if set(payload) != {"schema_version", "aliases"} or payload["schema_version"] != 1 or not isinstance(payload["aliases"], list):
        raise LedgerError("别名规则格式无效")
    result = {}
    for item in payload["aliases"]:
        if not isinstance(item, dict) or set(item) != {"from", "to", "basis"}:
            raise LedgerError("每条别名必须包含 from/to/basis")
        if any(not isinstance(v, str) or not v.strip() for v in item.values()):
            raise LedgerError("别名名称和依据不能为空")
        key, target = name_key(item["from"]), name_key(item["to"])
        if key == target or key in result:
            raise LedgerError("别名自指或存在多个规则")
        result[key] = item
    if any(name_key(item["to"]) in result for item in result.values()):
        raise LedgerError("第一版不接受别名链或循环，请直接映射到最终名称")
    return result


def normalize_name(raw: str, aliases: dict[str, dict[str, str]]) -> tuple[str, list[dict[str, str]]]:
    if aliases:
        raise LedgerError("不接受全局公司名称映射")
    # 两端均为汉字的单企业折行不应向全称插入空格。
    folded = re.sub(r"(?<=[\u3400-\u9fff])[\r\n]+\s*(?=[\u3400-\u9fff])", "", raw)
    name = clean(folded)
    actions = []
    if name != raw:
        actions.append({"rule": "format_normalization", "before": raw, "after": name,
                        "basis": "统一全半角、首尾空白和零宽格式字符"})
    if name.endswith("投标") and COMPANY_END.search(name[:-2]):
        changed = name[:-2]
        actions.append({"rule": "trailing_bid_annotation", "before": name, "after": changed,
                        "basis": "完整企业后缀后附加的投标标注"})
        name = changed
    return name, actions


def split_names(value: str, separator: str) -> list[str]:
    # 返回未规范化的名称片段，空白清理交给 normalize_name 留痕。
    value = str(value) if value is not None else ""
    if text(value) in EMPTY_NAMES:
        return []
    if "联合体" in value:
        raise LedgerError("联合体名单需要确认成员与共同投标关系，不能拆成独立投标企业")
    tokens, current, depth = [], [], 0
    for char in value:
        if char in "(（[【":
            depth += 1
        elif char in ")）]】":
            depth -= 1
            if depth < 0:
                raise LedgerError("企业名单括号不配对")
        split = depth == 0 and (char in "、，,；;" or (separator == "lines" and char in "\r\n"))
        if split:
            if separator == "single":
                raise LedgerError("单企业字段出现名单分隔符，请核对 bidder_separator")
            if text("".join(current)):
                tokens.append("".join(current))
            current = []
        else:
            current.append(char)
    if depth:
        raise LedgerError("企业名单括号不配对")
    if text("".join(current)):
        tokens.append("".join(current))
    result = []
    for token in tokens:
        lines = [part.strip() for part in token.splitlines() if part.strip()]
        if len(lines) > 1 and all(COMPANY_END.search(part) for part in lines):
            if separator == "single":
                raise LedgerError("单企业字段含多家企业，请核对映射")
            result.extend(part for part in token.splitlines() if part.strip())
        else:
            # 单家公司因排版换行：合并折行，不拆出伪企业。
            result.append(token)
    if any(text(v) in EMPTY_NAMES for v in result):
        raise LedgerError("企业名单混有空值占位，需核对完整性")
    return result


def _values(sheet: Sheet, row: int, columns: dict[str, str]) -> tuple[dict[str, str], dict[str, str]]:
    values, cells = {}, {}
    for role, label in columns.items():
        value, r, c = sheet.resolved(row, column_number(label))
        values[role] = "" if value.formula or value.error else text(value.value)
        if role in {"project_code", "lot_code"} and re.fullmatch(r"0+", value.number_format):
            if isinstance(value.value, (int, float)) and not isinstance(value.value, bool) and float(value.value).is_integer():
                values[role] = str(int(value.value)).zfill(len(value.number_format))
        cells[role] = coordinate(r, c)
    return values, cells


def _raw_values(sheet: Sheet, row: int, columns: dict[str, str]) -> dict[str, Any]:
    result = {}
    for role, label in columns.items():
        value = sheet.resolved(row, column_number(label))[0].value
        result[role] = value.isoformat() if isinstance(value, datetime) else value
    return result


def _issue(code: str, message: str, group: dict[str, Any] | None, **extra: Any) -> dict[str, Any]:
    return {"code": code, "message": message, "group_id": group["id"] if group else None,
            "record_ids": [], "final_sequences": [], "review_sequences": [], **extra}


def _vertical_merge(sheet: Sheet, row: int, role: str, columns: dict) -> bool:
    if role not in columns:
        return False
    col = column_number(columns[role])
    return any(top < bottom and top <= row <= bottom for top, _, bottom, _ in sheet.merge_columns.get(col, []))


def _award_completeness(table: dict[str, Any]) -> dict[str, str]:
    if "award_completeness" in table:
        return dict(table["award_completeness"])
    complete = table.get("award_list_complete", False)
    return {"status": "complete" if complete else "unknown", "basis_type": "structural" if complete else "none",
            "basis": "兼容旧版布尔映射；重新 inspect 后应提供区域级依据"}


def collect_table(book: Workbook, sheet: Sheet, table: dict[str, Any], table_index: int,
                  aliases: dict[str, dict[str, str]], project_scope_index: int | None = None) -> tuple[list, list, list, list]:
    projects, groups, issues, row_audit = {}, {}, [], []
    current_project = None
    current_group = None
    current_group_context = {}
    columns = table["columns"]
    table_id = table.get("table_id") or table_identity(book.sha256, sheet.index, table_index)
    table_kind = table.get("table_kind") or infer_table_kind(columns)
    award_mode = table.get("award_mode", "auto")
    completeness = _award_completeness(table)
    context_specs = [item for item in table.get("column_dispositions", [])
                     if item["disposition"] == "group_context"]
    audit_specs = [item for item in table.get("column_dispositions", [])
                   if item["disposition"] in {"context", "evidence", "group_context"}]
    price_units = {}
    for role in ("bidder_price", "award_price"):
        if role in columns:
            header = " ".join(text(sheet.resolved(r, column_number(columns[role]))[0].value) for r in table["header_rows"])
            price_units[role] = next((u for u in ("亿元", "万元", "元") if u in header), "")
    project_role = "project_name" if "project_name" in columns else "project_code" if "project_code" in columns else None
    project_col = column_number(columns[project_role]) if project_role else None
    context_roles = ("project_name", "project_code", "project_year", "project_owner", "project_serial")
    if table["group_mode"] == "anchor" and table.get("group_start_field") == "bidder_count" and "bidder_name" in columns:
        bidder_col, count_col = column_number(columns["bidder_name"]), column_number(columns["bidder_count"])
        bidder_rows = [row for row in range(table["data_start_row"], table["data_end_row"] + 1)
                       if text(sheet.raw(row, bidder_col).value)]
        anchor_rows = [row for row in bidder_rows if text(sheet.raw(row, count_col).value)]
        repeated_counts = [clean(sheet.raw(row, count_col).value) for row in anchor_rows]
        if len(bidder_rows) > 1 and len(anchor_rows) == len(bidder_rows) and any(
                re.fullmatch(r"[2-9]\d*(?:\.0+)?", value) for value in repeated_counts):
            raise MappingRevisionRequired(f"{sheet.name} 的投标数量在每行重复，不能作为组起点", {
                "source": book.path.name, "sheet": sheet.name, "rows": anchor_rows[:12],
                "field": "bidder_count", "values": repeated_counts[:12],
            })
    for row in range(table["data_start_row"], table["data_end_row"] + 1):
        if not any(text(sheet.raw(row, c).value) for c in range(1, sheet.max_col + 1)):
            continue
        values, cells = _values(sheet, row, columns)
        raw_values = _raw_values(sheet, row, columns)
        column_evidence = {}
        for item in audit_specs:
            cell, source_row, source_col = sheet.resolved(row, column_number(item["column"]))
            column_evidence[item["column"]] = {
                "disposition": item["disposition"], "header": item.get("header", ""),
                "value": "" if cell.formula or cell.error else text(cell.value),
                "cell": coordinate(source_row, source_col), "reason": item["reason"],
            }
        if clean(values.get("bidder_name")) in HEADER_NAMES["bidder_name"]:
            raise MappingRevisionRequired(f"{sheet.name} 第 {row} 行疑似重复/新表头，请拆成多个 table 映射", {
                "source": book.path.name, "sheet": sheet.name, "row": row, "row_values": sheet.row_view(row),
            })
        raw_project = text(sheet.raw(row, project_col).value) if project_col else ""
        mode = table["project_mode"]
        if mode == "blocks":
            if raw_project:
                project_anchor = cells[project_role]
            elif current_project:
                project_anchor = current_project["anchor"]
                for role in context_roles:
                    values[role] = current_project["values"].get(role, "")
                    if role in current_project["cells"]:
                        cells[role] = current_project["cells"][role]
            else:
                project_anchor = ""
        elif mode == "repeated":
            project_anchor = [clean(values.get(k, "")) for k in ("project_code", "project_name", "project_year", "project_owner")]
        else:
            project_anchor = cells.get(project_role, "")
        project_name = clean(values.get("project_name"))
        if project_name or clean(values.get("project_code")):
            pid = stable_id("project", book.sha256, sheet.index,
                            table_index if project_scope_index is None else project_scope_index, project_anchor)
            if current_project is None or current_project["id"] != pid:
                current_group = None
                current_group_context = {}
            current_project = projects.setdefault(pid, {
                "id": pid, "source_id": book.source_id, "sheet": sheet.name, "anchor": project_anchor,
                "table_id": table_id, "table_kind": table_kind,
                "values": {k: values.get(k, "") for k in context_roles},
                "cells": {k: cells[k] for k in context_roles if k in cells},
            })
        elif mode != "blocks":
            if current_project is not None:
                current_group = None
                current_group_context = {}
            current_project = None
        is_summary = any(clean(values.get(k)) in table["summary_markers"] for k in ("bidder_name", "agent", "project_name"))
        audit = {"source_id": book.source_id, "sheet": sheet.name, "row": row,
                 "project_id": current_project["id"] if current_project else None,
                 "column_evidence": column_evidence}
        if is_summary:
            row_audit.append({**audit, "type": "summary", "values": values, "raw_values": raw_values, "cells": cells})
            current_group = None
            continue
        unavailable = [role for role, label in columns.items()
                       if sheet.resolved(row, column_number(label))[0].formula or sheet.resolved(row, column_number(label))[0].error]
        bidder_raw = str(raw_values.get("bidder_name") or "") if "bidder_name" not in unavailable else ""
        award_raw = str(raw_values.get("award_name") or "") if "award_name" not in unavailable else ""
        company_role = "bidder_name" if "bidder_name" in columns else "award_name" if "award_name" in columns else None
        source_company = bidder_raw if company_role == "bidder_name" else award_raw
        explicit_context = any(role in columns and text(sheet.raw(row, column_number(columns[role])).value)
                               for role in ("project_name", "project_code", "lot_name", "lot_code"))
        context_record = not company_role or (not bidder_raw and not award_raw and bool(explicit_context))
        substantive = bool(source_company or award_raw or values.get("bidder_count") or unavailable or context_record)
        if not substantive:
            row_audit.append({**audit, "type": "context_only", "values": values, "raw_values": raw_values, "cells": cells})
            continue
        project_id = current_project["id"] if current_project else None
        group_context = {}
        group_context_cells = {}
        for item in context_specs:
            label = item["column"]
            col = column_number(label)
            cell, source_row, source_col = sheet.resolved(row, col)
            if cell.formula or cell.error:
                raise MappingRevisionRequired(f"{sheet.name} 第 {row} 行分组上下文不可用", {
                    "source": book.path.name, "sheet": sheet.name, "row": row,
                    "column": label, "cell": coordinate(source_row, source_col),
                })
            raw = "" if cell.formula or cell.error else clean(cell.value)
            if item.get("mode", "repeated") == "blocks":
                direct = sheet.raw(row, col)
                if text(direct.value) and not direct.formula and not direct.error:
                    current_group_context[label] = clean(direct.value)
                raw = current_group_context.get(label, "")
            if not raw:
                raise MappingRevisionRequired(f"{sheet.name} 第 {row} 行缺少分组上下文", {
                    "source": book.path.name, "sheet": sheet.name, "row": row, "column": label,
                })
            group_context[label] = raw
            group_context_cells[label] = coordinate(source_row, source_col)
        group_mode = table["group_mode"]
        if group_mode == "anchor":
            anchor_field = table["group_start_field"]
            new_group = bool(text(sheet.raw(row, column_number(columns[anchor_field])).value))
            group_anchor = row if new_group else (current_group["anchor_row"] if current_group else None)
            if group_anchor is None:
                group_anchor = row  # 缺少组锚点时按来源行隔离，不虚构项目或阻断企业提取。
        elif group_mode == "row":
            group_anchor = row
        elif group_mode == "lot":
            group_anchor = [clean(values.get("lot_code")), clean(values.get("lot_name"))]
            if not any(group_anchor):
                group_anchor = ["unresolved_row", row]
        elif group_mode == "source":
            group_anchor = "source_roster"
        else:
            group_anchor = "project" if project_id else row
        if group_context:
            group_anchor = [group_anchor, sorted(group_context.items())]
        scope_key = project_id or stable_id("scope", book.sha256, sheet.index, table_index)
        gid = stable_id("group", scope_key, table_index, group_mode, group_anchor)
        current_group = groups.setdefault(gid, {
            "id": gid, "project_id": project_id, "source_id": book.source_id,
            "table_id": table_id, "table_kind": table_kind,
            "sheet": sheet.name, "anchor_row": row, "lot_name": clean(values.get("lot_name")),
            "lot_code": clean(values.get("lot_code")), "records": [], "awards": [], "counts": [],
            "non_tender": False, "procurement_signals": [], "award_completeness_declared": completeness,
            "award_completeness": {}, "source_rows": [],
            "scope_type": "roster" if group_mode == "source" else "context" if context_record else "business",
            "company_role": company_role,
            "award_mode": award_mode, "group_mode": group_mode, "project_mode": mode, "table_index": table_index,
            "group_context": group_context, "group_context_cells": group_context_cells,
            "price_units": price_units,
            "price_unit_assumption": "未标单位的一侧沿用另一侧" if len(price_units) == 2 and bool(price_units.get("bidder_price")) != bool(price_units.get("award_price")) else "",
        })
        group = current_group
        for role in ("lot_name", "lot_code"):
            if clean(values.get(role)) and clean(values[role]) != group[role]:
                raise MappingRevisionRequired(f"{sheet.name} 第 {row} 行组内标段信息不一致，请核对 group_mode", {
                    "source": book.path.name, "sheet": sheet.name, "row": row,
                    "group_start_row": group["anchor_row"], "field": role,
                    "existing": group[role], "current": clean(values[role]),
                })
        group["source_rows"].append(row)
        for role in ("award_status", "rank"):
            if values.get(role) and _vertical_merge(sheet, row, role, columns):
                issues.append(_issue("AWARD_STATUS_SCOPE_AMBIGUOUS" if role == "award_status" else "RANK_SCOPE_AMBIGUOUS",
                                     "跨行合并的中标状态不能分配给多家企业" if role == "award_status" else
                                     "跨行合并的排名不能分配给多家企业", group,
                                     source_row=row, field=role, source_cell=cells[role], raw_value=values[role]))
                values[role] = ""
        for role in unavailable:
            issues.append(_issue("FIELD_UNAVAILABLE", f"{role} 的来源为未计算公式或错误值，该字段留空", group,
                                 source_row=row, field=role, source_cell=cells[role], raw_value=raw_values[role]))
        signal = procurement_signal(values.get("notes", ""), table["non_tender_markers"])
        if signal and signal not in group["procurement_signals"]:
            group["procurement_signals"].append(signal)
        if values.get("bidder_count"):
            count = clean(values["bidder_count"])
            if re.fullmatch(r"\d+(?:\.0+)?", count):
                group["counts"].append(int(float(count)))
            else:
                issues.append(_issue("INVALID_BIDDER_COUNT", "投标数量不是可确认的整数", group, source_row=row, raw_value=count))
        occurrence = {"row": row, "cells": cells, "values": values, "raw_values": raw_values,
                      "column_evidence": column_evidence}
        if award_raw:
            try:
                award_names = split_names(award_raw, "delimited")
            except LedgerError as exc:
                issues.append(_issue("AWARD_LIST_AMBIGUOUS", str(exc), group, source_row=row, raw_award=award_raw))
                award_names = []
            for raw in award_names:
                name, changes = normalize_name(raw, {})
                if not any(name_key(a["name"]) == name_key(name) and a["cell"] == cells.get("award_name") for a in group["awards"]):
                    group["awards"].append({"raw": raw, "name": name, "cell": cells.get("award_name"),
                                            "changes": changes, "legal_person": values.get("award_legal_person", ""),
                                            "price": values.get("award_price", ""),
                                            "row": sheet.resolved(row, column_number(columns["award_name"]))[1],
                                            "merged": _vertical_merge(sheet, row, "award_name", columns),
                                            "single_name": len(award_names) == 1})
        try:
            tokens = split_names(source_company, table["bidder_separator"])
        except LedgerError as exc:
            issues.append(_issue("BIDDER_LIST_AMBIGUOUS", str(exc), group, source_row=row,
                                 standalone=True, occurrence=occurrence))
            tokens = []
        if context_record or (unavailable and not tokens):
            tokens = [""]
        for index, raw in enumerate(tokens, 1):
            name, changes = normalize_name(raw, {})
            if name and not any(char.isalpha() for char in name):
                issues.append(_issue("INVALID_COMPANY_NAME", "企业字段只有数字或符号，不能形成企业名称", group,
                                     source_row=row, standalone=True, occurrence=occurrence))
                continue
            group["records"].append({
                "id": stable_id("record", book.sha256, sheet.index, table_index, row, index),
                "name": name, "project_id": project_id, "group_id": gid,
                "occurrences": [{**occurrence, "fragment_index": index, "raw_company": raw, "company_role": company_role,
                                 "single_bidder_row": company_role == "bidder_name" and len(tokens) == 1 and not _vertical_merge(sheet, row, "bidder_name", columns)}],
                "changes": changes, "award_status": "", "rank": values.get("rank", "") if len(tokens) == 1 else "",
                "explicit_award_status": values.get("award_status", "") if len(tokens) == 1 else "",
                "context_only": context_record,
            })
        if len(tokens) > 1 and values.get("rank"):
            issues.append(_issue("RANK_SCOPE_AMBIGUOUS", "同一名单包含多家企业，单个排名无法分配；排名留空", group, source_row=row))
        row_audit.append({**audit, "group_id": gid, "type": "business", "bidder_mentions": len(tokens)})
    return list(projects.values()), list(groups.values()), issues, row_audit


AWARD_STATUS_VALUES = {"是": "是", "否": "否", "中标": "是", "未中标": "否",
                       "true": "是", "false": "否", "1": "是", "0": "否"}


def _canonical_award_status(value: object) -> str | None:
    raw = clean(value)
    return AWARD_STATUS_VALUES.get(raw.casefold())


def finish_group(group: dict[str, Any], issues: list[dict[str, Any]], resolutions: dict[str, dict]) -> None:
    signals = set(group["procurement_signals"])
    uncertain = "uncertain" in signals or {"not_non_tender", "non_tender"}.issubset(signals)
    group["procurement_status"] = "uncertain" if uncertain else "non_tender" if "non_tender" in signals else (
        group["scope_type"] if group["scope_type"] in {"roster", "context"} else "bidding")
    group["non_tender"] = group["procurement_status"] == "non_tender"
    # 采购方式是原文上下文，不覆盖明确中标事实，不独立制造数据复核。
    deduplicated = {}
    for record in group["records"]:
        key = name_key(record["name"]) or record["id"]
        previous = deduplicated.get(key)
        if previous:
            previous["occurrences"].extend(record["occurrences"])
            previous["changes"].extend(record["changes"])
        else:
            deduplicated[key] = record
    group["records"] = list(deduplicated.values())
    for record in group["records"]:
        conflict_fields = []
        for field in ("bidder_price", "bidder_legal_person", "rank"):
            values = {clean(o["values"].get(field)) for o in record["occurrences"] if clean(o["values"].get(field))}
            if len(values) > 1:
                conflict_fields.append(field)
            if field == "rank":
                record["rank"] = next(iter(values)) if len(values) == 1 else ""
        raw_statuses = [clean(o["values"].get("award_status")) for o in record["occurrences"]
                        if clean(o["values"].get("award_status"))]
        canonical_statuses = {_canonical_award_status(value) for value in raw_statuses}
        recognized_statuses = canonical_statuses - {None}
        if len(recognized_statuses) > 1:
            conflict_fields.append("award_status")
            record["explicit_award_conflict"] = True
            record["explicit_award_status"] = ""
        elif len(recognized_statuses) == 1 and all(value is not None for value in canonical_statuses):
            record["explicit_award_status"] = next(iter(recognized_statuses))
        elif raw_statuses:
            record["explicit_award_status"] = raw_statuses[0] if len(set(raw_statuses)) == 1 else ""
            if len(set(raw_statuses)) > 1:
                conflict_fields.append("award_status")
                record["explicit_award_conflict"] = True
        else:
            record["explicit_award_status"] = ""
        if conflict_fields:
            issues.append(_issue("DUPLICATE_CONFLICT",
                                 "同组重复记录存在不一致字段；一致非空值已合并，冲突字段留空复核", group,
                                 record_name=record["name"], conflicting_fields=sorted(set(conflict_fields))))
    counts = set(group["counts"])
    actual_count = sum(bool(record["name"]) and not record["context_only"] for record in group["records"])
    if len(counts) > 1 or (counts and counts != {actual_count}):
        issues.append(_issue("BIDDER_COUNT_MISMATCH", "声明投标数量与整理后的企业数量不一致", group,
                             declared_counts=sorted(counts), actual_count=actual_count))
    matched, unresolved = match_group(group, resolutions)
    for problem in unresolved:
        details = dict(problem)
        issues.append(_issue(details.pop("code"), details.pop("message"), group, **details))
    if not group["records"] and not any(i.get("standalone") and i["group_id"] == group["id"] for i in issues):
        issues.append(_issue("BIDDER_MISSING", "存在结果信息，但缺少企业名单", group, standalone=True))
    incomplete = bool(unresolved) or any(
        i["group_id"] == group["id"] and (
            i["code"] == "AWARD_LIST_AMBIGUOUS" or
            (i["code"] == "FIELD_UNAVAILABLE" and i.get("field") == "award_name")
        ) for i in issues
    )
    declared_completeness = group["award_completeness_declared"]
    completeness_reasons = []
    if declared_completeness["status"] != "complete":
        completeness_reasons.append("区域未声明完整最终中标结果")
    if not group["awards"]:
        completeness_reasons.append("本组没有中标结果")
    if incomplete:
        completeness_reasons.append("本组存在未解决的中标读取或对应问题")
    group["award_completeness"] = {
        **declared_completeness,
        "verified": not completeness_reasons,
        "verification_reasons": completeness_reasons,
    }
    for record in group["records"]:
        declared = clean(record.get("explicit_award_status"))
        explicit = _canonical_award_status(declared)
        if declared and explicit is None and declared not in {"未知", "待定", "/", "-"}:
            issues.append(_issue("AWARD_STATUS_UNRECOGNIZED", "原表中标状态无法对应是/否，留空并保留原值", group,
                                 raw_value=declared, source_row=record["occurrences"][0]["row"]))
        if record["name"]:
            if record["id"] in matched:
                record["award_status"] = "是"
            elif matched and group["award_completeness"]["verified"]:
                record["award_status"] = "否"
        if explicit is not None:
            if record["award_status"] and record["award_status"] != explicit:
                record["award_status"] = ""
                issues.append(_issue("AWARD_STATUS_CONFLICT", "原表是否中标与中标名单冲突，留空供复核", group))
            else:
                record["award_status"] = explicit
        if record.get("explicit_award_conflict"):
            record["award_status"] = ""


def _issue_identity(issue: dict[str, Any]) -> str:
    if issue.get("field") or issue.get("source_cell"):
        scope = ("field", issue.get("field"), issue.get("source_cell") or issue.get("source_row"))
    elif issue.get("award_cell") or issue.get("original_award"):
        scope = ("award", issue.get("award_cell"), issue.get("original_award"))
    elif issue.get("record_name"):
        scope = ("record", issue.get("record_name"), tuple(issue.get("conflicting_fields", [])))
    elif issue.get("source_id") and not issue.get("group_id"):
        scope = ("source", issue.get("source_id"), issue.get("sheet"), issue.get("start_row"), issue.get("end_row"),
                 issue.get("source_project_id"), issue.get("target_project_id"))
    else:
        scope = ("group",)
    return stable_id("issue", issue.get("group_id"), issue["code"], scope)


def _record_confidence(record: dict[str, Any], group: dict[str, Any], related: list[dict[str, Any]]) -> dict[str, Any]:
    company_unavailable = any(issue["code"] == "FIELD_UNAVAILABLE" and issue.get("field") == group["company_role"]
                              for issue in related)
    extraction = 0.25 if company_unavailable else 1.0 if record["name"] else 0.70
    grouping = {"row": 1.0, "source": 1.0, "lot": 0.95, "anchor": 0.92,
                "project": 0.88}.get(group["group_mode"], 0.80)
    if group["group_context"]:
        grouping = max(grouping, 0.97)
    if group["project_mode"] == "blocks":
        grouping = min(grouping, 0.90)
    selected = next((match for match in group["award_matches"]
                     if match["selected_record_id"] == record["id"]), None)
    if selected:
        award = {"exact_name": 1.0, "user_selection": 1.0}.get(selected["basis"], 0.75)
        award_basis = selected["basis"]
    elif record["award_status"] == "否":
        basis_type = group["award_completeness"].get("basis_type")
        award = 1.0 if basis_type == "explicit" else 0.90
        award_basis = f"complete_{basis_type}"
    elif record.get("explicit_award_status") and record["award_status"]:
        award, award_basis = 1.0, "explicit_status"
    elif group["awards"]:
        award, award_basis = 0.60, "unresolved_or_incomplete"
    else:
        award, award_basis = 1.0, "no_award_claim"
    score = min(extraction, grouping, award)
    if related:
        score = min(score, 0.60)
    return {"score": round(score, 2), "meaning": "rule_reliability_not_probability",
            "components": {"extraction": extraction, "grouping": grouping, "award_mapping": award},
            "basis": {"group_mode": group["group_mode"], "project_mode": group["project_mode"],
                      "award": award_basis, "issue_codes": sorted({issue["code"] for issue in related})}}


def build_outputs(books: list[Workbook], plan: dict[str, Any], alias_payload: dict[str, Any] | None = None,
                  generated_at: str | None = None,
                  source_failures: list[dict[str, Any]] | None = None,
                  award_resolutions: dict[str, dict] | None = None,
                  relation_resolutions: dict[str, dict] | None = None) -> tuple[list[dict], list[dict], dict]:
    validate_plan(plan, books)
    alias_payload = alias_payload or {"schema_version": 1, "aliases": []}
    if aliases_from_json(alias_payload):
        raise LedgerError("不接受全局公司名映射；输出保留投标单位名称")
    aliases = {}
    timestamp = generated_at or datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")
    try:
        if datetime.fromisoformat(timestamp).utcoffset() is None:
            raise ValueError()
    except ValueError as exc:
        raise LedgerError("生成时间必须为带时区的 ISO 8601") from exc
    source_failures = source_failures or []
    award_resolutions = award_resolutions or {}
    relation_resolutions = relation_resolutions or {}
    projects, groups, issues, row_audit, skipped = [], [], [], [], []
    for failure in source_failures:
        issues.append(_issue("SOURCE_UNREADABLE", failure["error"], None, standalone=True,
                             source_id=failure["id"], source_file=failure["file_name"]))
    for book, source in zip(books, plan["sources"]):
        for sheet, spec in zip(book.sheets, source["sheets"]):
            if spec["action"] == "skip":
                skipped.append({"source_id": book.source_id, "sheet": sheet.name, "reason": spec["reason"]})
                if sheet.row_numbers:
                    issues.append(_issue("SOURCE_REGION_SKIPPED", "非空工作表未解析，已转入复核", None,
                                         standalone=True, source_id=book.source_id, source_file=book.path.name,
                                         sheet=sheet.name, start_row=min(sheet.row_numbers), end_row=max(sheet.row_numbers),
                                         reason=spec["reason"]))
                continue
            if spec["action"] == "review":
                issues.append(_issue("SOURCE_REGION_UNRESOLVED", "工作表结构未能可靠映射，已转入复核", None,
                                     standalone=True, source_id=book.source_id, source_file=book.path.name,
                                     sheet=sheet.name, start_row=min(sheet.row_numbers, default=1),
                                     end_row=max(sheet.row_numbers, default=1), reason=spec["reason"]))
                continue
            header_scopes = {}
            for table_index, table in enumerate(spec["tables"]):
                scope_index = header_scopes.setdefault(tuple(sorted(table["header_rows"])), table_index)
                p, g, i, audit = collect_table(book, sheet, table, table_index, aliases, scope_index)
                projects.extend(p); groups.extend(g); issues.extend(i); row_audit.extend(audit)
            for region in spec.get("review_regions", []):
                issues.append(_issue("SOURCE_REGION_UNRESOLVED", "局部结构未能可靠映射，已转入复核", None,
                                     standalone=True, source_id=book.source_id, source_file=book.path.name,
                                     sheet=sheet.name, start_row=region["start"], end_row=region["end"],
                                     reason=region["reason"]))
    projects = list({p["id"]: p for p in projects}.values())
    projects, groups, issues, relationships = apply_relationships(
        projects, groups, issues, plan, relation_resolutions)
    for group in groups:
        finish_group(group, issues, award_resolutions)
    project_by_id = {p["id"]: p for p in projects}
    book_by_id = {b.source_id: b for b in books}
    # 分组问题对应整个组，避免只标候选而把其他企业误写成已确认未中标。
    group_by_id = {group["id"]: group for group in groups}
    for issue in issues:
        group = group_by_id.get(issue["group_id"])
        issue["record_ids"] = [r["id"] for r in group["records"]] if group else []
        issue["id"] = _issue_identity(issue)
    # 字段问题按单元格保留；同一记录或组的同类问题按明确身份归并。
    issues = list({i["id"]: i for i in issues}.values())
    final, review, audit_records = [], [], []
    for group in groups:
        project = project_by_id.get(group["project_id"], {"id": None, "values": {}, "cells": {}})
        source = book_by_id[group["source_id"]]
        related = [i for i in issues if i["group_id"] == group["id"]]
        for record in group["records"]:
            sequence = len(final) + 1
            occurrences = record["occurrences"]
            evidence = [f"Sheet={group['sheet']}"]
            locations = [o["cells"][o["company_role"]] + f"#{o['fragment_index']}" if o.get("company_role") else f"行{o['row']}"
                         for o in occurrences]
            evidence.append("企业来源=" + ",".join(locations) if record["name"] else "来源=" + ",".join(locations))
            if project["cells"]:
                evidence.append("项目来源=" + str(project["cells"].get("project_name") or project["cells"].get("project_code") or ""))
            if group["scope_type"] == "roster":
                evidence.append("纯企业名单，原表未提供项目或标段归属")
            else:
                evidence.append(f"组起始行={group['anchor_row']}")
            if group["group_context"]:
                evidence.append("分组上下文=" + ",".join(
                    f"{label}:{value}@{group['group_context_cells'][label]}"
                    for label, value in group["group_context"].items()))
            if group["awards"]:
                evidence.append("原中标字段=" + "；".join(f"{a['cell']}:{a['raw']}" for a in group["awards"]))
            if record["changes"]:
                evidence.append("清洗=" + "；".join(f"{c['before']}→{c['after']}({c['rule']})" for c in record["changes"]))
            for matched in group["award_matches"]:
                if matched["selected_record_id"] == record["id"]:
                    evidence.append(f"中标对应={matched['original_award']}→{record['name']}({matched['basis']})")
            if group["non_tender"]:
                evidence.append("原表注明未招投标")
            confidence = _record_confidence(record, group, related)
            confidence_text = f"{confidence['score']:.2f}".rstrip("0").rstrip(".")
            row = dict.fromkeys(FIELDS, "")
            row.update({"序号": str(sequence), "项目名称": clean(project["values"].get("project_name")),
                        "项目编号": clean(project["values"].get("project_code")), "标段名称": group["lot_name"],
                        "标段编号": group["lot_code"], "公司名称": record["name"], "中标与否": record["award_status"],
                        "投标排名": record["rank"], "文件类别": "excel_ledger", "依据文件路径": source.path.name,
                        "提取方式": "台账整理", "证据文本": "；".join(evidence), "置信度": confidence_text,
                        "复核状态": "待复核" if related else "通过", "解析结果生成日期时间": timestamp})
            final.append(row)
            review_sequence = None
            if related:
                review_sequence = len(review) + 1
                reason = "；".join(dict.fromkeys(i["message"] for i in related))
                review.append({**row, "序号": str(review_sequence), "证据文本": reason + "\n" + row["证据文本"]})
                for issue in related:
                    issue["final_sequences"].append(sequence)
                    issue["review_sequences"].append(review_sequence)
            audit_records.append({"id": record["id"], "final_sequence": sequence, "review_sequence": review_sequence,
                                  "project_id": project["id"], "group_id": group["id"], "source_id": source.source_id,
                                  "sheet": group["sheet"], "company_name": record["name"],
                                  "participation_type": "context" if record["context_only"] else
                                      (("award_company" if group["company_role"] == "award_name" else "bidder")
                                       if group["procurement_status"] == "bidding" else group["procurement_status"]),
                                  "occurrences": occurrences, "corrections": record["changes"],
                                  "confidence": confidence,
                                  "issue_ids": [i["id"] for i in related]})
    source_files = {book.source_id: book.path.name for book in books}
    source_files.update({failure["id"]: failure["file_name"] for failure in source_failures})
    for issue in [item for item in issues if item.get("standalone")]:
        group = group_by_id.get(issue["group_id"])
        project = project_by_id.get(group["project_id"], {"values": {}}) if group else {"values": {}}
        source_id = group["source_id"] if group else issue.get("source_id")
        location = []
        if issue.get("sheet") or group:
            location.append("Sheet=" + str(issue.get("sheet") or group["sheet"]))
        if issue.get("start_row"):
            location.append(f"区域={issue['start_row']}:{issue.get('end_row', issue['start_row'])}")
        elif issue.get("source_row") or group:
            location.append("来源行=" + str(issue.get("source_row") or group["anchor_row"]))
        if issue.get("reason"):
            location.append("原因=" + issue["reason"])
        row = dict.fromkeys(FIELDS, "")
        row.update({"序号": str(len(review) + 1), "项目名称": clean(project["values"].get("project_name")),
                    "标段名称": group["lot_name"] if group else "", "文件类别": "excel_ledger",
                    "依据文件路径": source_files.get(source_id, issue.get("source_file", "")),
                    "提取方式": "台账整理", "复核状态": "待复核", "置信度": "0",
                    "证据文本": issue["message"] + ("\n" + "；".join(location) if location else ""),
                    "解析结果生成日期时间": timestamp})
        review.append(row)
        issue["review_sequences"].append(len(review))
    if not final and not review:
        raise LedgerError("没有识别到可整理的数据或复核项；未发布空的成功产物")
    unique = {}
    for row in final:
        if name_key(row["公司名称"]):
            unique.setdefault(name_key(row["公司名称"]), row["公司名称"])
    ledger = {
        "schema_version": 1, "parser_version": VERSION, "generated_at": timestamp,
        "sources": ([{"id": b.source_id, "file_name": b.path.name, "sha256": b.sha256, "size": b.size,
                      "format": b.format, "sheets": [s.name for s in b.sheets], "status": "parsed"} for b in books] +
                    [dict(failure) for failure in source_failures]),
        "mapping": plan, "matching_policy": dict(POLICY),
        "resolutions": [dict({"review_task_id": task_id}, **resolution)
                        for task_id, resolution in sorted(award_resolutions.items())],
        "summary": {
            "project_count": len(projects), "explicit_lot_count": sum(bool(g["lot_name"] or g["lot_code"]) for g in groups),
            "group_count": len(groups), "bidding_group_count": sum(g["procurement_status"] == "bidding" for g in groups),
            "non_tender_group_count": sum(g["non_tender"] for g in groups), "record_count": len(final),
            "uncertain_group_count": sum(g["procurement_status"] == "uncertain" for g in groups),
            "roster_group_count": sum(g["scope_type"] == "roster" for g in groups),
            "context_group_count": sum(g["scope_type"] == "context" for g in groups),
            "pending_record_count": sum(r["复核状态"] == "待复核" for r in final),
            "review_record_count": len(review), "issue_count": len(issues), "unique_company_count": len(unique),
            "duplicate_mentions_removed": sum(len(r["occurrences"]) - 1 for r in audit_records),
            "corrected_record_count": sum(bool(r["corrections"]) for r in audit_records),
        },
        "projects": projects, "groups": [{k: v for k, v in g.items() if k != "records"} for g in groups],
        "records": audit_records, "issues": issues, "row_audit": row_audit, "skipped_sheets": skipped,
        "unique_companies": list(unique.values()),
        "notes": ["置信度是由提取、分组和中标对应规则计算的可靠程度，不是正确概率；逐条组成和依据保存在 records.confidence。",
                  "非精确中标名称只生成推荐候选，必须由用户选择或保持不确定；法人和金额仅作来源审计。",
                  "存在投标列时公司名称仅来自投标列；中标名和对应依据保留在 groups.award_matches，不回写或纠正投标全称。",
                  "unique_companies 按 Gitee 客户端的 NFKC、去空白、casefold 规则从 final 非空公司名称去重；包括待复核名称，可供后续采集编排读取。",
                  "业务字段缺失仅留空，不伪造项目/标段归属，也不因缺少中标字段要求用户补充。",
                  "XLS 读取缓存值，xlrd 不提供可靠公式文本标记；本工具不计算公式。",
                  "内部项目/组 ID 仅对同一文件内容和同一映射稳定；不是官方编号。"],
    }
    if plan.get("relationships"):
        ledger["relationships"] = relationships
        ledger["relationship_resolutions"] = [dict({"review_task_id": task_id}, **resolution)
                                               for task_id, resolution in sorted(relation_resolutions.items())]
        ledger["summary"].update({
            "relationship_count": len(relationships),
            "unresolved_relationship_count": sum(item["status"] != "matched" for item in relationships),
        })
    return final, review, ledger
