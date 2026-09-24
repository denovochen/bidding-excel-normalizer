"""从来源单元格建立投标块证据，不使用 plan 的分组结果。 @author denovochen"""
from __future__ import annotations

import re
from typing import Any

from .contract import MappingRevisionRequired, clean, column_number, coordinate, name_key, stable_id, text


def project_text(value: object) -> bool:
    """项目样文本仅是冲突信号；采用它作为项目来源还需要语义确认。"""
    return bool(re.search(r"项目|工程|\bproject\b", clean(value), re.IGNORECASE))


def simple_lot(value: object) -> bool:
    value = clean(value).replace(" ", "")
    return value in {"", "/", "-", "—", "主标段", "施工标段", "本标段", "全部标段"} or bool(re.fullmatch(
        r"(?:施工|监理|设计)?第?[一二三四五六七八九十百零\dA-Za-z]+(?:标段|标包|包件|合同段|标|包)|(?:LOT|lot)[-_]?\d+", value))


def source_layout(book, sheet, table: dict[str, Any], strict: bool = True) -> dict[str, Any]:
    from .workbook import _row_header_matches

    columns = table["columns"]
    detected = {}
    for row in table["header_rows"]:
        detected.update(_row_header_matches(sheet, row))
    # 结构信号不因调用者把已识别列降成 evidence 而消失。
    signals = {**detected, **columns}
    if strict and detected.get("bidder_serial") and columns.get("project_serial") == detected["bidder_serial"]:
        raise MappingRevisionRequired("投标企业序号不能映射为 project_serial", {
            "source": book.path.name, "sheet": sheet.name, "code": "BIDDER_SERIAL_ROLE_CONFLICT",
            "column": detected["bidder_serial"],
        })
    bidder = columns.get("bidder_name")
    if not bidder:
        return {"blocks": [], "row_blocks": {}, "protected": False}
    for role in ("project_name", "project_code"):
        label = detected.get(role)
        if strict and label and label not in {columns.get("project_name"), columns.get("project_code")}:
            raise MappingRevisionRequired("具有项目身份的来源列不能被降为辅助证据", {
                "source": book.path.name, "sheet": sheet.name, "code": "SOURCE_PROJECT_ROLE_REQUIRED", "column": label,
            })
    start, end = table["data_start_row"], table["data_end_row"]
    summaries = set(table["summary_markers"])
    rows = [row for row in sheet.row_numbers if start <= row <= end]
    bidder_rows = [row for row in rows if text(sheet.resolved(row, column_number(bidder))[0].value)
                   and clean(sheet.resolved(row, column_number(bidder))[0].value) not in summaries]
    project_role = next((role for role in ("project_name", "project_code") if role in columns), None)
    context_roles = [role for role in ("lot_name", "lot_code", "bidder_count", "bidder_serial") if role in signals]
    sparse = {}
    for role in ([project_role] if project_role else []) + context_roles:
        label = signals[role]
        sparse[role] = any(not text(sheet.raw(row, column_number(label)).value) for row in bidder_rows)
    list_rows = table["bidder_separator"] != "single"
    blocks, row_blocks = [], {}
    last_project, last_lots, last_serial = "", {}, None
    last_context = {}
    pending_evidence = []
    pending_rows = []
    current = None
    for row in rows:
        evidence = []
        for item in table.get("column_dispositions", []):
            if item["disposition"] != "group_context":
                continue
            col = column_number(item["column"])
            value = clean(sheet.raw(row, col).value)
            if value and (item.get("mode") == "blocks" or value != last_context.get(col)):
                evidence.append({"role": "group_context", "cell": coordinate(row, col), "value": value})
                last_context[col] = value
        if project_role:
            col = column_number(columns[project_role])
            direct = sheet.raw(row, col)
            value = "" if direct.formula or direct.error else clean(direct.value)
            if value in summaries:
                continue
            if value:
                if sparse[project_role] or name_key(value) != last_project:
                    evidence.append({"role": project_role, "cell": coordinate(row, col), "value": value})
                last_project = name_key(value)
        for role in ("lot_name", "lot_code"):
            if role not in signals:
                continue
            col = column_number(signals[role])
            direct = sheet.raw(row, col)
            value = "" if direct.formula or direct.error else clean(direct.value)
            if value:
                if sparse[role] or value != last_lots.get(role):
                    evidence.append({"role": role, "cell": coordinate(row, col), "value": value})
                last_lots[role] = value
                if strict and signals[role] not in {columns.get("lot_name"), columns.get("lot_code"), columns.get("project_name")}:
                    raise MappingRevisionRequired("具有标段语义的来源列不能被降为辅助证据", {
                        "source": book.path.name, "sheet": sheet.name, "column": signals[role],
                        "row": row, "code": "SOURCE_LOT_ROLE_REQUIRED", "value": value,
                    })
                if project_role and not text(sheet.resolved(row, column_number(columns[project_role]))[0].value):
                    if (strict and project_text(value) and name_key(value) != last_project and
                            role not in table.get("project_context_fields", [])):
                        raise MappingRevisionRequired("新块含项目文本，不能静默继承上一个项目；需确认项目上下文来源", {
                            "source": book.path.name, "sheet": sheet.name, "row": row,
                            "code": "PROJECT_CONTEXT_CONFLICT", "field": role, "cell": coordinate(row, col),
                            "value": value,
                        })
                    if (strict and not simple_lot(value) and name_key(value) != last_project and
                            role not in table.get("project_context_fields", []) and
                            role not in table.get("lot_context_fields", [])):
                        raise MappingRevisionRequired("项目字段缺失，新块上下文含义不明确；需判断它是项目还是标段", {
                            "source": book.path.name, "sheet": sheet.name, "row": row,
                            "code": "BLOCK_CONTEXT_REVIEW_REQUIRED", "field": role,
                            "cell": coordinate(row, col), "value": value,
                        })
        company = sheet.resolved(row, column_number(bidder))[0]
        if clean(company.value) in summaries:
            pending_evidence = []
            pending_rows = []
            continue
        if "bidder_count" in signals and sparse["bidder_count"]:
            col = column_number(signals["bidder_count"])
            direct = sheet.raw(row, col)
            if not direct.formula and not direct.error and re.fullmatch(r"\d+(?:\.0+)?", clean(direct.value)):
                evidence.append({"role": "bidder_count", "cell": coordinate(row, col), "value": clean(direct.value)})
        if "bidder_serial" in signals:
            col = column_number(signals["bidder_serial"])
            raw_serial = clean(sheet.raw(row, col).value)
            if re.fullmatch(r"\d+(?:\.0+)?", raw_serial):
                serial = int(float(raw_serial))
                if last_serial is not None and serial < last_serial:
                    evidence.append({"role": "bidder_serial", "cell": coordinate(row, col), "value": raw_serial})
                last_serial = serial
        if not (text(company.value) or company.formula or company.error):
            pending_evidence.extend(evidence)
            if evidence:
                pending_rows.append(row)
            continue
        if list_rows:
            evidence.append({"role": "bidder_list", "cell": coordinate(row, column_number(bidder)), "value": ""})
        evidence = pending_evidence + evidence
        pending_evidence = []
        if current is None or evidence:
            anchor = min([row, *pending_rows])
            current = {
                "id": stable_id("source_block", book.sha256, sheet.index, anchor, bidder),
                "source_id": book.source_id, "sheet": sheet.name,
                "anchor_row": anchor, "bidder_column": bidder, "rows": [], "evidence": evidence,
                "context_rows": list(pending_rows),
            }
            blocks.append(current)
        current["rows"].append(row)
        row_blocks[row] = current["id"]
        for context_row in pending_rows:
            row_blocks[context_row] = current["id"]
        pending_rows = []
    protected = bool(list_rows or any(any(e["role"] in {"bidder_count", "lot_name", "lot_code", "bidder_serial"}
                                         for e in block["evidence"]) for block in blocks))
    return {"blocks": blocks, "row_blocks": row_blocks, "protected": protected}


def verify_group_boundaries(book, sheet, groups: list[dict], layout: dict) -> None:
    block_groups: dict[str, set[str]] = {}
    for group in groups:
        blocks = {layout["row_blocks"][row] for row in group["source_rows"] if row in layout["row_blocks"]}
        if len(blocks) > 1:
            raise MappingRevisionRequired("分组吞并了多个来源投标块，已在去重和推荐前阻断", {
                "source": book.path.name, "sheet": sheet.name, "code": "SOURCE_BLOCKS_COLLAPSED",
                "group_start_row": group["anchor_row"],
                "blocks": [_block_summary(block) for block in layout["blocks"] if block["id"] in blocks][:5],
            })
        for block in blocks:
            block_groups.setdefault(block, set()).add(group["id"])
    split = {block for block, ids in block_groups.items() if len(ids) > 1}
    if layout["protected"] and split:
        raise MappingRevisionRequired("来源投标块被拆成不完整候选组，已在关系解析和推荐前阻断", {
            "source": book.path.name, "sheet": sheet.name, "code": "SOURCE_BLOCK_TRUNCATED",
            "blocks": [_block_summary(block) for block in layout["blocks"] if block["id"] in split][:5],
        })


def _block_summary(block: dict) -> dict:
    return {"id": block["id"], "rows": [min(block["rows"]), max(block["rows"])],
            "row_count": len(block["rows"]), "evidence": [
                {**item, "value": item["value"][:160]} for item in block["evidence"][:6]]}
