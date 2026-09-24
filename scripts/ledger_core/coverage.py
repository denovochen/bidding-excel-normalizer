"""以来源行清单核验提取、隔离和去重的守恒关系。 @author denovochen"""
from __future__ import annotations

from .contract import LedgerError


def validate_source_coverage(ledger: dict) -> dict:
    from .relations import table_identity
    basis = {}
    for source in ledger["mapping"]["sources"]:
        for sheet_index, sheet in enumerate(source["sheets"]):
            for table_index, table in enumerate(sheet.get("tables", [])):
                table_id = table.get("table_id") or table_identity(source["sha256"], sheet_index, table_index)
                basis[table_id] = {item["column"] for item in table.get("column_dispositions", [])}
    groups = {group["id"]: group for group in ledger["groups"]}
    blocks = {block["id"]: block for block in ledger["source_blocks"]}
    if len(blocks) != len(ledger["source_blocks"]):
        raise LedgerError("来源投标块 ID 重复")
    expected = {}
    for row in ledger["row_audit"]:
        if row.get("company_role") == "bidder_name" and row.get("company_present"):
            key = (row["source_id"], row["sheet"], row["row"], row["company_cell"])
            if key in expected:
                raise LedgerError("投标来源行重复覆盖")
            expected[key] = row
    actual = {}
    for record in ledger["records"]:
        origin_blocks = set()
        for occurrence in record["occurrences"]:
            for column, item in occurrence.get("column_evidence", {}).items():
                if "reason_ref" in item and column not in basis.get(item["reason_ref"], set()):
                    raise LedgerError("辅助列证据引用了不存在的映射依据")
            if occurrence.get("company_role") != "bidder_name":
                continue
            key = (record["source_id"], record["sheet"], occurrence["row"], occurrence["cells"]["bidder_name"])
            # 无企业的项目上下文记录不属于企业提取覆盖。
            if not occurrence.get("raw_company") and key not in expected:
                continue
            if key not in expected:
                raise LedgerError("投标记录没有来源行清单依据")
            block_id = occurrence.get("source_block_id")
            block = blocks.get(block_id)
            if (not block or block["source_id"] != record["source_id"] or
                    block["sheet"] != record["sheet"] or occurrence["row"] not in block["rows"] or
                    expected[key]["source_block_id"] != block_id):
                raise LedgerError("投标记录与独立来源块边界不一致")
            origin_blocks.add(block_id)
            fragments = actual.setdefault(key, set())
            if occurrence["fragment_index"] in fragments:
                raise LedgerError("同一来源企业片段被重复输出")
            fragments.add(occurrence["fragment_index"])
        if len(origin_blocks) > 1:
            raise LedgerError("跨来源投标块参与被错误去重")
    quarantined = set()
    for issue in ledger["issues"]:
        occurrence = issue.get("occurrence")
        group = groups.get(issue.get("group_id"))
        if issue.get("standalone") and occurrence and group:
            quarantined.add((group["source_id"], group["sheet"], occurrence["row"],
                             occurrence.get("cells", {}).get("bidder_name")))
    for key, row in expected.items():
        fragments = actual.get(key, set())
        if fragments != set(range(1, row["company_fragments"] + 1)) or not fragments:
            if key not in quarantined:
                raise LedgerError(f"投标来源未完整输出且未进入独立复核: {key[1]}!{key[3]}")
    for group in groups.values():
        if not group.get("candidate_coverage", {}).get("complete", True):
            if group.get("award_completeness", {}).get("verified"):
                raise LedgerError("不完整投标组不得声称中标结果覆盖已通过")
            if any(match.get("candidates") or match.get("recommended_record_id") or match.get("selected_record_id")
                   for match in group["award_matches"]):
                raise LedgerError("不完整投标组不得输出中标候选或推荐")
    return {"verified": True, "bidder_source_row_count": len(expected),
            "extracted_bidder_row_count": len(actual),
            "quarantined_bidder_row_count": len(set(expected) & quarantined),
            "unaccounted_bidder_row_count": 0, "source_block_count": len(blocks)}
