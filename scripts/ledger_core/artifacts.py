"""原子发布三个产物，验证 CSV 与审计快照的一致性。 @author denovochen"""
from __future__ import annotations

import csv
import hashlib
import json
import math
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from .contract import FIELDS, OUTPUTS, VERSION, LedgerError, name_key


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            h.update(block)
    return h.hexdigest()


def _write_csv(path: Path, rows: list[dict]) -> list[dict[str, Any]]:
    escaped = []
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\r\n", extrasaction="raise")
        writer.writeheader()
        for row in rows:
            exported = dict(row)
            for key, value in row.items():
                # CSV 在 Excel 中打开时不可把原表的文本当作公式执行。
                if key not in {"序号", "置信度"} and value.lstrip().startswith(("=", "+", "-", "@")):
                    exported[key] = "'" + value
                    escaped.append({"sequence": int(row["序号"]), "field": key, "original": value})
            writer.writerow(exported)
    return escaped


def publish(output: Path, final: list[dict], review: list[dict], ledger: dict) -> dict[str, Any]:
    _validate_name_decisions(ledger)
    _validate_award_matches(ledger)
    output = output.expanduser().absolute()
    if output.is_symlink():
        raise LedgerError("输出目录不能是符号链接")
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise LedgerError("输出目录必须不存在或为空；不会覆盖已有产物")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=".excel-ledger-", dir=output.parent))
    try:
        ledger["csv_text_escapes"] = {
            "final.csv": _write_csv(temporary / "final.csv", final),
            "review_queue.csv": _write_csv(temporary / "review_queue.csv", review),
        }
        ledger["artifacts"] = {name: {"sha256": digest(temporary / name), "size": (temporary / name).stat().st_size}
                               for name in ("final.csv", "review_queue.csv")}
        (temporary / "ledger.json").write_text(json.dumps(ledger, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
        result = validate_outputs(temporary)
        if output.exists():
            output.rmdir()  # 仅移除仍为空的目标目录；若被并发写入则失败。
        temporary.rename(output)
        return {"kind": "result", "output": str(output), "files": sorted(OUTPUTS), **result}
    finally:
        if temporary.exists():
            shutil.rmtree(temporary)


def _read_csv(path: Path, escape_entries: list[dict]) -> list[dict[str, str]]:
    with path.open("rb") as stream:
        if stream.read(3) != b"\xef\xbb\xbf":
            raise LedgerError(f"{path.name} 缺少 UTF-8 BOM")
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, strict=True)
        if reader.fieldnames != FIELDS:
            raise LedgerError(f"{path.name} 列名或顺序不符合 16 列契约")
        rows = list(reader)
    for index, row in enumerate(rows, 1):
        if set(row) != set(FIELDS) or any(v is None for v in row.values()) or row["序号"] != str(index):
            raise LedgerError(f"{path.name} 行宽或序号错误: {index}")
        if row["中标与否"] not in {"", "是", "否"} or row["复核状态"] not in {"通过", "待复核"}:
            raise LedgerError("结果枚举无效")
        if row["文件类别"] != "excel_ledger" or row["提取方式"] != "台账整理" or row["来源页码"]:
            raise LedgerError("Excel 固定字段不符合契约")
        try:
            confidence = float(row["置信度"])
            parsed = datetime.fromisoformat(row["解析结果生成日期时间"])
            if not math.isfinite(confidence) or not 0 <= confidence <= 1 or parsed.utcoffset() is None:
                raise ValueError()
        except ValueError as exc:
            raise LedgerError("置信度或带时区时间格式无效") from exc
    seen = set()
    for item in escape_entries:
        sequence, field = item["sequence"], item["field"]
        if type(sequence) is not int or not 1 <= sequence <= len(rows) or field not in FIELDS or (sequence, field) in seen:
            raise LedgerError("CSV 文本转义索引无效")
        seen.add((sequence, field))
        if rows[sequence - 1][field] != "'" + item["original"]:
            raise LedgerError("CSV 文本转义与审计不一致")
        rows[sequence - 1][field] = item["original"]
    return rows


def validate_outputs(output: Path) -> dict[str, Any]:
    if not output.is_dir() or {p.name for p in output.iterdir()} != OUTPUTS:
        raise LedgerError("产物目录必须恰好包含 final.csv、review_queue.csv、ledger.json")
    if any(p.is_symlink() or not p.is_file() for p in output.iterdir()):
        raise LedgerError("产物必须为普通文件")
    # ledger 含逐条来源，可大于映射文件的 8 MiB 限制。
    if (output / "ledger.json").stat().st_size > 128 * 1024 * 1024:
        raise LedgerError("ledger.json 超限")
    ledger = json.loads((output / "ledger.json").read_text(encoding="utf-8"))
    if ledger.get("schema_version") != 1:
        raise LedgerError("ledger 版本不受支持")
    _validate_name_decisions(ledger)
    _validate_award_matches(ledger)
    for name in ("final.csv", "review_queue.csv"):
        path = output / name
        expected = ledger.get("artifacts", {}).get(name, {})
        if expected.get("size") != path.stat().st_size or expected.get("sha256") != digest(path):
            raise LedgerError(f"{name} 的大小/哈希与 ledger 不一致")
    final = _read_csv(output / "final.csv", ledger.get("csv_text_escapes", {}).get("final.csv", []))
    review = _read_csv(output / "review_queue.csv", ledger.get("csv_text_escapes", {}).get("review_queue.csv", []))
    if not final and not review:
        raise LedgerError("缺少整理记录和复核项")
    if any(row["复核状态"] != "待复核" for row in review):
        raise LedgerError("review_queue 中存在非待复核记录")
    records = ledger["records"]
    if len(records) != len(final) or [r["final_sequence"] for r in records] != list(range(1, len(final) + 1)):
        raise LedgerError("ledger 记录与 final 序号不一致")
    collections = [records, ledger["projects"], ledger["groups"], ledger["issues"]]
    if "relationships" in ledger:
        collections.append(ledger["relationships"])
    for collection in collections:
        if len({r["id"] for r in collection}) != len(collection):
            raise LedgerError("ledger 存在重复内部 ID")
    project_ids = {p["id"] for p in ledger["projects"]}
    groups = {g["id"]: g for g in ledger["groups"]}
    source_ids = {s["id"] for s in ledger["sources"]}
    issues = {i["id"]: i for i in ledger["issues"]}
    linked_reviews = set()
    for record, row in zip(records, final):
        if record["company_name"] != row["公司名称"] or (record["project_id"] is not None and record["project_id"] not in project_ids) or record["group_id"] not in groups:
            raise LedgerError("记录名称/项目/分组关联无效")
        if record["source_id"] not in source_ids or not record["occurrences"]:
            raise LedgerError("记录缺少原始来源")
        group = groups[record["group_id"]]
        if group["project_id"] != record["project_id"] or group["source_id"] != record["source_id"]:
            raise LedgerError("跨项目或跨文件串组")
        if any(i not in issues for i in record["issue_ids"]):
            raise LedgerError("记录关联未知问题")
        if bool(record["issue_ids"]) != (row["复核状态"] == "待复核"):
            raise LedgerError("复核状态与问题集合不一致")
        confidence = record.get("confidence")
        if ledger.get("parser_version") == VERSION:
            if (not isinstance(confidence, dict) or confidence.get("meaning") != "rule_reliability_not_probability" or
                    confidence.get("score") != float(row["置信度"])):
                raise LedgerError("置信度与逐条规则依据不一致")
        sequence = record["review_sequence"]
        if sequence is None:
            if row["复核状态"] == "待复核":
                raise LedgerError("待复核记录未进入队列")
            continue
        if type(sequence) is not int or not 1 <= sequence <= len(review) or sequence in linked_reviews:
            raise LedgerError("review 序号关联无效或重复")
        linked_reviews.add(sequence)
        reviewed = review[sequence - 1]
        if any(row[k] != reviewed[k] for k in FIELDS if k not in {"序号", "证据文本"}):
            raise LedgerError("复核副本与 final 业务字段不一致")
        if not reviewed["证据文本"].endswith("\n" + row["证据文本"]):
            raise LedgerError("复核副本未保留原始证据文本")
    all_issue_reviews = set()
    standalone_reviews = set()
    for issue in issues.values():
        for seq in issue["final_sequences"]:
            if type(seq) is not int or not 1 <= seq <= len(final) or issue["id"] not in records[seq - 1]["issue_ids"]:
                raise LedgerError("问题指向无效 final 记录")
        for seq in issue["review_sequences"]:
            if type(seq) is not int or not 1 <= seq <= len(review):
                raise LedgerError("问题指向无效 review 记录")
            all_issue_reviews.add(seq)
            if issue.get("standalone"):
                standalone_reviews.add(seq)
    if all_issue_reviews != set(range(1, len(review) + 1)) or linked_reviews | standalone_reviews != all_issue_reviews:
        raise LedgerError("存在没有审计关联的复核记录")
    names = {}
    for row in final:
        if name_key(row["公司名称"]):
            names.setdefault(name_key(row["公司名称"]), row["公司名称"])
    if ledger["unique_companies"] != list(names.values()):
        raise LedgerError("unique_companies 不等于 final 的规范去重名单")
    expected_counts = {
        "project_count": len(project_ids), "group_count": len(groups),
        "explicit_lot_count": sum(bool(g["lot_name"] or g["lot_code"]) for g in groups.values()),
        "bidding_group_count": sum(g["procurement_status"] == "bidding" for g in groups.values()),
        "non_tender_group_count": sum(g["non_tender"] for g in groups.values()),
        "uncertain_group_count": sum(g["procurement_status"] == "uncertain" for g in groups.values()),
        "roster_group_count": sum(g["scope_type"] == "roster" for g in groups.values()),
        "context_group_count": sum(g["scope_type"] == "context" for g in groups.values()),
        "record_count": len(final), "pending_record_count": sum(r["复核状态"] == "待复核" for r in final),
        "review_record_count": len(review), "issue_count": len(issues), "unique_company_count": len(names),
        "duplicate_mentions_removed": sum(len(r["occurrences"]) - 1 for r in records),
        "corrected_record_count": sum(bool(r["corrections"]) for r in records),
    }
    if ledger.get("parser_version") == VERSION:
        for group in groups.values():
            coverage = group.get("candidate_coverage")
            if (group.get("company_role") == "bidder_name" and
                    (not isinstance(coverage, dict) or type(coverage.get("complete")) is not bool or
                     not isinstance(coverage.get("reasons"), list))):
                raise LedgerError("投标候选覆盖审计无效")
        bidder_projectless = sum(
            record.get("company_role") == "bidder_name" and record.get("project_required") and
            bool(record["company_name"]) and record["project_id"] is None
            for record in records)
        cross_block_names = {}
        cross_block_deduplication = 0
        for record in records:
            source_blocks = {occurrence.get("source_block_id") for occurrence in record["occurrences"]}
            if None in source_blocks:
                raise LedgerError("记录缺少来源投标块审计")
            if len(source_blocks) > 1:
                cross_block_deduplication += 1
            if record["project_id"] and record["company_name"]:
                cross_block_names.setdefault((record["project_id"], name_key(record["company_name"])), set()).update(
                    source_blocks)
        expected_counts.update({
            "bidder_roster_projectless_record_count": bidder_projectless,
            "cross_block_duplicate_participation_count": sum(max(0, len(blocks) - 1)
                                                               for blocks in cross_block_names.values()),
            "cross_block_deduplication_count": cross_block_deduplication,
            "incomplete_bidder_group_count": sum(
                group.get("company_role") == "bidder_name" and
                not group.get("candidate_coverage", {}).get("complete", True)
                for group in groups.values()),
        })
        if bidder_projectless:
            raise LedgerError("bidder roster 存在无项目归属记录")
        if cross_block_deduplication:
            raise LedgerError("检测到跨来源投标块去重")
    if "relationships" in ledger:
        relationships = ledger["relationships"]
        expected_counts.update({
            "relationship_count": len(relationships),
            "unresolved_relationship_count": sum(item["status"] != "matched" for item in relationships),
        })
        target_ids = {item.get("target_project_id") for item in relationships if item.get("target_project_id")}
        if not target_ids <= project_ids:
            raise LedgerError("跨表关系引用未知目标项目")
        candidate_ids = {candidate["project_id"] for item in relationships for candidate in item.get("candidates", [])}
        if not candidate_ids <= project_ids:
            raise LedgerError("跨表关系候选引用未知目标项目")
        source_ids = {source["id"] for source in ledger["sources"]}
        for relation in relationships:
            snapshot = relation.get("source_project")
            if (not isinstance(snapshot, dict) or snapshot.get("id") != relation.get("source_project_id") or
                    snapshot.get("source_id") not in source_ids or not isinstance(snapshot.get("values"), dict) or
                    not isinstance(snapshot.get("cells"), dict)):
                raise LedgerError("跨表关系缺少可追溯的来源项目快照")
        resolutions = ledger.get("relationship_resolutions")
        if not isinstance(resolutions, list) or len({item.get("review_task_id") for item in resolutions}) != len(resolutions):
            raise LedgerError("项目关系人工决定审计无效")
        relations_by_task = {item.get("review_task_id"): item for item in relationships}
        for resolution in resolutions:
            relation = relations_by_task.get(resolution.get("review_task_id"))
            if not relation or resolution.get("decision") not in {"select_project", "deferred"}:
                raise LedgerError("项目关系决定缺少对应任务")
            if resolution["decision"] == "select_project":
                if (relation["status"] != "matched" or relation["basis"] != "user_selection" or
                        relation["target_project_id"] != resolution.get("project_id")):
                    raise LedgerError("项目关系人工选择未正确应用")
            elif relation["status"] == "matched":
                raise LedgerError("标记不确定的项目关系不得自动匹配")
    if ledger["summary"] != expected_counts:
        raise LedgerError("summary 与明细计数不一致")
    return {"validated": True, "summary": expected_counts}


def _validate_name_decisions(ledger: dict) -> None:
    if "name_pairs" not in ledger:
        return  # 兼容只读校验已有 1.1/1.2 产物。
    pairs = {p["id"]: p for p in ledger["name_pairs"]}
    decisions = ledger.get("name_decisions", {})
    if len(pairs) != len(ledger["name_pairs"]) or set(pairs) != set(decisions) or any(d not in {"use_a", "use_b", "different", "uncertain"} for d in decisions.values()):
        raise LedgerError("名称对尚未全部作出有效决定，不能发布最终产物")
    for record in ledger["records"]:
        for change in record["corrections"]:
            if change["rule"] != "model_name_pair":
                continue
            pair = pairs.get(change.get("pair_id"))
            decision = decisions.get(change.get("pair_id"))
            if not pair or decision not in {"use_a", "use_b"} or decision != change.get("decision"):
                raise LedgerError("纠错动作缺少已接受的名称对决定")
            selected = pair["name_a"] if decision == "use_a" else pair["name_b"]
            if name_key(change["after"]) != name_key(selected):
                raise LedgerError("纠错后名称不属于模型比较的两个名称")


def _validate_award_matches(ledger: dict) -> None:
    if "matching_policy" not in ledger:
        if ledger.get("parser_version") == VERSION:
            raise LedgerError("缺少确定性匹配策略")
        return  # 旧版本产物仍可只读核验。
    from .normalize import normalize_name
    records = {r["id"]: r for r in ledger["records"]}
    for record in records.values():
        if not record["occurrences"] or any(
            name_key(normalize_name(o["raw_company"], {})[0]) != name_key(record["company_name"])
            for o in record["occurrences"]
        ):
            raise LedgerError("公司名称不等于原企业字段的格式清洗结果")
        if any(c["rule"] not in {"format_normalization", "trailing_bid_annotation"} for c in record["corrections"]):
            raise LedgerError("不允许通过中标匹配改写企业全称")
    for group in ledger["groups"]:
        if group["award_matching"]["mode"] not in {"group_match", "row_aligned"}:
            raise LedgerError("中标匹配结构无效")
        completeness = group.get("award_completeness")
        if ledger.get("parser_version") == VERSION and (
                not isinstance(completeness, dict) or completeness.get("status") not in {"complete", "partial", "unknown"} or
                completeness.get("basis_type") not in {"explicit", "structural", "none"} or
                type(completeness.get("verified")) is not bool):
            raise LedgerError("中标完整性审计无效")
        selected_ids = set()
        for match in group["award_matches"]:
            if match["status"] not in {"matched", "unresolved", "conflict", "blocked"}:
                raise LedgerError("中标对应状态无效")
            if not match["award_cells"] or any(cell not in {a["cell"] for a in group["awards"]} for cell in match["award_cells"]):
                raise LedgerError("中标对应缺少原始单元格")
            for candidate in match["candidates"]:
                record = records.get(candidate["record_id"])
                if not record or record["group_id"] != group["id"] or candidate["name"] != record["company_name"]:
                    raise LedgerError("中标候选跨组或名称不一致")
            chosen = match["selected_record_id"]
            if match["status"] == "matched":
                record = records.get(chosen)
                if (not record or chosen in selected_ids or record["group_id"] != group["id"] or
                        match["selected_bidder_name"] != record["company_name"] or
                        chosen not in {c["record_id"] for c in match["candidates"]} or
                        (ledger.get("parser_version") == VERSION and
                         match["basis"] not in {"exact_name", "user_selection"})):
                    raise LedgerError("中标对应重复、跨组或名称不一致")
                selected_ids.add(chosen)
            elif chosen is not None or match["selected_bidder_name"] is not None:
                raise LedgerError("未解决的中标信息不得指定企业")
    if ledger.get("parser_version") == VERSION:
        resolutions = ledger.get("resolutions")
        if not isinstance(resolutions, list) or len({item.get("review_task_id") for item in resolutions}) != len(resolutions):
            raise LedgerError("人工中标决定审计无效")
        matches = {match.get("review_task_id"): match for group in ledger["groups"] for match in group["award_matches"]}
        for resolution in resolutions:
            task_id = resolution.get("review_task_id")
            match = matches.get(task_id)
            if not match or resolution.get("decision") not in {"select_bidder", "deferred"}:
                raise LedgerError("人工中标决定缺少对应任务")
            if resolution["decision"] == "select_bidder":
                record = records.get(resolution.get("record_id"))
                if (not record or match["status"] != "matched" or match["basis"] != "user_selection" or
                        match["selected_record_id"] != record["id"] or
                        resolution.get("company_name") != record["company_name"]):
                    raise LedgerError("人工选择未正确应用到投标记录")
            elif match["status"] == "matched":
                raise LedgerError("标记不确定的中标任务不得自动匹配")
