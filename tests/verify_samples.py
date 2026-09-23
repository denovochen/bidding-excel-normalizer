"""确定性匹配的真实样本验收，不调用模型或改写原表。 @author denovochen"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from ledger_core.artifacts import publish, validate_outputs
from ledger_core.contract import FIELDS, column_number, name_key
from ledger_core.normalize import build_outputs, normalize_name
from ledger_core.review import apply_answers, pending_questions, split_decisions
from ledger_core.workbook import inspect_workbooks, read_workbook


def _reviewed_dispositions(table: dict, columns: dict[str, str]) -> list[dict]:
    mapped = set(columns.values())
    dispositions = {item["column"]: item for item in table.get("column_dispositions", [])}
    for role, label in table.get("columns", {}).items():
        if label not in mapped and label not in dispositions:
            dispositions[label] = {"column": label, "header": role}
    return [{
        "column": label,
        "disposition": "evidence",
        "header": item.get("header", ""),
        "reason": "真实样本结构验收已核对为非输出辅助列",
    } for label, item in sorted(dispositions.items(), key=lambda pair: column_number(pair[0]))
    if label not in mapped]


def _configure_table(table: dict, columns: dict[str, str], table_kind: str, project_mode: str,
                     group_mode: str, completeness: str) -> None:
    table.update({
        "table_kind": table_kind,
        "columns": columns,
        "column_dispositions": _reviewed_dispositions(table, columns),
        "project_mode": project_mode,
        "group_mode": group_mode,
        "bidder_separator": "single",
        "award_completeness": {
            "status": completeness,
            "basis_type": "structural" if completeness != "unknown" else "none",
            "basis": ("真实样本结构验收已核对为完整最终中标结果区域" if completeness == "complete" else
                      "真实样本投标名册不独立声明完整中标结果"),
        },
        "award_mode": "name_match",
        "structure_warnings": [],
    })
    table.pop("group_start_field", None)


def yixing_fixture_plan(book) -> dict:
    """仅按 inspection 的标准角色配置真实样本 plan；生产代码不识别文件或 Sheet 名。"""
    plan = inspect_workbooks([book])["suggested_plan"]
    tables = [table for sheet in plan["sources"][0]["sheets"] for table in sheet.get("tables", [])]
    construction = next(table for table in tables
                        if table.get("table_kind") == "award_summary" and "project_name" in table["columns"])
    rosters = [table for table in tables if table.get("table_kind") == "bidder_roster"]
    award_only = [table for table in tables
                  if table.get("table_kind") == "award_summary" and table is not construction]
    for sheet in plan["sources"][0]["sheets"]:
        if sheet.get("tables"):
            sheet["action"] = "parse"
    _configure_table(construction, {
        "project_year": "C", "project_name": "D", "lot_name": "G", "project_owner": "H",
        "agent": "I", "bidder_count": "M", "award_name": "O", "award_price": "S", "notes": "U",
    }, "award_summary", "merged", "lot", "complete")
    for table in rosters:
        _configure_table(table, {
            "project_year": "C", "project_name": "D", "lot_name": "G", "bidder_name": "I",
            "bidder_price": "K",
        }, "bidder_roster", "merged", "anchor", "partial")
        table["group_start_field"] = "lot_name"
    for table in award_only:
        existing = table["columns"]
        _configure_table(table, {
            "project_year": existing["project_year"], "project_name": "D",
            "award_name": existing["award_name"], "award_price": existing["award_price"],
            "notes": existing["notes"],
        }, "award_summary", "repeated", "row", "complete")
    plan["relationships"] = [{
        "id": "construction_awards_to_rosters",
        "kind": "award_to_bidder_roster",
        "award_tables": [construction["table_id"]],
        "bidder_tables": [table["table_id"] for table in rosters],
        "project_keys": ["project_code", "project_name"],
    }]
    return plan


def reviewed_fixture_plan(book) -> dict:
    """真实样本回归使用已核对结构；不按文件名、列号或企业名称写生产规则。"""
    plan = inspect_workbooks([book])["suggested_plan"]
    for source in plan["sources"]:
        for sheet in source["sheets"]:
            if sheet.get("tables"):
                sheet["action"] = "parse"
            for table in sheet.get("tables", []):
                table["structure_warnings"] = []
                for item in table.get("column_dispositions", []):
                    if item["disposition"] == "unrecognized":
                        item.update({"disposition": "evidence", "reason": "真实样本结构回归已核对为非输出辅助列"})
                table["award_completeness"] = {
                    "status": "complete", "basis_type": "structural",
                    "basis": "真实样本回归已核对为完整最终中标结果区域",
                }
    return plan


def verify_yixing(path: Path, output: Path) -> dict:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    book = read_workbook(path)
    plan = yixing_fixture_plan(book)
    state = {"decisions": {}}
    draft = build_outputs([book], plan)
    task_types = {"relationship": 0, "award": 0}
    while True:
        questions, _, remaining = pending_questions(draft[2], state["decisions"])
        if not remaining:
            break
        if not questions:
            raise AssertionError("宜兴样本存在无法生成有界问题的复核任务")
        kind = "relationship" if questions[0]["question_id"].startswith("relation_review_") else "award"
        task_types[kind] += len(questions)
        answers = {question["question_id"]: question["options"][0]["value"] for question in questions}
        errors = apply_answers(state, draft[2], answers)
        if errors:
            deferred = {error["question_id"]: "unresolved" for error in errors}
            retry_errors = apply_answers(state, draft[2], deferred)
            if retry_errors:
                raise AssertionError(f"宜兴样本冲突问题无法安全保留为不确定: {retry_errors}")
        relation, award = split_decisions(state["decisions"])
        draft = build_outputs([book], plan, relation_resolutions=relation, award_resolutions=award)
    relation, award = split_decisions(state["decisions"])
    final, review, ledger = build_outputs(
        [book], plan, generated_at=draft[2]["generated_at"],
        relation_resolutions=relation, award_resolutions=award)
    expected = {
        "project_count": 205,
        "group_count": 214,
        "record_count": 10244,
        "review_record_count": 785,
        "issue_count": 9,
        "relationship_count": 67,
        "unresolved_relationship_count": 0,
    }
    for key, value in expected.items():
        if ledger["summary"][key] != value:
            raise AssertionError(f"yixing: {key}: {ledger['summary'][key]} != {value}")
    if task_types != {"relationship": 19, "award": 11} or len(state["decisions"]) != 30:
        raise AssertionError("宜兴样本复核任务统计不符合已核验基线")
    if any(record["sheet"] == "施工招标汇总" for record in ledger["records"]):
        raise AssertionError("施工中标汇总应关联到年度投标名册，不应作为重复企业记录发布")
    lot_issue_ids = {issue["id"] for issue in ledger["issues"] if issue["code"] == "LOT_SCOPE_UNRESOLVED"}
    for record, row in zip(ledger["records"], final):
        if lot_issue_ids.intersection(record["issue_ids"]) and (row["标段名称"] or row["标段编号"]):
            raise AssertionError("标段范围未解决时必须降级为项目级结果")
    if (final, review, ledger) != build_outputs(
            [book], plan, generated_at=ledger["generated_at"],
            relation_resolutions=relation, award_resolutions=award):
        raise AssertionError("宜兴样本相同输入/plan/决定/时间不稳定")
    result = publish(output / "yixing", final, review, ledger)
    validate_outputs(output / "yixing")
    after = hashlib.sha256(path.read_bytes()).hexdigest()
    if after != before:
        raise AssertionError("宜兴原文件发生变化")
    return {
        "sample": "yixing", "source_sha256": before, "summary": result["summary"],
        "relationship_review_count": task_types["relationship"],
        "award_review_count": task_types["award"], "resolution_count": len(state["decisions"]),
        "source_unchanged": True, "bidder_names_preserved": True,
        "deterministic": True, "validated": True,
    }


def verify(path: Path, kind: str, output: Path) -> dict:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    book = read_workbook(path)
    plan = reviewed_fixture_plan(book)
    draft = build_outputs([book], plan)
    questions, task_count, remaining = pending_questions(draft[2], {})
    expected_draft = ({"review_record_count": 201, "issue_count": 10, "task_count": 10}
                      if kind == "jiangyin" else
                      {"review_record_count": 15, "issue_count": 4, "task_count": 4})
    if (draft[2]["summary"]["review_record_count"] != expected_draft["review_record_count"] or
            draft[2]["summary"]["issue_count"] != expected_draft["issue_count"] or
            task_count != expected_draft["task_count"] or remaining != task_count):
        raise AssertionError(f"{kind}: 初始组级复核任务或记录数不符合预期")
    state = {"decisions": {}}
    if kind == "jiangyin":
        while questions:
            answers = {question["question_id"]: question["options"][0]["value"] for question in questions}
            if apply_answers(state, draft[2], answers):
                raise AssertionError("江阴样本推荐答案未能应用")
            draft = build_outputs([book], plan, award_resolutions=state["decisions"])
            questions, _, _ = pending_questions(draft[2], state["decisions"])
    else:
        groups = {group["id"]: group for group in draft[2]["groups"]}
        issues = {issue["review_task_id"]: issue for issue in draft[2]["issues"] if issue.get("review_task_id")}
        answers = {}
        for question in questions:
            group = groups[issues[question["question_id"]]["group_id"]]
            answers[question["question_id"]] = (
                "unresolved" if group["anchor_row"] == 51 else question["options"][0]["value"])
        if apply_answers(state, draft[2], answers):
            raise AssertionError("汇总表样本人工答案未能应用")
    final, review, ledger = build_outputs([book], plan, award_resolutions=state["decisions"])
    expected = ({"project_count": 108, "group_count": 124, "record_count": 1768,
                 "review_record_count": 0, "issue_count": 0}
                if kind == "jiangyin" else
                {"project_count": 7, "group_count": 54, "record_count": 289,
                 "review_record_count": 4, "issue_count": 1})
    for key, value in expected.items():
        if ledger["summary"][key] != value:
            raise AssertionError(f"{kind}: {key}: {ledger['summary'][key]} != {value}")
    if any(r["项目编号"] or r["标段编号"] or r["投标排名"] for r in final):
        raise AssertionError("原表缺少的官方编号/排名被补造")
    for record, row in zip(ledger["records"], final):
        for occurrence in record["occurrences"]:
            if name_key(row["公司名称"]) != name_key(normalize_name(occurrence["raw_company"], {})[0]):
                raise AssertionError("投标名称被中标名称覆盖")
    if "name_pairs" in ledger or "name_decisions" in ledger:
        raise AssertionError("不应产生模型名称交接")
    records_by_row = {}
    for record in ledger["records"]:
        records_by_row.setdefault(record["occurrences"][0]["row"], []).append(record)
    if kind == "jiangyin":
        if final[-1]["公司名称"] or not final[-1]["项目名称"]:
            raise AssertionError("设计阶段项目应保留项目字段并留空企业")
        for source_row in (12, 125, 251, 289, 348, 354, 425, 656):
            record = records_by_row[source_row][0]
            row = final[record["final_sequence"] - 1]
            if row["中标与否"] != "是" or row["复核状态"] != "通过":
                raise AssertionError(f"用户确认的中标行未被应用: {source_row}")
    else:
        if any(g["award_matching"]["mode"] != "group_match" for g in ledger["groups"]):
            raise AssertionError("一格多家企业的名单表误用同行判定")
        for source_row in (4,):
            source_records = records_by_row[source_row]
            winners = [r for r in source_records if final[r["final_sequence"] - 1]["中标与否"] == "是"]
            if len(winners) != 1 or winners[0]["id"] != source_records[0]["id"]:
                raise AssertionError(f"原表已核实的名单对应错误: {source_row}")
        for source_row in (16, 35, 44):
            source_records = records_by_row[source_row]
            winners = [record for record in source_records
                       if final[record["final_sequence"] - 1]["中标与否"] == "是"]
            if len(winners) != 1 or winners[0]["id"] != source_records[0]["id"]:
                raise AssertionError(f"用户选择的名单候选未被应用: {source_row}")
        for source_row in (51,):
            source_records = records_by_row[source_row]
            if (any(final[r["final_sequence"] - 1]["中标与否"] for r in source_records) or
                    any(final[r["final_sequence"] - 1]["复核状态"] != "待复核" for r in source_records)):
                raise AssertionError("用户选择不确定的组必须继续保留复核")
    if (final, review, ledger) != build_outputs([book], plan, generated_at=ledger["generated_at"],
                                                award_resolutions=state["decisions"]):
        raise AssertionError("相同输入/映射/时间不稳定")
    result = publish(output / kind, final, review, ledger)
    validate_outputs(output / kind)
    if hashlib.sha256(path.read_bytes()).hexdigest() != before:
        raise AssertionError("原文件发生变化")
    return {"sample": kind, "source_sha256": before, "summary": result["summary"],
            "review_task_count": task_count, "resolution_count": len(state["decisions"]),
            "source_unchanged": True, "bidder_names_preserved": True, "deterministic": True, "validated": True}


def main() -> None:
    parser = argparse.ArgumentParser(description="两份已核实样本的确定性验收，不是生产入口")
    parser.add_argument("--jiangyin", type=Path, required=True)
    parser.add_argument("--xinhe", type=Path, required=True)
    parser.add_argument("--yixing", type=Path)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--mineru-sample", type=Path)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=False)
    results = [verify(args.jiangyin, "jiangyin", args.output_root),
               verify(args.xinhe, "xinhe", args.output_root)]
    if args.yixing:
        results.append(verify_yixing(args.yixing, args.output_root))
    if args.mineru_sample:
        for name in ("final.csv", "review_queue.csv"):
            path = args.mineru_sample / name
            with path.open(encoding="utf-8-sig", newline="") as stream:
                if next(csv.reader(stream)) != FIELDS or not path.read_bytes().startswith(bytes([0xEF, 0xBB, 0xBF])):
                    raise AssertionError("与真实 MinerU CSV 列/编码不一致")
    report = {
        "acceptance_date": "2026-09-23",
        "samples": results,
        "mineru_headers_checked": bool(args.mineru_sample),
        "model_calls": 0,
        "selection_policy": "验收脚本选择脚本提供的首个有界候选；候选冲突时保留为不确定",
        "unverified": ["QM/Qwen 部署环境", "Spider/数据库/报告服务", "全部非精确候选的外部业务真值"],
    }
    (args.output_root / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
