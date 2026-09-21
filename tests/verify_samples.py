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
from ledger_core.contract import FIELDS, name_key
from ledger_core.normalize import build_outputs, normalize_name
from ledger_core.workbook import inspect_workbooks, read_workbook


def verify(path: Path, kind: str, output: Path) -> dict:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    book = read_workbook(path)
    plan = inspect_workbooks([book])["suggested_plan"]
    final, review, ledger = build_outputs([book], plan)
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
                raise AssertionError(f"原表已核实的中标行未被识别: {source_row}")
    else:
        if any(g["award_matching"]["mode"] != "group_match" for g in ledger["groups"]):
            raise AssertionError("一格多家企业的名单表误用同行判定")
        for source_row in (4, 16, 35, 44):
            source_records = records_by_row[source_row]
            winners = [r for r in source_records if final[r["final_sequence"] - 1]["中标与否"] == "是"]
            if len(winners) != 1 or winners[0]["id"] != source_records[0]["id"]:
                raise AssertionError(f"原表已核实的名单对应错误: {source_row}")
        if any(final[r["final_sequence"] - 1]["中标与否"] for r in records_by_row[51]):
            raise AssertionError("主体名称不同且缺乏辅助证据时不能强行选最高分")
    if (final, review, ledger) != build_outputs([book], plan, generated_at=ledger["generated_at"]):
        raise AssertionError("相同输入/映射/时间不稳定")
    result = publish(output / kind, final, review, ledger)
    validate_outputs(output / kind)
    if hashlib.sha256(path.read_bytes()).hexdigest() != before:
        raise AssertionError("原文件发生变化")
    return {"sample": kind, "source_sha256": before, "summary": result["summary"],
            "source_unchanged": True, "bidder_names_preserved": True, "deterministic": True, "validated": True}


def main() -> None:
    parser = argparse.ArgumentParser(description="两份已核实样本的确定性验收，不是生产入口")
    parser.add_argument("--jiangyin", type=Path, required=True)
    parser.add_argument("--xinhe", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--mineru-sample", type=Path)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=False)
    results = [verify(args.jiangyin, "jiangyin", args.output_root),
               verify(args.xinhe, "xinhe", args.output_root)]
    if args.mineru_sample:
        for name in ("final.csv", "review_queue.csv"):
            path = args.mineru_sample / name
            with path.open(encoding="utf-8-sig", newline="") as stream:
                if next(csv.reader(stream)) != FIELDS or not path.read_bytes().startswith(bytes([0xEF, 0xBB, 0xBF])):
                    raise AssertionError("与真实 MinerU CSV 列/编码不一致")
    report = {"samples": results, "mineru_headers_checked": bool(args.mineru_sample),
              "name_model_calls": 0, "qm_environment_tested": False}
    (args.output_root / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
