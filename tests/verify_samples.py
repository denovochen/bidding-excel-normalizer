"""真实样本验收：可读取宿主模型已完成任务的决定；默认只测保守基线。 @author denovochen"""
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
from ledger_core.contract import FIELDS
from ledger_core.normalize import build_outputs
from ledger_core.workbook import inspect_workbooks, read_workbook


def verify(path: Path, kind: str, output: Path, state: dict | None) -> dict:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    book = read_workbook(path)
    plan = inspect_workbooks([book])["suggested_plan"]
    base_final, base_review, baseline = build_outputs([book], plan)
    expected = ({"project_count": 108, "group_count": 124, "record_count": 1768,
                 "review_record_count": 201, "issue_count": 10}
                if kind == "jiangyin" else
                {"project_count": 7, "group_count": 54, "record_count": 289,
                 "review_record_count": 15, "issue_count": 4})
    for key, value in expected.items():
        if baseline["summary"][key] != value:
            raise AssertionError(f"{kind}: baseline {key}: {baseline['summary'][key]} != {value}")
    if state:
        if before not in {s["sha256"] for s in state["sources"]} or state.get("result") is None:
            raise AssertionError("决定状态必须属于已完成的同一输入任务")
        decisions = {p["id"]: state["decisions"][p["id"]] for p in baseline["name_pairs"]}
        origin = "completed_host_workflow"
    else:
        # 这是测试桩，明确全部无法确定，不冒充真实模型判断。
        decisions = {p["id"]: "uncertain" for p in baseline["name_pairs"]}
        origin = "test_all_uncertain"
    final, review, ledger = build_outputs([book], plan, generated_at=baseline["generated_at"], name_decisions=decisions)
    if len(final) != len(base_final):
        raise AssertionError("名称判断不应删除项目参与记录")
    fixed = ("序号", "项目名称", "项目编号", "标段名称", "标段编号", "投标排名", "依据文件路径")
    if any(any(a[k] != b[k] for k in fixed) for a, b in zip(base_final, final)):
        raise AssertionError("名称判断改变了无关业务字段")
    if any(r["项目编号"] or r["标段编号"] or r["投标排名"] for r in final):
        raise AssertionError("样本不存在的官方编号/排名被补造")
    if kind == "jiangyin":
        if final[-1]["公司名称"] or not final[-1]["项目名称"]:
            raise AssertionError("设计阶段项目应保留项目字段并留空企业")
        for source_row in (12, 125):
            rec = next(r for r in ledger["records"] if r["occurrences"][0]["row"] == source_row)
            row = final[rec["final_sequence"] - 1]
            if row["中标与否"] != "是" or row["复核状态"] != "通过":
                raise AssertionError("未招投标备注不应清空原表明确填写的结果")
    if (final, review, ledger) != build_outputs([book], plan, generated_at=ledger["generated_at"], name_decisions=decisions):
        raise AssertionError("相同输入/映射/决定/时间不稳定")
    result = publish(output / kind, final, review, ledger)
    validate_outputs(output / kind)
    if hashlib.sha256(path.read_bytes()).hexdigest() != before:
        raise AssertionError("原文件发生变化")
    return {"sample": kind, "source_sha256": before, "summary": result["summary"],
            "decision_origin": origin, "source_unchanged": True, "deterministic": True, "validated": True}


def main() -> None:
    parser = argparse.ArgumentParser(description="两份已核实样本的基线/模型闭环验收，不是生产入口")
    parser.add_argument("--jiangyin", type=Path, required=True)
    parser.add_argument("--xinhe", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--decision-state", type=Path, help="可选：宿主模型已完成 resolve 的 state.json")
    parser.add_argument("--mineru-sample", type=Path)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=False)
    state = json.loads(args.decision_state.read_text()) if args.decision_state else None
    results = [verify(args.jiangyin, "jiangyin", args.output_root, state),
               verify(args.xinhe, "xinhe", args.output_root, state)]
    if args.mineru_sample:
        for name in ("final.csv", "review_queue.csv"):
            path = args.mineru_sample / name
            with path.open(encoding="utf-8-sig", newline="") as f:
                if next(csv.reader(f)) != FIELDS or not path.read_bytes().startswith(bytes([0xEF, 0xBB, 0xBF])):
                    raise AssertionError("与真实 MinerU CSV 列/编码不一致")
    report = {"samples": results, "mineru_headers_checked": bool(args.mineru_sample),
              "qm_model_tested": False, "note": "仅实际指定决定来源；未假称已在 QM 本地模型验收。"}
    (args.output_root / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
