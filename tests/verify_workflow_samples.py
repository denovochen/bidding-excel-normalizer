"""从原 Excel 回放公开 CLI；结构答案为测试夹具，业务确认保持不确定。 @author denovochen"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import shlex
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from ledger_core.artifacts import validate_outputs
from ledger_core.contract import name_key


def semantic_fixture(question: dict) -> dict:
    """只提供领域字段语义，绝不设置物理行范围、group_mode 或 project_mode。"""
    answer = copy.deepcopy(question["answer_template"])
    columns = answer["columns"]
    bidder = "bidder_name" in columns.values()
    for profile in question["columns"]:
        label, header = profile["column"], profile["header"]
        if header in {"监理标段划分", "项目标段"} and not bidder:
            columns[label] = "project_name" if "project_name" not in columns.values() else "lot_name"
        elif columns[label] == "unresolved":
            columns[label] = "evidence"
    answer["project_context_columns"] = list(dict.fromkeys(
        re.match(r"[A-Z]+", item["cell"])[0] for item in question["project_context_candidates"]))
    answer["award_completeness"] = "complete" if "award_name" in columns.values() else "unknown"
    answer["basis"] = "离线样本夹具语义：明确投标/中标角色；其余财政、人员、验收列为来源证据"
    return answer


def verify(path: Path, destination: Path, kind: str, agent_mode: bool = False) -> dict:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    entry = ROOT / ("scripts/excel_agent.py" if agent_mode else "scripts/excel_ledger.py")
    command = [sys.executable, "-B", str(entry), "run", str(path),
               "--output", str(destination)]
    totals = {"cli_calls": 0, "response_characters": 0, "largest_response_characters": 0,
              "structure_questions": 0, "table_relationship_questions": 0,
              "project_questions": 0, "deferred_award_questions": 0}
    call_log = []
    started = time.monotonic()

    def invoke(args):
        process = subprocess.run(args, cwd=ROOT, capture_output=True, text=True, timeout=120)
        if process.returncode not in {0, 3, 4, 5}:
            raise AssertionError(process.stderr + process.stdout)
        if agent_mode and (process.returncode != 0 or len(process.stdout.rstrip("\n")) > 6000):
            raise AssertionError("平台入口未遵守成功交接或响应预算")
        reply = json.loads(process.stdout)
        totals["cli_calls"] += 1
        totals["response_characters"] += len(process.stdout)
        totals["largest_response_characters"] = max(totals["largest_response_characters"], len(process.stdout))
        call_log.append({"kind": reply["kind"], "exit_code": process.returncode,
                         "response_characters": len(process.stdout)})
        return reply

    def detail(reply, question, section):
        collected, offset = [], 0
        while True:
            page = invoke([sys.executable, "-B", str(entry), "question", "--state", reply["state"],
                           "--id", question["question_id"], "--section", section, "--offset", str(offset)])
            collected.extend(page["items"])
            offset += len(page["items"])
            if not page["next_command"]:
                return collected

    while totals["cli_calls"] < 100:
        reply = invoke(command)
        if reply["kind"] == "result":
            break
        if reply.get("validation_errors"):
            raise AssertionError(reply["validation_errors"])
        answers = {}
        envelope = json.loads(Path(reply["answer_file"]).read_text(encoding="utf-8")) if agent_mode else None
        for question in reply["questions"]:
            key = question["question_id"]
            if agent_mode:
                question = dict(question)
                if reply["kind"] == "mapping_required":
                    question["task_type"] = detail(reply, question, "task_type")[0]
                    question["answer_template"] = envelope["answers"][key]
                    sections = (["columns", "project_context_candidates"] if question["task_type"] == "structure"
                                else ["candidate_bidder_sets", "sources"])
                    for section in sections:
                        if section not in question:
                            question[section] = detail(reply, question, section)
                elif "options" not in question:
                    question["options"] = detail(reply, question, "options")
            if question.get("task_type") == "structure":
                if question.get("previous_error"):
                    raise AssertionError(question["previous_error"])
                totals["structure_questions"] += 1
                answers[key] = semantic_fixture(question)
            elif question.get("task_type") == "table_relationship":
                totals["table_relationship_questions"] += 1
                # 测试夹具的区域语义判断；生产逻辑不使用 Sheet 名称分支。
                answers[key] = {"bidder_sets": [item["set_id"] for item in question["candidate_bidder_sets"]]
                                if any("施工" in label for label in question["sources"]) else [],
                                "basis": "样本夹具语义：施工结果关联施工投标名册，其他专业独立保留"}
            elif reply["kind"] == "relationship_review_required":
                totals["project_questions"] += 1
                selected = question["options"][0]["value"]
                if agent_mode and selected != "unresolved":
                    if "evidence_refs" not in question:
                        question["evidence_refs"] = {item["key"]: item["value"] for item in detail(reply, question, "evidence_refs")}
                    answers[key] = {"value": selected, "basis": "测试夹具模拟候选选择，非业务真值验收",
                                    "evidence": question["evidence_refs"][selected]}
                else:
                    answers[key] = selected
            elif reply["kind"] == "award_review_required":
                totals["deferred_award_questions"] += 1
                answers[key] = "unresolved"
            else:
                raise AssertionError("未知交接类型")
        command = shlex.split(reply["next_command"])
        answers_path = Path(command[command.index("--answers") + 1])
        if agent_mode:
            envelope["answers"] = answers
        answers_path.write_text(json.dumps(envelope if agent_mode else answers, ensure_ascii=False), encoding="utf-8")
    else:
        raise AssertionError("超过100次 CLI 调用预算")
    checked = validate_outputs(destination)
    ledger = json.loads((destination / "ledger.json").read_text(encoding="utf-8"))
    expected_records = {"jiangyin": 1768, "xinhe": 289, "yixing": 10241}
    if checked["summary"]["record_count"] != expected_records[kind]:
        raise AssertionError((kind, checked["summary"]))
    if checked["summary"]["cross_block_deduplication_count"] or ledger["source_coverage"]["unaccounted_bidder_row_count"]:
        raise AssertionError("来源覆盖或跨块去重不变量失败")
    if any(decision["decision"] == "select_bidder" for decision in ledger["resolutions"]):
        raise AssertionError("离线回放不能冒充用户选择")
    if kind == "yixing":
        if any(record["company_name"] in {"中标企业名称", "中标单位", "单位名称"} for record in ledger["records"]):
            raise AssertionError("合并表头不能成为企业记录")
        for sheet_year, row, winner in [("2019", 461, "南京长城建设发展有限公司"),
                                        ("2023", 2012, "江苏必和必拓建设有限公司"),
                                        ("2023", 2333, "江苏鑫慧达建设工程有限公司")]:
            matches = [record for record in ledger["records"] if sheet_year in record["sheet"] and
                       name_key(record["company_name"]) == name_key(winner) and
                       any(occurrence["row"] == row for occurrence in record["occurrences"])]
            if len(matches) != 1 or not any(match["selected_record_id"] == matches[0]["id"] and match["basis"] == "exact_name"
                                           for group in ledger["groups"] for match in group["award_matches"]):
                raise AssertionError("已确认准确名称必须通过完整来源块精确匹配")
    if hashlib.sha256(path.read_bytes()).hexdigest() != before:
        raise AssertionError("原文件发生变化")
    return {"sample": kind, "source_sha256": before, "source_unchanged": True,
            "summary": checked["summary"], "source_coverage": ledger["source_coverage"],
            **totals, "elapsed_seconds": round(time.monotonic() - started, 2), "calls": call_log,
            "selection_policy": "结构语义为测试夹具，项目关系模拟首个候选，所有非精确企业保持不确定",
            "model_calls": 0, "not_platform_blind_test": True, "agent_protocol": agent_mode}


def main():
    parser = argparse.ArgumentParser(description="从原始文件回放公开 CLI，无预先 plan")
    for name in ("jiangyin", "xinhe", "yixing"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--agent", action="store_true", help="验证平台短响应入口与公开证据分页")
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=False)
    reports = []
    for name in ("jiangyin", "xinhe", "yixing"):
        report = verify(getattr(args, name), args.output_root / name, name, args.agent)
        reports.append(report)
        print(json.dumps(report, ensure_ascii=False), flush=True)
    (args.output_root / "workflow-verification.json").write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
