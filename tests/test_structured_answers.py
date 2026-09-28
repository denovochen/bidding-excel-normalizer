"""结构化提交、业务范围和原核心兼容回归。 @author denovochen"""
from __future__ import annotations

import json
import copy
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import excel_agent as agent
from ledger_core.contract import LedgerError
from ledger_core.artifacts import digest, validate_outputs
from ledger_core.workbook import load_json
from ledger_core.normalize import build_outputs
from ledger_core.coverage import validate_source_coverage
from ledger_core.workbook import read_workbook


class StructuredAnswerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def workbook(self, sheets):
        book = openpyxl.Workbook()
        book.remove(book.active)
        for name, rows in sheets:
            sheet = book.create_sheet(name)
            for row in rows:
                sheet.append(row)
        source = self.root / "业务 输入.xlsx"
        book.save(source)
        book.close()
        return source

    def send(self, reply, patch, *, draft=False):
        return agent.answer(Path(reply["state"]), reply["batch_id"], reply["state_version"],
                            reply["questions"][0]["question_id"], patch, draft=draft)

    def test_structured_submit_ignores_json_indentation_and_replays_once(self):
        source = self.workbook([("清单", [["项目名称", "投标单位", "备注"], ["甲项目", "甲公司", "原件"]])])
        reply = agent.start([source], self.root / "out", 6000, scope=None)
        answers = Path(reply["answer_file"])
        answers.write_text(json.dumps(load_json(answers), separators=(",", ":")), encoding="utf-8")
        patch = {"basis": "逐企业投标，备注仅为来源证据"}
        result = self.send(reply, patch)
        self.assertEqual(result["kind"], "result")
        self.assertEqual(self.send(reply, patch), result)
        self.assertEqual(len(load_json(Path(reply["state"]))["receipts"]), 1)
        self.assertEqual(result["files"], ["final.csv", "review_queue.csv"])
        self.assertEqual(result["internal_files"], ["ledger.json"])
        self.assertTrue(validate_outputs(Path(result["output"]))["validated"])
        with self.assertRaisesRegex(LedgerError, "不同答案"):
            self.send(reply, {"basis": "另一个决定"})

    def test_draft_merges_fields_and_refuses_unresolved_without_committing(self):
        source = self.workbook([("清单", [["项目名称", "投标单位", "附加资料"], ["甲项目", "甲公司", "说明"]])])
        reply = agent.start([source], self.root / "out", 6000, scope=None)
        state = Path(reply["state"])
        before = state.read_bytes()
        with self.assertRaisesRegex(LedgerError, "未解释"):
            self.send(reply, {"basis": "投标明细"})
        self.assertEqual(state.read_bytes(), before)
        for invalid in ({"columns": {"Z": "evidence"}}, {"group_mode": "row"}):
            with self.assertRaises(LedgerError):
                self.send(reply, invalid, draft=True)
        saved = self.send(reply, {"columns": {"C": "evidence"}}, draft=True)
        self.assertEqual(saved["kind"], "answer_saved")
        self.assertEqual(state.read_bytes(), before)
        checked = agent.checked_status(state, resume=True)
        self.assertEqual(checked["state_version"], 1)
        result = self.send(reply, {"basis": "逐企业投标，附加资料为说明"})
        self.assertEqual(result["summary"]["record_count"], 1)

    def test_scope_excludes_before_mapping_even_with_identical_headers_and_changed_order(self):
        source = self.workbook([
            ("后勤服务", [["项目名称", "投标单位"], ["后勤采购", "乙公司"]]),
            ("目标明细", [["项目名称", "投标单位"], ["设备采购", "甲公司"]]),
            ("空白", []),
        ])
        fingerprint = digest(source)
        reply = agent.start([source], self.root / "out", 6000, scope="设备采购投标明细及关联汇总")
        question = load_json(Path(reply["state"]))["reply"]["questions"][0]
        excluded = question["sheets"][0]["sheet_id"]
        self.assertEqual(question["task_type"], "scope")
        self.assertEqual(reply["workflow"]["phase"], "scope")
        with self.assertRaises(LedgerError):
            self.send(reply, {"exclude": {"unknown": "无关"}, "basis": "限定设备采购"})
        reply = self.send(reply, {"exclude": {excluded: "后勤服务采购不属于设备采购"}, "basis": "仅设备采购业务"})
        self.assertEqual(reply["questions"][0]["region_count"], 1)
        self.assertEqual(reply["questions"][0]["regions"][0]["sheet"], "目标明细")
        agent.checked_status(Path(reply["state"]), resume=True)
        result = self.send(reply, {"basis": "设备采购逐企业投标记录"})
        ledger = load_json(Path(result["output"]) / "ledger.json")
        self.assertEqual(result["summary"]["record_count"], 1)
        self.assertEqual(result["summary"]["review_record_count"], 0)
        skipped = next(s for s in ledger["skipped_sheets"] if s["sheet"] == "后勤服务")
        self.assertIn("设备采购", skipped["reason"])
        self.assertEqual(result["delivery"]["excluded_sheet_count"], 1)
        self.assertEqual(digest(source), fingerprint)
        self.assertTrue(validate_outputs(Path(result["output"]))["validated"])
        changed = copy.deepcopy(ledger)
        changed["skipped_sheets"] = []
        with self.assertRaisesRegex(LedgerError, "排除清单"):
            validate_source_coverage(changed)
        legacy_plan = copy.deepcopy(ledger["mapping"])
        legacy_plan["sources"][0]["sheets"][0].pop("scope_exclusion")
        _, _, legacy_ledger = build_outputs([read_workbook(source)], legacy_plan)
        self.assertIn("SOURCE_REGION_SKIPPED", {issue["code"] for issue in legacy_ledger["issues"]})

    def test_scope_keeps_standalone_awards_and_cannot_exclude_everything(self):
        source = self.workbook([("结果", [["项目名称", "中标单位"], ["甲项目", "甲公司"]])])
        reply = agent.start([source], self.root / "out", 6000)
        question = load_json(Path(reply["state"]))["reply"]["questions"][0]
        with self.assertRaisesRegex(LedgerError, "至少保留"):
            self.send(reply, {"exclude": {question["sheets"][0]["sheet_id"]: "没有投标明细"}, "basis": "排除全部"})
        reply = self.send(reply, {"basis": "仅提供中标结果，独立保留"})
        result = self.send(reply, {"basis": "逐项目列出最终中标企业", "award_completeness": "complete"})
        self.assertEqual(result["summary"]["record_count"], 1)
        self.assertEqual(result["kind"], "result")

    def test_resume_rejects_changed_scope_and_keeps_original_batch(self):
        source = self.workbook([("清单", [["项目名称", "投标单位"], ["甲项目", "甲公司"]])])
        reply = agent.start([source], self.root / "out", 6000)
        reply = self.send(reply, {"basis": "全部来源属于目标范围"})
        session = Path(reply["state"])
        inner_path = Path(load_json(session)["reply"]["state"])
        inner = load_json(inner_path)
        inner["scope"]["goal"] = "篡改目标"
        inner_path.write_text(json.dumps(inner), encoding="utf-8")
        with self.assertRaisesRegex(LedgerError, "业务范围"):
            agent.checked_status(session, resume=True)

    def test_scope_evidence_is_pageable_and_draft_exclusion_can_be_corrected(self):
        rows = [["项目名称", "投标单位"], ["甲项目", "甲公司"]]
        source = self.workbook([(f"业务{i}", rows) for i in range(12)])
        reply = agent.start([source], self.root / "out", 4000)
        state = Path(reply["state"])
        self.assertLessEqual(len(agent.encoded(reply)), 4000)
        q = load_json(state)["reply"]["questions"][0]
        total, offset = [], 0
        while True:
            page = agent.question_page(state, q["question_id"], "sheets", offset, 8)
            self.assertLessEqual(len(agent.encoded(page)), 4000)
            total.extend(page["items"])
            offset += len(page["items"])
            if not page["next_command"]:
                break
        self.assertEqual(len(total), 12)
        region = total[0]["regions"][0]
        self.assertEqual(agent.inspect_page(state, region["id"], region["start"], 2, None)["returned_row_count"], 2)
        sheet_id = total[0]["sheet_id"]
        self.send(reply, {"exclude": {sheet_id: "待修正决定"}}, draft=True)
        self.send(reply, {"exclude": {sheet_id: None}}, draft=True)
        result = self.send(reply, {"basis": "各表均为目标范围内的独立业务"})
        q = load_json(state)["reply"]["questions"][0]
        self.assertEqual(q["region_count"], 12)
        with self.assertRaisesRegex(LedgerError, "不同答案"):
            self.send(reply, {"exclude": {sheet_id: "重新排除"}, "basis": "改变已提交范围"})
        self.assertEqual(agent.checked_status(state)["batch_id"], result["batch_id"])

    def test_cli_answer_limits_and_bad_json_do_not_advance_batch(self):
        source = self.workbook([("清单", [["项目名称", "投标单位"], ["甲项目", "甲公司"]])])
        reply = agent.start([source], self.root / "out", 6000, scope=None)
        state = Path(reply["state"])
        before = state.read_bytes()
        command = [sys.executable, str(ROOT / "scripts/excel_agent.py"), "answer", "--state", str(state),
                   "--batch-id", reply["batch_id"], "--state-version", str(reply["state_version"]),
                   "--id", reply["questions"][0]["question_id"], "--json"]
        for value in ("{}" * 1001, "{bad}", '{"basis":NaN}', '{"columns":{"A":[]}}'):
            process = subprocess.run(command + [value], capture_output=True, text=True, timeout=30)
            self.assertEqual(process.returncode, 2, process.stdout)
            error = json.loads(process.stdout)
            self.assertEqual(error["kind"], "error")
            self.assertEqual(state.read_bytes(), before)
            self.assertEqual(process.stderr, "")
        process = subprocess.run(command + ['{"basis":"逐企业投标记录"}'], capture_output=True, text=True, timeout=30)
        self.assertEqual(process.returncode, 0, process.stdout)
        self.assertEqual(json.loads(process.stdout)["kind"], "result")

    def test_batch_answers_validate_together_and_advance_only_once(self):
        source = self.workbook([
            ("项目名册", [["项目名称", "投标单位"], ["甲项目", "甲公司"]]),
            ("标段名册", [["标段名称", "投标单位"], ["甲标段", "乙公司"]]),
        ])
        reply = agent.start([source], self.root / "out", 6000, scope=None)
        self.assertEqual(len(reply["questions"]), 2)
        state = Path(reply["state"])
        args = (state, reply["batch_id"], reply["state_version"], None)
        values = {q["question_id"]: {"basis": "每行一个企业投标记录"} for q in reply["questions"]}
        invalid = copy.deepcopy(values)
        invalid[next(reversed(invalid))]["basis"] = ""
        before = state.read_bytes()
        with self.assertRaisesRegex(LedgerError, "basis"):
            agent.answer(*args, invalid)
        self.assertEqual(state.read_bytes(), before)
        result = agent.answer(*args, values)
        self.assertEqual(result["kind"], "result")
        self.assertEqual(result["summary"]["record_count"], 2)
        self.assertEqual(len(load_json(state)["receipts"]), 1)
        self.assertEqual(agent.answer(*args, values), result)


if __name__ == "__main__":
    unittest.main()
