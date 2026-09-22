"""组级中标确认问题、人工输入和不确定决定的回归测试。 @author denovochen"""
from __future__ import annotations

import tempfile
import sys
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from ledger_core.normalize import build_outputs
from ledger_core.review import apply_answers, pending_questions
from ledger_core.workbook import inspect_workbooks, read_workbook


class AwardReviewWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def build(self, rows):
        path = self.root / "input.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        workbook.save(path)
        workbook.close()
        book = read_workbook(path)
        plan = inspect_workbooks([book])["suggested_plan"]
        table = plan["sources"][0]["sheets"][0]["tables"][0]
        table["award_completeness"] = {
            "status": "complete", "basis_type": "structural",
            "basis": "测试夹具声明为完整最终结果区域",
        }
        return book, plan, build_outputs([book], plan)

    def test_questions_are_batched_and_keep_the_minimal_three_paths(self):
        rows = [["项目名称", "投标企业名单", "中标单位"]]
        rows.extend([[f"项目{index}", f"企业{index}甲公司、企业{index}乙公司", f"旧企业{index}"]
                     for index in range(1, 7)])
        _, _, prepared = self.build(rows)
        questions, total, remaining = pending_questions(prepared[2], {})
        self.assertEqual((total, remaining, len(questions)), (6, 6, 5))
        for question in questions:
            self.assertEqual(len(question["options"]), 2)
            self.assertTrue(question["options"][0]["label"].endswith(" (Recommended)"))
            self.assertEqual(question["options"][1], {"label": "不确定", "value": "unresolved"})
            self.assertTrue(question["allow_other"])
            self.assertFalse(question["multi_select"])
            self.assertNotIn("法人", question["question"])
            self.assertNotIn("金额", question["question"])

    def test_other_text_selects_a_unique_bidder_and_python_rebuilds_results(self):
        book, plan, prepared = self.build([
            ["项目名称", "投标企业名单", "中标单位"],
            ["项目甲", "甲公司、乙公司", "旧企业名称"],
        ])
        questions, _, _ = pending_questions(prepared[2], {})
        task_id = questions[0]["question_id"]
        state = {"decisions": {}}
        errors = apply_answers(state, prepared[2], {task_id: {"other": "乙公司"}})
        self.assertFalse(errors)
        decision = state["decisions"][task_id]
        self.assertEqual(decision["decision"], "select_bidder")
        self.assertEqual(decision["company_name"], "乙公司")
        self.assertEqual(decision["source"], "manual")
        final, review, ledger = build_outputs([book], plan, award_resolutions=state["decisions"])
        self.assertEqual([row["中标与否"] for row in final], ["否", "是"])
        self.assertFalse(review)
        self.assertEqual(ledger["resolutions"][0]["user_input"], "乙公司")

    def test_unknown_other_text_stays_pending_without_creating_a_company(self):
        _, _, prepared = self.build([
            ["项目名称", "投标企业名单", "中标单位"],
            ["项目甲", "甲公司、乙公司", "旧企业名称"],
        ])
        questions, _, _ = pending_questions(prepared[2], {})
        task_id = questions[0]["question_id"]
        state = {"decisions": {}}
        errors = apply_answers(state, prepared[2], {task_id: {"other": "名单外公司"}})
        self.assertEqual(errors[0]["question_id"], task_id)
        self.assertFalse(state["decisions"])
        next_questions, _, remaining = pending_questions(prepared[2], state["decisions"])
        self.assertEqual(remaining, 1)
        self.assertEqual(next_questions[0]["question_id"], task_id)

    def test_uncertain_decision_is_not_asked_again_and_keeps_review(self):
        book, plan, prepared = self.build([
            ["项目名称", "投标企业名单", "中标单位"],
            ["项目甲", "甲公司、乙公司", "旧企业名称"],
        ])
        questions, _, _ = pending_questions(prepared[2], {})
        task_id = questions[0]["question_id"]
        state = {"decisions": {}}
        self.assertFalse(apply_answers(state, prepared[2], {task_id: "不确定"}))
        final, review, ledger = build_outputs([book], plan, award_resolutions=state["decisions"])
        next_questions, _, remaining = pending_questions(ledger, state["decisions"])
        self.assertEqual((next_questions, remaining), ([], 0))
        self.assertEqual([row["中标与否"] for row in final], ["", ""])
        self.assertEqual(len(review), 2)
        self.assertEqual(ledger["resolutions"][0]["decision"], "deferred")


if __name__ == "__main__":
    unittest.main()
