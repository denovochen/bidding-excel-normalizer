"""跨 Sheet 项目关系、关系复核与标段降级回归。 @author denovochen"""
from __future__ import annotations

import copy
import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from ledger_core.artifacts import publish, validate_outputs
from ledger_core.contract import LedgerError
from ledger_core.normalize import build_outputs
from ledger_core.review import apply_answers, pending_questions, split_decisions
from ledger_core.workbook import apply_plan_patch, inspect_workbooks, read_workbook


class CrossSheetRelationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def book(self, summary_rows, roster_rows):
        path = self.root / "input.xlsx"
        workbook = openpyxl.Workbook()
        summary = workbook.active
        summary.title = "结果汇总"
        for row in summary_rows:
            summary.append(row)
        roster = workbook.create_sheet("投标明细")
        for row in roster_rows:
            roster.append(row)
        workbook.save(path)
        workbook.close()
        return read_workbook(path)

    def plan(self, book):
        inspection = inspect_workbooks([book])
        plan = inspection["suggested_plan"]
        tables = [table for sheet in plan["sources"][0]["sheets"] for table in sheet.get("tables", [])]
        summary = next(table for table in tables if table["table_kind"] == "award_summary")
        roster = next(table for table in tables if table["table_kind"] == "bidder_roster")
        summary["award_completeness"] = {
            "status": "complete", "basis_type": "structural",
            "basis": "测试汇总区域完整列出最终中标结果",
        }
        roster["award_completeness"] = {
            "status": "partial", "basis_type": "none", "basis": "投标名单不含中标结果",
        }
        plan["relationships"] = [{
            "id": "summary_to_roster", "kind": "award_to_bidder_roster",
            "award_tables": [summary["table_id"]], "bidder_tables": [roster["table_id"]],
            "project_keys": ["project_code", "project_name"],
        }]
        return plan

    def test_exact_project_relation_moves_award_to_bidder_records(self):
        book = self.book(
            [["工程名称", "标段", "中标单位"], ["项目甲", "1标段", "甲公司"]],
            [["项目名称", "标段名称", "投标单位名称"],
             ["项目甲", "一标段", "甲公司"], ["项目甲", "一标段", "乙公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in final], [("甲公司", "是"), ("乙公司", "否")])
        self.assertFalse(review)
        self.assertEqual(ledger["summary"]["relationship_count"], 1)
        self.assertEqual(ledger["summary"]["unresolved_relationship_count"], 0)
        self.assertEqual(ledger["relationships"][0]["basis"], "exact_project_key")
        self.assertTrue(all(record["sheet"] == "投标明细" for record in ledger["records"]))
        publish(self.root / "output", final, review, ledger)

    def test_structure_roles_ids_and_plan_patch_are_stable(self):
        book = self.book(
            [["项目编号", "项目名称", "中标单位"], ["P-001", "项目甲", "甲公司"]],
            [["项目编号", "项目名称", "投标单位名称"],
             ["P-001", "项目甲", "甲公司"], ["P-001", "项目甲", "乙公司"]],
        )
        first = inspect_workbooks([book])
        second = inspect_workbooks([book])
        first_tables = [table for sheet in first["suggested_plan"]["sources"][0]["sheets"]
                        for table in sheet.get("tables", [])]
        second_tables = [table for sheet in second["suggested_plan"]["sources"][0]["sheets"]
                         for table in sheet.get("tables", [])]
        self.assertEqual([table["table_id"] for table in first_tables],
                         [table["table_id"] for table in second_tables])
        self.assertEqual([table["table_kind"] for table in first_tables],
                         ["award_summary", "bidder_roster"])
        patch = {
            "schema_version": 1,
            "table_updates": [{
                "table_ids": [table["table_id"] for table in first_tables],
                "set": {"structure_warnings": []},
            }],
            "relationships": [{
                "kind": "award_to_bidder_roster",
                "award_tables": [first_tables[0]["table_id"]],
                "bidder_tables": [first_tables[1]["table_id"]],
            }],
        }
        planned = apply_plan_patch(first, patch)
        self.assertNotIn("relationships", first["suggested_plan"])
        self.assertEqual(planned["relationships"][0]["award_tables"], [first_tables[0]["table_id"]])
        with self.assertRaisesRegex(LedgerError, "不可修改字段"):
            apply_plan_patch(first, {
                "schema_version": 1,
                "table_updates": [{"table_id": first_tables[0]["table_id"], "set": {"table_id": "changed"}}],
            })

    def test_old_plan_without_table_metadata_preserves_single_table_behavior(self):
        book = self.book(
            [["项目名称", "投标单位名称", "中标单位"], ["项目甲", "甲公司", "甲公司"]],
            [],
        )
        modern = inspect_workbooks([book])["suggested_plan"]
        modern["sources"][0]["sheets"][1] = {"name": "投标明细", "action": "skip", "reason": "空白工作表"}
        table = modern["sources"][0]["sheets"][0]["tables"][0]
        table["award_completeness"] = {
            "status": "complete", "basis_type": "structural", "basis": "测试区域完整",
        }
        legacy = copy.deepcopy(modern)
        legacy_table = legacy["sources"][0]["sheets"][0]["tables"][0]
        legacy_table.pop("table_id")
        legacy_table.pop("table_kind")
        modern_result = build_outputs([book], modern, generated_at="2026-09-23T12:00:00+08:00")
        legacy_result = build_outputs([book], legacy, generated_at="2026-09-23T12:00:00+08:00")
        self.assertEqual(modern_result[:2], legacy_result[:2])
        for key in ("projects", "groups", "records", "issues", "summary"):
            self.assertEqual(modern_result[2][key], legacy_result[2][key])

    def test_project_code_precedes_name_and_lot_text_does_not_block_relation(self):
        book = self.book(
            [["项目编号", "项目名称", "标段", "中标单位"],
             ["P-001", "旧项目名称", "施工一标段", "甲公司"]],
            [["项目编号", "项目名称", "标段名称", "投标单位名称"],
             ["P-001", "新项目名称", "第一合同段", "甲公司"],
             ["P-001", "新项目名称", "第一合同段", "乙公司"],
             ["P-002", "旧项目名称", "施工一标段", "甲公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual(ledger["relationships"][0]["target_project_name"], "新项目名称")
        self.assertEqual(ledger["relationships"][0]["basis"], "exact_project_key")
        selected = [row for row in final if row["项目编号"] == "P-001"]
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in selected],
                         [("甲公司", "是"), ("乙公司", "否")])
        self.assertFalse(review)

    def test_nonexact_candidates_are_bounded_year_is_auxiliary_and_bidders_do_not_reverse_match(self):
        roster = [
            ["项目名称", "年度", "投标单位名称"],
            ["东片灌溉改造工程（财政补助）", "2023", "唯一中标企业"],
            ["西片道路提升项目", "2024", "乙公司"],
            ["南片渠道项目", "2024", "丙公司"],
            ["北片泵站项目", "2024", "丁公司"],
            ["中片土地整理项目", "2024", "戊公司"],
            ["河西生态项目", "2024", "己公司"],
            ["河东生态项目", "2024", "庚公司"],
        ]
        book = self.book(
            [["项目名称", "年度", "中标单位"],
             ["东片灌溉改造项目（财政补助）", "2024", "唯一中标企业"]],
            roster,
        )
        plan = self.plan(book)
        _, _, ledger = build_outputs([book], plan)
        relation = ledger["relationships"][0]
        self.assertEqual(relation["status"], "unresolved")
        self.assertEqual(len(relation["candidates"]), 5)
        self.assertEqual(relation["candidates"][0]["project_name"], "东片灌溉改造工程(财政补助)")
        self.assertFalse(relation["candidates"][0]["year_match"])
        self.assertNotIn("bidder", json.dumps(relation["candidates"], ensure_ascii=False).lower())
        questions, total, remaining = pending_questions(ledger, {})
        self.assertEqual((total, remaining), (1, 1))
        self.assertEqual(len(questions[0]["options"]), 6)
        self.assertFalse(questions[0]["allow_other"])

    def test_multilevel_unit_name_uses_bidder_header_context(self):
        path = self.root / "headers.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["项目名称", "投标单位信息", None])
        sheet.append([None, "序号", "单位名称"])
        sheet.append(["项目甲", 1, "甲公司"])
        sheet.merge_cells("A1:A2")
        sheet.merge_cells("B1:C1")
        workbook.save(path)
        workbook.close()
        plan = inspect_workbooks([read_workbook(path)])["suggested_plan"]
        table = plan["sources"][0]["sheets"][0]["tables"][0]
        self.assertEqual(table["columns"]["bidder_name"], "C")
        self.assertEqual(table["table_kind"], "bidder_roster")

    def test_nonexact_project_requires_bounded_relation_selection(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲（财政补助）", "甲公司"]],
            [["工程名称", "投标企业名称"], ["项目甲", "甲公司"], ["项目甲", "乙公司"]],
        )
        plan = self.plan(book)
        final, review, draft = build_outputs([book], plan)
        questions, total, remaining = pending_questions(draft, {})
        self.assertEqual((total, remaining, len(questions)), (1, 1, 1))
        self.assertEqual(len(questions[0]["options"]), 2)
        self.assertFalse(questions[0]["allow_other"])
        state = {"decisions": {}}
        selected = questions[0]["options"][0]["value"]
        self.assertFalse(apply_answers(state, draft, {questions[0]["question_id"]: selected}))
        relation, award = split_decisions(state["decisions"])
        final, review, ledger = build_outputs([book], plan, award_resolutions=award,
                                              relation_resolutions=relation)
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in final], [("甲公司", "是"), ("乙公司", "否")])
        self.assertFalse(review)
        self.assertEqual(ledger["relationships"][0]["basis"], "user_selection")

    def test_deferred_relation_publishes_roster_and_standalone_review(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲（旧称）", "甲公司"]],
            [["项目名称", "投标单位名称"], ["项目甲", "甲公司"], ["项目甲", "乙公司"]],
        )
        plan = self.plan(book)
        _, _, draft = build_outputs([book], plan)
        question = pending_questions(draft, {})[0][0]
        state = {"decisions": {}}
        self.assertFalse(apply_answers(state, draft, {question["question_id"]: "unresolved"}))
        relation, award = split_decisions(state["decisions"])
        final, review, ledger = build_outputs([book], plan, award_resolutions=award,
                                              relation_resolutions=relation)
        self.assertEqual([row["中标与否"] for row in final], ["", ""])
        self.assertEqual(len(review), 1)
        self.assertEqual(ledger["summary"]["unresolved_relationship_count"], 1)
        publish(self.root / "deferred", final, review, ledger)

    def test_clear_multilot_routes_each_award_without_scope_issue(self):
        book = self.book(
            [["项目名称", "标段名称", "中标单位"],
             ["项目甲", "1标段", "甲公司"], ["项目甲", "2标段", "乙公司"]],
            [["项目名称", "标段", "投标单位名称"],
             ["项目甲", "一标段", "甲公司"], ["项目甲", "一标段", "乙公司"],
             ["项目甲", "二标段", "乙公司"], ["项目甲", "二标段", "丙公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([(row["标段名称"], row["公司名称"], row["中标与否"]) for row in final],
                         [("一标段", "甲公司", "是"), ("一标段", "乙公司", "否"),
                          ("二标段", "乙公司", "是"), ("二标段", "丙公司", "否")])
        self.assertNotIn("LOT_SCOPE_UNRESOLVED", [issue["code"] for issue in ledger["issues"]])
        self.assertFalse(review)

    def test_sparse_lot_anchor_keeps_continuation_bidders_in_one_group(self):
        book = self.book(
            [["项目名称", "标段名称", "中标单位"], ["项目甲", "一标段", "甲公司"]],
            [["项目名称", "标段名称", "投标单位名称"],
             ["项目甲", "第一合同段", "甲公司"],
             ["项目甲", None, "乙公司"],
             ["项目甲", None, "丙公司"]],
        )
        plan = self.plan(book)
        roster = next(table for sheet in plan["sources"][0]["sheets"] for table in sheet.get("tables", [])
                      if table["table_kind"] == "bidder_roster")
        roster["project_mode"] = "repeated"
        roster["group_mode"] = "anchor"
        roster["group_start_field"] = "lot_name"
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in final],
                         [("甲公司", "是"), ("乙公司", "否"), ("丙公司", "否")])
        self.assertEqual(len(ledger["groups"]), 1)
        self.assertFalse(review)

    def test_multilot_summary_with_single_roster_group_keeps_project_level_results_under_review(self):
        book = self.book(
            [["项目名称", "标段名称", "中标单位"],
             ["项目甲", "1标段", "甲公司"], ["项目甲", "2标段", "乙公司"]],
            [["项目名称", "投标单位名称"],
             ["项目甲", "甲公司"], ["项目甲", "乙公司"], ["项目甲", "丙公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([row["中标与否"] for row in final], ["是", "是", "否"])
        self.assertTrue(all(not row["标段名称"] and not row["标段编号"] for row in final))
        self.assertTrue(all(row["复核状态"] == "待复核" for row in final))
        self.assertEqual({issue["code"] for issue in ledger["issues"]}, {"LOT_SCOPE_UNRESOLVED"})
        self.assertEqual(len(review), 3)

    def test_unclear_multilot_relation_collapses_to_project_level_without_fake_lot_losses(self):
        book = self.book(
            [["项目名称", "标段名称", "中标单位"],
             ["项目甲", "采购一包", "甲公司"], ["项目甲", "采购二包", "乙公司"]],
            [["项目名称", "标段名称", "投标单位名称"],
             ["项目甲", "东片", "甲公司"], ["项目甲", "东片", "乙公司"], ["项目甲", "东片", "丙公司"],
             ["项目甲", "西片", "甲公司"], ["项目甲", "西片", "乙公司"], ["项目甲", "西片", "丙公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in final],
                         [("甲公司", ""), ("乙公司", ""), ("丙公司", ""),
                          ("甲公司", ""), ("乙公司", ""), ("丙公司", "")])
        self.assertTrue(all(not row["标段名称"] and not row["标段编号"] for row in final))
        self.assertTrue(all(row["复核状态"] == "待复核" for row in final))
        self.assertEqual(len(review), 6)
        self.assertIn("LOT_SCOPE_UNRESOLVED", {issue["code"] for issue in ledger["issues"]})
        self.assertEqual(ledger["summary"]["cross_block_duplicate_participation_count"], 3)
        self.assertEqual(ledger["summary"]["cross_block_deduplication_count"], 0)
        self.assertTrue(ledger["relationships"][0].get("project_level_fallback_group_id"))

    def test_incomplete_bidder_block_stops_relation_and_award_recommendation(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲", "甲公司"]],
            [["项目名称", "投标单位名称"], ["项目甲", "甲公司"], ["项目甲", "坏)公司"]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([(row["公司名称"], row["中标与否"]) for row in final], [("甲公司", "")])
        self.assertEqual(ledger["relationships"][0]["status"], "blocked")
        self.assertIn("PROJECT_RELATION_COVERAGE_INCOMPLETE", {issue["code"] for issue in ledger["issues"]})
        self.assertEqual(pending_questions(ledger, {})[2], 0)
        self.assertTrue(review)

    def test_bidder_count_gap_blocks_relation_before_award_consumption(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲", "甲公司"]],
            [["项目名称", "投标单位名称", "投标企业数量"],
             ["项目甲", "甲公司", 3], ["项目甲", "乙公司", 3]],
        )
        plan = self.plan(book)
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([row["中标与否"] for row in final], ["", ""])
        self.assertEqual(ledger["relationships"][0]["status"], "blocked")
        self.assertIn("BIDDER_COUNT_MISMATCH", {issue["code"] for issue in ledger["issues"]})
        self.assertIn("PROJECT_RELATION_COVERAGE_INCOMPLETE", {issue["code"] for issue in ledger["issues"]})
        self.assertEqual(pending_questions(ledger, {})[2], 0)
        self.assertTrue(review)

    def test_output_validation_requires_relation_source_snapshot(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲", "甲公司"]],
            [["项目名称", "投标单位名称"], ["项目甲", "甲公司"], ["项目甲", "乙公司"]],
        )
        plan = self.plan(book)
        output = self.root / "audited-output"
        publish(output, *build_outputs([book], plan))
        ledger_path = output / "ledger.json"
        ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
        ledger["relationships"][0]["source_project"]["source_id"] = "source_missing"
        ledger_path.write_text(json.dumps(ledger, ensure_ascii=False), encoding="utf-8")
        with self.assertRaisesRegex(LedgerError, "来源项目快照"):
            validate_outputs(output)

    def test_cli_relation_review_precedes_award_review(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲（财政）", "甲建设工程有限公司"]],
            [["项目名称", "投标单位名称"], ["项目甲", "甲建设有限公司"], ["项目甲", "乙公司"]],
        )
        plan = self.plan(book)
        plan_path = self.root / "plan.json"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        output = self.root / "cli-output"
        first = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "run",
                                str(book.path), "--plan", str(plan_path), "--output", str(output)],
                               text=True, capture_output=True)
        self.assertEqual(first.returncode, 5, first.stderr)
        relation_reply = json.loads(first.stdout)
        self.assertEqual(relation_reply["kind"], "relationship_review_required")
        answers = self.root / "relation-answers.json"
        question = relation_reply["questions"][0]
        answers.write_text(json.dumps({question["question_id"]: question["options"][0]["value"]}), encoding="utf-8")
        second = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "resolve",
                                 "--state", relation_reply["state"], "--answers", str(answers)],
                                text=True, capture_output=True)
        self.assertEqual(second.returncode, 4, second.stderr)
        award_reply = json.loads(second.stdout)
        self.assertEqual(award_reply["kind"], "award_review_required")
        award_question = award_reply["questions"][0]
        answers.write_text(json.dumps({award_question["question_id"]: award_question["options"][0]["value"]}),
                           encoding="utf-8")
        final = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "resolve",
                                "--state", award_reply["state"], "--answers", str(answers)],
                               text=True, capture_output=True)
        self.assertEqual(final.returncode, 0, final.stderr)
        with (output / "final.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual([row["中标与否"] for row in rows], ["是", "否"])

    def test_compact_inspection_plan_patch_and_internal_cleanup(self):
        book = self.book(
            [["项目名称", "中标单位"], ["项目甲", "甲公司"]],
            [["项目名称", "投标单位名称"], ["项目甲", "甲公司"], ["项目甲", "乙公司"]],
        )
        output = self.root / "patched-output"
        initial = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "run",
                                  str(book.path), "--output", str(output)], text=True, capture_output=True)
        self.assertEqual(initial.returncode, 3, initial.stderr)
        reply = json.loads(initial.stdout)
        self.assertEqual(reply["kind"], "mapping_required")
        self.assertNotIn("suggested_plan", reply["inspection"])
        inspection_path = Path(reply["inspection"]["inspection_path"])
        self.assertTrue(inspection_path.is_file())
        candidate = reply["inspection"]["relationship_candidates"][0]
        stored = json.loads(inspection_path.read_text(encoding="utf-8"))
        tables = {table["table_id"]: table for source in stored["suggested_plan"]["sources"]
                  for sheet in source["sheets"] for table in sheet.get("tables", [])}
        summary_id = candidate["award_tables"][0]
        roster_id = candidate["bidder_tables"][0]
        patch = {
            "schema_version": 1,
            "table_updates": [
                {"table_id": summary_id, "set": {"award_completeness": {
                    "status": "complete", "basis_type": "structural", "basis": "测试汇总结果完整"}}},
                {"table_ids": [roster_id], "set": {"award_completeness": {
                    "status": "partial", "basis_type": "none", "basis": "测试明细不含结果"}}},
            ],
            "relationships": [{
                "id": "patched_relation", "kind": "award_to_bidder_roster",
                "award_tables": [summary_id], "bidder_tables": [roster_id],
                "project_keys": ["project_code", "project_name"],
            }],
        }
        patch_path = inspection_path.parent / "patch.json"
        plan_path = inspection_path.parent / "plan.json"
        patch_path.write_text(json.dumps(patch, ensure_ascii=False), encoding="utf-8")
        planned = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "plan",
                                  "--inspection", str(inspection_path), "--patch", str(patch_path),
                                  "--output", str(plan_path)], text=True, capture_output=True)
        self.assertEqual(planned.returncode, 0, planned.stderr)
        self.assertTrue(plan_path.is_file())
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/excel_ledger.py"), "run",
                                 str(book.path), "--plan", str(plan_path), "--output", str(output)],
                                text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(inspection_path.parent.exists())
        self.assertEqual({path.name for path in output.iterdir()}, {"final.csv", "review_queue.csv", "ledger.json"})


if __name__ == "__main__":
    unittest.main()
