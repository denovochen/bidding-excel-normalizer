"""验证表格变体与产物业务不变量。 @author denovochen"""
from __future__ import annotations

import copy
import csv
import hashlib
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
from ledger_core.contract import FIELDS, LedgerError
from ledger_core.normalize import aliases_from_json, build_outputs, normalize_name, split_names
from ledger_core.workbook import inspect_workbooks, load_json, read_workbook, validate_plan

TIME = "2026-09-20T12:00:00+08:00"


class LedgerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.aliases = {"schema_version": 1, "aliases": []}

    def tearDown(self):
        self.temp.cleanup()

    def book(self, rows, merges=(), extra_sheet=None):
        w = openpyxl.Workbook()
        s = w.active
        s.title = "台账"
        for row in rows:
            s.append(row)
        for merge in merges:
            s.merge_cells(merge)
        if extra_sheet:
            other = w.create_sheet("补充")
            for row in extra_sheet:
                other.append(row)
        path = self.directory / "input.xlsx"
        w.save(path)
        w.close()
        b = read_workbook(path)
        plan = inspect_workbooks([b])["suggested_plan"]
        for source in plan["sources"]:
            for sheet in source["sheets"]:
                for table in sheet.get("tables", []):
                    table["award_list_complete"] = True
                    # 既有名称匹配测试显式使用该模式；逐行模式另用 inspect 的实际建议验证。
                    table["award_mode"] = "name_match"
        return b, plan

    def build(self, book, plan):
        return build_outputs([book], plan, self.aliases, TIME)

    def build_auto(self, book):
        return self.build(book, inspect_workbooks([book])["suggested_plan"])

    def test_group_match_rejects_global_alias_configuration(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"],
                          ["项目甲", "甲建设有限公司", "甲建设工程有限公司"]])
        self.aliases["aliases"] = [{"from": "甲建设有限公司", "to": "甲建设工程有限公司", "basis": "测试全局映射"}]
        with self.assertRaisesRegex(LedgerError, "不接受全局"):
            self.build(b, p)

    def test_row_presence_does_not_assume_first_row_or_only_one_winner(self):
        b, _ = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["项目甲", "甲公司", 3, None], [None, "乙公司", None, "乙公司"],
                          [None, "丙公司", None, "丙公司"]], ("A2:A4",))
        f, r, ledger = self.build_auto(b)
        self.assertEqual([x["中标与否"] for x in f], ["否", "是", "是"])
        self.assertFalse(r)

    def test_summary_is_not_an_extra_record_and_non_tender_keeps_stated_winner(self):
        b, _ = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位", "备注"],
                          ["项目甲", "项目汇总", 5, "甲公司、乙公司、丙公司"],
                          [None, "甲公司", 2, "甲公司"], [None, "丙公司", None, None],
                          [None, "乙公司", 2, "乙公司"], [None, "丁公司", None, None],
                          [None, "丙公司", 1, "丙公司", "未招投标"]], ("A2:A7",))
        f, r, ledger = self.build_auto(b)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "是", "否", "是"])
        self.assertEqual(len(f), 5)
        self.assertFalse(r)
        self.assertEqual(ledger["records"][-1]["participation_type"], "non_tender")
        self.assertEqual(ledger["summary"]["group_count"], 3)
        publish(self.directory / "out", f, r, ledger)

    def test_unmatched_singleton_does_not_win_just_because_award_cell_is_nonempty(self):
        b, _ = self.book([["投标单位名称", "中标单位"], ["石家石家庄甲公司", "石家庄甲公司"]])
        f, r, ledger = self.build_auto(b)
        self.assertEqual(f[0]["公司名称"], "石家石家庄甲公司")
        self.assertEqual(f[0]["中标与否"], "")
        self.assertTrue(r)
        self.assertTrue(ledger["name_pairs"])

    def test_row_presence_formula_is_unknown_not_a_win(self):
        b, _ = self.book([["投标单位名称", "中标单位"], ["甲公司", "=A2"], ["乙公司", None]])
        f, r, ledger = self.build_auto(b)
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 1)
        self.assertEqual(ledger["issues"][0]["code"], "FIELD_UNAVAILABLE")

    def test_merged_award_uses_group_matching_not_top_row_presence(self):
        b, _ = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["项目甲", "甲公司", 3, "丙公司"], [None, "乙公司"], [None, "丙公司"]], ("A2:A4", "D2:D4"))
        p = inspect_workbooks([b])["suggested_plan"]
        self.assertEqual(p["sources"][0]["sheets"][0]["tables"][0]["award_mode"], "name_match")
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["否", "否", "是"])
        self.assertFalse(r)
        p["sources"][0]["sheets"][0]["tables"][0]["award_mode"] = "row_presence"
        with self.assertRaisesRegex(LedgerError, "已停用"):
            self.build(b, p)

    def test_mixed_structure_beyond_first_twelve_rows_requests_mapping(self):
        rows = [["项目名称", "投标单位名称", "中标单位"]]
        rows.extend([["项目甲", "甲公司", "甲公司"] for _ in range(13)])
        rows.append(["项目甲", "甲公司、乙公司", "甲公司"])
        b, _ = self.book(rows)
        p = inspect_workbooks([b])["suggested_plan"]
        self.assertEqual(p["sources"][0]["sheets"][0]["action"], "needs_mapping")

    def test_mixed_sections_can_reuse_header_and_preserve_project_identity(self):
        b, p = self.book([["项目名称", "标段", "投标单位名称", "中标单位", "投标单位数量"],
                          ["项目甲", "一标", "甲公司", "甲公司", 1],
                          [None, "二标", "乙公司", "丙公司", 2], [None, "二标", "丙公司"],
                          [None, "三标", "丁公司、戊公司", "丁公司", 2],
                          [None, "四标", "己公司", None, 1]], ("A2:A6", "D3:D4"))
        spec = p["sources"][0]["sheets"][0]
        self.assertEqual(spec["action"], "needs_mapping")
        base = {"header_rows": [1], "columns": {"project_name": "A", "lot_name": "B", "bidder_name": "C", "award_name": "D", "bidder_count": "E"},
                "project_mode": "merged", "group_mode": "lot", "award_list_complete": True,
                "summary_markers": ["项目汇总"], "non_tender_markers": []}
        spec.update({"action": "parse", "tables": [
            {**base, "data_start_row": 2, "data_end_row": 2, "award_mode": "name_match", "bidder_separator": "single"},
            {**base, "data_start_row": 3, "data_end_row": 4, "award_mode": "name_match", "bidder_separator": "single"},
            {**base, "data_start_row": 5, "data_end_row": 5, "award_mode": "name_match", "bidder_separator": "delimited"},
            {**base, "data_start_row": 6, "data_end_row": 6, "award_mode": "name_match", "bidder_separator": "single"},
        ]})
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "是", "是", "否", ""])
        self.assertEqual(ledger["summary"]["project_count"], 1)
        self.assertEqual(ledger["summary"]["group_count"], 4)
        self.assertFalse(r)
        publish(self.directory / "out", f, r, ledger)

    def test_shared_header_does_not_allow_overlapping_data(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"], ["项目甲", "甲公司", "甲公司"], ["项目乙", "乙公司", "乙公司"]])
        spec = p["sources"][0]["sheets"][0]
        spec["tables"].append(copy.deepcopy(spec["tables"][0]))
        with self.assertRaisesRegex(LedgerError, "重叠"):
            self.build(b, p)

    def test_blank_duplicate_row_does_not_conflict_with_shared_group_award(self):
        b, _ = self.book([["项目名称", "投标单位名称", "中标单位"],
                          ["项目甲", "甲公司", "甲公司"], ["项目甲", "甲公司", None]])
        f, r, ledger = self.build_auto(b)
        self.assertEqual(len(f), 1)
        self.assertEqual(f[0]["中标与否"], "是")
        self.assertFalse(r)
        self.assertEqual(len(ledger["records"][0]["occurrences"]), 2)
        publish(self.directory / "out", f, r, ledger)

    def test_row_presence_with_direct_status_conflict_is_not_silently_overwritten(self):
        b, _ = self.book([["投标单位名称", "中标单位", "是否中标"], ["甲公司", "甲公司", "否"]])
        f, r, ledger = self.build_auto(b)
        self.assertEqual(f[0]["中标与否"], "")
        self.assertEqual(len(r), 1)
        self.assertIn("AWARD_STATUS_CONFLICT", [i["code"] for i in ledger["issues"]])

    def test_exact_winner_and_three_artifact_contract(self):
        b, p = self.book([["项目名称", "标段", "投标企业名单", "中标单位"], ["项目甲", "一标", "甲公司、乙公司", "甲公司"]])
        before = hashlib.sha256(b.path.read_bytes()).hexdigest()
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertEqual(r, [])
        out = self.directory / "output"
        publish(out, f, r, ledger)
        self.assertEqual({x.name for x in out.iterdir()}, {"final.csv", "review_queue.csv", "ledger.json"})
        for name in ("final.csv", "review_queue.csv"):
            self.assertTrue((out / name).read_bytes().startswith(b"\xef\xbb\xbf"))
            with (out / name).open(encoding="utf-8-sig", newline="") as stream:
                self.assertEqual(next(csv.reader(stream)), FIELDS)
        self.assertTrue(validate_outputs(out)["validated"])
        self.assertEqual(hashlib.sha256(b.path.read_bytes()).hexdigest(), before)

    def test_model_name_decision_is_applied_with_audit(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"],
                          ["项目甲", "甲建设有限公司、乙公司", "甲建设工程有限公司"]])
        _, _, draft = self.build(b, p)
        decisions = {x["id"]: "use_a" if x["name_b"] == "甲建设有限公司" else "different" for x in draft["name_pairs"]}
        f, r, ledger = build_outputs([b], p, generated_at=TIME, name_decisions=decisions)
        self.assertEqual(f[0]["公司名称"], "甲建设工程有限公司")
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)
        self.assertEqual(ledger["records"][0]["occurrences"][0]["raw_company"], "甲建设有限公司")
        self.assertEqual(ledger["records"][0]["corrections"][0]["rule"], "model_name_pair")
        publish(self.directory / "out", f, r, ledger)

    def test_model_award_decision_keeps_bidder_spelling_and_original_award(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"],
                          ["项目甲", "甲建筑工程有限公司、乙公司、丙公司", "甲建筑公司有限公司"]])
        _, _, draft = self.build(b, p)
        decisions = {x["id"]: "use_b" if x["name_b"] == "甲建筑工程有限公司" else "different" for x in draft["name_pairs"]}
        f, r, ledger = build_outputs([b], p, generated_at=TIME, name_decisions=decisions)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "否"])
        self.assertFalse(r)
        self.assertEqual(ledger["groups"][0]["awards"][0]["raw"], "甲建筑公司有限公司")
        self.assertEqual(ledger["groups"][0]["awards"][0]["name"], "甲建筑工程有限公司")
        self.assertEqual(ledger["summary"]["record_count"], 3)
        publish(self.directory / "out", f, r, ledger)

    def test_fuzzy_candidate_never_becomes_winner_or_rewrites_name(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"], ["甲项目", "甲建设有限公司、乙公司", "甲建设工程有限公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(f[0]["公司名称"], "甲建设有限公司")
        self.assertEqual(len(r), 2)
        self.assertEqual(len(ledger["issues"]), 1)
        self.assertGreater(ledger["issues"][0]["candidates"][0]["name_similarity"], .7)
        f, r, ledger = build_outputs([b], p, generated_at=TIME,
                                    name_decisions={x["id"]: "uncertain" for x in ledger["name_pairs"]})
        publish(self.directory / "out", f, r, ledger)

    def test_repeated_prefix_requires_a_model_decision(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"],
                          ["项目甲", "石家石家庄甲公司、乙公司", "石家庄甲公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(f[0]["公司名称"], "石家石家庄甲公司")
        self.assertEqual(f[0]["中标与否"], "")
        self.assertTrue(r)
        self.assertTrue(any(x["name_a"] == "石家庄甲公司" and x["name_b"] == "石家石家庄甲公司" for x in ledger["name_pairs"]))

    def test_repeated_prefix_collision_does_not_merge_two_bidders(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"],
                          ["甲项目", "石家石家庄甲公司、石家庄甲公司", "石家庄甲公司"]])
        f, _, _ = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(f[0]["公司名称"], "石家石家庄甲公司")
        self.assertEqual(f[0]["中标与否"], "否")
        self.assertEqual(f[0]["复核状态"], "通过")

    def test_trailing_annotation_cleanup(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"], ["甲项目", "甲公司投标、乙公司", "甲公司"]])
        f, r, _ = self.build(b, p)
        self.assertEqual(f[0]["公司名称"], "甲公司")
        self.assertFalse(r)

    def test_missing_award_and_rank_one_do_not_infer_award(self):
        b, p = self.book([["项目名称", "投标单位名称", "排名"], ["甲项目", "甲公司", 1]])
        f, r, _ = self.build(b, p)
        self.assertEqual(f[0]["投标排名"], "1")
        self.assertEqual(f[0]["中标与否"], "")
        self.assertEqual(len(r), 0)

    def test_incomplete_award_list_never_marks_losers(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"], ["甲项目", "甲公司、乙公司", "甲公司"]])
        p["sources"][0]["sheets"][0]["tables"][0]["award_list_complete"] = False
        f, _, _ = self.build(b, p)
        self.assertEqual([r["中标与否"] for r in f], ["是", ""])

    def test_non_tender_note_does_not_clear_stated_award(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位", "备注"], ["甲项目", "甲公司", "甲公司", "该项目未招投标"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(f[0]["中标与否"], "是")
        self.assertEqual(len(r), 0)
        self.assertEqual(ledger["records"][0]["participation_type"], "non_tender")

    def test_negated_non_tender_note_does_not_override_explicit_award(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位", "备注"],
                          ["甲项目", "甲公司、乙公司", "甲公司", "不存在未招投标情形，已完成公开招标"],
                          ["乙项目", "丙公司", "丙公司", "该项目未招投标"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "是"])
        self.assertEqual(len(r), 0)
        self.assertEqual(ledger["summary"]["non_tender_group_count"], 1)

    def test_uncertain_procurement_is_not_claimed_as_non_tender(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位", "备注"],
                          ["甲项目", "甲公司、乙公司", "甲公司", "是否未招投标，尚待核实"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertEqual(len(r), 0)
        self.assertFalse(ledger["groups"][0]["non_tender"])
        self.assertEqual(ledger["summary"]["uncertain_group_count"], 1)
        self.assertEqual(ledger["summary"]["bidding_group_count"], 0)
        publish(self.directory / "out", f, r, ledger)

    def test_dedup_only_within_group_preserves_every_occurrence(self):
        b, p = self.book([["项目名称", "标段", "投标企业名单", "中标单位"],
                          ["甲项目", "一标", "甲公司、甲公司、乙公司", "甲公司"],
                          ["甲项目", "二标", "甲公司、乙公司", "乙公司"]])
        f, _, ledger = self.build(b, p)
        self.assertEqual(len(f), 4)
        self.assertEqual(ledger["summary"]["duplicate_mentions_removed"], 1)
        self.assertEqual(len(ledger["records"][0]["occurrences"]), 2)

    def test_duplicate_conflicting_prices_are_preserved_in_review(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位", "投标报价"],
                          ["甲项目", "甲公司", "甲公司", 1], ["甲项目", "甲公司", "甲公司", 2]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 1)
        self.assertEqual(len(r), 1)
        self.assertIn("DUPLICATE_CONFLICT", [x["code"] for x in ledger["issues"]])
        self.assertEqual(len(ledger["records"][0]["occurrences"]), 2)

    def test_merged_same_name_projects_are_not_collapsed(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["同名项目", "甲公司", 2, "甲公司"], [None, "乙公司"],
                          ["同名项目", "甲公司", 2, "甲公司"], [None, "丙公司"]], ("A2:A3", "A4:A5"))
        f, _, ledger = self.build(b, p)
        self.assertEqual(ledger["summary"]["project_count"], 2)
        self.assertEqual(len(f), 4)
        self.assertNotEqual(ledger["records"][0]["project_id"], ledger["records"][2]["project_id"])

    def test_blocks_inherit_project_but_not_company(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["甲项目", "甲公司", 2, "甲公司"], [None, "乙公司"], [None, None]])
        p["sources"][0]["sheets"][0]["tables"][0]["project_mode"] = "blocks"
        f, _, ledger = self.build(b, p)
        self.assertEqual([r["项目名称"] for r in f], ["甲项目", "甲项目"])
        self.assertEqual(ledger["records"][1]["occurrences"][0]["cells"]["project_name"], "A2")

    def test_summary_and_agent_are_not_enterprises(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位", "代理公司"],
                          ["甲项目", "项目汇总", 2, "甲公司", "代理甲公司"],
                          [None, "甲公司", 2, "甲公司", "代理乙公司"], [None, "乙公司"]], ("A2:A4",))
        f, _, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(ledger["unique_companies"], ["甲公司", "乙公司"])
        self.assertEqual(ledger["summary"]["group_count"], 1)

    def test_count_mismatch_is_reviewed(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"], ["甲项目", "甲公司", 3, "甲公司"]])
        _, r, ledger = self.build(b, p)
        self.assertEqual(len(r), 1)
        self.assertIn("BIDDER_COUNT_MISMATCH", [x["code"] for x in ledger["issues"]])

    def test_award_only_at_group_first_row_applies_to_whole_group(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["同一项目", "甲公司", 2, "甲公司"], [None, "乙公司"],
                          [None, "丙公司", 2, "丙公司"], [None, "丁公司"]], ("A2:A5",))
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "是", "否"])
        self.assertFalse(r)
        self.assertEqual(ledger["summary"]["project_count"], 1)
        self.assertEqual(ledger["summary"]["group_count"], 2)
        self.assertIn("原中标字段=D2:甲公司", f[1]["证据文本"])
        self.assertIn("原中标字段=D4:丙公司", f[3]["证据文本"])

    def test_missing_award_in_next_group_does_not_inherit_previous_winner(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["同一项目", "甲公司", 2, "甲公司"], [None, "乙公司"],
                          [None, "甲公司", 2, None], [None, "丙公司"]], ("A2:A5",))
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否", "", ""])
        self.assertFalse(r)
        self.assertEqual(ledger["groups"][1]["awards"], [])

    def test_first_row_name_typo_not_blank_continuations_causes_review(self):
        b, p = self.book([["项目名称", "投标单位名称", "投标单位数量", "中标单位"],
                          ["项目甲", "甲建设工程有限公司", 3, "甲建设公司有限公司"],
                          [None, "乙公司"], [None, "丙公司"]], ("A2:A4",))
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(r), 3)
        self.assertEqual([i["code"] for i in ledger["issues"]], ["AWARD_NAME_MISMATCH"])
        self.assertTrue(all("原中标字段=D2:" in x["证据文本"] for x in f))

    def test_reordered_columns_shifted_headers_and_two_sheets(self):
        rows = [["标题"], ["中标企业", "工程名称", "投标人名称"], ["甲公司", "甲项目", "甲公司"]]
        b, p = self.book(rows, extra_sheet=[["项目名称", "投标单位名称", "中标单位"], ["乙项目", "乙公司", "乙公司"]])
        f, _, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual([r["项目名称"] for r in f], ["甲项目", "乙项目"])
        self.assertEqual(ledger["summary"]["project_count"], 2)

    def test_multilevel_header_and_parenthesized_delimiter(self):
        b, p = self.book([["项目基本情况", None, "招投标情况", None],
                          ["项目名称", "标段", "投标企业名单", "中标单位"],
                          ["甲项目", "一标", "甲(上海,中国)公司、乙公司", "甲(上海,中国)公司"]], ("A1:B1", "C1:D1"))
        p["sources"][0]["sheets"][0]["tables"][0]["header_rows"] = [1, 2]
        p["sources"][0]["sheets"][0]["ignored_rows"] = []
        f, _, _ = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(f[0]["公司名称"], "甲(上海,中国)公司")

    def test_line_wrapping_not_a_fake_enterprise(self):
        tokens = split_names("河北甲建设\n工程有限公司", "delimited")
        self.assertEqual(tokens, ["河北甲建设\n工程有限公司"])
        name, actions = normalize_name(tokens[0], {})
        self.assertEqual(name, "河北甲建设工程有限公司")
        self.assertEqual(actions[0]["before"], tokens[0])
        self.assertEqual(split_names("甲公司\n乙公司", "delimited"), ["甲公司", "乙公司"])

    def test_raw_cell_whitespace_and_split_normalization_are_audited(self):
        original = " 甲公司、 乙公司\n"
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"], ["甲项目", original, "甲公司"]])
        f, _, ledger = self.build(b, p)
        self.assertEqual([r["公司名称"] for r in f], ["甲公司", "乙公司"])
        for record in ledger["records"]:
            self.assertEqual(record["occurrences"][0]["raw_values"]["bidder_name"], original)
            self.assertTrue(record["corrections"])
        self.assertEqual(ledger["records"][1]["occurrences"][0]["raw_company"], " 乙公司\n")

    def test_wrong_group_mode_cannot_silently_merge_distinct_lots(self):
        b, p = self.book([["项目名称", "标段", "投标单位名称", "中标单位"],
                          ["甲项目", "一标", "甲公司", "甲公司"], ["甲项目", "二标", "乙公司", "乙公司"]])
        p["sources"][0]["sheets"][0]["tables"][0]["group_mode"] = "project"
        with self.assertRaisesRegex(LedgerError, "标段信息不一致"):
            self.build(b, p)

    def test_consortium_creates_standalone_review_not_fake_bidder(self):
        b, p = self.book([["项目名称", "标段", "投标企业名单", "中标单位"],
                          ["甲项目", "一标", "甲公司、乙公司", "甲公司"],
                          ["乙项目", "二标", "甲公司与乙公司联合体", "甲公司与乙公司联合体"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(len(r), 1)
        self.assertEqual(r[0]["公司名称"], "")
        publish(self.directory / "out", f, r, ledger)

    def test_formula_field_is_blank_and_reviewed_without_blocking(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"], ["甲项目", "=A1", "甲公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(f[0]["公司名称"], "")
        self.assertEqual(f[0]["项目名称"], "甲项目")
        self.assertEqual(len(r), 1)
        self.assertEqual(ledger["unique_companies"], [])
        self.assertIn("FIELD_UNAVAILABLE", [i["code"] for i in ledger["issues"]])
        self.assertEqual(ledger["records"][0]["occurrences"][0]["raw_values"]["bidder_name"], "=A1")
        publish(self.directory / "out", f, r, ledger)

    def test_roster_without_projects_exports_blank_business_fields(self):
        b, p = self.book([["序号", "投标单位名称"], [1, "甲公司"], [2, "甲 公司"], [3, "乙公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        for row in f:
            for key in ("项目名称", "项目编号", "标段名称", "标段编号", "中标与否", "投标排名"):
                self.assertEqual(row[key], "")
        self.assertEqual(r, [])
        self.assertEqual(ledger["projects"], [])
        self.assertEqual(ledger["summary"]["project_count"], 0)
        self.assertEqual(ledger["summary"]["bidding_group_count"], 0)
        self.assertEqual(ledger["summary"]["roster_group_count"], 1)
        self.assertTrue(all(rec["project_id"] is None for rec in ledger["records"]))
        self.assertEqual(ledger["unique_companies"], ["甲公司", "乙公司"])
        publish(self.directory / "out", f, r, ledger)

    def test_only_award_names_are_kept_as_explicit_winners(self):
        b, p = self.book([["中标单位"], ["甲公司"], ["乙公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["公司名称"] for x in f], ["甲公司", "乙公司"])
        self.assertEqual([x["中标与否"] for x in f], ["是", "是"])
        self.assertFalse(r)
        self.assertEqual(ledger["projects"], [])
        self.assertTrue(all(x["participation_type"] == "award_company" for x in ledger["records"]))
        publish(self.directory / "out", f, r, ledger)

    def test_project_only_table_keeps_fields_without_fake_company(self):
        b, p = self.book([["项目名称", "项目编号"], ["甲项目", "P001"], ["乙项目", "P002"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["项目名称"] for x in f], ["甲项目", "乙项目"])
        self.assertTrue(all(x["公司名称"] == "" for x in f))
        self.assertFalse(r)
        self.assertEqual(ledger["unique_companies"], [])
        publish(self.directory / "out", f, r, ledger)

    def test_explicit_award_status_is_preserved_without_award_name(self):
        b, p = self.book([["投标单位名称", "是否中标"], ["甲公司", "是"], ["乙公司", "否"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)
        publish(self.directory / "out", f, r, ledger)

    def test_missing_project_cell_does_not_block_known_company(self):
        b, p = self.book([["项目名称", "投标单位名称"], [None, "甲公司"], ["乙项目", "乙公司"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual([x["项目名称"] for x in f], ["", "乙项目"])
        self.assertEqual([x["公司名称"] for x in f], ["甲公司", "乙公司"])
        self.assertEqual(ledger["summary"]["project_count"], 1)
        self.assertFalse(r)
        publish(self.directory / "out", f, r, ledger)

    def test_project_row_with_empty_company_is_kept_without_fabrication(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位", "备注"],
                          ["甲项目", "甲公司", "甲公司", "已完成"], ["乙项目", None, None, "设计状态"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(f[1]["项目名称"], "乙项目")
        self.assertEqual(f[1]["公司名称"], "")
        self.assertEqual(f[1]["中标与否"], "")
        self.assertFalse(r)
        self.assertEqual(ledger["unique_companies"], ["甲公司"])
        self.assertEqual(ledger["summary"]["context_group_count"], 1)
        publish(self.directory / "out", f, r, ledger)

    def test_empty_company_in_existing_group_is_not_marked_a_loser(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"],
                          ["甲项目", "甲公司", "甲公司"], ["甲项目", None, None]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 2)
        self.assertEqual(f[1]["中标与否"], "")
        self.assertEqual(ledger["records"][1]["participation_type"], "context")
        publish(self.directory / "out", f, r, ledger)

    def test_conflicting_direct_status_is_not_resolved_by_first_row(self):
        b, p = self.book([["项目名称", "投标单位名称", "是否中标"],
                          ["甲项目", "甲公司", "是"], ["甲项目", "甲公司", "否"]])
        f, r, ledger = self.build(b, p)
        self.assertEqual(len(f), 1)
        self.assertEqual(f[0]["中标与否"], "")
        self.assertEqual(len(r), 1)
        publish(self.directory / "out", f, r, ledger)

    def test_multisource_collection_list_dedups_without_deleting_participation(self):
        b, p = self.book([["项目名称", "投标单位名称"], ["项目甲", "ＡＣＭＥ公司"], ["项目乙", "乙公司"]])
        second = self.directory / "roster.xlsx"
        w = openpyxl.Workbook()
        w.active.append(["投标单位名称"])
        w.active.append(["acme公司"])
        w.active.append(["丙公司"])
        w.save(second); w.close()
        books = [b, read_workbook(second)]
        plan = inspect_workbooks(books)["suggested_plan"]
        f, r, ledger = build_outputs(books, plan, self.aliases, TIME)
        self.assertEqual(len(f), 4)
        self.assertEqual(ledger["unique_companies"], ["ACME公司", "乙公司", "丙公司"])
        self.assertEqual(f[2]["项目名称"], "")
        self.assertFalse(r)
        publish(self.directory / "out", f, r, ledger)

    def test_run_cli_completes_roster_without_a_plan_or_question(self):
        b, _ = self.book([["序号", "投标单位名称"], [1, "甲公司"], [2, "甲公司"], [3, "乙公司"]])
        output = self.directory / "direct-output"
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "run",
                                 str(b.path), "--output", str(output)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["kind"], "result")
        names = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "companies", str(output)],
                                capture_output=True, text=True)
        self.assertEqual(names.returncode, 0, names.stderr)
        self.assertEqual(json.loads(names.stdout)["companies"], ["甲公司", "乙公司"])
        self.assertEqual(len(list(output.iterdir())), 3)

    def test_progress_reports_only_completed_script_stages_before_result(self):
        b, _ = self.book([["投标单位名称"], ["甲公司"]])
        out = self.directory / "progress-output"
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "run",
                                 str(b.path), "--output", str(out), "--progress"], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        messages = [json.loads(line) for line in result.stdout.splitlines()]
        events = messages[:-1]
        self.assertEqual([(e["step"], e["status"]) for e in events],
                         [(1, "in_progress"), (1, "completed"), (2, "in_progress"), (2, "completed"),
                          (3, "in_progress"), (3, "completed")])
        self.assertTrue(all(e["kind"] == "progress" for e in events))
        self.assertEqual(messages[-1]["kind"], "result")
        self.assertTrue((out / "ledger.json").is_file())
        self.assertFalse(any(e.get("step") == 4 for e in messages))

    def test_progress_failure_does_not_claim_success_or_delivery(self):
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "run",
                                 str(self.directory / "missing.xlsx"), "--progress"], capture_output=True, text=True)
        self.assertEqual(result.returncode, 2)
        events = [json.loads(line) for line in result.stdout.splitlines()]
        self.assertEqual([(e["step"], e["status"]) for e in events], [(1, "in_progress"), (1, "failed")])
        self.assertEqual(json.loads(result.stderr)["kind"], "error")

    def test_mapping_required_keeps_cleaning_open_and_no_artifacts(self):
        b, _ = self.book([["自定义组织字段", "说明"], ["甲公司", "示例"]])
        out = self.directory / "unknown-output"
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "run",
                                 str(b.path), "--output", str(out), "--progress"], capture_output=True, text=True)
        self.assertEqual(result.returncode, 3)
        messages = [json.loads(line) for line in result.stdout.splitlines()]
        self.assertEqual(messages[-1]["kind"], "mapping_required")
        self.assertEqual([(e["step"], e["status"]) for e in messages[:-1]],
                         [(1, "in_progress"), (1, "completed"), (2, "in_progress")])
        self.assertFalse(out.exists())

    def test_official_code_zero_format_is_preserved(self):
        b, p = self.book([["项目名称", "项目编号", "投标单位名称", "中标单位"], ["甲项目", 42, "甲公司", "甲公司"]])
        w = openpyxl.load_workbook(b.path)
        w.active["B2"].number_format = "00000"
        w.save(b.path); w.close()
        b = read_workbook(b.path)
        p["sources"][0]["sha256"] = b.sha256
        f, _, _ = self.build(b, p)
        self.assertEqual(f[0]["项目编号"], "00042")

    def test_source_hash_drift_rejected(self):
        b, p = self.book([["项目名称", "投标单位名称"], ["甲项目", "甲公司"]])
        p["sources"][0]["sha256"] = "0" * 64
        with self.assertRaisesRegex(LedgerError, "哈希"):
            self.build(b, p)

    def test_unmapped_sheet_and_uncovered_rows_rejected(self):
        b, p = self.book([["项目名称", "投标单位名称"], ["甲项目", "甲公司"], ["乙项目", "乙公司"]])
        bad = copy.deepcopy(p); bad["sources"][0]["sheets"] = []
        with self.assertRaises(LedgerError):
            self.build(b, bad)
        p["sources"][0]["sheets"][0]["tables"][0]["data_end_row"] = 2
        with self.assertRaisesRegex(LedgerError, "未解释"):
            self.build(b, p)

    def test_unknown_mapping_code_is_rejected(self):
        b, p = self.book([["项目名称", "投标单位名称"], ["甲项目", "甲公司"]])
        p["sources"][0]["sheets"][0]["tables"][0]["python"] = "print(123)"
        with self.assertRaises(LedgerError):
            self.build(b, p)

    def test_multiple_tables_and_repeated_header_require_explicit_scope(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"], ["甲项目", "甲公司", "甲公司"],
                          ["项目名称", "投标单位名称", "中标单位"], ["乙项目", "乙公司", "乙公司"]])
        spec = p["sources"][0]["sheets"][0]
        self.assertEqual(spec["action"], "needs_mapping")
        spec["action"] = "parse"
        template = {"columns": {"project_name": "A", "bidder_name": "B", "award_name": "C"},
                    "project_mode": "repeated", "group_mode": "project", "bidder_separator": "single",
                    "award_list_complete": True, "summary_markers": ["合计"], "non_tender_markers": ["未招投标"]}
        first = {**template, "header_rows": [1], "data_start_row": 2, "data_end_row": 4}
        spec["tables"] = [first]; spec["ignored_rows"] = []
        with self.assertRaisesRegex(LedgerError, "表头"):
            self.build(b, p)
        first["data_end_row"] = 2
        second = {**template, "header_rows": [3], "data_start_row": 4, "data_end_row": 4}
        spec["tables"] = [first, second]
        f, _, _ = self.build(b, p)
        self.assertEqual(len(f), 2)

    def test_ids_and_rows_repeat_deterministically(self):
        b, p = self.book([["项目名称", "投标企业名单", "中标单位"], ["甲项目", "甲公司、乙公司", "甲公司"]])
        a = self.build(b, p)
        other = build_outputs([b], p, self.aliases, "2026-09-21T12:00:00+08:00")
        self.assertEqual([r["id"] for r in a[2]["records"]], [r["id"] for r in other[2]["records"]])
        self.assertEqual(a[2]["groups"], other[2]["groups"])

    def test_tampering_and_existing_output_are_rejected(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"], ["甲项目", "甲公司", "甲公司"]])
        f, r, ledger = self.build(b, p)
        out = self.directory / "out"
        publish(out, f, r, ledger)
        with self.assertRaisesRegex(LedgerError, "覆盖"):
            publish(out, f, r, ledger)
        with (out / "final.csv").open("ab") as stream:
            stream.write(b"tampered")
        with self.assertRaisesRegex(LedgerError, "哈希"):
            validate_outputs(out)

    def test_csv_dangerous_literal_is_escaped_and_audited(self):
        b, p = self.book([["项目名称", "投标单位名称", "中标单位"], ["@项目", "+甲公司", "+甲公司"]])
        f, r, ledger = self.build(b, p)
        out = self.directory / "out"
        publish(out, f, r, ledger)
        self.assertIn("'+甲公司", (out / "final.csv").read_text(encoding="utf-8-sig"))
        self.assertTrue(validate_outputs(out)["validated"])

    def test_malformed_and_fake_extension_are_rejected(self):
        path = self.directory / "fake.xls"
        path.write_bytes(b"not an excel workbook")
        with self.assertRaises(LedgerError):
            read_workbook(path)

    def test_alias_cycle_and_duplicate_rule_are_rejected(self):
        payload = {"schema_version": 1, "aliases": [{"from": "甲公司", "to": "乙公司", "basis": "用户确认"},
                                                    {"from": "乙公司", "to": "甲公司", "basis": "用户确认"}]}
        with self.assertRaises(LedgerError):
            aliases_from_json(payload)

    def test_cli_error_has_nonzero_exit_and_no_success_files(self):
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "export",
                                 str(self.directory / "missing.xlsx"), "--plan", str(self.directory / "missing.json"),
                                 "--output", str(self.directory / "out")], capture_output=True, text=True)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(json.loads(result.stderr)["kind"], "error")
        self.assertFalse((self.directory / "out").exists())


if __name__ == "__main__":
    unittest.main()
