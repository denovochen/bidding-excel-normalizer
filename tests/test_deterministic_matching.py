"""结构分流、证据冲突和投标名称保留的回归验证。 @author denovochen"""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from ledger_core.artifacts import publish
from ledger_core.contract import LedgerError
from ledger_core.normalize import build_outputs
from ledger_core.workbook import inspect_workbooks, read_workbook


class DeterministicMatchingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def prepare(self, rows, merges=(), mode="auto"):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        for area in merges:
            sheet.merge_cells(area)
        path = self.root / "input.xlsx"
        workbook.save(path); workbook.close()
        book = read_workbook(path)
        plan = inspect_workbooks([book])["suggested_plan"]
        plan["sources"][0]["sheets"][0]["tables"][0]["award_mode"] = mode
        return book, plan

    def run_case(self, rows, merges=(), mode="auto"):
        book, plan = self.prepare(rows, merges, mode)
        return build_outputs([book], plan)

    def test_repeated_winner_on_every_row_selects_only_actual_bidder(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "甲公司", "乙公司"], ["项目甲", "乙公司", "乙公司"], ["项目甲", "丙公司", "乙公司"]])
        self.assertEqual([x["中标与否"] for x in f], ["否", "是", "否"])
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")
        self.assertFalse(r)

    def test_first_row_can_name_a_winner_on_another_row(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "甲公司", "乙公司"], ["项目甲", "乙公司", None], ["项目甲", "丙公司", None]])
        self.assertEqual([x["中标与否"] for x in f], ["否", "是", "否"])
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")
        self.assertFalse(r)

    def test_sparse_row_alignment_handles_changed_person_without_renaming(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                 ["项目甲", "海岳水利市政工程有限公司", 120, "张甲", "海岳建设有限公司", "李乙", 120],
                                 ["项目甲", "远洋建设有限公司", 121, "王丙", None]])
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertEqual(f[0]["公司名称"], "海岳水利市政工程有限公司")
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "row_aligned")
        self.assertEqual(l["groups"][0]["award_matches"][0]["basis"], "row_alignment")
        self.assertEqual(l["records"][0]["corrections"], [])
        self.assertFalse(r)

    def test_sparse_nonempty_cell_alone_is_not_a_row_relationship(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "海岳建设有限公司", "远洋水务有限公司"], ["项目甲", "华洲建筑有限公司", None]])
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")

    def test_explicitly_mapped_row_layout_keeps_left_name(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "海岳建设有限公司", "远洋水务有限公司"], ["项目甲", "华洲建筑有限公司", None]], mode="row_aligned")
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)
        self.assertEqual(f[0]["公司名称"], "海岳建设有限公司")

    def test_explicit_row_layout_cannot_override_cross_row_exact_match(self):
        f, r, _ = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "甲公司", "乙公司"], ["项目甲", "乙公司", None]], mode="row_aligned")
        self.assertEqual([x["中标与否"] for x in f], ["否", "是"])
        self.assertFalse(r)

    def test_multiple_names_in_one_cell_never_use_row_presence(self):
        f, r, l = self.run_case([["项目名称", "标段", "投标企业名单", "中标单位"],
                                 ["项目甲", "一标", "甲公司、乙公司、丙公司", "乙公司"]], mode="row_aligned")
        self.assertEqual([x["中标与否"] for x in f], ["否", "是", "否"])
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")
        self.assertFalse(r)

    def test_all_rows_can_legitimately_have_different_winners(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "甲公司", "甲公司"], ["项目甲", "乙公司", "乙公司"]])
        self.assertEqual([x["中标与否"] for x in f], ["是", "是"])
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")
        self.assertFalse(r)

    def test_multiple_winners_inside_single_merged_award_cell(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "中标单位"],
                                 ["项目甲", "甲公司", "乙公司、丙公司"], ["项目甲", "乙公司", None], ["项目甲", "丙公司", None]], ("C2:C4",))
        self.assertEqual([x["中标与否"] for x in f], ["否", "是", "是"])
        self.assertFalse(r)
        self.assertEqual(l["groups"][0]["award_matching"]["mode"], "group_match")

    def test_person_and_price_find_renamed_bidder_that_is_not_most_similar(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                 ["项目甲", "远洋建设有限公司", 100, "张甲", "海岳建设有限公司", "张甲", 100],
                                 ["项目甲", "海月建设有限公司", 101, "李乙", None]], mode="name_match")
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)
        self.assertEqual(l["groups"][0]["award_matches"][0]["basis"], "unique_person_amount")

    def test_exact_name_and_auxiliary_evidence_conflict_is_reviewed(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                 ["项目甲", "甲公司", 100, "张甲", "乙公司", "张甲", 100], ["项目甲", "乙公司", 101, "李乙", None]])
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)
        self.assertEqual(l["issues"][0]["code"], "AWARD_MATCH_CONFLICT")

    def test_equal_person_amount_for_two_bidders_does_not_pick_first(self):
        f, r, _ = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                ["项目甲", "甲公司", 100, "张甲", "旧名称", "张甲", 100], ["项目甲", "乙公司", 100, "张甲", None]], mode="name_match")
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)

    def test_short_legal_person_headers_preserve_conflict_detection(self):
        f, r, l = self.run_case([["项目名称", "投标单位", "投标报价", "投标法人", "中标单位", "中标金额", "中标法人"],
                                 ["项目甲", "海岳公司", 100, "张甲", "远洋公司", 100, "张甲"],
                                 ["项目甲", "远洋公司", 200, "李乙", None]])
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)
        self.assertEqual(l["issues"][0]["code"], "AWARD_MATCH_CONFLICT")

    def test_declared_currency_units_are_compared_numerically(self):
        f, r, l = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                 ["项目甲", "甲公司", "100000元", "张甲", "旧名称", "张甲", "10万元"],
                                 ["项目甲", "乙公司", "12万元", "李乙", None]], mode="name_match")
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)

    def test_placeholder_in_award_field_does_not_make_a_winner(self):
        f, r, _ = self.run_case([["项目名称", "投标单位名称", "中标单位"], ["项目甲", "甲公司", "未中标"], ["项目甲", "乙公司", None]], mode="row_aligned")
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertFalse(r)

    def test_placeholder_person_and_equal_price_do_not_prove_a_match(self):
        f, r, _ = self.run_case([["项目名称", "投标单位名称", "投标报价", "投标单位法人", "中标单位", "中标单位法人", "中标金额"],
                                 ["项目甲", "甲公司", 100, "/", "旧名称", "/", 100], ["项目甲", "乙公司", 101, "/", None]], mode="name_match")
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)

    def test_repeated_award_amount_formats_do_not_create_a_false_conflict(self):
        f, r, _ = self.run_case([["项目名称", "投标单位名称", "中标单位", "中标金额"],
                                 ["项目甲", "甲公司", "乙公司", "100.00元"], ["项目甲", "乙公司", "乙公司", "100元"]])
        self.assertEqual([x["中标与否"] for x in f], ["否", "是"])
        self.assertFalse(r)

    def test_unique_abbreviated_name_keeps_short_bidder_name(self):
        f, r, l = self.run_case([["项目名称", "投标企业名单", "中标单位"],
                                 ["项目甲", "海岳建设工程有限公司、北方远洋建筑工程有限公司", "北方海岳建设工程有限公司"]])
        self.assertEqual(f[0]["公司名称"], "海岳建设工程有限公司")
        self.assertEqual([x["中标与否"] for x in f], ["是", "否"])
        self.assertFalse(r)

    def test_ambiguous_abbreviations_keep_review(self):
        f, r, _ = self.run_case([["项目名称", "投标企业名单", "中标单位"],
                                 ["项目甲", "北方海岳建设工程有限公司、南方海岳建设工程有限公司", "海岳建设工程有限公司"]])
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertEqual(len(r), 2)

    def test_two_award_names_cannot_silently_claim_the_same_bidder(self):
        f, r, l = self.run_case([["项目名称", "投标企业名单", "中标单位"],
                                 ["项目甲", "海岳建设有限公司、远洋公司", "海岳建设有限公司、海岳建设工程有限公司"]])
        self.assertEqual([x["中标与否"] for x in f], ["", ""])
        self.assertTrue(r)
        self.assertTrue(all(m["status"] == "conflict" for m in l["groups"][0]["award_matches"]))

    def test_only_shared_generic_words_cannot_establish_a_match(self):
        f, r, _ = self.run_case([["项目名称", "投标企业名单", "中标单位"],
                                 ["项目甲", "北方海岳建设工程有限公司、北方远洋建设工程有限公司", "北方长风建设工程有限公司"]])
        self.assertTrue(r)
        self.assertEqual([x["中标与否"] for x in f], ["", ""])

    def test_output_validation_rejects_rewritten_bidder_or_cross_group_match(self):
        f, r, ledger = self.run_case([["项目名称", "投标企业名单", "中标单位"],
                                    ["项目甲", "海岳建设有限公司、远洋公司", "海岳建设工程有限公司"],
                                    ["项目乙", "长风公司", "长风公司"]])
        altered = copy.deepcopy(ledger)
        altered["records"][0]["company_name"] = "海岳建设工程有限公司"
        with self.assertRaisesRegex(LedgerError, "原企业字段"):
            publish(self.root / "bad-name", f, r, altered)
        altered = copy.deepcopy(ledger)
        altered["groups"][0]["award_matches"][0]["selected_record_id"] = ledger["records"][-1]["id"]
        with self.assertRaises(LedgerError):
            publish(self.root / "bad-group", f, r, altered)

    def test_cli_finishes_unresolved_names_in_one_run_without_decision_files(self):
        book, _ = self.prepare([["项目名称", "投标企业名单", "中标单位"], ["项目甲", "甲公司、乙公司", "丙公司"]])
        out = self.root / "output"
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py"), "run", str(book.path), "--output", str(out)], text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        reply = json.loads(result.stdout)
        self.assertEqual(reply["kind"], "result")
        self.assertEqual(reply["summary"]["review_record_count"], 2)
        self.assertEqual({p.name for p in out.iterdir()}, {"final.csv", "review_queue.csv", "ledger.json"})
        self.assertEqual(list(self.root.glob(".excel-ledger-work-*")), [])


if __name__ == "__main__":
    unittest.main()
