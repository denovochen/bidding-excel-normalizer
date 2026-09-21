"""名称判断闭环、范围隔离与失败恢复的行为验证。 @author denovochen"""
from __future__ import annotations

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
from ledger_core.artifacts import validate_outputs
from ledger_core.contract import LedgerError
from ledger_core.normalize import build_outputs
from ledger_core.workbook import inspect_workbooks, read_workbook
from ledger_core.workflow import MAX_BATCH_CHARS, MAX_PAIRS, resolve, start


class NameWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.events = []

    def tearDown(self):
        self.temporary.cleanup()

    def progress(self, step, status):
        self.events.append((step, status))

    def setup_book(self, rows):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "台账"
        for row in rows:
            sheet.append(row)
        path = self.root / "input.xlsx"
        book.save(path); book.close()
        b = read_workbook(path)
        plan = inspect_workbooks([b])["suggested_plan"]
        return path, b, plan

    def begin(self, count=1):
        rows = [["项目名称", "投标单位名称", "中标单位"]]
        rows.extend([[f"项目{i}", f"甲{i}建设有限公司", f"甲{i}建设工程有限公司"] for i in range(count)])
        path, book, plan = self.setup_book(rows)
        output = self.root / "out"
        handoff = start(output, [book], plan, build_outputs([book], plan), self.progress)
        return path, output, handoff

    def answer(self, handoff, decision="use_b"):
        path = self.root / (handoff["batch_id"] + ".json")
        payload = {"batch_id": handoff["batch_id"], "decisions": [{"id": p["id"], "decision": decision} for p in handoff["pairs"]]}
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return path

    def test_repeated_winner_cells_do_not_make_every_bidder_a_winner(self):
        _, book, plan = self.setup_book([["项目名称", "投标单位名称", "中标单位"],
                                        ["项目甲", "甲公司", "甲公司"], ["项目甲", "乙公司", "甲公司"], ["项目甲", "丙公司", "甲公司"]])
        final, review, ledger = build_outputs([book], plan)
        self.assertEqual([r["中标与否"] for r in final], ["是", "否", "否"])
        self.assertFalse(review)
        self.assertFalse(ledger["name_pairs"])

    def test_unresolved_handoff_contains_only_names_and_no_final_artifacts(self):
        _, output, handoff = self.begin()
        self.assertEqual(handoff["kind"], "name_review_required")
        self.assertFalse(output.exists())
        self.assertTrue(Path(handoff["state"]).is_file())
        self.assertTrue(all(set(p) == {"id", "name_a", "name_b"} for p in handoff["pairs"]))
        self.assertNotIn((2, "completed"), self.events)
        self.assertNotIn((3, "in_progress"), self.events)

    def test_batches_are_bounded_and_each_pair_is_asked_once(self):
        _, output, handoff = self.begin(14)
        seen = set()
        state = Path(handoff["state"])
        while handoff["kind"] == "name_review_required":
            self.assertLessEqual(len(handoff["pairs"]), MAX_PAIRS)
            self.assertLessEqual(len(json.dumps(handoff["pairs"], ensure_ascii=False)), MAX_BATCH_CHARS)
            identifiers = {p["id"] for p in handoff["pairs"]}
            self.assertFalse(seen & identifiers)
            seen |= identifiers
            handoff = resolve(state, self.answer(handoff), self.progress)
        self.assertEqual(len(seen), 14)
        self.assertEqual(handoff["kind"], "result")
        self.assertTrue(validate_outputs(output)["validated"])
        self.assertNotIn((4, "completed"), self.events)

    def test_uncertain_finishes_with_review_without_asking_user(self):
        _, output, handoff = self.begin()
        result = resolve(Path(handoff["state"]), self.answer(handoff, "uncertain"), self.progress)
        self.assertEqual(result["kind"], "result")
        self.assertEqual(result["summary"]["review_record_count"], 1)
        ledger = json.loads((output / "ledger.json").read_text())
        self.assertEqual(set(ledger["name_decisions"].values()), {"uncertain"})
        self.assertEqual(ledger["records"][0]["corrections"], [])

    def test_duplicate_decision_batch_is_idempotent(self):
        _, output, handoff = self.begin(8)
        state = Path(handoff["state"])
        answer = self.answer(handoff)
        second = resolve(state, answer, self.progress)
        replay = resolve(state, answer, self.progress)
        self.assertEqual(second, replay)
        final = resolve(state, self.answer(second), self.progress)
        replay_final = resolve(state, answer, self.progress)
        self.assertEqual(final, replay_final)
        self.assertTrue(validate_outputs(output)["validated"])

    def test_conflicting_replay_is_rejected(self):
        _, _, handoff = self.begin(8)
        state = Path(handoff["state"])
        resolve(state, self.answer(handoff), self.progress)
        with self.assertRaisesRegex(LedgerError, "不可改写"):
            resolve(state, self.answer(handoff, "different"), self.progress)

    def test_incomplete_unknown_and_invented_decisions_leave_state_unchanged(self):
        _, output, handoff = self.begin(2)
        state = Path(handoff["state"])
        before = state.read_bytes()
        path = self.answer(handoff)
        valid = json.loads(path.read_text())
        bad = [
            {"batch_id": valid["batch_id"], "decisions": valid["decisions"][:1]},
            {"batch_id": "other", "decisions": valid["decisions"]},
            {"batch_id": valid["batch_id"], "decisions": [{"id": "unknown", "decision": "use_a"}]},
            {"batch_id": valid["batch_id"], "decisions": [{"id": valid["decisions"][0]["id"], "decision": "invent_company"}]},
        ]
        for payload in bad:
            path.write_text(json.dumps(payload), encoding="utf-8")
            with self.assertRaises(LedgerError):
                resolve(state, path, self.progress)
            self.assertEqual(state.read_bytes(), before)
            self.assertFalse(output.exists())

    def test_source_change_rejects_old_decisions(self):
        path, output, handoff = self.begin()
        state = Path(handoff["state"])
        before = state.read_bytes()
        book = openpyxl.load_workbook(path)
        book.active["B2"] = "其他企业"
        book.save(path); book.close()
        with self.assertRaisesRegex(LedgerError, "已改变"):
            resolve(state, self.answer(handoff), self.progress)
        self.assertEqual(state.read_bytes(), before)
        self.assertFalse(output.exists())

    def test_resume_returns_same_batch_without_restarting(self):
        _, _, handoff = self.begin()
        recovered = resolve(Path(handoff["state"]), None, self.progress)
        self.assertEqual(recovered, handoff)

    def test_publish_state_gap_can_be_recovered(self):
        _, output, handoff = self.begin()
        state_path = Path(handoff["state"])
        first = resolve(state_path, self.answer(handoff), self.progress)
        state = json.loads(state_path.read_text())
        state["result"] = None  # 模拟最终发布已完成而状态尚未记录。
        state_path.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
        files_before = {p.name: p.read_bytes() for p in output.iterdir()}
        resumed = resolve(state_path, None, self.progress)
        self.assertEqual(first, resumed)
        self.assertEqual(files_before, {p.name: p.read_bytes() for p in output.iterdir()})

    def test_decisions_are_scoped_to_group_not_global_rename(self):
        _, book, plan = self.setup_book([["项目名称", "投标单位名称", "中标单位"],
                                        ["项目甲", "甲建设有限公司", "甲建设工程有限公司"],
                                        ["项目乙", "甲建设有限公司", None]])
        _, _, draft = build_outputs([book], plan)
        decisions = {p["id"]: "use_a" for p in draft["name_pairs"]}
        final, review, ledger = build_outputs([book], plan, name_decisions=decisions)
        self.assertEqual([r["公司名称"] for r in final], ["甲建设工程有限公司", "甲建设有限公司"])
        self.assertEqual([r["中标与否"] for r in final], ["是", ""])
        self.assertFalse(review)
        self.assertEqual(ledger["alias_rules"]["aliases"], [])

    def test_multiple_positive_candidates_are_not_merged(self):
        _, book, plan = self.setup_book([["项目名称", "投标企业名单", "中标单位"],
                                        ["项目甲", "甲建设有限公司、甲建设有限责任公司", "甲建设工程有限公司"]])
        _, _, draft = build_outputs([book], plan)
        final, review, ledger = build_outputs([book], plan, name_decisions={p["id"]: "use_a" for p in draft["name_pairs"]})
        self.assertEqual(len(final), 2)
        self.assertEqual(len(review), 2)
        self.assertEqual([r["公司名称"] for r in final], ["甲建设有限公司", "甲建设有限责任公司"])
        self.assertTrue(all(r["中标与否"] == "" for r in final))

    def test_single_pair_reused_in_different_groups_is_only_one_model_request(self):
        _, book, plan = self.setup_book([["项目名称", "投标单位名称", "中标单位"],
                                        ["项目甲", "甲建设有限公司", "甲建设工程有限公司"],
                                        ["项目乙", "甲建设有限公司", "甲建设工程有限公司"]])
        _, _, ledger = build_outputs([book], plan)
        self.assertEqual(len(ledger["name_pairs"]), 1)

    def test_cli_runs_complete_name_round_trip(self):
        path, _, _ = self.setup_book([["项目名称", "投标单位名称", "中标单位"], ["项目甲", "甲建设有限公司", "甲建设工程有限公司"]])
        cli = [sys.executable, "-B", str(ROOT / "scripts" / "excel_ledger.py")]
        output = self.root / "cli-out"
        first = subprocess.run(cli + ["run", str(path), "--output", str(output), "--progress"], text=True, capture_output=True)
        self.assertEqual(first.returncode, 4, first.stderr)
        messages = [json.loads(line) for line in first.stdout.splitlines()]
        handoff = messages[-1]
        self.assertEqual(messages[-2]["status"], "in_progress")
        second = subprocess.run(cli + ["resolve", "--state", handoff["state"], "--decisions", str(self.answer(handoff)), "--progress"], text=True, capture_output=True)
        self.assertEqual(second.returncode, 0, second.stderr)
        self.assertEqual(json.loads(second.stdout.splitlines()[-1])["kind"], "result")
        self.assertEqual({p.name for p in output.iterdir()}, {"final.csv", "review_queue.csv", "ledger.json"})


if __name__ == "__main__":
    unittest.main()
