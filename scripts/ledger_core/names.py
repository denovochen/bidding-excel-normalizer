"""本组双名称候选与任务内决策应用，不保存全局公司映射。 @author denovochen"""
from __future__ import annotations

from collections import Counter
from difflib import SequenceMatcher
from typing import Any

from .contract import clean, name_key, stable_id

DECISIONS = {"use_a", "use_b", "different", "uncertain"}
MAX_NAME_CHARS = 200
MAX_CANDIDATES = 3


def candidate_pairs(award: dict, records: list[dict]) -> list[dict]:
    candidates = []
    for record in records:
        if not record["name"]:
            continue
        similarity = SequenceMatcher(None, name_key(award["name"]), name_key(record["name"]), autojunk=False).ratio()
        values = record["occurrences"][0]["values"]
        same_person = bool(award["legal_person"] and name_key(award["legal_person"]) == name_key(values.get("bidder_legal_person")))
        same_amount = bool(award["price"] and clean(award["price"]) == clean(values.get("bidder_price")))
        if similarity >= 0.55 or same_person or same_amount:
            candidates.append({"record_id": record["id"], "name": record["name"], "name_similarity": round(similarity, 6),
                               "legal_person_text_equal": same_person, "amount_text_equal": same_amount,
                               "pair_id": stable_id("pair", name_key(award["name"]), name_key(record["name"]))})
    # 辅助字段仅用于脚本选候选，模型只接收两个名称；相同法人/金额不是同一实体证明。
    candidates.sort(key=lambda c: (-int(c["legal_person_text_equal"] and c["amount_text_equal"]),
                                   -int(c["legal_person_text_equal"]), -int(c["amount_text_equal"]),
                                   -c["name_similarity"], c["name"]))
    return candidates[:MAX_CANDIDATES]


def match_group(group: dict, decisions: dict[str, str], registry: dict[str, dict]) -> tuple[set[str], list[dict]]:
    records = {r["id"]: r for r in group["records"] if r["name"]}
    by_name = {name_key(r["name"]): r for r in records.values()}
    awards = {}
    for award in group["awards"]:
        awards.setdefault(name_key(award["name"]), award)
    matched = {by_name[k]["id"] for k in awards if k in by_name}
    cases = []
    for award_key, award in awards.items():
        if award_key in by_name:
            continue
        candidates = candidate_pairs(award, list(records.values()))
        for candidate in candidates:
            pair = {"id": candidate["pair_id"], "name_a": award["name"], "name_b": candidate["name"]}
            if max(len(pair["name_a"]), len(pair["name_b"])) <= MAX_NAME_CHARS:
                registry.setdefault(pair["id"], pair)
            else:
                candidate["not_sent_reason"] = "name_too_long"
        positive = [c for c in candidates if decisions.get(c["pair_id"]) in {"use_a", "use_b"}]
        others_certain = all(decisions.get(c["pair_id"]) == "different" for c in candidates if c not in positive)
        selected = positive[0] if len(positive) == 1 and others_certain else None
        cases.append({"award_key": award_key, "award": award, "candidates": candidates, "selected": selected})

    target_counts = Counter(c["selected"]["record_id"] for c in cases if c["selected"])
    audits, unresolved = [], []
    for case in cases:
        award, candidates, selected = case["award"], case["candidates"], case["selected"]
        status = "unresolved"
        if selected:
            record = records[selected["record_id"]]
            decision = decisions[selected["pair_id"]]
            canonical = award["name"] if decision == "use_a" else record["name"]
            collision = any(r["id"] != record["id"] and name_key(r["name"]) == name_key(canonical) for r in records.values())
            if target_counts[record["id"]] != 1 or record["id"] in matched or collision:
                status = "conflict"
            else:
                before = record["name"] if decision == "use_a" else award["name"]
                change = {"rule": "model_name_pair", "field": "company_name" if decision == "use_a" else "award_name",
                          "before": before, "after": canonical, "basis": "本次任务双名称比较决定，仅作用于本组",
                          "pair_id": selected["pair_id"], "decision": decision}
                record["name"] = canonical
                record["changes"].append(change)
                for original in group["awards"]:
                    if name_key(original["name"]) == case["award_key"]:
                        original["name"] = canonical
                        original["changes"].append(change)
                matched.add(record["id"])
                status = "applied"
        audits.append({"original_award": award["raw"], "award_cell": award["cell"], "status": status,
                       "candidates": [{**c, "decision": decisions.get(c["pair_id"], "pending")} for c in candidates],
                       "selected_pair_id": selected["pair_id"] if selected else None})
        if status != "applied":
            unresolved.append({"code": "MODEL_MATCH_CONFLICT" if status == "conflict" else "AWARD_NAME_MISMATCH",
                               "message": "名称对应决定发生碰撞，保留原值供复核" if status == "conflict" else "原中标名称未能唯一对应本组投标企业，保留复核",
                               "original_award": award["raw"], "award_cell": award["cell"], "candidates": candidates})
    group["name_resolutions"] = audits
    return matched, unresolved
