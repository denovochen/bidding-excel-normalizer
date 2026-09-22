"""按结构和组内证据对应投标记录；不纠正企业全称。 @author denovochen"""
from __future__ import annotations

import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher

from .contract import clean, name_key

# 通用组织/行业词仅在比较时降权，绝不据此改写输出名称或删除参与记录。
COMMON_WORDS = re.compile(r"有限责任|股份|有限|公司|集团|建设|建筑|工程|水利|市政|节水|器材")
POLICY = {"version": 2, "company_name": "bidder_preferred", "method": "deterministic",
          "min_full_similarity": 0.80, "min_core_similarity": 0.75,
          "min_score": 0.85, "min_margin": 0.10, "core_weight": 0.65,
          "min_abbreviation_full": 0.90, "min_abbreviation_margin": 0.15,
          "name_similarity_usage": "candidate_order_only",
          "single_sided_amount_unit": "review_only"}


def person_key(value: str) -> str:
    key = name_key(value)
    return "" if key in {"", "/", "-", "—", "无", "暂无", "不详", "未知", "未提供", "待定"} else key


def similarities(a: str, b: str) -> tuple[float, float, float]:
    a, b = name_key(a), name_key(b)
    core_a, core_b = COMMON_WORDS.sub("", a), COMMON_WORDS.sub("", b)
    full = SequenceMatcher(None, a, b, autojunk=False).ratio()
    core = SequenceMatcher(None, core_a, core_b, autojunk=False).ratio() if core_a and core_b else 0.0
    return full, core, POLICY["core_weight"] * core + (1 - POLICY["core_weight"]) * full


def amount(value: str, unit: str = "") -> Decimal | None:
    raw = clean(value).replace(",", "")
    match = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)\s*(亿元|万元|元)?", raw)
    if not match:
        return None
    try:
        result = Decimal(match[1]) * {"": 1, "元": 1, "万元": 10000, "亿元": 100000000}[match[2] or unit]
        return result if result.is_finite() and result > 0 else None
    except (InvalidOperation, KeyError):
        return None


def _declared_unit(value: str, header_unit: str) -> str:
    raw = clean(value).replace(",", "")
    match = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)\s*(亿元|万元|元)?", raw)
    return (match[2] if match else "") or header_unit


def candidate_scores(award: dict, records: list[dict], units: dict) -> list[dict]:
    bidder_header_unit = units.get("bidder_price", "")
    award_header_unit = units.get("award_price", "")
    award_declared_unit = _declared_unit(award["price"], award_header_unit)
    candidates = []
    for record in records:
        if not record["name"]:
            continue
        full, core, score = similarities(award["name"], record["name"])
        stems = sorted((COMMON_WORDS.sub("", name_key(award["name"])), COMMON_WORDS.sub("", name_key(record["name"]))), key=len)
        prefix_omitted = 2 <= len(stems[0]) < len(stems[1]) and stems[1].endswith(stems[0])
        evidence = []
        for occurrence in record["occurrences"]:
            values = occurrence["values"]
            person = bool(person_key(award["legal_person"]) and
                          person_key(award["legal_person"]) == person_key(values.get("bidder_legal_person", "")))
            bidder_value = values.get("bidder_price", "")
            bidder_declared_unit = _declared_unit(bidder_value, bidder_header_unit)
            award_amount = amount(award["price"], award_declared_unit or bidder_declared_unit)
            bidder_amount = amount(bidder_value, bidder_declared_unit or award_declared_unit)
            price = award_amount is not None and award_amount == bidder_amount
            unit_assumed = price and bool(award_declared_unit) != bool(bidder_declared_unit)
            evidence.append((person, price, unit_assumed))
        candidates.append({"record_id": record["id"], "name": record["name"],
                           "name_similarity": round(full, 6), "core_similarity": round(core, 6),
                           "score": round(score, 6), "same_row": any(o["row"] == award["row"] for o in record["occurrences"]),
                           "prefix_omitted": prefix_omitted,
                           "legal_person_equal": any(e[0] for e in evidence) if award["single_name"] else False,
                           "amount_equal": any(e[1] for e in evidence) if award["single_name"] else False,
                           "amount_unit_assumed": any(e[1] and e[2] for e in evidence) if award["single_name"] else False,
                           "person_amount_equal": any(e[0] and e[1] for e in evidence) if award["single_name"] else False,
                           "person_amount_confirming": any(e[0] and e[1] and not e[2] for e in evidence)
                           if award["single_name"] else False})
    return sorted(candidates, key=lambda c: (-c["score"], c["record_id"]))


def _row_mode(group: dict, cases: list[dict]) -> dict:
    fallback = {"mode": "group_match", "reason": "按组内中标名单对应"}
    if group["award_mode"] == "name_match" or group["company_role"] != "bidder_name" or not cases:
        return fallback
    rows = Counter(o["row"] for r in group["records"] for o in r["occurrences"])
    if any(n != 1 for n in rows.values()) or any(not o["single_bidder_row"] for r in group["records"] for o in r["occurrences"]):
        return {"mode": "group_match", "reason": "投标单元格为名单或跨行合并"}
    awards = group["awards"]
    award_rows = [a["row"] for a in awards]
    if (any(a["merged"] or not a["single_name"] for a in awards) or len(set(award_rows)) != len(award_rows) or
            len({name_key(a["name"]) for a in awards}) != len(awards)):
        return {"mode": "group_match", "reason": "中标字段合并、列多家或重复展示"}
    if set(award_rows) >= set(rows):
        return {"mode": "group_match", "reason": "全部投标行均有中标字段，逐名匹配"}
    for case in cases:
        same = [c for c in case["candidates"] if c["same_row"]]
        if len(same) != 1 or any(c["record_id"] != same[0]["record_id"] for c in case["exact"]):
            return {"mode": "group_match", "reason": "中标名称对应其他行或同行不是单家企业"}
        own = same[0]
        if any(c["record_id"] != own["record_id"] for c in case["joint"]):
            return {"mode": "group_match", "reason": "法人和金额证据指向其他行"}
        supported = bool(case["exact"]) or own["person_amount_confirming"]
        if not supported and group["award_mode"] != "row_aligned":
            return {"mode": "group_match", "reason": "同行关系缺少独立支持，不能仅凭非空推断"}
    return {"mode": "row_aligned", "reason": "单企业行与单家中标信息对应，且无跨行匹配冲突",
            "declared": group["award_mode"] == "row_aligned"}


def match_group(group: dict) -> tuple[set[str], list[dict]]:
    awards = {}
    for award in group["awards"]:
        awards.setdefault(name_key(award["name"]), []).append(award)
    cases = []
    for entries in awards.values():
        award = entries[0]
        candidates = candidate_scores(award, group["records"], group["price_units"])
        cases.append({"award": award, "entries": entries, "candidates": candidates,
                      "exact": [c for c in candidates if name_key(c["name"]) == name_key(award["name"])],
                      "joint": [c for c in candidates if c["person_amount_confirming"]]})
    group["award_matching"] = _row_mode(group, cases)
    for case in cases:
        candidates, exact, joint = case["candidates"], case["exact"], case["joint"]
        selected, basis, conflict = None, "unresolved", False
        # 同名中标信息重复填写时，不能忽略不同的非空辅助信息。
        people = {person_key(a["legal_person"]) for a in case["entries"]} - {""}
        unit = group["price_units"].get("award_price") or group["price_units"].get("bidder_price", "")
        prices = {amount(a["price"], unit) for a in case["entries"]} - {None}
        conflict = len(people) > 1 or len(prices) > 1
        if exact:
            if len(exact) == 1:
                selected, basis = exact[0], "exact_name"
                conflict |= any(c["record_id"] != selected["record_id"] for c in joint)
            else:
                conflict = True
        elif group["award_matching"]["mode"] == "row_aligned":
            selected = next(c for c in candidates if c["same_row"])
            basis = "row_alignment"
        elif joint:
            if len(joint) == 1:
                selected, basis = joint[0], "unique_person_amount"
            else:
                conflict = True
        # 名称相似度和简称只用于候选排序；没有结构或独立证据时保留复核。
        plausible = None
        if not selected and candidates:
            best = candidates[0]
            margin = best["score"] - (candidates[1]["score"] if len(candidates) > 1 else 0)
            if ((best["name_similarity"] >= POLICY["min_full_similarity"] and
                 best["core_similarity"] >= POLICY["min_core_similarity"] and
                 best["score"] >= POLICY["min_score"] and margin >= POLICY["min_margin"]) or
                    (best["prefix_omitted"] and best["name_similarity"] >= POLICY["min_abbreviation_full"] and
                     margin >= POLICY["min_abbreviation_margin"])):
                plausible = best["record_id"]
        case.update(selected=selected, basis=basis, conflict=conflict, plausible_record_id=plausible)
    targets = Counter(c["selected"]["record_id"] for c in cases if c["selected"])
    contested = {case["plausible_record_id"] for case in cases
                 if case["plausible_record_id"] and case["plausible_record_id"] in targets}
    matched, unresolved, audits = set(), [], []
    for case in cases:
        award, selected = case["award"], case["selected"]
        conflict = (case["conflict"] or bool(selected and targets[selected["record_id"]] > 1) or
                    bool(selected and selected["record_id"] in contested) or
                    bool(not selected and case["plausible_record_id"] in contested))
        applied = selected is not None and not conflict
        displayed = case["candidates"][:3]
        if selected and selected not in displayed:
            displayed = displayed + [selected]
        audits.append({"original_award": award["raw"], "award_cells": [a["cell"] for a in case["entries"]],
                       "status": "matched" if applied else "conflict" if conflict else "unresolved",
                       "basis": case["basis"], "selected_record_id": selected["record_id"] if applied else None,
                       "selected_bidder_name": selected["name"] if applied else None,
                       "candidate_count": len(case["candidates"]), "candidates": displayed})
        if applied:
            matched.add(selected["record_id"])
        else:
            unresolved.append({"code": "AWARD_MATCH_CONFLICT" if conflict else "AWARD_NAME_MISMATCH",
                               "message": "中标对应证据冲突，保留复核" if conflict else "中标信息未能唯一对应本组投标企业，保留复核",
                               "original_award": award["raw"], "award_cell": award["cell"], "candidates": displayed})
    group["award_matches"] = audits
    return matched, unresolved
