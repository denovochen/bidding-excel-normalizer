"""按组内名称和结构生成确定性对应及人工复核候选。 @author denovochen"""
from __future__ import annotations

import re
from collections import Counter
from difflib import SequenceMatcher

from .contract import clean, name_key, stable_id

COMMON_WORDS = re.compile(r"有限责任|股份|有限|公司|集团|建设|建筑|工程|水利|市政|节水|器材")
POLICY = {
    "version": 4,
    "company_name": "bidder_preferred",
    "method": "deterministic_with_user_resolution",
    "automatic_match": "exact_name_only",
    "recommendation": "safe_same_row_then_unique_name_similarity",
    "core_weight": 0.65,
    "auxiliary_fields": "audit_only",
}


def similarities(a: str, b: str) -> tuple[float, float, float]:
    a, b = name_key(a), name_key(b)
    core_a, core_b = COMMON_WORDS.sub("", a), COMMON_WORDS.sub("", b)
    full = SequenceMatcher(None, a, b, autojunk=False).ratio()
    core = SequenceMatcher(None, core_a, core_b, autojunk=False).ratio() if core_a and core_b else 0.0
    return full, core, POLICY["core_weight"] * core + (1 - POLICY["core_weight"]) * full


def candidate_scores(award: dict, records: list[dict]) -> list[dict]:
    candidates = []
    for record in records:
        if not record["name"]:
            continue
        full, core, score = similarities(award["name"], record["name"])
        candidates.append({
            "record_id": record["id"],
            "name": record["name"],
            "name_similarity": round(full, 6),
            "core_similarity": round(core, 6),
            "score": round(score, 6),
            "same_row": any(occurrence["row"] == award["row"] for occurrence in record["occurrences"]),
        })
    return sorted(candidates, key=lambda candidate: (-candidate["score"], candidate["record_id"]))


def _row_recommendation(group: dict, cases: list[dict]) -> dict:
    fallback = {"safe": False, "reason": "按组内名称接近程度推荐"}
    if group["award_mode"] == "name_match" or group["company_role"] != "bidder_name" or not cases:
        return fallback
    rows = Counter(occurrence["row"] for record in group["records"] for occurrence in record["occurrences"])
    if any(count != 1 for count in rows.values()) or any(
        not occurrence["single_bidder_row"]
        for record in group["records"]
        for occurrence in record["occurrences"]
    ):
        return {"safe": False, "reason": "投标字段为名单或跨行合并，不能按同行推荐"}
    awards = group["awards"]
    award_rows = [award["row"] for award in awards]
    if (any(award["merged"] or not award["single_name"] for award in awards) or
            len(set(award_rows)) != len(award_rows) or
            len({name_key(award["name"]) for award in awards}) != len(awards)):
        return {"safe": False, "reason": "中标字段合并、含多家或重复展示，不能按同行推荐"}
    if set(award_rows) >= set(rows) and group["award_mode"] != "row_aligned":
        return {"safe": False, "reason": "全部投标行均有中标字段，按组内名称推荐"}
    for case in cases:
        same_row = [candidate for candidate in case["candidates"] if candidate["same_row"]]
        if len(same_row) != 1:
            return {"safe": False, "reason": "中标行未唯一对应一家投标企业"}
        if any(candidate["record_id"] != same_row[0]["record_id"] for candidate in case["exact"]):
            return {"safe": False, "reason": "中标名称精确对应组内其他行，不能按同行推荐"}
    return {"safe": True, "reason": "中标信息与推荐投标企业位于同一行"}


def _task_id(group: dict, award: dict, entries: list[dict]) -> str:
    return stable_id(
        "award_review",
        group["id"],
        name_key(award["name"]),
        sorted({entry["cell"] for entry in entries}),
    )


def match_group(group: dict, resolutions: dict[str, dict] | None = None,
                candidate_coverage_complete: bool = True) -> tuple[set[str], list[dict]]:
    resolutions = resolutions or {}
    awards = {}
    for award in group["awards"]:
        awards.setdefault(name_key(award["name"]), []).append(award)
    cases = []
    for entries in awards.values():
        award = entries[0]
        candidates = candidate_scores(award, group["records"]) if candidate_coverage_complete else []
        cases.append({
            "award": award,
            "entries": entries,
            "candidates": candidates,
            "exact": [candidate for candidate in candidates if name_key(candidate["name"]) == name_key(award["name"])],
            "review_task_id": _task_id(group, award, entries),
        })
    if not candidate_coverage_complete:
        group["award_matching"] = {
            "mode": "group_match",
            "reason": "投标候选范围不完整，停止名称对应和候选推荐",
            "row_recommendation": {"safe": False, "reason": "投标候选范围不完整"},
        }
        group["award_matches"] = [{
            "original_award": case["award"]["raw"],
            "award_cells": [entry["cell"] for entry in case["entries"]],
            "status": "blocked", "basis": "candidate_coverage_incomplete",
            "selected_record_id": None, "selected_bidder_name": None,
            "candidate_count": 0, "candidates": [], "review_task_id": None,
            "recommended_record_id": None, "recommended_bidder_name": None,
            "recommendation_basis": "blocked", "recommendation_reason": "投标候选范围不完整",
            "resolution": None,
        } for case in cases]
        return set(), []
    row_recommendation = _row_recommendation(group, cases)
    group["award_matching"] = {
        "mode": "group_match",
        "reason": "精确名称自动对应；非精确名称由用户确认",
        "row_recommendation": row_recommendation,
    }
    for case in cases:
        candidates, exact = case["candidates"], case["exact"]
        selected, basis, conflict = None, "unresolved", False
        if len(exact) == 1:
            selected, basis = exact[0], "exact_name"
        elif len(exact) > 1:
            conflict = True
        else:
            resolution = resolutions.get(case["review_task_id"])
            if resolution and resolution.get("decision") == "select_bidder":
                selected = next(
                    (candidate for candidate in candidates if candidate["record_id"] == resolution.get("record_id")),
                    None,
                )
                if selected:
                    basis = "user_selection"
                else:
                    conflict = True
        same_row = [candidate for candidate in candidates if candidate["same_row"]]
        recommended = same_row[0] if row_recommendation["safe"] and len(same_row) == 1 else (
            candidates[0] if candidates else None
        )
        recommendation_basis = "same_row" if recommended in same_row and row_recommendation["safe"] else "name_similarity"
        if recommendation_basis == "name_similarity" and len(candidates) > 1 and candidates[0]["score"] == candidates[1]["score"]:
            recommended = None
            recommendation_basis = "ambiguous_name"
        case.update(
            selected=selected,
            basis=basis,
            conflict=conflict,
            recommended=recommended,
            recommendation_basis=recommendation_basis,
        )
    targets = Counter(case["selected"]["record_id"] for case in cases if case["selected"])
    matched, unresolved, audits = set(), [], []
    for case in cases:
        award, selected = case["award"], case["selected"]
        conflict = case["conflict"] or bool(selected and targets[selected["record_id"]] > 1)
        applied = selected is not None and not conflict
        displayed = case["candidates"][:5]
        for candidate in (selected, case["recommended"]):
            if candidate and candidate not in displayed:
                displayed.append(candidate)
        resolution = resolutions.get(case["review_task_id"])
        audits.append({
            "original_award": award["raw"],
            "award_cells": [entry["cell"] for entry in case["entries"]],
            "status": "matched" if applied else "conflict" if conflict else "unresolved",
            "basis": case["basis"],
            "selected_record_id": selected["record_id"] if applied else None,
            "selected_bidder_name": selected["name"] if applied else None,
            "candidate_count": len(case["candidates"]),
            "candidates": displayed,
            "review_task_id": case["review_task_id"],
            "recommended_record_id": case["recommended"]["record_id"] if case["recommended"] else None,
            "recommended_bidder_name": case["recommended"]["name"] if case["recommended"] else None,
            "recommendation_basis": case["recommendation_basis"],
            "recommendation_reason": row_recommendation["reason"] if case["recommendation_basis"] == "same_row" else
                                     "候选名称评分并列，需额外依据，不设推荐项" if case["recommendation_basis"] == "ambiguous_name" else "组内名称最接近",
            "resolution": resolution if resolution and resolution.get("decision") == "select_bidder" else None,
        })
        if applied:
            matched.add(selected["record_id"])
        else:
            unresolved.append({
                "code": "AWARD_MATCH_CONFLICT" if conflict else "AWARD_NAME_MISMATCH",
                "message": "中标对应证据冲突，保留复核" if conflict else "中标名称需要用户确认对应的投标企业",
                "original_award": award["raw"],
                "award_cell": award["cell"],
                "candidates": displayed,
                "review_task_id": case["review_task_id"],
                "recommended_record_id": case["recommended"]["record_id"] if case["recommended"] else None,
            })
    group["award_matches"] = audits
    return matched, unresolved
