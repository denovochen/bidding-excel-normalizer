"""跨表项目关系、标段降级与有界候选。 @author denovochen"""
from __future__ import annotations

import re
from copy import deepcopy
from collections import defaultdict
from difflib import SequenceMatcher
from typing import Any

from .contract import LedgerError, clean, name_key, stable_id

TABLE_KINDS = {"award_summary", "bidder_roster", "complete_results", "other"}
RELATION_KIND = "award_to_bidder_roster"
PROJECT_ROLES = ("project_code", "project_name")
_PROJECT_PUNCTUATION = re.compile(r"[\s()（）\[\]【】，,。、“”‘’：:；;·—_\-]+")
_YEAR = re.compile(r"(?:19|20)\d{2}")
_LOT_NUMBER = {
    "一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
    "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",
}
_EMPTY_LOTS = {"", "/", "-", "—", "主标段", "施工标段", "本标段", "全部标段"}


def infer_table_kind(columns: dict[str, str]) -> str:
    bidder, award = "bidder_name" in columns, "award_name" in columns
    if bidder and award:
        return "complete_results"
    if bidder:
        return "bidder_roster"
    if award:
        return "award_summary"
    return "other"


def table_identity(source_sha256: str, sheet_index: int, table_index: int) -> str:
    return stable_id("table", source_sha256, sheet_index, table_index)


def project_name_key(value: object) -> str:
    normalized = clean(value).replace("年度", "年")
    return _PROJECT_PUNCTUATION.sub("", normalized).casefold()


def project_similarity(left: object, right: object) -> float:
    a, b = project_name_key(left), project_name_key(right)
    return SequenceMatcher(None, a, b, autojunk=False).ratio() if a and b else 0.0


def project_year(value: object) -> str:
    found = _YEAR.search(clean(value))
    return found.group(0) if found else ""


def lot_key(value: object) -> str:
    raw = clean(value)
    if raw in _EMPTY_LOTS:
        return ""
    compact = re.sub(r"\s+", "", raw)
    match = re.fullmatch(r"第?([一二三四五六七八九十]|\d+)(?:标段|标包|包件|合同段|标)?", compact)
    if match:
        token = match.group(1)
        return _LOT_NUMBER[token] if token in _LOT_NUMBER else str(int(token))
    return project_name_key(compact)


def _issue(code: str, message: str, group_id: str | None = None, **extra: Any) -> dict[str, Any]:
    return {"code": code, "message": message, "group_id": group_id,
            "record_ids": [], "final_sequences": [], "review_sequences": [], **extra}


def _project_label(project: dict[str, Any]) -> str:
    values = project.get("values", {})
    return clean(values.get("project_name")) or clean(values.get("project_code")) or "（原表未提供项目名称）"


def _candidate_evidence(source: dict[str, Any], target: dict[str, Any], score: float) -> list[str]:
    evidence = []
    source_values, target_values = source.get("values", {}), target.get("values", {})
    if clean(source_values.get("project_code")) and name_key(source_values.get("project_code")) == name_key(
            target_values.get("project_code")):
        evidence.append("项目编号一致")
    source_year = project_year(source_values.get("project_year") or source_values.get("project_name"))
    target_year = project_year(target_values.get("project_year") or target_values.get("project_name"))
    if source_year and source_year == target_year:
        evidence.append("年度一致")
    evidence.append(f"项目名称相似度={score:.3f}")
    origin = target.get("cells", {}).get("project_name") or target.get("cells", {}).get("project_code")
    if origin:
        evidence.append(f"项目来源={target['sheet']}!{origin}")
    return evidence


def selection_blockers(source: dict[str, Any], target: dict[str, Any], peers: list[dict[str, Any]]) -> list[str]:
    """Only explicit identity conflicts block selection; missing fields are not conflicts."""
    left, right = source.get("values", {}), target.get("values", {})
    a, b = clean(left.get("project_name")), clean(right.get("project_name"))
    reasons = []
    for field, label, normalize in (("project_code", "项目编号", name_key),
                                     ("project_year", "年度", project_year)):
        av = normalize(left.get(field) or (a if field == "project_year" else ""))
        bv = normalize(right.get(field) or (b if field == "project_year" else ""))
        if av and bv and av != bv:
            reasons.append(label + "冲突")
    extension = re.compile(r"增做|追加工程|新增工程|节余资金|结余资金")
    if bool(extension.search(a)) != bool(extension.search(b)):
        reasons.append("主工程与追加/节余资金工程范围不能直接等同")
    locality = re.compile(r"([^省市县区年度\d（）()，,\s]{1,12}(?:镇|街道))")
    places = [list(locality.finditer(value)) for value in (a, b)]
    if all(len(matches) == 1 for matches in places):
        if places[0][0].group(0) != places[1][0].group(0):
            reasons.append("乡镇/街道冲突")
        else:
            areas = [re.match(r"([\u4e00-\u9fff]{1,10}?(?:村|片))", value[matches[0].end():])
                     for value, matches in zip((a, b), places)]
            if all(areas) and areas[0].group(0) != areas[1].group(0):
                reasons.append("村/片区冲突")
    funding = ("增发国债", "国家专项债", "财政补助")
    av, bv = {token for token in funding if token in a}, {token for token in funding if token in b}
    if av and bv and not av.intersection(bv):
        reasons.append("资金来源冲突")
    kinds = [{label for pattern, label in ((r"改造|提升", "改造"), (r"新建|新增建设", "新建"))
              if re.search(pattern, value)} for value in (a, b)]
    if all(kinds) and not kinds[0].intersection(kinds[1]):
        reasons.append("建设类型冲突")
    same_name = [peer for peer in peers if project_name_key(_project_label(peer)) == project_name_key(b)
                 and project_year(peer.get("values", {}).get("project_year") or _project_label(peer)) ==
                 project_year(right.get("project_year") or b)]
    if len(same_name) > 1:
        # A source row address distinguishes candidates, but does not identify which one the summary means.
        discriminators = [field for field in ("project_code", "project_owner") if clean(left.get(field))]
        unique = any(name_key(left[field]) == name_key(right.get(field)) and
                     sum(name_key(peer.get("values", {}).get(field)) == name_key(left[field])
                         for peer in same_name) == 1 for field in discriminators)
        if not unique:
            reasons.append("同名同年度候选有多个独立来源，缺少唯一项目编号或实施主体依据")
    return reasons


def _rank_candidates(source: dict[str, Any], targets: list[dict[str, Any]],
                     peers: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    source_values = source.get("values", {})
    source_year = project_year(source_values.get("project_year") or source_values.get("project_name"))
    result = []
    peers_by_name = defaultdict(list)
    for peer in peers if peers is not None else targets:
        peers_by_name[project_name_key(_project_label(peer))].append(peer)
    for target in targets:
        target_year = project_year(
            target.get("values", {}).get("project_year") or target.get("values", {}).get("project_name"))
        score = project_similarity(source_values.get("project_name"), target.get("values", {}).get("project_name"))
        result.append({
            "project_id": target["id"],
            "project_name": _project_label(target),
            "project_code": clean(target.get("values", {}).get("project_code")),
            "project_year": clean(target.get("values", {}).get("project_year")),
            "sheet": target["sheet"],
            "table_id": target["table_id"],
            "source_cell": f"{target['sheet']}!{target.get('cells', {}).get('project_name') or target.get('cells', {}).get('project_code', '')}",
            "name_similarity": round(score, 6),
            "year_match": source_year == target_year if source_year and target_year else None,
            "selection_blockers": selection_blockers(source, target, peers_by_name[project_name_key(_project_label(target))]),
            "evidence": _candidate_evidence(source, target, score),
        })
    return sorted(result, key=lambda item: (
        bool(item["selection_blockers"]),
        -int(item["year_match"] is True),
        -item["name_similarity"],
        item["project_id"],
    ))[:5]


def _exact_candidates(source: dict[str, Any], targets: list[dict[str, Any]], keys: list[str]) -> list[dict[str, Any]]:
    values = source.get("values", {})
    if "project_code" in keys and clean(values.get("project_code")):
        code = name_key(values["project_code"])
        matches = [target for target in targets
                   if name_key(target.get("values", {}).get("project_code")) == code]
        if len(matches) <= 1:
            return matches
        if "project_name" in keys and clean(values.get("project_name")):
            name = project_name_key(values["project_name"])
            named = [target for target in matches
                     if project_name_key(target.get("values", {}).get("project_name")) == name]
            if named:
                return named
        return matches
    if "project_name" in keys and clean(values.get("project_name")):
        name = project_name_key(values["project_name"])
        return [target for target in targets
                if project_name_key(target.get("values", {}).get("project_name")) == name]
    return []


def _relation_task_id(relation_id: str, source_project_id: str) -> str:
    return stable_id("relation_review", relation_id, source_project_id)


def _copy_award(award: dict[str, Any], source_group: dict[str, Any], target_group: dict[str, Any]) -> None:
    copied = dict(award)
    copied["relation_source"] = {
        "source_id": source_group["source_id"],
        "group_id": source_group["id"], "sheet": source_group["sheet"],
        "row": source_group["anchor_row"], "lot_name": source_group["lot_name"],
        "lot_code": source_group["lot_code"],
    }
    if not any(name_key(item["name"]) == name_key(copied["name"]) and
               item.get("relation_source", {}).get("group_id") == source_group["id"]
               for item in target_group["awards"]):
        target_group["awards"].append(copied)


def _award_target_groups(source_group: dict[str, Any], target_groups: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if len(target_groups) == 1:
        return target_groups
    source_lot = lot_key(source_group.get("lot_code") or source_group.get("lot_name"))
    if source_lot:
        matching = [group for group in target_groups
                    if lot_key(group.get("lot_code") or group.get("lot_name")) == source_lot]
        if len(matching) == 1:
            return matching
    awards = {name_key(award["name"]) for award in source_group["awards"] if name_key(award["name"])}
    containing = []
    for group in target_groups:
        names = {name_key(record["name"]) for record in group["records"] if name_key(record["name"])}
        if awards and awards <= names:
            containing.append(group)
    return containing if len(containing) == 1 else []


def _project_level_completeness(source_groups: list[dict[str, Any]]) -> dict[str, str]:
    declarations = [group["award_completeness_declared"] for group in source_groups]
    if declarations and all(item.get("status") == "complete" for item in declarations):
        return {
            "status": "complete",
            "basis_type": "structural",
            "basis": "跨表项目关系已确认；标段无法可靠拆分，按项目级合并完整中标结果",
        }
    if any(item.get("status") == "partial" for item in declarations):
        return {
            "status": "partial",
            "basis_type": "structural",
            "basis": "跨表项目关系已确认；标段无法可靠拆分，项目级中标结果仅部分完整",
        }
    return {
        "status": "unknown",
        "basis_type": "none",
        "basis": "跨表项目关系已确认；标段无法可靠拆分，中标结果完整性未知",
    }


def _project_level_group(relation_id: str, target_project: dict[str, Any], source_groups: list[dict[str, Any]],
                         target_groups: list[dict[str, Any]], issues: list[dict[str, Any]]) -> dict[str, Any]:
    """将无法可靠拆分的多标段名册降级为单一项目级组。"""
    template = min(target_groups, key=lambda group: (group["sheet"], group["anchor_row"], group["id"]))
    group = deepcopy(template)
    group_id = stable_id("group", target_project["id"], relation_id, "project_level_fallback")
    original_ids = {item["id"] for item in target_groups}
    group.update({
        "id": group_id,
        "anchor_row": min(item["anchor_row"] for item in target_groups),
        "lot_name": "",
        "lot_code": "",
        "records": [],
        "awards": [],
        "counts": [],
        "non_tender": False,
        "procurement_signals": list(dict.fromkeys(
            signal for item in target_groups for signal in item.get("procurement_signals", []))),
        "award_completeness_declared": _project_level_completeness(source_groups),
        "award_completeness": {},
        "source_rows": sorted({row for item in target_groups for row in item.get("source_rows", [])}),
        "scope_type": "business",
        "company_role": "bidder_name",
        "award_mode": "name_match",
        "group_mode": "project",
        "group_context": {},
        "group_context_cells": {},
        "source_block_ids": sorted({block_id for item in target_groups
                                    for block_id in item.get("source_block_ids", [item["id"]])}),
        "candidate_coverage": {
            "complete": all(item.get("candidate_coverage", {}).get("complete", True) for item in target_groups),
            "reasons": list(dict.fromkeys(reason for item in target_groups
                                           for reason in item.get("candidate_coverage", {}).get("reasons", []))),
        },
    })
    if len(group["source_block_ids"]) > 1:
        group["candidate_coverage"] = {
            "complete": False,
            "reasons": [*group["candidate_coverage"]["reasons"], "中标标段无法唯一分配到独立来源投标块"],
        }
    for target in target_groups:
        for record in target["records"]:
            copied = deepcopy(record)
            copied["group_id"] = group_id
            copied["project_id"] = target_project["id"]
            group["records"].append(copied)
        for award in target["awards"]:
            _copy_award(award, target, group)
    for source in source_groups:
        for award in source["awards"]:
            _copy_award(award, source, group)
    for issue in issues:
        if issue.get("group_id") in original_ids:
            issue["group_id"] = group_id
    return group


def _apply_project_relation(source_project: dict[str, Any], target_project: dict[str, Any],
                            source_groups: list[dict[str, Any]], target_groups: list[dict[str, Any]],
                            issues: list[dict[str, Any]], audit: dict[str, Any]
                            ) -> tuple[dict[str, Any] | None, dict[str, str | None]]:
    unresolved = []
    linked = []
    linked_sources_by_target: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for source_group in source_groups:
        targets = _award_target_groups(source_group, target_groups)
        if not targets:
            unresolved.append(source_group)
            continue
        target = targets[0]
        for award in source_group["awards"]:
            _copy_award(award, source_group, target)
        linked.append({"source_group_id": source_group["id"], "target_group_id": target["id"],
                       "award_count": len(source_group["awards"])})
        linked_sources_by_target[target["id"]].append(source_group)
        completeness = source_group["award_completeness_declared"]
        if completeness.get("status") == "complete":
            target["award_completeness_declared"] = {
                "status": "complete", "basis_type": "structural",
                "basis": "跨表中标汇总与投标名单项目关系已确认，且中标结果来源声明完整",
            }
    ambiguous_targets = []
    for target in target_groups:
        sources = linked_sources_by_target.get(target["id"], [])
        source_lots = {lot_key(group.get("lot_code") or group.get("lot_name")) for group in sources}
        if len(sources) > 1 and len(source_lots - {""}) > 1:
            ambiguous_targets.append(target)
    redirects: dict[str, str | None] = {
        source["id"]: next((item["target_group_id"] for item in linked
                            if item["source_group_id"] == source["id"]), None)
        for source in source_groups
    }
    fallback = None
    if (unresolved or ambiguous_targets) and target_groups:
        fallback = _project_level_group(
            audit["relation_id"], target_project, source_groups, target_groups, issues)
        for group in source_groups + target_groups:
            redirects[group["id"]] = fallback["id"]
        linked = [{
            "source_group_id": source["id"],
            "target_group_id": fallback["id"],
            "award_count": len(source["awards"]),
            "basis": "project_level_fallback",
        } for source in source_groups]
        issues.append(_issue(
            "LOT_SCOPE_UNRESOLVED",
            "项目关系已确认，但标段无法唯一对应；结果已降级为项目级，未声明具体标段归属",
            fallback["id"], source_project_id=source_project["id"],
            target_project_id=target_project["id"],
            source_group_ids=[item["id"] for item in source_groups],
            target_group_ids=[item["id"] for item in target_groups],
            standalone=False,
        ))
    elif unresolved or ambiguous_targets:
        affected = list({group["id"]: group for group in (target_groups if unresolved else ambiguous_targets)}.values())
        for group in affected:
            issues.append(_issue(
                "LOT_SCOPE_UNRESOLVED", "项目已关联，但中标汇总与投标明细的标段范围无法唯一对应",
                group["id"], source_project_id=source_project["id"], target_project_id=target_project["id"],
                source_group_ids=[item["id"] for item in unresolved] or
                                 [item["id"] for item in linked_sources_by_target.get(group["id"], [])],
                standalone=False,
            ))
        if not affected:
            issues.append(_issue(
                "LOT_SCOPE_UNRESOLVED", "项目已关联，但目标项目没有可用投标分组",
                None, standalone=True, source_id=source_project["source_id"], source_file="",
                sheet=source_project["sheet"], source_project_id=source_project["id"],
                target_project_id=target_project["id"],
            ))
    audit["group_links"] = linked
    audit["unresolved_source_group_ids"] = [group["id"] for group in unresolved]
    if fallback:
        audit["project_level_fallback_group_id"] = fallback["id"]
    return fallback, redirects


def apply_relationships(projects: list[dict[str, Any]], groups: list[dict[str, Any]],
                        issues: list[dict[str, Any]], plan: dict[str, Any],
                        resolutions: dict[str, dict[str, Any]] | None = None
                        ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
                                   list[dict[str, Any]], dict[str, str | None]]:
    """仅在 plan 显式声明关系时消费中标汇总组并关联到投标名单组。"""
    relationships = plan.get("relationships", [])
    if not relationships:
        return projects, groups, issues, [], {}
    resolutions = resolutions or {}
    projects_by_table: dict[str, list[dict[str, Any]]] = defaultdict(list)
    groups_by_project: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for project in projects:
        projects_by_table[project["table_id"]].append(project)
    for group in groups:
        if group.get("project_id"):
            groups_by_project[group["project_id"]].append(group)
    consumed_groups = set()
    group_redirects: dict[str, str | None] = {}
    audits = []
    for index, relation in enumerate(relationships):
        relation_id = relation.get("id") or stable_id(
            "relation", index, relation["kind"], relation["award_tables"], relation["bidder_tables"])
        source_projects = [project for table_id in relation["award_tables"]
                           for project in projects_by_table[table_id] if groups_by_project[project["id"]]]
        target_projects = [project for table_id in relation["bidder_tables"]
                           for project in projects_by_table[table_id] if groups_by_project[project["id"]]]
        keys = relation.get("project_keys", list(PROJECT_ROLES))
        for source in source_projects:
            consumed_groups.update(group["id"] for group in groups_by_project[source["id"]])
            task_id = _relation_task_id(relation_id, source["id"])
            exact = _exact_candidates(source, target_projects, keys)
            eligible_targets = [project for project in target_projects if all(
                group.get("candidate_coverage", {}).get("complete", True)
                for group in groups_by_project[project["id"]])]
            exact_ids = {project["id"] for project in exact}
            ranked = _rank_candidates(source, [project for project in eligible_targets
                                               if not exact_ids or project["id"] in exact_ids], target_projects)
            decision = resolutions.get(task_id)
            selected = exact[0] if len(exact) == 1 and not selection_blockers(source, exact[0], target_projects) else None
            basis = "exact_project_key" if selected else "unresolved"
            if not selected and decision and decision.get("decision") == "select_project":
                selected = next((project for project in eligible_targets
                                 if project["id"] == decision.get("project_id") and
                                 project["id"] in {item["project_id"] for item in ranked
                                                   if not item["selection_blockers"]}), None)
                basis = ("model_selection" if decision.get("actor") == "model" else "user_selection") if selected else "invalid_resolution"
            audit = {
                "id": stable_id("project_relation", relation_id, source["id"]),
                "relation_id": relation_id, "kind": RELATION_KIND,
                "source_project_id": source["id"], "source_project_name": _project_label(source),
                "status": "matched" if selected else "unresolved",
                "basis": basis, "target_project_id": selected["id"] if selected else None,
                "target_project_name": _project_label(selected) if selected else None,
                "review_task_id": task_id, "candidates": ranked,
                "resolution": decision if decision else None,
                "source_project": {
                    "id": source["id"], "source_id": source["source_id"], "sheet": source["sheet"],
                    "table_id": source["table_id"], "values": dict(source.get("values", {})),
                    "cells": dict(source.get("cells", {})),
                },
            }
            blocked_groups = ([group for group in groups_by_project[selected["id"]]
                               if not group.get("candidate_coverage", {}).get("complete", True)]
                              if selected else [])
            if selected and blocked_groups:
                audit["status"] = "blocked"
                audit["basis"] = "candidate_coverage_incomplete"
                audit["coverage_blocked_group_ids"] = [group["id"] for group in blocked_groups]
                for group in blocked_groups:
                    issues.append(_issue(
                        "PROJECT_RELATION_COVERAGE_INCOMPLETE",
                        "投标候选范围不完整，已停止跨表项目关系消费和中标企业推荐",
                        group["id"], source_project_id=source["id"], target_project_id=selected["id"],
                        coverage_reasons=group.get("candidate_coverage", {}).get("reasons", []),
                    ))
                group_redirects.update({group["id"]: None for group in groups_by_project[source["id"]]})
            elif selected:
                target_groups_for_project = groups_by_project[selected["id"]]
                fallback, redirects = _apply_project_relation(
                    source, selected, groups_by_project[source["id"]], target_groups_for_project, issues, audit)
                group_redirects.update(redirects)
                if fallback:
                    consumed_groups.update(group["id"] for group in target_groups_for_project)
                    groups.append(fallback)
                    groups_by_project[selected["id"]] = [fallback]
            else:
                code = "PROJECT_RELATION_MISSING" if not ranked else "PROJECT_RELATION_REVIEW_REQUIRED"
                message = ("中标汇总项目没有投标明细候选" if not ranked else
                           "中标汇总项目名称未唯一对应投标明细项目")
                issues.append(_issue(
                    code, message, None, standalone=True, source_id=source["source_id"], source_file="",
                    sheet=source["sheet"], source_project_id=source["id"],
                    source_project_name=_project_label(source), candidates=ranked,
                    review_task_id=task_id if ranked else None,
                ))
                group_redirects.update({group["id"]: None for group in groups_by_project[source["id"]]})
            audits.append(audit)
    for issue in issues:
        group_id = issue.get("group_id")
        if group_id in group_redirects and group_redirects[group_id]:
            issue["group_id"] = group_redirects[group_id]
    issues = [issue for issue in issues if issue.get("group_id") not in consumed_groups]
    remaining_groups = [group for group in groups if group["id"] not in consumed_groups]
    remaining_project_ids = {group["project_id"] for group in remaining_groups if group.get("project_id")}
    remaining_projects = [project for project in projects if project["id"] in remaining_project_ids]
    return remaining_projects, remaining_groups, issues, audits, group_redirects


def validate_relation_plan(relationships: Any, tables: dict[str, dict[str, Any]]) -> None:
    if relationships is None:
        return
    if not isinstance(relationships, list):
        raise LedgerError("relationships 必须为列表")
    ids = set()
    used_award_tables = set()
    for relation in relationships:
        allowed = {"id", "kind", "award_tables", "bidder_tables", "project_keys"}
        required = {"kind", "award_tables", "bidder_tables"}
        if not isinstance(relation, dict) or set(relation) - allowed or required - set(relation):
            raise LedgerError("跨表关系字段无效")
        if relation["kind"] != RELATION_KIND:
            raise LedgerError("不支持的跨表关系类型")
        relation_id = relation.get("id")
        if relation_id is not None and (not isinstance(relation_id, str) or not relation_id.strip() or relation_id in ids):
            raise LedgerError("跨表关系 id 必须非空且唯一")
        ids.add(relation_id)
        award_tables, bidder_tables = relation["award_tables"], relation["bidder_tables"]
        if (not isinstance(award_tables, list) or not award_tables or len(award_tables) != len(set(award_tables)) or
                not isinstance(bidder_tables, list) or not bidder_tables or len(bidder_tables) != len(set(bidder_tables)) or
                set(award_tables) & set(bidder_tables)):
            raise LedgerError("跨表关系的来源和目标表必须为不重复且不相交的非空列表")
        if any(table_id not in tables for table_id in award_tables + bidder_tables):
            raise LedgerError("跨表关系引用未知 table_id")
        if any(tables[table_id]["table_kind"] != "award_summary" for table_id in award_tables):
            raise LedgerError("award_tables 只能引用 award_summary")
        if used_award_tables & set(award_tables):
            raise LedgerError("同一 award_summary 不能重复用于多个跨表关系")
        used_award_tables.update(award_tables)
        if any(tables[table_id]["table_kind"] not in {"bidder_roster", "complete_results"}
               for table_id in bidder_tables):
            raise LedgerError("bidder_tables 必须引用 bidder_roster/complete_results")
        if any(not ({"project_name", "project_code"} & set(tables[table_id]["columns"]))
               for table_id in award_tables + bidder_tables):
            raise LedgerError("跨表关系两侧都必须映射 project_name 或 project_code")
        keys = relation.get("project_keys", list(PROJECT_ROLES))
        if (not isinstance(keys, list) or not keys or len(keys) != len(set(keys)) or
                any(key not in PROJECT_ROLES for key in keys)):
            raise LedgerError("project_keys 仅支持 project_code/project_name")
