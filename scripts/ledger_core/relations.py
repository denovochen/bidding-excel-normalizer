"""跨表项目关系、标段降级与有界候选。 @author denovochen"""
from __future__ import annotations

import re
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
    return evidence


def _rank_candidates(source: dict[str, Any], targets: list[dict[str, Any]]) -> list[dict[str, Any]]:
    source_values = source.get("values", {})
    source_year = project_year(source_values.get("project_year") or source_values.get("project_name"))
    narrowed = [target for target in targets if not source_year or project_year(
        target.get("values", {}).get("project_year") or target.get("values", {}).get("project_name")) == source_year]
    candidates = narrowed or targets
    result = []
    for target in candidates:
        score = project_similarity(source_values.get("project_name"), target.get("values", {}).get("project_name"))
        result.append({
            "project_id": target["id"],
            "project_name": _project_label(target),
            "project_code": clean(target.get("values", {}).get("project_code")),
            "project_year": clean(target.get("values", {}).get("project_year")),
            "sheet": target["sheet"],
            "table_id": target["table_id"],
            "name_similarity": round(score, 6),
            "evidence": _candidate_evidence(source, target, score),
        })
    return sorted(result, key=lambda item: (-item["name_similarity"], item["project_id"]))[:5]


def _exact_candidates(source: dict[str, Any], targets: list[dict[str, Any]], keys: list[str]) -> list[dict[str, Any]]:
    values = source.get("values", {})
    for role in keys:
        value = clean(values.get(role))
        if not value:
            continue
        key = name_key(value) if role == "project_code" else project_name_key(value)
        matches = [target for target in targets if (
            name_key(target.get("values", {}).get(role)) if role == "project_code" else
            project_name_key(target.get("values", {}).get(role))) == key]
        if matches:
            return matches
    return []


def _relation_task_id(relation_id: str, source_project_id: str) -> str:
    return stable_id("relation_review", relation_id, source_project_id)


def _copy_award(award: dict[str, Any], source_group: dict[str, Any], target_group: dict[str, Any]) -> None:
    copied = dict(award)
    copied["relation_source"] = {
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


def _apply_project_relation(source_project: dict[str, Any], target_project: dict[str, Any],
                            source_groups: list[dict[str, Any]], target_groups: list[dict[str, Any]],
                            issues: list[dict[str, Any]], audit: dict[str, Any]) -> None:
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
    if unresolved or ambiguous_targets:
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


def apply_relationships(projects: list[dict[str, Any]], groups: list[dict[str, Any]],
                        issues: list[dict[str, Any]], plan: dict[str, Any],
                        resolutions: dict[str, dict[str, Any]] | None = None
                        ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """仅在 plan 显式声明关系时消费中标汇总组并关联到投标名单组。"""
    relationships = plan.get("relationships", [])
    if not relationships:
        return projects, groups, issues, []
    resolutions = resolutions or {}
    projects_by_table: dict[str, list[dict[str, Any]]] = defaultdict(list)
    groups_by_project: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for project in projects:
        projects_by_table[project["table_id"]].append(project)
    for group in groups:
        if group.get("project_id"):
            groups_by_project[group["project_id"]].append(group)
    consumed_groups = set()
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
            ranked = _rank_candidates(source, target_projects)
            decision = resolutions.get(task_id)
            selected = exact[0] if len(exact) == 1 else None
            basis = "exact_project_key" if selected else "unresolved"
            if not selected and decision and decision.get("decision") == "select_project":
                selected = next((project for project in target_projects
                                 if project["id"] == decision.get("project_id")), None)
                basis = "user_selection" if selected else "invalid_resolution"
            audit = {
                "id": stable_id("project_relation", relation_id, source["id"]),
                "relation_id": relation_id, "kind": RELATION_KIND,
                "source_project_id": source["id"], "source_project_name": _project_label(source),
                "status": "matched" if selected else "unresolved",
                "basis": basis, "target_project_id": selected["id"] if selected else None,
                "target_project_name": _project_label(selected) if selected else None,
                "review_task_id": task_id, "candidates": ranked,
                "resolution": decision if decision else None,
            }
            if selected:
                _apply_project_relation(source, selected, groups_by_project[source["id"]],
                                        groups_by_project[selected["id"]], issues, audit)
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
            audits.append(audit)
    issues = [issue for issue in issues if issue.get("group_id") not in consumed_groups]
    remaining_groups = [group for group in groups if group["id"] not in consumed_groups]
    remaining_project_ids = {group["project_id"] for group in remaining_groups if group.get("project_id")}
    remaining_projects = [project for project in projects if project["id"] in remaining_project_ids]
    return remaining_projects, remaining_groups, issues, audits


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
