"""持久化项目关系与组级中标确认，并生成有界复核问题。 @author denovochen"""
from __future__ import annotations

import json
import shutil
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .contract import VERSION, LedgerError, clean, name_key
from .workbook import load_json

STATE_SCHEMA = 1
QUESTION_BATCH_SIZE = 5
DEFERRED_VALUES = {"unresolved", "不确定", "暂不确定", "不知道", "无法确定"}


def _timestamp() -> str:
    return datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")


def _award_tasks(ledger: dict[str, Any]) -> list[dict[str, Any]]:
    groups = {group["id"]: group for group in ledger["groups"]}
    projects = {project["id"]: project for project in ledger["projects"]}
    source_files = {source["id"]: source["file_name"] for source in ledger["sources"]}
    tasks = []
    for issue in ledger["issues"]:
        task_id = issue.get("review_task_id")
        group = groups.get(issue.get("group_id"))
        if not task_id or not group or issue["code"] not in {"AWARD_NAME_MISMATCH", "AWARD_MATCH_CONFLICT"}:
            continue
        match = next((item for item in group["award_matches"] if item.get("review_task_id") == task_id), None)
        if not match or not match.get("recommended_record_id") or not match.get("recommended_bidder_name"):
            continue
        project = projects.get(group.get("project_id"), {"values": {}})
        project_name = clean(project.get("values", {}).get("project_name")) or clean(
            project.get("values", {}).get("project_code")) or "（原表未提供项目名称）"
        source_rows = group.get("source_rows") or [group["anchor_row"]]
        scope = group.get("lot_name") or group.get("lot_code") or "原表未提供标段名称"
        location = (f"{source_files.get(group['source_id'], '')} / {group['sheet']} / "
                    f"投标行{min(source_rows)}:{max(source_rows)}")
        award_locations = []
        for award in group["awards"]:
            if award["cell"] in match["award_cells"]:
                source = award.get("relation_source", {})
                origin = (f"{source_files.get(source.get('source_id', group['source_id']), '')} / "
                          f"{source.get('sheet', group['sheet'])}!{award['cell']}")
                if origin not in award_locations:
                    award_locations.append(origin)
        question = (f"项目：{project_name}\n招标范围：{scope}\n投标来源：{location}\n"
                    f"中标来源：{'; '.join(award_locations)}\n"
                    f"原中标企业：{match['original_award']}\n\n请选择对应的投标企业")
        tasks.append({
            "task_type": "award",
            "review_task_id": task_id,
            "issue_id": issue["id"],
            "group_id": group["id"],
            "source_id": group["source_id"],
            "sheet": group["sheet"],
            "anchor_row": group["anchor_row"],
            "project_name": project_name,
            "original_award": match["original_award"],
            "recommended_record_id": match["recommended_record_id"],
            "recommended_bidder_name": match["recommended_bidder_name"],
            "recommendation_basis": match["recommendation_basis"],
            "question": {
                "question_id": task_id,
                "question": question,
                "options": [
                    {"label": match["recommended_bidder_name"] + " (Recommended)",
                     "value": match["recommended_record_id"]},
                    {"label": "不确定", "value": "unresolved"},
                ],
                "multi_select": False,
                "allow_other": True,
            },
        })
    return sorted(tasks, key=lambda task: (
        task["source_id"], task["sheet"], task["anchor_row"], task["review_task_id"]))


def _relation_tasks(ledger: dict[str, Any]) -> list[dict[str, Any]]:
    issues = {issue.get("review_task_id"): issue for issue in ledger["issues"] if issue.get("review_task_id")}
    tasks = []
    for relation in ledger.get("relationships", []):
        task_id = relation.get("review_task_id")
        candidates = relation.get("candidates", [])
        if relation.get("status") in {"matched", "blocked"} or not task_id or not candidates:
            continue
        issue = issues.get(task_id, {})
        options = [{
            "label": f"{candidate['project_name']}（{candidate['sheet']}）",
            "value": candidate["project_id"],
        } for candidate in candidates]
        options.append({"label": "不确定", "value": "unresolved"})
        evidence = "\n".join(
            f"{index}. {candidate['project_name']}：{'；'.join(candidate['evidence'])}"
            for index, candidate in enumerate(candidates, 1))
        tasks.append({
            "task_type": "relation", "review_task_id": task_id,
            "issue_id": issue.get("id"), "source_project_id": relation["source_project_id"],
            "source_project_name": relation["source_project_name"], "candidates": candidates,
            "question": {
                "question_id": task_id,
                "question": (f"中标汇总项目：{relation['source_project_name']}\n\n候选依据：\n{evidence}"
                             "\n\n请选择对应的投标明细项目；不能可靠确认时选择不确定。"),
                "options": options, "multi_select": False, "allow_other": False,
            },
        })
    return sorted(tasks, key=lambda task: (task["source_project_name"], task["review_task_id"]))


def _all_tasks(ledger: dict[str, Any]) -> list[dict[str, Any]]:
    return _relation_tasks(ledger) + _award_tasks(ledger)


def pending_questions(ledger: dict[str, Any], decisions: dict[str, dict]) -> tuple[list[dict], int, int]:
    tasks = _all_tasks(ledger)
    pending = [task for task in tasks if task["review_task_id"] not in decisions]
    relation_pending = [task for task in pending if task["task_type"] == "relation"]
    selected = relation_pending if relation_pending else pending
    questions = [task["question"] for task in selected[:QUESTION_BATCH_SIZE]]
    return questions, len(tasks), len(pending)


def split_decisions(decisions: dict[str, dict]) -> tuple[dict[str, dict], dict[str, dict]]:
    relation = {key: value for key, value in decisions.items() if key.startswith("relation_review_")}
    award = {key: value for key, value in decisions.items() if not key.startswith("relation_review_")}
    return relation, award


def create_state(inputs: list[Path], plan: dict, output: Path, generated_at: str,
                 review_task_count: int, work_directory: Path | None = None) -> Path:
    resolved_output = output.expanduser().absolute()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    if work_directory is None:
        directory = resolved_output.parent / (".excel-ledger-work-" + uuid.uuid4().hex)
        directory.mkdir(parents=False, exist_ok=False)
    else:
        directory = work_directory.expanduser().resolve(strict=True)
        if not directory.is_dir() or not directory.name.startswith(".excel-ledger-work-"):
            raise LedgerError("内部工作目录无效")
    state_path = directory / "state.json"
    if state_path.exists():
        raise LedgerError("复核状态已存在")
    state = {
        "schema_version": STATE_SCHEMA,
        "parser_version": VERSION,
        "created_at": _timestamp(),
        "generated_at": generated_at,
        "inputs": [str(path.expanduser().resolve(strict=True)) for path in inputs],
        "output": str(resolved_output),
        "plan": plan,
        "review_task_count": review_task_count,
        "decisions": {},
    }
    save_state(state_path, state)
    return state_path


def load_state(path: Path) -> dict[str, Any]:
    path = path.expanduser().resolve(strict=True)
    if path.name != "state.json" or not path.is_file():
        raise LedgerError("复核状态路径必须指向 state.json")
    state = load_json(path)
    required = {"schema_version", "parser_version", "created_at", "generated_at", "inputs", "output",
                "plan", "review_task_count", "decisions"}
    if set(state) != required or state["schema_version"] != STATE_SCHEMA or state["parser_version"] != VERSION:
        raise LedgerError("复核状态版本或字段无效")
    if (not isinstance(state["inputs"], list) or not state["inputs"] or
            any(not isinstance(value, str) or not value for value in state["inputs"]) or
            not isinstance(state["output"], str) or not state["output"] or
            type(state["review_task_count"]) is not int or state["review_task_count"] < 1 or
            not isinstance(state["decisions"], dict)):
        raise LedgerError("复核状态内容无效")
    state["state_path"] = str(path)
    return state


def save_state(path: Path, state: dict[str, Any]) -> None:
    stored = {key: value for key, value in state.items() if key != "state_path"}
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(stored, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    temporary.replace(path)


def load_answers(path: Path) -> dict[str, Any]:
    payload = load_json(path.expanduser().resolve(strict=True))
    if set(payload) == {"answer"} and isinstance(payload["answer"], dict):
        payload = payload["answer"]
    if not isinstance(payload, dict):
        raise LedgerError("提问答案必须是 question_id 到 answer 的对象")
    return payload


def _answer_value(answer: Any) -> str:
    if isinstance(answer, str):
        return clean(answer)
    if isinstance(answer, list) and len(answer) == 1:
        return _answer_value(answer[0])
    if isinstance(answer, dict):
        for key in ("value", "text", "other", "answer", "input"):
            if key in answer:
                return _answer_value(answer[key])
        if len(answer) == 1:
            return _answer_value(next(iter(answer.values())))
    raise LedgerError("提问答案必须是单选值或 Other 文本")


def apply_answers(state: dict[str, Any], ledger: dict[str, Any], answers: dict[str, Any]) -> list[dict[str, str]]:
    tasks = {task["review_task_id"]: task for task in _all_tasks(ledger)}
    issued = {question["question_id"] for question in pending_questions(ledger, state["decisions"])[0]}
    records_by_group: dict[str, list[dict]] = {}
    for record in ledger["records"]:
        records_by_group.setdefault(record["group_id"], []).append(record)
    selected_by_group: dict[str, set[str]] = {}
    for group in ledger["groups"]:
        selected_by_group[group["id"]] = {
            match["selected_record_id"] for match in group["award_matches"] if match["selected_record_id"]
        }
    decisions = dict(state["decisions"])
    errors = []
    for task_id, answer in answers.items():
        task = tasks.get(task_id)
        if not task or task_id in decisions or task_id not in issued:
            errors.append({"question_id": task_id, "message": "问题不存在、已处理或不属于当前复核状态"})
            continue
        try:
            value = _answer_value(answer)
        except LedgerError as exc:
            errors.append({"question_id": task_id, "message": str(exc)})
            continue
        if name_key(value) in {name_key(item) for item in DEFERRED_VALUES}:
            decisions[task_id] = {"decision": "deferred", "decided_at": _timestamp()}
            continue
        if task["task_type"] == "relation":
            candidate = next((item for item in task["candidates"] if item["project_id"] == value), None)
            if not candidate:
                errors.append({"question_id": task_id, "message": "所选项目不属于当前关系候选"})
                continue
            decisions[task_id] = {
                "decision": "select_project", "project_id": candidate["project_id"],
                "project_name": candidate["project_name"], "decided_at": _timestamp(),
            }
            continue
        group_records = records_by_group.get(task["group_id"], [])
        selected = next((record for record in group_records if record["id"] == value), None)
        source = "recommended" if selected and selected["id"] == task["recommended_record_id"] else "manual"
        if not selected:
            matches = [record for record in group_records if name_key(record["company_name"]) == name_key(value)]
            if len(matches) == 1:
                selected = matches[0]
            else:
                errors.append({"question_id": task_id, "message": "输入企业名称未唯一对应当前招标组的投标企业"})
                continue
        if selected["id"] in selected_by_group.setdefault(task["group_id"], set()):
            errors.append({"question_id": task_id, "message": "该投标企业已对应本组另一条中标信息"})
            continue
        selected_by_group[task["group_id"]].add(selected["id"])
        decision = {
            "decision": "select_bidder",
            "record_id": selected["id"],
            "company_name": selected["company_name"],
            "source": source,
            "decided_at": _timestamp(),
            "confirmation_provenance": "local_answer_file_not_host_verified",
        }
        if source == "manual":
            decision["user_input"] = value
        decisions[task_id] = decision
    state["decisions"] = decisions
    return errors


def review_result(state_path: Path, ledger: dict[str, Any], state: dict[str, Any],
                  validation_errors: list[dict[str, str]] | None = None) -> dict[str, Any]:
    questions, total, remaining = pending_questions(ledger, state["decisions"])
    state["review_task_count"] = max(state["review_task_count"], total)
    relation_ids = {task["review_task_id"] for task in _relation_tasks(ledger)}
    relation_review = bool(questions and questions[0]["question_id"] in relation_ids)
    remaining_by_type = {kind: sum(task["task_type"] == kind and task["review_task_id"] not in state["decisions"]
                                   for task in _all_tasks(ledger)) for kind in ("relation", "award")}
    from .workflow import command
    return {
        "kind": "relationship_review_required" if relation_review else "award_review_required",
        "message": ("请确认非精确项目名称对应的投标明细项目。" if relation_review else
                    "请使用内置向用户提问工具确认非精确中标名称对应的投标企业。"),
        "state": str(state_path),
        "review_task_count": total,
        "remaining_task_count": remaining,
        "remaining_by_type": remaining_by_type,
        "batch_task_type": "relation" if relation_review else "award",
        "questions": questions,
        "decision_owner": "model" if relation_review else "user",
        "next_command": command("resolve", state=state_path, answers=state_path.with_name("answers.json")),
        "validation_errors": validation_errors or [],
        "summary": ledger["summary"],
    }


def cleanup_state(path: Path) -> None:
    resolved = path.expanduser().resolve(strict=True)
    parent = resolved.parent
    if resolved.name == "state.json" and parent.name.startswith(".excel-ledger-work-"):
        shutil.rmtree(parent)


def create_work_directory(output: Path) -> Path:
    resolved = output.expanduser().absolute()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    directory = resolved.parent / (".excel-ledger-work-" + uuid.uuid4().hex)
    directory.mkdir(parents=False, exist_ok=False)
    return directory


def cleanup_work_directory(directory: Path) -> None:
    resolved = directory.expanduser().resolve(strict=True)
    if resolved.is_dir() and resolved.name.startswith(".excel-ledger-work-"):
        shutil.rmtree(resolved)
