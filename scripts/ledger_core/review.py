"""持久化组级中标确认，并生成内置提问工具所需的有界问题。 @author denovochen"""
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
        question = f"项目：{project_name}\n原中标企业：{match['original_award']}\n\n请选择对应的投标企业"
        tasks.append({
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


def pending_questions(ledger: dict[str, Any], decisions: dict[str, dict]) -> tuple[list[dict], int, int]:
    tasks = _award_tasks(ledger)
    pending = [task for task in tasks if task["review_task_id"] not in decisions]
    questions = [task["question"] for task in pending[:QUESTION_BATCH_SIZE]]
    return questions, len(tasks), len(pending)


def create_state(inputs: list[Path], plan: dict, output: Path, generated_at: str,
                 review_task_count: int) -> Path:
    resolved_output = output.expanduser().absolute()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    directory = resolved_output.parent / (".excel-ledger-work-" + uuid.uuid4().hex)
    directory.mkdir(parents=False, exist_ok=False)
    state_path = directory / "state.json"
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
    tasks = {task["review_task_id"]: task for task in _award_tasks(ledger)}
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
        if not task or task_id in decisions:
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
        }
        if source == "manual":
            decision["user_input"] = value
        decisions[task_id] = decision
    state["decisions"] = decisions
    return errors


def review_result(state_path: Path, ledger: dict[str, Any], state: dict[str, Any],
                  validation_errors: list[dict[str, str]] | None = None) -> dict[str, Any]:
    questions, _, remaining = pending_questions(ledger, state["decisions"])
    return {
        "kind": "award_review_required",
        "message": "请使用内置向用户提问工具确认非精确中标名称对应的投标企业。",
        "state": str(state_path),
        "review_task_count": state["review_task_count"],
        "remaining_task_count": remaining,
        "questions": questions,
        "validation_errors": validation_errors or [],
        "summary": ledger["summary"],
    }


def cleanup_state(path: Path) -> None:
    resolved = path.expanduser().resolve(strict=True)
    parent = resolved.parent
    if resolved.name == "state.json" and parent.name.startswith(".excel-ledger-work-"):
        shutil.rmtree(parent)
