#!/usr/bin/env python3
"""面向通用 Agent 的短响应、按批次恢复入口。 @author denovochen"""
from __future__ import annotations

import argparse
import copy
import json
import os
import shlex
import sys
import uuid
from contextlib import contextmanager
from collections import Counter
from pathlib import Path

import excel_ledger
from ledger_core.contract import VERSION, LedgerError, stable_id
from ledger_core.artifacts import digest
from ledger_core.review import save_state
from ledger_core.workbook import load_json
from ledger_core.workflow import DEFAULT_SCOPE, validate_scope_answer

RELEASE_VERSION = "1.8.2"
DEFAULT_RESPONSE_CHARS = 6000
MAX_ANSWER_CHARS = 2000
HANDOFFS = {"mapping_required", "relationship_review_required", "award_review_required"}


def workflow_progress(reply: dict, *, paused: bool = False) -> dict:
    """Project actual engine handoffs onto the three host todo items."""
    kind = reply.get("kind")
    step = 3 if kind == "result" else 2 if kind in HANDOFFS else 1
    phase = reply.get("stage") or reply.get("batch_task_type") or kind
    if kind == "mapping_required" and reply.get("questions"):
        phase = reply["questions"][0].get("task_type", phase)
    completed = kind == "result" and reply.get("validated") is True
    activity = "complete" if completed else "paused" if paused or kind == "error" else "running"
    return {"name": "bidding-excel-normalizer", "schema_version": 1, "step": step,
            "sync_tool": "write_todos", "sync_before_next_command": True,
            "phase": phase, "activity": activity,
            "todos": [{"content": title, "status": "completed" if completed or index < step else
                       "in_progress" if index == step and activity == "running" else "pending"}
                      for index, title in enumerate(excel_ledger.PROGRESS_TITLES, 1)]}


def checked_status(path: Path, *, resume: bool = False) -> dict:
    """Verify a resume point without replaying run/resolve or modifying answers."""
    state = load_session(path)
    if path.with_suffix(".lock").exists() or state.get("inflight"):
        raise LedgerError("提交仍在运行或上次提交中断；先核对引擎状态，不自动重新 run/resolve")
    snapshot = state["snapshot"]
    if (type(state["state_version"]) is not int or state["state_version"] < 1 or
            Path(snapshot["state"]).resolve() != path.resolve() or
            snapshot["batch_id"] != state["batch_id"] or snapshot["state_version"] != state["state_version"] or
            state["batch_id"] != stable_id("batch", state["session_id"], state["state_version"])):
        raise LedgerError("会话快照与当前批次不一致，不能恢复")
    if len(state["receipts"]) != state["state_version"] - 1:
        raise LedgerError("提交回执数量与会话版本不一致")
    checks = {"batch": "verified", "committed_batches": len(state["receipts"])}
    committed_scope = None
    for version in range(1, state["state_version"]):
        batch = stable_id("batch", state["session_id"], version)
        answers = load_json(path.with_name(f"submitted-{version}.json"))
        payload = {"batch_id": batch, "state_version": version, "answers": answers}
        receipt = stable_id("answers", json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False))
        if state["receipts"].get(batch) != receipt:
            raise LedgerError("已提交答案与回执不一致，不能猜测已完成的批次")
        if "scope_selection" in answers:
            committed_scope = answers["scope_selection"]
    if snapshot["kind"] == "result":
        checked = engine(["validate", snapshot["output"]])
        if checked["summary"] != snapshot["summary"]:
            raise LedgerError("已发布产物与会话摘要不一致")
        checks["artifacts"] = "verified"
    else:
        inner = load_json(Path(state["reply"]["state"]))
        if inner.get("parser_version") != state["parser_version"]:
            raise LedgerError("引擎状态版本不一致")
        if inner.get("stage") == "structure":
            expected = list(inner["source_hashes"].values()) + [x["sha256"] for x in inner.get("failures", [])]
            if not set(state["issued_questions"]) <= set(inner["issued_questions"]):
                raise LedgerError("引擎问题与当前批次不一致")
            if committed_scope is not None and (inner.get("scope", {}).get("decision") != committed_scope or
                    inner["scope"]["goal"] != state.get("scope_summary", {}).get("goal")):
                raise LedgerError("业务范围与已提交回执不一致，不能恢复")
        else:
            expected = [source["sha256"] for source in inner["plan"]["sources"]]
            if set(state["issued_questions"]).intersection(inner["decisions"]):
                raise LedgerError("当前问题已被引擎处理，需要核对中断提交")
        actual = [digest(Path(value)) for value in inner["inputs"]]
        if Counter(actual) != Counter(expected):
            raise LedgerError("输入指纹变化，不能继续使用旧决定，也不能自动重新 run")
        pending = load_json(Path(snapshot["answer_file"]))
        if (pending.get("batch_id") != state["batch_id"] or
                pending.get("state_version") != state["state_version"] or
                not isinstance(pending.get("answers"), dict) or
                not set(pending["answers"]) <= set(state["issued_questions"])):
            raise LedgerError("当前答案文件与待处理批次不一致")
        checks.update(inputs="fingerprints_verified", engine_state="verified", answer_batch="verified")
    result = copy.deepcopy(snapshot)
    result["workflow"] = workflow_progress(state["reply"], paused=not resume)
    result["resume_checks"] = checks
    if snapshot["kind"] in HANDOFFS:
        result["answer_command"] = command("answer", state=path, batch_id=state["batch_id"],
                                           state_version=state["state_version"])
        result["questions"] = [{"question_id": key, "details_required": True,
                                "detail_command": command("question", state=path, id=key)}
                               for key in state["issued_questions"]]
    return result


class ProtocolParser(argparse.ArgumentParser):
    def error(self, message):
        raise LedgerError(message)


def encoded(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def command(action: str, **args) -> str:
    return shlex.join([sys.executable, str(Path(__file__).resolve()), action,
                       *[part for key, value in args.items() if value is not None
                         for part in ("--" + key.replace("_", "-"), str(value))]])


def engine(args: list[str]) -> dict:
    replies = []
    code = excel_ledger.main(args, emit=replies.append)
    if len(replies) != 1:
        raise LedgerError("解析引擎未返回单个结构化结果")
    reply = replies[0]
    if code not in {0, 3, 4, 5} or reply.get("kind") == "error":
        raise LedgerError(reply.get("message", "解析引擎失败"))
    if reply.get("kind") in HANDOFFS and not reply.get("state"):
        raise LedgerError(reply.get("message", "解析引擎缺少可恢复状态"))
    return reply


def load_session(path: Path) -> dict:
    state = load_json(path)
    if state.get("protocol_version") not in {1, 2} or state.get("parser_version") != VERSION:
        raise LedgerError("Agent 会话版本不匹配，请使用创建该会话的 Skill 版本")
    if state["snapshot"]["kind"] == "result":
        state["snapshot"]["files"] = ["final.csv", "review_queue.csv"]
        state["snapshot"]["internal_files"] = ["ledger.json"]
    return state


@contextmanager
def session_lock(path: Path):
    lock = path.with_suffix(".lock")
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError as exc:
        raise LedgerError("会话正在提交或上次进程异常退出；请核对会话后重试") from exc
    try:
        os.close(descriptor)
        yield
    finally:
        lock.unlink()


def brief(question: dict) -> dict:
    result = {key: value for key, value in question.items()
              if key not in {"answer_template", "header_preview", "source_block_examples", "inspect_command"}}
    if isinstance(question.get("answer_template"), dict):
        result["answer_defaults"] = {k: v for k, v in question["answer_template"].items()
                                     if k not in {"columns", "basis"}}
    if "columns" in result:
        result["columns"] = [{key: item[key] for key in ("column", "header", "suggested_role", "samples")
                               if key in item} for item in result["columns"]]
        for compact, full in zip(result["columns"], question["columns"]):
            compact.update({key: full[key] for key in ("formula_count", "error_count") if full.get(key)})
    return result


def save_reply(path: Path, state: dict, reply: dict) -> dict:
    state["reply"] = reply
    state["state_version"] += 1
    version = state["state_version"]
    batch = stable_id("batch", state["session_id"], version)
    state["batch_id"] = batch
    answers_path = path.with_name(f"answers-{version}.json")
    result = {"kind": reply["kind"], "state": str(path), "state_version": version, "batch_id": batch,
              "workflow": workflow_progress(reply)}
    if reply["kind"] == "result":
        result.update({key: reply[key] for key in ("output", "validated", "summary", "checks")})
        result["files"] = ["final.csv", "review_queue.csv"]
        result["internal_files"] = ["ledger.json"]
        result["delivery"] = copy.deepcopy(reply["delivery"])
        if "scope_summary" in state:
            result["delivery"]["excluded_sheet_count"] = state["scope_summary"]["excluded_sheet_count"]
            result["delivery"]["message"] += f"按业务范围排除 {state['scope_summary']['excluded_sheet_count']} 个工作表，原因保留在内部审计。"
        reasons = result["delivery"]["review_reasons"]
        total = len(reasons)
        result["delivery"]["omitted_reason_types"] = 0
        result["delivery"]["details_file"] = str(Path(reply["output"]) / "ledger.json")
        while len(encoded(result)) > state["response_chars"] and reasons:
            reasons.pop()
            result["delivery"]["omitted_reason_types"] = total - len(reasons)
        result["delivery"]["omitted_reason_types"] = total - len(reasons)
        state["issued_questions"] = []
    else:
        result.update({"stage": reply.get("stage", reply.get("batch_task_type")),
                       "decision_owner": reply.get("decision_owner", "model"),
                       "remaining_task_count": reply["remaining_task_count"],
                       "answer_file": str(answers_path),
                       "answer_command": command("answer", state=path, batch_id=batch, state_version=version),
                       "answer_roles": reply.get("answer_roles", []),
                       "validation_errors": reply.get("validation_errors", []), "questions": []})
        selected = []
        for question in reply["questions"]:
            preview = brief(question)
            details = {"question_id": question["question_id"],
                       "detail_command": command("question", state=path, id=question["question_id"])}
            if "task_type" in question:
                details["task_type"] = question["task_type"]
            if question.get("task_type") == "scope":
                details["goal"] = question["goal"]
            if result["decision_owner"] != "user":
                preview.update(details)
            result["questions"].append(preview)
            if len(encoded(result)) > state["response_chars"]:
                result["questions"].pop()
                if selected:
                    break
                preview = details
                preview["details_required"] = True
                result["questions"].append(preview)
            selected.append(question)
        state["issued_questions"] = [question["question_id"] for question in selected]
        templates = {question["question_id"]: question.get("answer_template", "unresolved") for question in selected}
        save_state(answers_path, {"batch_id": batch, "state_version": version, "answers": templates})
    if len(encoded(result)) > state["response_chars"]:
        raise LedgerError("协议控制信息超出响应预算，请缩短输出路径或增大 --response-chars")
    state["snapshot"] = result
    state.pop("inflight", None)
    save_state(path, state)
    return result


def start(inputs: list[Path], output: Path, response_chars: int, scope: str | None = DEFAULT_SCOPE) -> dict:
    output = output.expanduser().absolute()
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise LedgerError("输出目录必须不存在或为空")
    if scope is not None and (not isinstance(scope, str) or not scope.strip() or len(scope) > 500):
        raise LedgerError("业务范围须为1到500字的目标描述")
    reply = engine(["run", *[str(path) for path in inputs], "--output", str(output),
                    *(["--scope", scope] if scope is not None else [])])
    directory = output.parent / (".excel-agent-" + uuid.uuid4().hex)
    directory.mkdir()
    state = {"protocol_version": 2, "parser_version": VERSION, "session_id": uuid.uuid4().hex,
             "state_version": 0, "response_chars": response_chars, "receipts": {}}
    return save_reply(directory / "session.json", state, reply)


def submit(path: Path, answers_path: Path) -> dict:
    with session_lock(path):
        return _submit_payload(path, load_session(path), load_json(answers_path))


def _submit_payload(path: Path, state: dict, payload: dict) -> dict:
    if (not isinstance(payload, dict) or set(payload) != {"batch_id", "state_version", "answers"} or
            not isinstance(payload["answers"], dict) or type(payload["state_version"]) is not int):
        raise LedgerError("答案须包含 batch_id、state_version、answers")
    batch = payload["batch_id"]
    if not isinstance(batch, str):
        raise LedgerError("batch_id 必须是字符串")
    receipt = stable_id("answers", json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False))
    if batch in state["receipts"]:
        if state["receipts"][batch] != receipt:
            raise LedgerError("该批次已提交，不能用不同答案覆盖")
        return state["snapshot"]
    if state.get("inflight"):
        raise LedgerError("上次引擎提交中断；请核对会话与引擎状态，不自动重复提交")
    if batch != state["batch_id"] or payload["state_version"] != state["state_version"]:
        raise LedgerError("答案批次已过期；运行 status 获取当前批次")
    if not payload["answers"] or not set(payload["answers"]) <= set(state["issued_questions"]):
        raise LedgerError("只能提交当前批次问题；历史答案由程序保存")
    if state["reply"]["kind"] not in HANDOFFS:
        raise LedgerError("会话已结束，无待提交的问题")
    scope_question = next((q for q in state["reply"].get("questions", [])
                           if q.get("task_type") == "scope" and q["question_id"] in payload["answers"]), None)
    if scope_question:
        validate_scope_answer(scope_question, payload["answers"][scope_question["question_id"]])
    if state["reply"]["kind"] == "relationship_review_required":
        questions = {q["question_id"]: q for q in state["reply"]["questions"]}
        for key, value in payload["answers"].items():
            if value == "unresolved" or isinstance(value, dict) and value.get("value") == "unresolved":
                continue
            if (not isinstance(value, dict) or set(value) != {"value", "basis", "evidence"} or
                    not isinstance(value["value"], str) or not isinstance(value["basis"], str) or
                    not value["basis"].strip() or len(value["basis"]) > 1000):
                raise LedgerError("项目关系选择须包含 value、basis 和 evidence；无法确认用 unresolved")
            required = questions[key]["evidence_refs"].get(value["value"])
            if not required or value["evidence"] != required:
                raise LedgerError("项目关系必须引用所选候选 evidence_refs 中的汇总与名册来源")
    engine_answers = path.with_name(f"submitted-{state['state_version']}.json")
    save_state(engine_answers, payload["answers"])
    state["inflight"] = {"batch_id": batch, "digest": receipt}
    save_state(path, state)
    reply = engine(["resolve", "--state", state["reply"]["state"], "--answers", str(engine_answers)])
    if scope_question:
        state["scope_summary"] = {"goal": scope_question["goal"],
                                  "excluded_sheet_count": len(payload["answers"][scope_question["question_id"]]["exclude"])}
    state["receipts"][batch] = receipt
    return save_reply(path, state, reply)


def _merge_answer(template, patch):
    if not isinstance(template, dict) or not isinstance(patch, dict):
        return copy.deepcopy(patch)
    if patch.get("action") == "review":
        if set(patch) != {"action", "basis"}:
            raise LedgerError("review 只接受 action 和 basis")
        return copy.deepcopy(patch)
    allowed = set(template) | ({"header_rows"} if "columns" in template else set())
    if set(patch) - allowed:
        raise LedgerError("答案包含当前模板之外的字段")
    result = copy.deepcopy(template)
    for key, value in patch.items():
        if key in {"columns", "exclude"}:
            if not isinstance(value, dict):
                raise LedgerError(f"{key} 必须是对象")
            if key == "columns" and set(value) - set(template[key]):
                raise LedgerError("列标不属于当前问题")
            result[key].update(value)
            if key == "exclude":
                result[key] = {label: reason for label, reason in result[key].items() if reason is not None}
        else:
            result[key] = copy.deepcopy(value)
    return result


def answer(path: Path, batch: str, version: int, question_id: str | None, patch, *, draft: bool = False) -> dict:
    """Apply semantic fields, never text replacements; commit bounded current answers."""
    patches = {question_id: patch} if question_id is not None else patch
    if not isinstance(patches, dict) or not patches:
        raise LedgerError("未指定 --id 时，--json 必须是非空的 question_id 到答案的对象")
    with session_lock(path):
        state = load_session(path)
        if version < 1 or batch != stable_id("batch", state["session_id"], version):
            raise LedgerError("答案批次与版本不匹配")
        if batch in state["receipts"]:
            payload = {"batch_id": batch, "state_version": version,
                       "answers": load_json(path.with_name(f"submitted-{version}.json"))}
            if draft or any(key not in payload["answers"] or
                            _merge_answer(payload["answers"][key], value) != payload["answers"][key]
                            for key, value in patches.items()):
                raise LedgerError("该批次已提交，不能用不同答案覆盖")
            return _submit_payload(path, state, payload)
        if batch != state["batch_id"] or version != state["state_version"] or state.get("inflight"):
            raise LedgerError("批次过期或上次提交中断；先用 status --resume 核验，不重新 run")
        questions = {q["question_id"]: q for q in state["reply"].get("questions", [])
                     if q["question_id"] in state["issued_questions"]}
        if not set(patches) <= set(questions):
            raise LedgerError("问题不属于当前批次")
        answers_path = Path(state["snapshot"]["answer_file"])
        pending = load_json(answers_path)
        if pending.get("batch_id") != batch or pending.get("state_version") != version:
            raise LedgerError("答案文件与当前批次不一致")
        answers = {}
        for key, fields in patches.items():
            question = questions[key]
            original = question.get("answer_template", "unresolved")
            value = _merge_answer(pending["answers"].get(key, original), fields)
            if isinstance(value, dict) and "basis" in value:
                if not isinstance(value["basis"], str) or len(value["basis"]) > 256:
                    raise LedgerError("basis 须为不超过256字的依据")
            if not draft:
                if isinstance(original, dict):
                    if (not isinstance(value, dict) or not isinstance(value.get("basis"), str) or
                            not value["basis"].strip() or value["basis"] == original.get("basis")):
                        raise LedgerError("提交前必须填写具体 basis；需要分段填写时使用 --draft")
                if question.get("task_type") == "scope":
                    validate_scope_answer(question, value)
                if isinstance(value, dict) and value.get("action") == "interpret":
                    roles = value.get("columns")
                    if not isinstance(roles, dict) or any(role not in state["reply"]["answer_roles"] for role in roles.values()):
                        raise LedgerError("仍有未解释或无效列角色；可继续 --draft 或提交 review")
            answers[key] = value
        if draft:
            pending["answers"].update(answers)
            save_state(answers_path, pending)
            return {"kind": "answer_saved", "state": str(path), "batch_id": batch,
                    "state_version": version, "question_ids": list(answers),
                    "workflow": workflow_progress(state["reply"]),
                    "answer_command": command("answer", state=path, batch_id=batch, state_version=version, id=question_id)}
        return _submit_payload(path, state, {"batch_id": batch, "state_version": version,
                                            "answers": answers})


def question_page(path: Path, question_id: str, section: str | None, offset: int, limit: int) -> dict:
    state = load_session(path)
    question = next((item for item in state["reply"].get("questions", [])
                     if item["question_id"] == question_id and question_id in state["issued_questions"]), None)
    if question is None:
        raise LedgerError("问题不属于当前批次")
    question = copy.deepcopy(question)
    if "inspect_command" in question:
        region = question["regions"][0]
        question["inspect_command"] = command("inspect", state=path, region=region["id"], start=region["start"], count=3)
    if section is None:
        result = {"kind": "question_index", "question_id": question_id,
                  "sections": {key: {"type": type(value).__name__, "size": len(value) if isinstance(value, (dict, list, str)) else 1}
                               for key, value in question.items()},
                  "next_command": command("question", state=path, id=question_id,
                                          section="columns" if "columns" in question else
                                          "sheets" if "sheets" in question else
                                          "candidate_bidder_sets" if "candidate_bidder_sets" in question else "question")}
        if "inspect_command" in question:
            result["inspect_command"] = question["inspect_command"]
            result["inspect_limits"] = {"rows": 20, "columns": 16}
        return result
    value = question
    try:
        for key in section.split("."):
            value = value[int(key)] if isinstance(value, list) else value[key]
    except (KeyError, IndexError, TypeError, ValueError) as exc:
        raise LedgerError("问题证据路径无效；使用 question 查看可用 sections") from exc
    if isinstance(value, list):
        items = value[offset:offset + limit]
    elif isinstance(value, dict):
        items = [{"key": key, "value": item} for key, item in list(value.items())[offset:offset + limit]]
    else:
        items = [value] if offset == 0 else []
    total = len(value) if isinstance(value, (list, dict)) else 1
    result = {"kind": "question_detail", "question_id": question_id, "section": section,
              "offset": offset, "total": total, "items": items, "next_command": None}
    while items:
        end = offset + len(items)
        result["next_command"] = command("question", state=path, id=question_id, section=section, offset=end, limit=limit) if end < total else None
        if len(encoded(result)) <= state["response_chars"]:
            return result
        items.pop()
    if offset < total:
        child = value[offset] if isinstance(value, list) else value
        children = list(child) if isinstance(child, dict) else []
        result.update(kind="question_index", message="单项较大，请用 section 子路径继续读取",
                      sections=[f"{section}.{offset}.{key}" if isinstance(value, list) else f"{section}.{key}" for key in children],
                      next_command=None)
        if not children:
            raise LedgerError("单项证据超出响应预算，请增大 --response-chars 后重新运行")
    return result


def inspect_page(path: Path, region_id: str, start_row: int, count: int, columns: str | None) -> dict:
    state = load_session(path)
    reply = state["reply"]
    if reply.get("stage") != "structure":
        raise LedgerError("结构补证据仅适用于结构阶段")
    regions = [region for q in reply["questions"] if q["question_id"] in state["issued_questions"]
               for region in q.get("regions", [r for sheet in q.get("sheets", []) for r in sheet["regions"]])]
    region = next((item for item in regions if item["id"] == region_id), None)
    if not region or not region["start"] <= start_row <= region["end"] or not 1 <= count <= 20:
        raise LedgerError("区域和起始行必须属于当前问题，count 须为1到20")
    end = min(start_row + count - 1, region["end"])
    args = ["inspect", "--state", reply["state"], "--region", region_id, "--rows", f"{start_row}:{end}"]
    if columns:
        args.extend(["--columns", columns])
    result = engine(args)
    while result["rows"]:
        next_row = start_row + len(result["rows"])
        result["next_command"] = command("inspect", state=path, region=region_id, start=next_row,
                                         count=count, columns=columns) if next_row <= region["end"] else None
        result["returned_row_count"] = len(result["rows"])
        result["next_column_command"] = command("inspect", state=path, region=region_id, start=start_row,
                                                count=len(result["rows"]), columns=result["next_column_page"]) if result["next_column_page"] else None
        if len(encoded(result)) <= state["response_chars"]:
            return result
        result["rows"].pop()
    raise LedgerError("当前列页超出预算；请用 --columns A:H 等更小列范围")


def main(argv: list[str] | None = None) -> int:
    parser = ProtocolParser(description="招投标 Excel 平台入口：正常交接退出0，按 kind 继续")
    sub = parser.add_subparsers(dest="action", required=True)
    run = sub.add_parser("run")
    run.add_argument("inputs", nargs="+", type=Path)
    run.add_argument("--output", required=True, type=Path)
    run.add_argument("--response-chars", type=int, default=DEFAULT_RESPONSE_CHARS)
    run.add_argument("--scope", default=DEFAULT_SCOPE, help="业务目标；默认投标明细及关联汇总，不按Sheet位置筛选")
    sub.add_parser("doctor")
    for action in ("status", "resolve", "answer", "question", "inspect"):
        child = sub.add_parser(action)
        child.add_argument("--state", type=Path, required=True)
        if action == "status":
            child.add_argument("--resume", action="store_true", help="核验成功后显示当前阶段恢复执行；不重新 run/resolve")
        elif action == "resolve":
            child.add_argument("--answers", type=Path, required=True)
        elif action == "answer":
            child.add_argument("--batch-id", required=True)
            child.add_argument("--state-version", type=int, required=True)
            child.add_argument("--id", help="单问题ID；省略时 --json 为当前问题ID到答案的映射")
            child.add_argument("--json", required=True, dest="answer_json")
            child.add_argument("--draft", action="store_true", help="分段保存字段，不提交本批次")
        elif action == "question":
            child.add_argument("--id", required=True)
            child.add_argument("--section")
            child.add_argument("--offset", type=int, default=0)
            child.add_argument("--limit", type=int, default=8)
        elif action == "inspect":
            child.add_argument("--region", required=True)
            child.add_argument("--start", type=int, required=True)
            child.add_argument("--count", type=int, default=3)
            child.add_argument("--columns")
    try:
        args = parser.parse_args(argv)
        if args.action == "doctor":
            result = engine(["doctor"])
            result["skill_version"] = RELEASE_VERSION
            result["parser_version"] = VERSION
            result["workflow"] = workflow_progress(result)
        elif args.action == "run":
            if not 4000 <= args.response_chars <= 12000:
                raise LedgerError("response-chars 须为4000到12000")
            result = start(args.inputs, args.output, args.response_chars, args.scope)
        else:
            args.state = args.state.expanduser().resolve(strict=True)
            if args.action == "status":
                result = checked_status(args.state, resume=args.resume)
            elif args.action == "resolve":
                result = submit(args.state, args.answers)
            elif args.action == "answer":
                if len(args.answer_json) > MAX_ANSWER_CHARS:
                    raise LedgerError("单次答案最多2000字符；按字段分段使用 --draft，最后一次去掉 --draft 提交")
                value = json.loads(args.answer_json)
                encoded(value)  # Reject NaN/Infinity before persisting a draft.
                result = answer(args.state, args.batch_id, args.state_version, args.id, value, draft=args.draft)
            elif args.action == "question":
                if args.offset < 0 or not 1 <= args.limit <= 20:
                    raise LedgerError("offset 不能为负数，limit 须为1到20")
                result = question_page(args.state, args.id, args.section, args.offset, args.limit)
            else:
                result = inspect_page(args.state, args.region, args.start, args.count, args.columns)
        print(encoded(result), flush=True)
        return 0
    except (LedgerError, OSError, ValueError, TypeError, KeyError) as exc:
        message = str(exc) if isinstance(exc, LedgerError) else f"Agent 输入或状态无效 ({type(exc).__name__})"
        error = {"kind": "error", "message": message[:2000], "message_truncated": len(message) > 2000}
        last = error
        try:
            if "args" in locals() and getattr(args, "state", None):
                last = load_session(args.state)["reply"]
        except (LedgerError, OSError, ValueError, TypeError, KeyError):
            pass
        if last.get("kind") == "result":
            last = {**last, "validated": False}
        error["workflow"] = workflow_progress(last, paused=True)
        error["workflow"]["activity"] = "failed"
        print(encoded(error), flush=True)
        return 2


if __name__ == "__main__":
    sys.exit(main())
