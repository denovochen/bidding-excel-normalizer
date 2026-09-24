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
from pathlib import Path

import excel_ledger
from ledger_core.contract import VERSION, LedgerError, stable_id
from ledger_core.review import save_state
from ledger_core.workbook import load_json

DEFAULT_RESPONSE_CHARS = 6000
HANDOFFS = {"mapping_required", "relationship_review_required", "award_review_required"}


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
    if state.get("protocol_version") != 1 or state.get("parser_version") != VERSION:
        raise LedgerError("Agent 会话版本不匹配，请使用创建该会话的 Skill 版本")
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
    result = {"kind": reply["kind"], "state": str(path), "state_version": version, "batch_id": batch}
    if reply["kind"] == "result":
        result.update({key: reply[key] for key in ("output", "files", "summary", "checks")})
        result["delivery"] = copy.deepcopy(reply["delivery"])
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
                       "next_command": command("resolve", state=path, answers=answers_path),
                       "answer_roles": reply.get("answer_roles", []),
                       "validation_errors": reply.get("validation_errors", []), "questions": []})
        selected = []
        for question in reply["questions"]:
            preview = brief(question)
            details = {"question_id": question["question_id"],
                       "detail_command": command("question", state=path, id=question["question_id"]),
                       "available_sections": list(question)}
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


def start(inputs: list[Path], output: Path, response_chars: int) -> dict:
    output = output.expanduser().absolute()
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise LedgerError("输出目录必须不存在或为空")
    reply = engine(["run", *[str(path) for path in inputs], "--output", str(output)])
    directory = output.parent / (".excel-agent-" + uuid.uuid4().hex)
    directory.mkdir()
    state = {"protocol_version": 1, "parser_version": VERSION, "session_id": uuid.uuid4().hex,
             "state_version": 0, "response_chars": response_chars, "receipts": {}}
    return save_reply(directory / "session.json", state, reply)


def submit(path: Path, answers_path: Path) -> dict:
    with session_lock(path):
        state = load_session(path)
        payload = load_json(answers_path)
        if (set(payload) != {"batch_id", "state_version", "answers"} or
                not isinstance(payload["answers"], dict) or type(payload["state_version"]) is not int):
            raise LedgerError("答案须包含 batch_id、state_version、answers，直接填写返回的 answer_file")
        batch = payload["batch_id"]
        if not isinstance(batch, str):
            raise LedgerError("batch_id 必须是字符串")
        digest = stable_id("answers", json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False))
        if batch in state["receipts"]:
            if state["receipts"][batch] != digest:
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
        if state["reply"]["kind"] == "relationship_review_required":
            questions = {q["question_id"]: q for q in state["reply"]["questions"]}
            for key, answer in payload["answers"].items():
                if answer == "unresolved" or isinstance(answer, dict) and answer.get("value") == "unresolved":
                    continue
                if (not isinstance(answer, dict) or set(answer) != {"value", "basis", "evidence"} or
                        not isinstance(answer["value"], str) or not isinstance(answer["basis"], str) or
                        not answer["basis"].strip() or len(answer["basis"]) > 1000):
                    raise LedgerError("项目关系选择须包含 value、basis 和 evidence；无法确认用 unresolved")
                required = questions[key]["evidence_refs"].get(answer["value"])
                if not required or answer["evidence"] != required:
                    raise LedgerError("项目关系必须引用所选候选 evidence_refs 中的汇总与名册来源")
        engine_answers = path.with_name(f"submitted-{state['state_version']}.json")
        save_state(engine_answers, payload["answers"])
        state["inflight"] = {"batch_id": batch, "digest": digest}
        save_state(path, state)
        reply = engine(["resolve", "--state", state["reply"]["state"], "--answers", str(engine_answers)])
        state["receipts"][batch] = digest
        return save_reply(path, state, reply)


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
                                          section="columns" if "columns" in question else "question")}
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
               for region in q.get("regions", [])]
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
    sub.add_parser("doctor")
    for action in ("status", "resolve", "question", "inspect"):
        child = sub.add_parser(action)
        child.add_argument("--state", type=Path, required=True)
        if action == "resolve":
            child.add_argument("--answers", type=Path, required=True)
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
        elif args.action == "run":
            if not 4000 <= args.response_chars <= 12000:
                raise LedgerError("response-chars 须为4000到12000")
            result = start(args.inputs, args.output, args.response_chars)
        else:
            args.state = args.state.expanduser().resolve(strict=True)
            if args.action == "status":
                state = load_session(args.state)
                if state.get("inflight"):
                    raise LedgerError("上次引擎提交中断，请核对会话与引擎状态")
                result = state["snapshot"]
            elif args.action == "resolve":
                result = submit(args.state, args.answers)
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
        print(encoded({"kind": "error", "message": message[:2000], "message_truncated": len(message) > 2000}), flush=True)
        return 2


if __name__ == "__main__":
    sys.exit(main())
