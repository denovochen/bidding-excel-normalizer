"""宿主模型双名称比较的有界交接、恢复与幂等完成。 @author denovochen"""
from __future__ import annotations

import hashlib
import json
import os
import tempfile
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable

from .artifacts import publish, validate_outputs
from .contract import VERSION, LedgerError
from .names import DECISIONS
from .normalize import build_outputs
from .workbook import load_json, read_workbook

MAX_PAIRS = 6
MAX_BATCH_CHARS = 4000
Progress = Callable[[int, str], None]


def _fingerprint(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _save(path: Path, payload: dict) -> None:
    fd, temporary = tempfile.mkstemp(prefix=".state-", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


@contextmanager
def _lock(path: Path):
    # QM 的 Python 运行于 Linux；同时支持本地 macOS 验证。
    import fcntl
    with (path.parent / "state.lock").open("a+b") as stream:
        fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _batch(state: dict) -> dict | None:
    pending = [p for p in state["pairs"] if p["id"] not in state["decisions"]]
    if not pending:
        return None
    selected = []
    for pair in pending:
        if len(selected) == MAX_PAIRS or len(json.dumps(selected + [pair], ensure_ascii=False)) > MAX_BATCH_CHARS:
            break
        selected.append(pair)
    if not selected:
        raise LedgerError("名称对超过交接长度限制")
    batch_id = _fingerprint([state["run_id"], state["input_fingerprint"], selected])
    return {"batch_id": batch_id, "pairs": selected}


def _handoff(path: Path, state: dict) -> dict:
    batch = _batch(state)
    if not batch:
        return {"kind": "ready_to_finalize", "state": str(path)}
    return {"kind": "name_review_required", "state": str(path), **batch,
            "remaining_pairs": sum(p["id"] not in state["decisions"] for p in state["pairs"]),
            "message": "宿主模型按 name-review.md 仅比较两个名字；不问用户，写决定文件后调用 resolve。"}


def start(output: Path, books: list, plan: dict, prepared: tuple, progress: Progress) -> dict:
    final, review, ledger = prepared
    output = output.expanduser().absolute()
    if output.is_symlink() or (output.exists() and (not output.is_dir() or any(output.iterdir()))):
        raise LedgerError("输出目录必须不存在或为空；不会覆盖已有产物")
    if not ledger["name_pairs"]:
        progress(2, "completed")
        progress(3, "in_progress")
        result = publish(output, final, review, ledger)
        progress(3, "completed")
        return result
    output.parent.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix=".excel-ledger-work-", dir=output.parent)).resolve()
    path = work / "state.json"
    sources = [{"path": str(b.path), "sha256": b.sha256} for b in books]
    state = {"schema_version": 1, "parser_version": VERSION, "run_id": uuid.uuid4().hex,
             "sources": sources, "plan": plan, "output": str(output), "generated_at": ledger["generated_at"],
             "pairs": ledger["name_pairs"], "decisions": {}, "batches": {}, "result": None,
             "input_fingerprint": _fingerprint([sources, plan])}
    _save(path, state)
    return _handoff(path, state)


def _load_state(path: Path) -> dict:
    state = load_json(path)
    required = {"schema_version", "parser_version", "run_id", "sources", "plan", "output", "generated_at",
                "pairs", "decisions", "batches", "result", "input_fingerprint"}
    if set(state) != required or state["schema_version"] != 1 or state["parser_version"] != VERSION:
        raise LedgerError("任务状态格式/解析器版本不一致")
    if state["input_fingerprint"] != _fingerprint([state["sources"], state["plan"]]):
        raise LedgerError("任务输入或映射快照发生变化")
    pairs = state["pairs"]
    if not isinstance(pairs, list) or any(set(p) != {"id", "name_a", "name_b"} for p in pairs):
        raise LedgerError("名称对状态无效")
    if len({p["id"] for p in pairs}) != len(pairs) or set(state["decisions"]) - {p["id"] for p in pairs}:
        raise LedgerError("名称对重复或决定指向未知名称对")
    if any(d not in DECISIONS for d in state["decisions"].values()):
        raise LedgerError("名称对决定无效")
    return state


def _prepared(state: dict) -> tuple:
    books = [read_workbook(Path(s["path"])) for s in state["sources"]]
    if [b.sha256 for b in books] != [s["sha256"] for s in state["sources"]]:
        raise LedgerError("原 Excel 已改变；拒绝把旧名称决定应用到新文件")
    outputs = build_outputs(books, state["plan"], generated_at=state["generated_at"], name_decisions=state["decisions"])
    if outputs[2]["name_pairs"] != state["pairs"]:
        raise LedgerError("名称候选与任务快照不一致")
    return outputs


def _complete_or_recover(path: Path, state: dict, outputs: tuple, progress: Progress) -> dict:
    output = Path(state["output"])
    # 发布成功但状态尚未落盘时也不重复发布；核对同一输入、映射、决定和 CSV 哈希后恢复。
    if output.exists() and any(output.iterdir()):
        checked = validate_outputs(output)
        existing = json.loads((output / "ledger.json").read_text(encoding="utf-8"))
        expected = outputs[2]
        for key in expected:
            if existing.get(key) != expected[key]:
                raise LedgerError("输出目录不是本任务的已完成结果")
        result = {"kind": "result", "output": str(output), "files": ["final.csv", "ledger.json", "review_queue.csv"], **checked}
    else:
        progress(3, "in_progress")
        result = publish(output, *outputs)
        progress(3, "completed")
    state["result"] = result
    _save(path, state)
    return result


def resolve(state_path: Path, decisions_path: Path | None, progress: Progress) -> dict:
    path = state_path.expanduser().resolve(strict=True)
    with _lock(path):
        state = _load_state(path)
        progress(2, "in_progress")
        # 先验证原文件/候选不变，再接受任何决定。
        outputs = _prepared(state)
        if decisions_path is not None:
            payload = load_json(decisions_path)
            if set(payload) != {"batch_id", "decisions"} or not isinstance(payload["batch_id"], str) or not isinstance(payload["decisions"], list):
                raise LedgerError("决定文件必须包含 batch_id 和 decisions 列表")
            if any(not isinstance(d, dict) or set(d) != {"id", "decision"} or not isinstance(d["id"], str)
                   or not isinstance(d["decision"], str) or d["decision"] not in DECISIONS for d in payload["decisions"]):
                raise LedgerError("决定只能为 use_a/use_b/different/uncertain，不能编造新名称")
            answers = {d["id"]: d["decision"] for d in payload["decisions"]}
            if len(answers) != len(payload["decisions"]):
                raise LedgerError("同一名称对不能重复给出决定")
            previous = state["batches"].get(payload["batch_id"])
            if previous is not None:
                if previous != answers:
                    raise LedgerError("已接受批次的决定不可改写；重新运行可创建独立任务")
            else:
                batch = _batch(state)
                if not batch or payload["batch_id"] != batch["batch_id"] or set(answers) != {p["id"] for p in batch["pairs"]}:
                    raise LedgerError("批次不一致或未完整回答本批名称对")
                state["decisions"].update(answers)
                state["batches"][payload["batch_id"]] = answers
                outputs = _prepared(state)
                _save(path, state)
        if _batch(state):
            return _handoff(path, state)
        progress(2, "completed")
        return _complete_or_recover(path, state, outputs, progress)
