#!/usr/bin/env python3
"""招投标 Excel 检查、导出与验证入口。 @author denovochen"""
from __future__ import annotations

import argparse
import csv
import json
import sys
import uuid
from pathlib import Path

from ledger_core.artifacts import publish, validate_outputs
from ledger_core.contract import LedgerError
from ledger_core.normalize import build_outputs
from ledger_core.workbook import inspect_workbooks, load_json, read_workbook
from ledger_core.workflow import resolve, start as start_run


PROGRESS_TITLES = ("读取并检查 Excel", "清洗并整理数据", "生成并校验结果", "交付结果文件")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="招投标 Excel 结构识别辅助与确定性清洗")
    commands = parser.add_subparsers(dest="command", required=True)
    run = commands.add_parser("run", help="直接整理上传文件；已识别结构自动完成，未知结构交由 Agent 内部映射")
    run.add_argument("inputs", type=Path, nargs="+")
    run.add_argument("--plan", type=Path, help="Agent 内部准备的可选结构映射")
    run.add_argument("--output", type=Path, help="省略时创建 outputs/excel-ledger-<唯一ID>")
    run.add_argument("--progress", action="store_true", help="可选：输出脚本实际阶段日志，不规定宿主进度计划")
    inspect = commands.add_parser("inspect", help="只读展示结构和建议映射，输出 JSON 到 stdout")
    inspect.add_argument("inputs", type=Path, nargs="+")
    inspect.add_argument("--sheet", help="仅在 profile 中展示指定 Sheet 的局部行，不改变完整映射建议")
    inspect.add_argument("--rows", help="局部查看的闭区间，如 10:20；须与 --sheet 一起使用")
    export = commands.add_parser("export", help="按已核对的 JSON 映射整理并发布三个文件")
    export.add_argument("inputs", type=Path, nargs="+")
    export.add_argument("--plan", type=Path, required=True)
    export.add_argument("--output", type=Path, required=True)
    export.add_argument("--progress", action="store_true", help="逐行输出真实阶段进度")
    validate = commands.add_parser("validate", help="校验现有三个产物的结构与关联")
    validate.add_argument("output", type=Path)
    companies = commands.add_parser("companies", help="只读输出全批次去重企业名单 JSON，供后续 Gateway 采集编排使用")
    companies.add_argument("output", type=Path)
    resolution = commands.add_parser("resolve", help="接受宿主模型的本批双名称决定，继续到下一批或交付")
    resolution.add_argument("--state", required=True, type=Path)
    resolution.add_argument("--decisions", required=True, type=Path)
    resolution.add_argument("--progress", action="store_true")
    resume = commands.add_parser("resume", help="从任务状态恢复未完成的名称判断/发布，不重新创建任务")
    resume.add_argument("--state", required=True, type=Path)
    resume.add_argument("--progress", action="store_true")
    args = parser.parse_args(argv)
    active_step = 0

    def progress(step: int, status: str) -> None:
        nonlocal active_step
        if not getattr(args, "progress", False):
            return
        active_step = step if status == "in_progress" else 0
        print(json.dumps({"kind": "progress", "step": step, "status": status,
                          "title": PROGRESS_TITLES[step - 1]}, ensure_ascii=False), flush=True)

    try:
        if args.command in {"resolve", "resume"}:
            result = resolve(args.state, getattr(args, "decisions", None), progress)
        elif args.command in {"validate", "companies"}:
            checked = validate_outputs(args.output)
            result = {"kind": "validation", **checked}
            if args.command == "companies":
                ledger = json.loads((args.output / "ledger.json").read_text(encoding="utf-8"))
                result = {"kind": "companies", "companies": ledger["unique_companies"],
                          "company_count": len(ledger["unique_companies"])}
        else:
            progress(1, "in_progress")
            books = [read_workbook(path) for path in args.inputs]
            progress(1, "completed")
            if args.command == "inspect":
                result = inspect_workbooks(books)
                if bool(args.sheet) != bool(args.rows):
                    raise LedgerError("--sheet 与 --rows 必须同时指定")
                if args.rows:
                    start, end = [int(part) for part in args.rows.split(":")]
                    if start < 1 or end < start or end - start >= 100:
                        raise LedgerError("局部查看最多 100 行")
                    views = [{"file_name": b.path.name, "sheet": s.name,
                              "rows": [s.row_view(r) for r in range(start, min(end, s.max_row) + 1)]}
                             for b in books for s in b.sheets if s.name == args.sheet]
                    if not views:
                        raise LedgerError("指定 Sheet 不存在")
                    result = {"kind": "inspection_detail", "views": views}
            else:
                progress(2, "in_progress")
                if args.plan:
                    plan = load_json(args.plan)
                else:
                    inspection = inspect_workbooks(books)
                    plan = inspection["suggested_plan"]
                    if any(s["action"] == "needs_mapping" for source in plan["sources"] for s in source["sheets"]):
                        print(json.dumps({"kind": "mapping_required", "message": "Agent 应在内部完成字段映射后用 run --plan 继续，不询问用户补充缺失字段。",
                                          "inspection": inspection}, ensure_ascii=False), flush=True)
                        return 3
                prepared = build_outputs(books, plan)
                output = args.output or Path.cwd() / "outputs" / ("excel-ledger-" + uuid.uuid4().hex)
                result = start_run(output, books, plan, prepared, progress)
        print(json.dumps(result, ensure_ascii=False, allow_nan=False), flush=True)
        return 4 if result["kind"] == "name_review_required" else 0
    except (LedgerError, OSError, ValueError, TypeError, KeyError, csv.Error) as exc:
        if active_step:
            progress(active_step, "failed")
        message = str(exc) if isinstance(exc, LedgerError) else f"输入或产物无效 ({type(exc).__name__})"
        print(json.dumps({"kind": "error", "message": message}, ensure_ascii=False), file=sys.stderr, flush=True)
        return 2

if __name__ == "__main__":
    sys.exit(main())
