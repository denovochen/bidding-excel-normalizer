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
from ledger_core.contract import LedgerError, MappingRevisionRequired, RecoverableWorkbookError
from ledger_core.normalize import build_outputs
from ledger_core.workbook import describe_source_failure, inspect_workbooks, load_json, read_workbook


PROGRESS_TITLES = ("读取并检查 Excel", "清洗并整理数据", "生成并校验结果")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="招投标 Excel 结构识别辅助与确定性清洗")
    commands = parser.add_subparsers(dest="command", required=True)
    run = commands.add_parser("run", help="整理上传文件；首次扫描后由 Agent 轻量审阅结构，再用 --plan 完成")
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
        if args.command in {"validate", "companies"}:
            checked = validate_outputs(args.output)
            result = {"kind": "validation", **checked}
            if args.command == "companies":
                ledger = json.loads((args.output / "ledger.json").read_text(encoding="utf-8"))
                result = {"kind": "companies", "companies": ledger["unique_companies"],
                          "company_count": len(ledger["unique_companies"])}
        else:
            progress(1, "in_progress")
            books, source_failures = [], []
            for path in args.inputs:
                try:
                    books.append(read_workbook(path))
                except RecoverableWorkbookError as exc:
                    source_failures.append(describe_source_failure(path, str(exc)))
            progress(1, "completed")
            if args.command == "inspect":
                result = inspect_workbooks(books, source_failures)
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
                    inspection = inspect_workbooks(books, source_failures)
                    plan = inspection["suggested_plan"]
                    if books:
                        print(json.dumps({"kind": "mapping_required", "message": "Agent 应轻量审阅每份文件及各结构区域后，用 run --plan 继续；不向用户补问缺失业务值。",
                                          "inspection": inspection}, ensure_ascii=False), flush=True)
                        return 3
                prepared = build_outputs(books, plan, source_failures=source_failures)
                progress(2, "completed")
                progress(3, "in_progress")
                output = args.output or Path.cwd() / "outputs" / ("excel-ledger-" + uuid.uuid4().hex)
                result = publish(output, *prepared)
                progress(3, "completed")
        print(json.dumps(result, ensure_ascii=False, allow_nan=False), flush=True)
        return 0
    except MappingRevisionRequired as exc:
        print(json.dumps({"kind": "mapping_required", "message": str(exc), "evidence": exc.evidence},
                         ensure_ascii=False), flush=True)
        return 3
    except (LedgerError, OSError, ValueError, TypeError, KeyError, csv.Error) as exc:
        if active_step:
            progress(active_step, "failed")
        message = str(exc) if isinstance(exc, LedgerError) else f"输入或产物无效 ({type(exc).__name__})"
        print(json.dumps({"kind": "error", "message": message}, ensure_ascii=False), file=sys.stderr, flush=True)
        return 2

if __name__ == "__main__":
    sys.exit(main())
