#!/usr/bin/env python3
"""招投标 Excel 检查、导出与验证入口。 @author denovochen"""
from __future__ import annotations

import argparse
import csv
import importlib
import importlib.metadata
import json
import re
import sys
import uuid
from pathlib import Path

from ledger_core.artifacts import publish, validate_outputs
from ledger_core.contract import LedgerError, MappingRevisionRequired, RecoverableWorkbookError
from ledger_core.normalize import build_outputs
from ledger_core.review import (apply_answers, cleanup_state, create_state, load_answers, load_state,
                                pending_questions, review_result, save_state, split_decisions,
                                create_work_directory, cleanup_work_directory)
from ledger_core.workbook import (apply_plan_patch, compact_inspection, describe_source_failure, inspect_workbooks,
                                  load_json, read_workbook)


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
    inspect.add_argument("--save", type=Path, help="将完整 inspection JSON 保存到内部工作目录")
    inspect.add_argument("--compact", action="store_true", help="stdout 仅输出有界摘要；建议与 --save 一起使用")
    plan_command = commands.add_parser("plan", help="将有界 plan patch 应用到完整 inspection 并生成 plan.json")
    plan_command.add_argument("--inspection", type=Path, required=True)
    plan_command.add_argument("--patch", type=Path, required=True)
    plan_command.add_argument("--output", type=Path, required=True)
    doctor = commands.add_parser("doctor", help="只读检查 Python 和 Excel 依赖，不安装依赖")
    export = commands.add_parser("export", help="按已核对的 JSON 映射整理并发布三个文件")
    export.add_argument("inputs", type=Path, nargs="+")
    export.add_argument("--plan", type=Path, required=True)
    export.add_argument("--output", type=Path, required=True)
    export.add_argument("--progress", action="store_true", help="逐行输出真实阶段进度")
    validate = commands.add_parser("validate", help="校验现有三个产物的结构与关联")
    validate.add_argument("output", type=Path)
    companies = commands.add_parser("companies", help="只读输出全批次去重企业名单 JSON，供后续 Gateway 采集编排使用")
    companies.add_argument("output", type=Path)
    resolve = commands.add_parser("resolve", help="接受项目关系或中标企业复核选择并继续处理")
    resolve.add_argument("--state", type=Path, required=True)
    resolve.add_argument("--answers", type=Path, required=True)
    resolve.add_argument("--progress", action="store_true", help="输出恢复、重建和发布阶段日志")
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
        if args.command == "doctor":
            required = {"openpyxl": "3.1.5", "xlrd": "2.0.2"}
            packages = {}
            for name, minimum in required.items():
                try:
                    version = importlib.metadata.version(name)
                    importlib.import_module(name)
                    current = tuple(int(value) for value in re.findall(r"\d+", version)[:3])
                    required_version = tuple(int(value) for value in minimum.split("."))
                    compatible = current >= required_version
                    packages[name] = {"version": version, "minimum": minimum,
                                      "available": True, "compatible": compatible}
                except (importlib.metadata.PackageNotFoundError, ImportError):
                    packages[name] = {"version": None, "minimum": minimum,
                                      "available": False, "compatible": False}
            result = {"kind": "environment", "python": sys.executable, "packages": packages,
                      "ready": all(item["compatible"] for item in packages.values())}
            if not result["ready"]:
                print(json.dumps(result, ensure_ascii=False), file=sys.stderr, flush=True)
                return 2
        elif args.command == "plan":
            inspection = load_json(args.inspection.expanduser().resolve(strict=True))
            patch = load_json(args.patch.expanduser().resolve(strict=True))
            plan = apply_plan_patch(inspection, patch)
            args.output.parent.mkdir(parents=True, exist_ok=True)
            if args.output.exists():
                raise LedgerError("plan 输出文件已存在，不会覆盖")
            args.output.write_text(json.dumps(plan, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
            result = {"kind": "plan", "output": str(args.output.absolute()),
                      "source_count": len(plan["sources"]), "relationship_count": len(plan.get("relationships", []))}
        elif args.command in {"validate", "companies"}:
            checked = validate_outputs(args.output)
            result = {"kind": "validation", **checked}
            if args.command == "companies":
                ledger = json.loads((args.output / "ledger.json").read_text(encoding="utf-8"))
                result = {"kind": "companies", "companies": ledger["unique_companies"],
                          "company_count": len(ledger["unique_companies"])}
        elif args.command == "resolve":
            state = load_state(args.state)
            state_path = Path(state["state_path"])
            progress(1, "in_progress")
            books, source_failures = [], []
            for input_path in [Path(value) for value in state["inputs"]]:
                try:
                    books.append(read_workbook(input_path))
                except RecoverableWorkbookError as exc:
                    source_failures.append(describe_source_failure(input_path, str(exc)))
            progress(1, "completed")
            progress(2, "in_progress")
            relation_decisions, award_decisions = split_decisions(state["decisions"])
            prepared = build_outputs(books, state["plan"], generated_at=state["generated_at"],
                                     source_failures=source_failures, award_resolutions=award_decisions,
                                     relation_resolutions=relation_decisions)
            errors = apply_answers(state, prepared[2], load_answers(args.answers))
            save_state(state_path, state)
            relation_decisions, award_decisions = split_decisions(state["decisions"])
            prepared = build_outputs(books, state["plan"], generated_at=state["generated_at"],
                                     source_failures=source_failures, award_resolutions=award_decisions,
                                     relation_resolutions=relation_decisions)
            questions, _, remaining = pending_questions(prepared[2], state["decisions"])
            progress(2, "completed")
            if remaining:
                result = review_result(state_path, prepared[2], state, errors)
            else:
                progress(3, "in_progress")
                result = publish(Path(state["output"]), *prepared)
                try:
                    cleanup_state(state_path)
                except OSError:
                    pass
                progress(3, "completed")
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
                elif args.save:
                    saved = args.save.expanduser().absolute()
                    saved.parent.mkdir(parents=True, exist_ok=True)
                    if saved.exists():
                        raise LedgerError("inspection 保存文件已存在，不会覆盖")
                    saved.write_text(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
                    if args.compact:
                        result = compact_inspection(result, saved)
                elif args.compact:
                    result = compact_inspection(result)
            else:
                progress(2, "in_progress")
                if args.plan:
                    plan = load_json(args.plan)
                else:
                    inspection = inspect_workbooks(books, source_failures)
                    if books:
                        output = args.output or Path.cwd() / "outputs" / ("excel-ledger-" + uuid.uuid4().hex)
                        work_directory = create_work_directory(output)
                        inspection_path = work_directory / "inspection.json"
                        inspection_path.write_text(json.dumps(inspection, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
                                                   encoding="utf-8")
                        print(json.dumps({
                            "kind": "mapping_required",
                            "message": "审阅有界摘要并编写 plan patch；完整 inspection 已保存到内部工作目录。",
                            "inspection": compact_inspection(inspection, inspection_path),
                        }, ensure_ascii=False), flush=True)
                        return 3
                    plan = inspection["suggested_plan"]
                prepared = build_outputs(books, plan, source_failures=source_failures)
                progress(2, "completed")
                output = args.output or Path.cwd() / "outputs" / ("excel-ledger-" + uuid.uuid4().hex)
                if args.plan:
                    plan_parent = args.plan.expanduser().absolute().parent
                    if plan_parent.name.startswith(".excel-ledger-work-") and output.expanduser().absolute().is_relative_to(plan_parent):
                        raise LedgerError("最终输出目录不能位于内部工作目录")
                questions, review_task_count, remaining = pending_questions(prepared[2], {})
                if remaining:
                    plan_parent = args.plan.expanduser().absolute().parent if args.plan else None
                    work_directory = plan_parent if plan_parent and plan_parent.name.startswith(".excel-ledger-work-") else None
                    state_path = create_state(args.inputs, plan, output, prepared[2]["generated_at"], review_task_count,
                                              work_directory=work_directory)
                    state = load_state(state_path)
                    result = review_result(state_path, prepared[2], state)
                else:
                    progress(3, "in_progress")
                    result = publish(output, *prepared)
                    if args.plan:
                        plan_parent = args.plan.expanduser().absolute().parent
                        if plan_parent.name.startswith(".excel-ledger-work-") and plan_parent.exists():
                            cleanup_work_directory(plan_parent)
                    progress(3, "completed")
        print(json.dumps(result, ensure_ascii=False, allow_nan=False), flush=True)
        return 5 if result["kind"] == "relationship_review_required" else 4 if result["kind"] == "award_review_required" else 0
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
