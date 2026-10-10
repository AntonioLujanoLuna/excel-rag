"""The ``excel-rag`` command line: index a workbook, inspect or render it, or serve the API.

``serve`` deliberately names the application by **string** (``"excel_rag.app:create_app"``) so this
module never imports the retrieval workstream: ingestion ships and runs on its own, and uvicorn
resolves the factory only when a server is actually started.
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .ingest import IngestedWorkbook, IngestError, build_client, ingest_workbook
from .ingest.indexer import Indexer
from .settings import Settings


def inspect_workbook(path: str | Path) -> dict[str, Any]:
    """A JSON-ready description of what ingestion detects: sheets, regions, tables and counts."""
    ingested = ingest_workbook(path, workbook_id="inspect", version=1)
    sheets: list[dict[str, Any]] = []
    regions: list[dict[str, Any]] = []
    for sheet in ingested.model.sheets:
        sheets.append(
            {
                "name": sheet.name,
                "visibility": sheet.visibility,
                "is_macro_sheet": sheet.is_macro_sheet,
                "used_range": sheet.a1_range.a1 if sheet.a1_range else None,
                "regions": len(sheet.regions),
                "tables": list(sheet.table_names),
                "formulas": len(sheet.formulas),
                "declared_dimension_flagged": sheet.declared_dimension_flagged,
            }
        )
        for region in sheet.regions:
            regions.append(
                {
                    "sheet": sheet.name,
                    "kind": region.kind.value,
                    "a1_range": region.a1_range.a1,
                    "node_id": region.node_id,
                    "title": region.title,
                    "table_name": region.table_name,
                    "columns": [column.name for column in region.columns],
                    "row_groups": len(region.row_groups),
                }
            )
    return {
        "workbook": ingested.model.source_file,
        "workbook_id": ingested.model.workbook_id,
        "sheets": sheets,
        "regions": regions,
        "named_ranges": [name.label for name in ingested.model.named_ranges],
        "macro_sheets": list(ingested.model.macro_sheet_names),
        "has_vba": ingested.model.has_vba,
        "documents": {
            "chunks": len(ingested.chunks),
            "structure": len(ingested.structure),
            "total": len(ingested.chunks) + len(ingested.structure),
        },
    }


def _index(args: argparse.Namespace) -> int:
    settings = Settings()
    indexer = Indexer(build_client(settings), settings)
    for path in args.paths:
        ingested: IngestedWorkbook = ingest_workbook(
            path,
            workbook_id=args.workbook_id,
            version=args.version,
            acl_scope=tuple(args.acl or ()),
        )
        result = indexer.index_workbook(ingested)
        print(json.dumps(asdict(result), indent=2, sort_keys=True))
    return 0


def _render(args: argparse.Namespace) -> int:
    from .context import render_workbook

    rendered = render_workbook(args.path, token_budget=args.budget, tools_hint=args.tools_hint)
    sys.stdout.write(rendered.text)
    print(
        f"excel-rag: ~{rendered.tokens} tokens, detail {rendered.detail}, "
        f"{'complete' if rendered.complete else 'partial'}",
        file=sys.stderr,
    )
    return 0


def _diff(args: argparse.Namespace) -> int:
    from .workbook.diff import diff_files, format_diff

    diff = diff_files(args.before, args.after, max_changes=args.max_changes)
    if args.json:
        payload = asdict(diff)
        payload["total"] = diff.total
        print(json.dumps(payload, indent=2, sort_keys=True, default=str))
    else:
        print(format_diff(diff))
    return 0


def _calc(args: argparse.Namespace) -> int:
    from .workbook import load_workbook
    from .workbook.calc import (
        CalcInputError,
        CalcUnavailable,
        calculate,
        check_saved_values,
        format_calculation,
        format_check,
    )

    changes: dict[str, str] = {}
    for item in args.set:
        cell, sep, value = item.partition("=")
        if not sep or not cell.strip():
            print(f"excel-rag calc: --set wants CELL=VALUE, got {item!r}", file=sys.stderr)
            return 2
        changes[cell.strip()] = value
    if not args.targets and not args.check:
        print("excel-rag calc: name the cells to calculate, or pass --check", file=sys.stderr)
        return 2
    try:
        model = load_workbook(args.path)
        if args.check:
            calculation = check_saved_values(model, timeout_seconds=args.timeout)
        else:
            calculation = calculate(
                model,
                args.targets,
                changes,
                recalculate_all=args.all,
                timeout_seconds=args.timeout,
            )
    except CalcUnavailable as error:
        print(f"excel-rag calc: {error}", file=sys.stderr)
        return 1
    except CalcInputError as error:
        print(f"excel-rag calc: {error}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(asdict(calculation), indent=2, sort_keys=True, default=str))
    elif args.check:
        print(format_check(calculation))
    else:
        print(format_calculation(calculation))
    return 3 if args.check and calculation.mismatches else 0


def _mcp(args: argparse.Namespace) -> int:
    try:
        from .mcp_server import run
    except ModuleNotFoundError as error:  # pragma: no cover - only without the extra
        print(f"excel-rag: the MCP server needs the 'mcp' extra ({error})", file=sys.stderr)
        return 1
    run(args.root or ["."], search=args.search)
    return 0


def _serve(args: argparse.Namespace) -> int:
    import uvicorn

    settings = Settings()
    host = args.host or settings.server.host
    port = args.port or settings.server.port
    uvicorn.run("excel_rag.app:create_app", factory=True, host=host, port=port)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excel-rag", description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    index_parser = subparsers.add_parser("index", help="ingest one or more workbooks")
    index_parser.add_argument("paths", nargs="+", help="paths to .xlsx/.xlsm/.xlsb/.csv files")
    index_parser.add_argument("--workbook-id", required=True)
    index_parser.add_argument("--version", type=int, required=True)
    index_parser.add_argument("--acl", action="append", default=[], help="repeatable ACL scope")

    inspect_parser = subparsers.add_parser("inspect", help="print detected regions and counts")
    inspect_parser.add_argument("path", help="path to an .xlsx/.xlsm/.xlsb/.csv file")

    render_parser = subparsers.add_parser(
        "render", help="print a workbook as text for a conversation's context window"
    )
    render_parser.add_argument("path", help="path to an .xlsx/.xlsm/.xlsb/.csv file")
    render_parser.add_argument(
        "--budget", type=int, default=8_000, help="token budget (estimated at 3 chars/token)"
    )
    render_parser.add_argument(
        "--tools-hint",
        action="store_true",
        help="tell the model omitted ranges can be read with the workbook tools",
    )

    diff_parser = subparsers.add_parser(
        "diff", help="what changed between two versions of a workbook (cells, formulas, impact)"
    )
    diff_parser.add_argument("before", help="the earlier .xlsx/.xlsm/.xlsb/.csv")
    diff_parser.add_argument("after", help="the later .xlsx/.xlsm/.xlsb/.csv")
    diff_parser.add_argument("--json", action="store_true", help="print the diff as JSON")
    diff_parser.add_argument(
        "--max-changes", type=int, default=500, help="list at most this many cell changes"
    )

    calc_parser = subparsers.add_parser(
        "calc",
        help="compute cells, optionally with changed inputs (what-if); needs the calc extra",
    )
    calc_parser.add_argument("path", help="path to an .xlsx/.xlsm/.xlsb/.csv file")
    calc_parser.add_argument(
        "targets", nargs="*", help="cells or ranges to compute, e.g. 'Forecast!E4:E7'"
    )
    calc_parser.add_argument(
        "--set",
        action="append",
        default=[],
        metavar="CELL=VALUE",
        help="change an input before computing, e.g. Assumptions!B4=7%%; repeatable",
    )
    calc_parser.add_argument(
        "--all",
        action="store_true",
        help="recompute every formula the targets need, not only those a change reaches",
    )
    calc_parser.add_argument(
        "--check",
        action="store_true",
        help="recompute every formula and compare with Excel's saved values (exit 3 on a "
        "difference)",
    )
    calc_parser.add_argument("--json", action="store_true", help="print the result as JSON")
    calc_parser.add_argument(
        "--timeout", type=float, default=30.0, help="wall-clock budget in seconds"
    )

    subparsers.add_parser(
        "evaluate",
        help="measure retrieval quality (hit@k, recall@k, MRR); see `excel-rag evaluate --help`",
        add_help=False,
    )

    mcp_parser = subparsers.add_parser(
        "mcp", help="serve the workbook tools to an MCP client over stdio (needs the mcp extra)"
    )
    mcp_parser.add_argument(
        "--root",
        action="append",
        help="a directory the tools may read workbooks from, repeatable (default: .)",
    )
    mcp_parser.add_argument(
        "--search",
        action="store_true",
        help="also serve search_index over the configured Elasticsearch",
    )

    serve_parser = subparsers.add_parser("serve", help="run the retrieval API")
    serve_parser.add_argument("--host")
    serve_parser.add_argument("--port", type=int)
    return parser


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments[:1] == ["evaluate"]:
        from .evaluate.__main__ import main as evaluate_main

        return evaluate_main(arguments[1:])
    args = build_parser().parse_args(arguments)
    try:
        if args.command == "index":
            return _index(args)
        if args.command == "inspect":
            print(json.dumps(inspect_workbook(args.path), indent=2, sort_keys=True))
            return 0
        if args.command == "render":
            return _render(args)
        if args.command == "diff":
            return _diff(args)
        if args.command == "calc":
            return _calc(args)
        if args.command == "mcp":
            return _mcp(args)
        if args.command == "serve":
            return _serve(args)
    except IngestError as error:
        print(f"excel-rag: {error}", file=sys.stderr)
        return 1
    return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
