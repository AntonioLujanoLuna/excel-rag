"""The ``excel-rag`` command line: index a workbook, inspect what ingestion detects, or serve.

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
        "named_ranges": [name.name for name in ingested.model.named_ranges],
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
    index_parser.add_argument("paths", nargs="+", help="paths to .xlsx/.xlsm files")
    index_parser.add_argument("--workbook-id", required=True)
    index_parser.add_argument("--version", type=int, required=True)
    index_parser.add_argument("--acl", action="append", default=[], help="repeatable ACL scope")

    inspect_parser = subparsers.add_parser("inspect", help="print detected regions and counts")
    inspect_parser.add_argument("path", help="path to an .xlsx/.xlsm file")

    serve_parser = subparsers.add_parser("serve", help="run the retrieval API")
    serve_parser.add_argument("--host")
    serve_parser.add_argument("--port", type=int)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "index":
            return _index(args)
        if args.command == "inspect":
            print(json.dumps(inspect_workbook(args.path), indent=2, sort_keys=True))
            return 0
        if args.command == "serve":
            return _serve(args)
    except IngestError as error:
        print(f"excel-rag: {error}", file=sys.stderr)
        return 1
    return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
