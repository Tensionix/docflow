from __future__ import annotations

import argparse
from pathlib import Path
import sys

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from docx_table_unifier_model import parse_docx_tables


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect DOCX tables and print signatures.")
    parser.add_argument("--input", required=True, help="Path to DOCX file")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = Path(args.input).resolve()
    tables = parse_docx_tables(source)
    if not tables:
        print("[ERROR] No tables were found.")
        return 1
    print(f"[INFO] Tables found: {len(tables)}")
    for table in tables:
        preview = table.preview_text() or "<empty>"
        print(f"[TABLE {table.table_index:02d}] width={table.width} height={table.height} merged={table.merged_cells} continuation_slots={table.continuation_slots} preview={preview}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
