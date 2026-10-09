"""``python -m excel_rag.bench`` -- run the benchmark and print the report."""

from __future__ import annotations

import sys

from .runner import main

if __name__ == "__main__":
    sys.exit(main())
