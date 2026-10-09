"""The benchmark harness: a synthetic corpus, latency measurement, and an honest report.

Runnable as ``python -m excel_rag.bench``. It builds a workbook corpus programmatically (no parser
involved -- the retrieval layer is what is under test), indexes it into the in-memory double, and
measures per-endpoint latency and Elasticsearch call counts.

Every number it prints is **milliseconds on the in-memory client**, which is a Python dict, not
Elasticsearch. It is not a production latency and the report says so.
"""

from __future__ import annotations

__all__ = ["Corpus", "build_corpus"]
