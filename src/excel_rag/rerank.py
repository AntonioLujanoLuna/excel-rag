"""Rerankers: score (query, chunk) pairs jointly, after fusion, to reorder the top of the list.

Fusion orders candidates by where two independent retrievers ranked them; neither retriever ever
reads the query and the chunk together. A cross-encoder does, which is why reranking the top few
dozen fused candidates is the usual cheap gain on table and column descriptions -- a chunk whose
words match but whose meaning does not (``Revenue`` in a cost table's notes) drops, and a chunk
that answers the question in other words rises.

Two implementations satisfy :class:`Reranker`:

* :class:`CrossEncoderReranker` runs a sentence-transformers ``CrossEncoder`` in-process. The
  default, ``cross-encoder/mmarco-mMiniLMv2-L12-H384-v1``, is multilingual like the default
  embedder, so a Spanish question still reranks an English workbook. It needs the ``embed`` extra,
  imported lazily, and loads the weights on first use, once, behind a lock.
* :class:`OverlapReranker` scores by the fraction of query tokens a chunk contains. It is the
  reranking counterpart of the hashing embedder -- deterministic, no weights -- so tests exercise
  the whole path; its ``model_name`` says it is not a model.

A reranked hit's ``score_kind`` is ``rerank``: its score is the cross-encoder's, not comparable with
a fused or a lexical one.
"""

from __future__ import annotations

import importlib
import re
import threading
from collections.abc import Sequence
from typing import Any, Protocol

from .embedding import EmbeddingError
from .settings import RerankSettings

_TOKEN_RE = re.compile(r"\w+", re.UNICODE)


class Reranker(Protocol):
    """What retrieval needs from a reranking model."""

    @property
    def model_name(self) -> str: ...

    def score(self, query: str, texts: Sequence[str]) -> list[float]: ...


class CrossEncoderReranker:
    """A sentence-transformers ``CrossEncoder``, run in-process."""

    def __init__(
        self,
        model_name: str,
        *,
        batch_size: int = 32,
        max_length: int | None = None,
        device: str | None = None,
    ) -> None:
        self._model_name = model_name
        self._batch_size = batch_size
        self._max_length = max_length
        self._device = device
        self._model: Any | None = None
        self._lock = threading.Lock()

    @property
    def model_name(self) -> str:
        return self._model_name

    def _loaded(self) -> Any:
        if self._model is not None:
            return self._model
        with self._lock:
            if self._model is None:
                try:
                    module = importlib.import_module("sentence_transformers")
                except ModuleNotFoundError as exc:
                    raise EmbeddingError(
                        "the cross-encoder reranker needs the 'embed' extra: install "
                        "excel-rag[embed], or set EXCEL_RAG_RERANK__PROVIDER=none"
                    ) from exc
                self._model = module.CrossEncoder(
                    self._model_name, max_length=self._max_length, device=self._device
                )
        return self._model

    def score(self, query: str, texts: Sequence[str]) -> list[float]:
        if not texts:
            return []
        scores = self._loaded().predict(
            [(query, text) for text in texts],
            batch_size=self._batch_size,
            show_progress_bar=False,
        )
        return [float(value) for value in scores]


class OverlapReranker:
    """The fraction of distinct query tokens a text contains. A test double, not a model."""

    model_name = "overlap-test-double"

    def score(self, query: str, texts: Sequence[str]) -> list[float]:
        wanted = set(_TOKEN_RE.findall(query.lower()))
        if not wanted:
            return [0.0 for _ in texts]
        return [len(wanted & set(_TOKEN_RE.findall(text.lower()))) / len(wanted) for text in texts]


def build_reranker(settings: RerankSettings) -> Reranker | None:
    """The reranker the settings name, or ``None`` when reranking is off (the default)."""
    if settings.provider == "none":
        return None
    if settings.provider == "overlap":
        return OverlapReranker()
    return CrossEncoderReranker(
        settings.model,
        batch_size=settings.batch_size,
        max_length=settings.max_length,
        device=settings.device,
    )


__all__ = ["CrossEncoderReranker", "OverlapReranker", "Reranker", "build_reranker"]
