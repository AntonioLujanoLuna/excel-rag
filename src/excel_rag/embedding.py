"""Embedders: text in, one dense vector per text out.

The default is `lightonai/mDenseOn <https://huggingface.co/lightonai/mDenseOn>`_, a 307M-parameter
multilingual (English, French, German, Italian, Spanish, Portuguese, Swedish, Norwegian, Arabic)
dense retriever: 768 dimensions, cosine similarity, ``[CLS]`` pooling, and asymmetric prompts --
``query: `` for the search text, ``document: `` for what is indexed. A query encoded with the
document prompt (or the reverse) lands in the wrong region of the space, so the two entry points
here are separate methods rather than one ``encode`` with a flag a caller can forget.

Two implementations satisfy :class:`Embedder`:

* :class:`SentenceTransformerEmbedder` runs a sentence-transformers model in-process. It needs the
  ``embed`` extra (``sentence-transformers`` and ``torch``), imported lazily so the package imports
  without them, and loads the weights on first use, once, behind a lock -- the service handles
  requests on a thread pool.
* :class:`HashingEmbedder` is deterministic feature hashing over word tokens. It is the embedding
  counterpart of the in-memory Elasticsearch double: no weights, no download, stable across runs,
  so tests exercise the whole vector path. It carries no semantics beyond shared words and must not
  be mistaken for a model; its ``model_name`` says so.

Vectors are L2-normalised, so cosine and dot product agree, and every vector records the model
that produced it (``ChunkDocument.embedding_model``): a query embedded by a different model is a
miss, never a silent comparison across incomparable spaces.
"""

from __future__ import annotations

import hashlib
import importlib
import math
import re
import threading
from collections.abc import Sequence
from typing import Any, Protocol

from .settings import EmbeddingSettings

#: The prompts mDenseOn was trained with (``config_sentence_transformers.json``).
QUERY_PROMPT_NAME = "query"
DOCUMENT_PROMPT_NAME = "document"

_TOKEN_RE = re.compile(r"\w+", re.UNICODE)


class Embedder(Protocol):
    """What ingestion and retrieval need from an embedding model."""

    @property
    def model_name(self) -> str: ...

    @property
    def dims(self) -> int: ...

    def embed_documents(self, texts: Sequence[str]) -> list[tuple[float, ...]]: ...

    def embed_query(self, text: str) -> tuple[float, ...]: ...


class EmbeddingError(RuntimeError):
    """The embedder cannot run as configured (missing extra, wrong dimensionality)."""


class SentenceTransformerEmbedder:
    """A sentence-transformers model, run in-process, with its query/document prompts.

    ``dims`` is the dimensionality the index was created with; the model's own output size is
    checked against it on load, because a mismatch would be refused by the cluster at write time
    (or, against the in-memory double, silently never match).
    """

    def __init__(
        self,
        model_name: str,
        *,
        dims: int,
        batch_size: int = 32,
        max_seq_length: int | None = None,
        device: str | None = None,
        query_prompt_name: str = QUERY_PROMPT_NAME,
        document_prompt_name: str = DOCUMENT_PROMPT_NAME,
    ) -> None:
        self._model_name = model_name
        self._dims = dims
        self._batch_size = batch_size
        self._max_seq_length = max_seq_length
        self._device = device
        self._query_prompt_name = query_prompt_name
        self._document_prompt_name = document_prompt_name
        self._model: Any | None = None
        self._lock = threading.Lock()

    @property
    def model_name(self) -> str:
        return self._model_name

    @property
    def dims(self) -> int:
        return self._dims

    def _loaded(self) -> Any:
        if self._model is not None:
            return self._model
        with self._lock:
            if self._model is None:
                self._model = self._load()
        return self._model

    def _load(self) -> Any:
        try:
            module = importlib.import_module("sentence_transformers")
        except ModuleNotFoundError as exc:
            raise EmbeddingError(
                "the sentence-transformers embedder needs the 'embed' extra: install "
                "excel-rag[embed], or set EXCEL_RAG_EMBEDDING__PROVIDER=none to index and search "
                "without vectors"
            ) from exc
        model: Any = module.SentenceTransformer(self._model_name, device=self._device)
        if self._max_seq_length is not None:
            model.max_seq_length = self._max_seq_length
        prompts = getattr(model, "prompts", {}) or {}
        for name in (self._query_prompt_name, self._document_prompt_name):
            if name not in prompts:
                raise EmbeddingError(
                    f"model {self._model_name!r} defines no {name!r} prompt "
                    f"(it has {sorted(prompts)}); an asymmetric retriever needs both"
                )
        produced = _model_dims(model)
        if produced is not None and produced != self._dims:
            raise EmbeddingError(
                f"model {self._model_name!r} produces {produced}-dimensional vectors but the "
                f"index is configured for {self._dims} (EXCEL_RAG_EMBEDDING__DIMS)"
            )
        return model

    def _encode(self, texts: Sequence[str], prompt_name: str) -> list[tuple[float, ...]]:
        if not texts:
            return []
        vectors = self._loaded().encode(
            list(texts),
            prompt_name=prompt_name,
            batch_size=self._batch_size,
            normalize_embeddings=True,
            convert_to_numpy=True,
            show_progress_bar=False,
        )
        return [tuple(float(value) for value in row) for row in vectors]

    def embed_documents(self, texts: Sequence[str]) -> list[tuple[float, ...]]:
        return self._encode(texts, self._document_prompt_name)

    def embed_query(self, text: str) -> tuple[float, ...]:
        return self._encode([text], self._query_prompt_name)[0]


def _model_dims(model: Any) -> int | None:
    for accessor in ("get_embedding_dimension", "get_sentence_embedding_dimension"):
        method = getattr(model, accessor, None)
        if callable(method):
            value = method()
            return int(value) if value is not None else None
    return None


class HashingEmbedder:
    """Deterministic feature hashing: each lowercased word token adds a signed unit to one bucket.

    Texts that share words land close together, which is all a test of the vector path needs. The
    last bucket always carries a small constant so no vector is all zeros -- Elasticsearch refuses
    a zero-magnitude vector under cosine similarity.
    """

    def __init__(self, dims: int, *, model_name: str = "hashing-test-double") -> None:
        if dims < 2:
            raise ValueError("a hashing embedder needs at least two dimensions")
        self._dims = dims
        self._model_name = model_name

    @property
    def model_name(self) -> str:
        return self._model_name

    @property
    def dims(self) -> int:
        return self._dims

    def _vector(self, text: str) -> tuple[float, ...]:
        buckets = [0.0] * self._dims
        for token in _TOKEN_RE.findall(text.lower()):
            digest = hashlib.blake2b(token.encode("utf-8"), digest_size=8).digest()
            value = int.from_bytes(digest, "big")
            index = value % (self._dims - 1)
            buckets[index] += 1.0 if (value >> 63) & 1 else -1.0
        buckets[-1] = 0.1
        norm = math.sqrt(sum(value * value for value in buckets))
        return tuple(value / norm for value in buckets)

    def embed_documents(self, texts: Sequence[str]) -> list[tuple[float, ...]]:
        return [self._vector(text) for text in texts]

    def embed_query(self, text: str) -> tuple[float, ...]:
        return self._vector(text)


def build_embedder(settings: EmbeddingSettings) -> Embedder | None:
    """The embedder the settings name, or ``None`` when vectors are switched off."""
    if settings.provider == "none":
        return None
    if settings.provider == "hashing":
        # Never labelled with `settings.model`: its vectors are not that model's.
        return HashingEmbedder(settings.dims)
    return SentenceTransformerEmbedder(
        settings.model,
        dims=settings.dims,
        batch_size=settings.batch_size,
        max_seq_length=settings.max_seq_length,
        device=settings.device,
    )


def chunk_text(title: str, content: str) -> str:
    """What a chunk is embedded as: its title, then its content."""
    return f"{title}\n{content}" if title else content


__all__ = [
    "DOCUMENT_PROMPT_NAME",
    "QUERY_PROMPT_NAME",
    "Embedder",
    "EmbeddingError",
    "HashingEmbedder",
    "SentenceTransformerEmbedder",
    "build_embedder",
    "chunk_text",
]
