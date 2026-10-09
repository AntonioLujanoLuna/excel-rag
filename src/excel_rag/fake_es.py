"""An in-memory implementation of :class:`~excel_rag.es.ElasticsearchLike`.

Both workstreams test against this, and the service runs against it when
``use_live_elasticsearch`` is false, so a developer needs no cluster to exercise the whole path.

**It supports a declared subset of the Elasticsearch query DSL and refuses the rest loudly.**
Silently accepting a query it does not understand would produce a passing test and a wrong service;
a ``NotImplementedError`` naming the unsupported clause is the only honest behaviour. What is
supported:

``match_all``, ``term``, ``terms``, ``match``, ``range`` (including intersection with
``integer_range`` fields), ``bool`` with ``must`` / ``filter`` / ``should`` / ``must_not`` and
``minimum_should_match``, ``nested``, and a top-level ``knn`` (``field``, ``query_vector``, ``k``,
``num_candidates``, ``filter``) issued on its own.

``knn`` here is **exact**: every filtered document's vector is compared with the query by cosine
similarity and scored ``(1 + cos) / 2``, as Elasticsearch scores a cosine ``dense_vector``. A
cluster's HNSW search is approximate, so this is the *ceiling* of what a cluster returns, and no
recall or latency measured against this double may be reported as the cluster's. A document whose
vector is missing or of another dimensionality is not a candidate.

Deliberately **not** supported: ``knn`` combined with a ``query`` in one request, ``script_score``,
``function_score``, aggregations, and anything requiring a real analyzer. Lexical scores returned by
this double are ordinal (counts of matched leaf clauses), not BM25, and must never be reported as
relevance quality.
"""

from __future__ import annotations

import copy
import math
from collections.abc import Iterable, Mapping, Sequence
from typing import Any

from .es import ElasticsearchLike

_SUPPORTED = frozenset({"match_all", "term", "terms", "match", "range", "bool", "nested"})
_UNSUPPORTED = frozenset(
    {
        "script",
        "script_score",
        "function_score",
        "aggs",
        "aggregations",
        "wildcard",
        "regexp",
        "fuzzy",
        "prefix",
        "match_phrase",
        "more_like_this",
        "rank_feature",
        "dis_max",
        "constant_score",
    }
)


class UnsupportedQueryError(NotImplementedError):
    """Raised for a clause this double does not implement, naming the clause."""


class InMemoryElasticsearch:
    """A dict-backed client. Deterministic ordering: score desc, then ``_id`` asc."""

    def __init__(self) -> None:
        self._store: dict[str, dict[str, dict[str, Any]]] = {}
        self._mappings: dict[str, Mapping[str, Any]] = {}
        self.calls: list[tuple[str, str]] = []

    # -- index management -------------------------------------------------------------------
    def indices_exists(self, index: str) -> bool:
        self.calls.append(("indices_exists", index))
        return index in self._store

    def create_index(self, index: str, mappings: Mapping[str, Any]) -> None:
        self.calls.append(("create_index", index))
        if index in self._store:
            raise ValueError(f"index already exists: {index}")
        self._store[index] = {}
        self._mappings[index] = mappings

    def delete_index(self, index: str) -> None:
        self.calls.append(("delete_index", index))
        self._store.pop(index, None)
        self._mappings.pop(index, None)

    def mapping(self, index: str) -> Mapping[str, Any]:
        return self._mappings.get(index, {})

    def refresh(self, index: str) -> None:
        """Accepted and ignored: this store has no near-real-time lag."""

    # -- documents --------------------------------------------------------------------------
    def index_document(
        self, index: str, document_id: str, document: Mapping[str, Any], *, refresh: bool = False
    ) -> None:
        self.calls.append(("index_document", index))
        self._require(index)
        self._store[index][document_id] = copy.deepcopy(dict(document))

    def bulk_index(
        self,
        index: str,
        documents: Iterable[tuple[str, Mapping[str, Any]]],
        *,
        refresh: bool = False,
    ) -> int:
        self.calls.append(("bulk_index", index))
        self._require(index)
        written = 0
        for document_id, document in documents:
            self._store[index][document_id] = copy.deepcopy(dict(document))
            written += 1
        return written

    def get_document(self, index: str, document_id: str) -> Mapping[str, Any] | None:
        self.calls.append(("get_document", index))
        stored = self._store.get(index, {}).get(document_id)
        return copy.deepcopy(stored) if stored is not None else None

    def mget_documents(
        self, index: str, document_ids: Sequence[str]
    ) -> Sequence[Mapping[str, Any]]:
        """Returns the documents in the order requested, skipping ids that are absent.

        Note for callers: a real ``_mget`` does not apply any ACL filter. Whatever this returns is
        returned by Elasticsearch too, so scope filtering has to happen on the caller's side.
        """
        self.calls.append(("mget_documents", index))
        found: list[Mapping[str, Any]] = []
        for document_id in document_ids:
            stored = self._store.get(index, {}).get(document_id)
            if stored is not None:
                found.append(copy.deepcopy(stored))
        return found

    def delete_by_query(self, index: str, query: Mapping[str, Any]) -> int:
        self.calls.append(("delete_by_query", index))
        self._require(index)
        doomed = [
            document_id
            for document_id, document in self._store[index].items()
            if self._matches(document, query)
        ]
        for document_id in doomed:
            del self._store[index][document_id]
        return len(doomed)

    def count(self, index: str, query: Mapping[str, Any] | None = None) -> int:
        self.calls.append(("count", index))
        self._require(index)
        if query is None:
            return len(self._store.get(index, {}))
        return sum(1 for document in self._store[index].values() if self._matches(document, query))

    # -- search -----------------------------------------------------------------------------
    def search(
        self,
        index: str,
        query: Mapping[str, Any],
        *,
        size: int = 10,
        source_includes: Sequence[str] | None = None,
    ) -> Mapping[str, Any]:
        self.calls.append(("search", index))
        self._require(index)
        knn = query.get("knn")
        if knn is not None and len(query) > 1:
            raise UnsupportedQueryError(
                "knn combined with a query in one request; issue them as separate searches"
            )
        scored: list[tuple[float, str, Mapping[str, Any]]] = []
        for document_id, document in self._store[index].items():
            score = (
                self._score_knn(document, knn) if knn is not None else self._score(document, query)
            )
            if score is not None:
                projected = (
                    {key: document[key] for key in source_includes if key in document}
                    if source_includes is not None
                    else document
                )
                scored.append((score, document_id, copy.deepcopy(projected)))
        scored.sort(key=lambda row: (-row[0], row[1]))
        if knn is not None:
            scored = scored[: int(knn.get("k", size))]
        window = scored[:size]
        return {
            "took": 0,
            "hits": {
                "total": {"value": len(scored), "relation": "eq"},
                "max_score": window[0][0] if window else None,
                "hits": [
                    {"_index": index, "_id": document_id, "_score": score, "_source": source}
                    for score, document_id, source in window
                ],
            },
        }

    # -- internals --------------------------------------------------------------------------
    def _score_knn(self, document: Mapping[str, Any], knn: Mapping[str, Any]) -> float | None:
        field = knn.get("field")
        query_vector = knn.get("query_vector")
        if not isinstance(field, str) or not isinstance(query_vector, list):
            raise UnsupportedQueryError("knn requires {'field': str, 'query_vector': [...], ...}")
        filters = knn.get("filter") or []
        if isinstance(filters, Mapping):
            filters = [filters]
        if not all(self._matches(document, clause) for clause in filters):
            return None
        stored = document.get(field)
        if not isinstance(stored, list) or len(stored) != len(query_vector) or not stored:
            return None
        cosine = _cosine(query_vector, stored)
        return None if cosine is None else (1.0 + cosine) / 2.0

    def _require(self, index: str) -> None:
        if index not in self._store:
            raise ValueError(f"no such index: {index}")

    def _score(self, document: Mapping[str, Any], query: Mapping[str, Any]) -> float | None:
        """None means no match; otherwise the count of matched leaf clauses (ordinal, not BM25)."""
        matched: float = 0
        for clause, body in query.items():
            if clause in _UNSUPPORTED:
                raise UnsupportedQueryError(
                    f"the in-memory client does not implement {clause!r}; supported clauses are "
                    f"{sorted(_SUPPORTED)}. Use a live cluster for this query."
                )
            if clause not in _SUPPORTED:
                raise UnsupportedQueryError(f"unknown or unsupported query clause: {clause!r}")
            if clause == "match_all":
                matched += 1
            elif clause == "bool":
                score = self._score_bool(document, body)
                if score is None:
                    return None
                matched += score
            elif clause == "nested":
                if not self._match_nested(document, body):
                    return None
                matched += 1
            else:
                if not self._match_leaf(document, clause, body):
                    return None
                matched += 1
        return float(matched)

    def _score_bool(self, document: Mapping[str, Any], body: Mapping[str, Any]) -> float | None:
        matched: float = 0
        for clause in body.get("must", []) or []:
            score = self._score(document, clause)
            if score is None:
                return None
            matched += score
        for clause in body.get("filter", []) or []:
            if self._score(document, clause) is None:
                return None
        must_not = body.get("must_not", []) or []
        for clause in must_not:
            if self._score(document, clause) is not None:
                return None
        should = body.get("should", []) or []
        hits = sum(1 for clause in should if self._score(document, clause) is not None)
        minimum = body.get("minimum_should_match")
        if minimum is None:
            # Elasticsearch's own rule: a `should` only has to match when it is the only clause.
            minimum = 1 if should and not (body.get("must") or body.get("filter")) else 0
        if hits < int(minimum):
            return None
        return float(matched + hits)

    def _matches(self, document: Mapping[str, Any], query: Mapping[str, Any]) -> bool:
        return self._score(document, query) is not None

    def _match_leaf(
        self, document: Mapping[str, Any], clause: str, body: Mapping[str, Any]
    ) -> bool:
        for field, condition in body.items():
            value = document.get(field)
            if clause == "term":
                if not _term_matches(value, condition):
                    return False
            elif clause == "terms":
                if not isinstance(condition, list):
                    raise UnsupportedQueryError("terms expects a list of values")
                if not any(_term_matches(value, candidate) for candidate in condition):
                    return False
            elif clause == "match":
                if not _text_matches(value, condition):
                    return False
            elif clause == "range":
                if not _range_matches(value, condition):
                    return False
            else:  # pragma: no cover - guarded by _SUPPORTED
                raise UnsupportedQueryError(clause)
        return True

    def _match_nested(self, document: Mapping[str, Any], body: Mapping[str, Any]) -> bool:
        path = body.get("path")
        inner = body.get("query")
        if not isinstance(path, str) or not isinstance(inner, Mapping):
            raise UnsupportedQueryError("nested requires {'path': str, 'query': {...}}")
        values = document.get(path)
        if not isinstance(values, list):
            return False
        # Real nested queries write the full path in the inner clause ("references.resolved");
        # inside the element the field is just "resolved".
        scoped = _strip_prefix(inner, path)
        return any(isinstance(item, Mapping) and self._matches(item, scoped) for item in values)


def _strip_prefix(query: Mapping[str, Any], path: str) -> dict[str, Any]:
    """Rewrite inner nested field names, dropping the redundant ``<path>.`` prefix."""
    prefix = f"{path}."
    rewritten: dict[str, Any] = {}
    for clause, body in query.items():
        if clause == "bool" and isinstance(body, Mapping):
            rewritten[clause] = {
                key: [_strip_prefix(clause_body, path) for clause_body in value]
                if isinstance(value, list)
                else _strip_prefix(value, path)
                for key, value in body.items()
            }
        elif isinstance(body, Mapping):
            rewritten[clause] = {
                (field[len(prefix) :] if field.startswith(prefix) else field): condition
                for field, condition in body.items()
            }
        else:  # pragma: no cover - defensive
            rewritten[clause] = body
    return rewritten


def _cosine(left: Sequence[float], right: Sequence[float]) -> float | None:
    """Cosine similarity, or ``None`` for a zero vector (a cluster refuses to index one)."""
    dot = sum(a * b for a, b in zip(left, right, strict=True))
    norm = math.sqrt(sum(a * a for a in left)) * math.sqrt(sum(b * b for b in right))
    return None if norm == 0.0 else dot / norm


def _term_matches(value: Any, condition: Any) -> bool:
    if isinstance(value, list):
        return any(_term_matches(item, condition) for item in value)
    if isinstance(value, Mapping) and isinstance(condition, Mapping):
        return all(value.get(key) == sub for key, sub in condition.items())
    return bool(value == condition)


def _text_matches(value: Any, condition: Any) -> bool:
    if not isinstance(condition, str):
        raise UnsupportedQueryError("match expects a string")
    if value is None:
        return False
    haystack = " ".join(str(part) for part in value) if isinstance(value, list) else str(value)
    needle = {token for token in condition.lower().split() if token}
    tokens = {token.strip(".,;:()?!\"'") for token in haystack.lower().split()}
    return bool(needle & tokens)


def _range_matches(value: Any, condition: Any) -> bool:
    if not isinstance(condition, Mapping):
        raise UnsupportedQueryError("range expects {'gte': ..., 'lte': ...}")
    bounds = {key: condition[key] for key in ("gte", "gt", "lte", "lt") if key in condition}
    if isinstance(value, Mapping) and {"gte", "lte"} <= set(value):
        # An `integer_range` field: match when the query bounds intersect the stored span.
        low = bounds.get("gte", bounds.get("gt"))
        high = bounds.get("lte", bounds.get("lt"))
        return not (low is not None and value["lte"] < low) and not (
            high is not None and value["gte"] > high
        )
    if value is None:
        return False
    for key, bound in bounds.items():
        try:
            if key == "gte" and not value >= bound:
                return False
            if key == "gt" and not value > bound:
                return False
            if key == "lte" and not value <= bound:
                return False
            if key == "lt" and not value < bound:
                return False
        except TypeError:
            return False
    return True


def in_memory_client(settings: Any = None) -> InMemoryElasticsearch:
    """Create a client and, when settings describe an index prefix, create the three indices.

    Import-free of :mod:`excel_rag.settings` on purpose: this module is the double used by tests
    of any layer, and it should not drag configuration in. Accepts either a :class:`Settings` or an
    ``ElasticsearchSettings``.
    """
    client = InMemoryElasticsearch()
    if settings is None:
        return client
    es_settings = getattr(settings, "elasticsearch", settings)
    if not hasattr(es_settings, "index_name"):
        raise TypeError("expected Settings or ElasticsearchSettings")
    from .es import INDEX_MAPPINGS

    for base, mapping in INDEX_MAPPINGS.items():
        client.create_index(es_settings.index_name(base), mapping)
    return client


__all__ = ["InMemoryElasticsearch", "UnsupportedQueryError", "in_memory_client"]


def _self_check() -> type[ElasticsearchLike]:  # pragma: no cover - typing assertion
    """Compile-time check that the double satisfies the protocol."""
    return InMemoryElasticsearch
