"""Elasticsearch is the only persistent store, so its two mappings are part of the contract.

The mappings are data (:data:`INDEX_MAPPINGS`) rather than prose: the ingestion side writes against
them, the retrieval side queries the fields they declare, and a test asserts they stay consistent
with :mod:`excel_rag.models`. A field the code uses but the mapping does not declare is the failure
mode this module exists to prevent -- Elasticsearch would accept the document and never match on it.

The client protocol is deliberately narrow. Anything not listed there is something this service
does not get to do, which is how "no additional persistent store" stays true.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import Any, Protocol

INDEX_CHUNKS = "excel_chunks"
INDEX_STRUCTURE = "excel_structure"

#: Keyword fields that identify a scope. Filtering on these is the only ACL mechanism in play, and
#: it has to be applied to primary searches *and* to every related-node lookup: `_mget` does not
#: enforce document-level ACLs on its own.
ACL_FIELD = "acl_scope"

#: The embedding fields. Dimensionality follows the existing retrieval stack; `embedding_model` is
#: carried per document so a query embedded by a different model is a miss rather than a silent
#: ranking change.
EMBEDDING_FIELD = "embedding"
EMBEDDING_MODEL_FIELD = "embedding_model"
COLBERT_FIELD = "colbert"

_RANGE_SPAN = {"type": "integer_range"}

INDEX_MAPPINGS: dict[str, dict[str, Any]] = {
    INDEX_CHUNKS: {
        "properties": {
            "id": {"type": "keyword"},
            "workbook_id": {"type": "keyword"},
            "version": {"type": "integer"},
            "node_id": {"type": "keyword"},
            "sheet_id": {"type": "keyword"},
            "sheet_name": {"type": "keyword"},
            "a1_range": {"type": "keyword"},
            "chunk_type": {"type": "keyword"},
            "title": {"type": "text", "fields": {"raw": {"type": "keyword"}}},
            "content": {"type": "text"},
            "headers": {"type": "text", "fields": {"raw": {"type": "keyword"}}},
            EMBEDDING_FIELD: {"type": "dense_vector", "index": True, "similarity": "cosine"},
            EMBEDDING_MODEL_FIELD: {"type": "keyword"},
            ACL_FIELD: {"type": "keyword"},
            "source_file": {"type": "keyword"},
            "source_sha256": {"type": "keyword"},
            "ingested_at": {"type": "date"},
        }
    },
    INDEX_STRUCTURE: {
        "properties": {
            "node_id": {"type": "keyword"},
            "workbook_id": {"type": "keyword"},
            "version": {"type": "integer"},
            "node_type": {"type": "keyword"},
            "sheet_id": {"type": "keyword"},
            "sheet_name": {"type": "keyword"},
            "a1_range": {"type": "keyword"},
            "row_span": _RANGE_SPAN,
            "column_span": _RANGE_SPAN,
            "parent_id": {"type": "keyword"},
            "child_ids": {"type": "keyword"},
            "formula": {"type": "text"},
            "cached_value": {"type": "object", "enabled": False},
            "display_value": {"type": "text"},
            "table_name": {"type": "keyword"},
            "column_name": {"type": "keyword"},
            "named_range": {"type": "keyword"},
            "references": {
                "type": "nested",
                "properties": {
                    "target_node_id": {"type": "keyword"},
                    "sheet_name": {"type": "keyword"},
                    "a1_range": {"type": "keyword"},
                    "kind": {"type": "keyword"},
                    "resolved": {"type": "boolean"},
                },
            },
            "unresolved_references": {
                "type": "nested",
                "properties": {
                    "reference_text": {"type": "keyword"},
                    "reason": {"type": "keyword"},
                    "detail": {"type": "text"},
                },
            },
            ACL_FIELD: {"type": "keyword"},
            "source_file": {"type": "keyword"},
        }
    },
}

#: One bookkeeping index for the active-version manifest.
INDEX_VERSIONS = "excel_versions"
INDEX_MAPPINGS[INDEX_VERSIONS] = {
    "properties": {
        "workbook_id": {"type": "keyword"},
        "active_version": {"type": "integer"},
        "replaced_at": {"type": "date"},
    }
}


class ElasticsearchLike(Protocol):
    """The slice of Elasticsearch this service uses.

    A real ``elasticsearch.Elasticsearch`` satisfies it; so does the in-memory double the tests
    run against. Anything outside this slice -- scripts, aggregations with side effects, reindex
    jobs -- is out of scope on purpose.
    """

    def indices_exists(self, index: str) -> bool: ...

    def create_index(self, index: str, mappings: Mapping[str, Any]) -> None: ...

    def delete_index(self, index: str) -> None: ...

    def index_document(
        self, index: str, document_id: str, document: Mapping[str, Any], *, refresh: bool = False
    ) -> None: ...

    def bulk_index(
        self,
        index: str,
        documents: Iterable[tuple[str, Mapping[str, Any]]],
        *,
        refresh: bool = False,
    ) -> int: ...

    def get_document(self, index: str, document_id: str) -> Mapping[str, Any] | None: ...

    def mget_documents(
        self, index: str, document_ids: Sequence[str]
    ) -> Sequence[Mapping[str, Any]]: ...

    def search(
        self,
        index: str,
        query: Mapping[str, Any],
        *,
        size: int = 10,
        source_includes: Sequence[str] | None = None,
    ) -> Mapping[str, Any]: ...

    def delete_by_query(self, index: str, query: Mapping[str, Any]) -> int: ...

    def count(self, index: str, query: Mapping[str, Any] | None = None) -> int: ...
