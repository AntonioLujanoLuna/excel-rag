"""Write an ingested workbook into Elasticsearch -- the only persistent store.

Index creation, bulk writes, the active-version manifest, a document-budget refusal and an atomic
version replacement all live here. Two invariants from the README shape it:

* **No cell-per-document explosion.** A workbook that would write more than
  ``budgets.max_documents_per_workbook`` documents is *refused*, never silently clipped.
* **Index before activate, then garbage-collect.** A replacement version's documents are written
  first, then the manifest flips to the new version, and only then are the previous version's
  documents removed by query. A reader that consults the manifest therefore never sees a half-built
  version, and ids never cross versions.

Embedding fields are left empty: producing vectors is the retrieval workstream's pipeline, not this
one's.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

from ..es import INDEX_MAPPINGS, ElasticsearchLike
from ..models import ActiveVersionManifest, ChunkDocument, StructureDocument
from ..settings import Settings
from .canonical import WorkbookModel
from .documents import IngestedWorkbook
from .errors import IngestError

_EMPTY_EMBEDDING_FIELDS = ("embedding", "embedding_model", "colbert")


@dataclass(frozen=True, slots=True)
class IndexResult:
    """What one indexing run did, so a caller can report or assert on it."""

    workbook_id: str
    version: int
    chunks_written: int
    structure_written: int
    documents_total: int
    replaced_version: int | None
    deleted_documents: int


def document_body(document: ChunkDocument | StructureDocument) -> dict[str, Any]:
    """Serialise a document, dropping embedding fields that are still empty."""
    body: dict[str, Any] = document.model_dump(mode="json")
    for field in _EMPTY_EMBEDDING_FIELDS:
        if body.get(field) is None:
            body.pop(field, None)
    return body


class Indexer:
    """Indexes ingested workbooks against the narrow :class:`ElasticsearchLike` protocol."""

    def __init__(self, client: ElasticsearchLike, settings: Settings) -> None:
        self.client = client
        self.settings = settings

    def _name(self, base: str) -> str:
        return self.settings.elasticsearch.index_name(base)

    def ensure_indices(self) -> None:
        """Create any of the three indices that does not yet exist. Idempotent."""
        for base, mapping in INDEX_MAPPINGS.items():
            name = self._name(base)
            if not self.client.indices_exists(name):
                self.client.create_index(name, mapping)

    def count_documents(self, ingested: IngestedWorkbook) -> int:
        return len(ingested.chunks) + len(ingested.structure)

    def index_workbook(self, ingested: IngestedWorkbook, *, replace: bool = True) -> IndexResult:
        model: WorkbookModel = ingested.model
        total = self.count_documents(ingested)
        ceiling = self.settings.budgets.max_documents_per_workbook
        if total > ceiling:
            raise IngestError(
                f"workbook {model.workbook_id!r} v{model.version} would write {total} documents, "
                f"above the {ceiling}-document ceiling; refusing rather than clipping"
            )
        self.ensure_indices()
        chunks_index = self._name("excel_chunks")
        structure_index = self._name("excel_structure")
        versions_index = self._name("excel_versions")

        previous: int | None = None
        existing = self.client.get_document(versions_index, model.workbook_id)
        if existing is not None and existing.get("active_version") is not None:
            previous = int(existing["active_version"])

        chunks = [(chunk.id, document_body(chunk)) for chunk in ingested.chunks]
        structure = [(node.node_id, document_body(node)) for node in ingested.structure]
        self.client.bulk_index(chunks_index, chunks, refresh=True)
        self.client.bulk_index(structure_index, structure, refresh=True)

        manifest = ActiveVersionManifest(
            workbook_id=model.workbook_id,
            active_version=model.version,
            replaced_at=datetime.now(UTC),
        )
        self.client.index_document(
            versions_index, model.workbook_id, manifest.model_dump(mode="json"), refresh=True
        )

        deleted = 0
        replaced_version: int | None = None
        if replace and previous is not None and previous != model.version:
            replaced_version = previous
            deleted += self._delete_version(chunks_index, model.workbook_id, previous)
            deleted += self._delete_version(structure_index, model.workbook_id, previous)

        return IndexResult(
            workbook_id=model.workbook_id,
            version=model.version,
            chunks_written=len(chunks),
            structure_written=len(structure),
            documents_total=total,
            replaced_version=replaced_version,
            deleted_documents=deleted,
        )

    def _delete_version(self, index: str, workbook_id: str, version: int) -> int:
        query = {
            "bool": {
                "filter": [
                    {"term": {"workbook_id": workbook_id}},
                    {"term": {"version": version}},
                ]
            }
        }
        return self.client.delete_by_query(index, query)

    def active_version(self, workbook_id: str) -> int | None:
        stored = self.client.get_document(self._name("excel_versions"), workbook_id)
        if stored is None or stored.get("active_version") is None:
            return None
        return int(stored["active_version"])


class _LiveElasticsearch:
    """A thin adapter from the real ``elasticsearch`` client to the narrow protocol.

    The protocol uses ``indices_exists``/``create_index`` and flat arguments; the real client uses
    ``indices.exists``/``indices.create`` and kwargs. Imported lazily so the package installs and
    tests without the ``es`` extra.
    """

    def __init__(self, client: Any) -> None:  # pragma: no cover - needs the es extra
        self._client = client

    def indices_exists(self, index: str) -> bool:  # pragma: no cover - needs the es extra
        return bool(self._client.indices.exists(index=index))

    def create_index(self, index: str, mappings: Mapping[str, Any]) -> None:  # pragma: no cover
        self._client.indices.create(index=index, mappings=mappings)

    def delete_index(self, index: str) -> None:  # pragma: no cover - needs the es extra
        self._client.indices.delete(index=index)

    def index_document(
        self, index: str, document_id: str, document: Mapping[str, Any], *, refresh: bool = False
    ) -> None:  # pragma: no cover - needs the es extra
        self._client.index(index=index, id=document_id, document=dict(document), refresh=refresh)

    def bulk_index(
        self,
        index: str,
        documents: Any,
        *,
        refresh: bool = False,
    ) -> int:  # pragma: no cover - needs the es extra
        operations: list[dict[str, Any]] = []
        written = 0
        for document_id, document in documents:
            operations.append({"index": {"_index": index, "_id": document_id}})
            operations.append(dict(document))
            written += 1
        if operations:
            self._client.bulk(operations=operations, refresh=refresh)
        return written

    def get_document(
        self, index: str, document_id: str
    ) -> Mapping[str, Any] | None:  # pragma: no cover
        if not self._client.exists(index=index, id=document_id):
            return None
        result: Mapping[str, Any] = self._client.get(index=index, id=document_id)["_source"]
        return result

    def mget_documents(self, index: str, document_ids: Any) -> Any:  # pragma: no cover
        response = self._client.mget(index=index, ids=list(document_ids))
        return [doc["_source"] for doc in response["docs"] if doc.get("found")]

    def search(
        self, index: str, query: Mapping[str, Any], **kwargs: Any
    ) -> Any:  # pragma: no cover
        return self._client.search(index=index, query=query, **kwargs)

    def delete_by_query(self, index: str, query: Mapping[str, Any]) -> int:  # pragma: no cover
        return int(self._client.delete_by_query(index=index, query=query)["deleted"])

    def count(self, index: str, query: Mapping[str, Any] | None = None) -> int:  # pragma: no cover
        return int(self._client.count(index=index, query=query)["count"])


def build_client(settings: Settings) -> ElasticsearchLike:
    """The in-memory double for local runs, the real client (wrapped) when configured."""
    if not settings.use_live_elasticsearch:
        from ..fake_es import in_memory_client

        return in_memory_client(settings)
    from elasticsearch import Elasticsearch  # type: ignore[import-not-found]  # pragma: no cover

    credentials = settings.elasticsearch
    auth: tuple[str, str] | None = None
    if credentials.username and credentials.password is not None:
        auth = (credentials.username, credentials.password.get_secret_value())
    client = Elasticsearch(
        list(credentials.urls),
        request_timeout=credentials.request_timeout_seconds,
        basic_auth=auth,
    )
    return _LiveElasticsearch(client)  # pragma: no cover - needs the es extra


__all__ = ["IndexResult", "Indexer", "build_client", "document_body"]
