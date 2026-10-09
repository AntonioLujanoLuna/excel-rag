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

from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

from ..es import ElasticsearchLike, index_mappings
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
        """Create any of the three indices that does not yet exist. Idempotent.

        The mappings are built from the configured embedding dimensionality: a cluster fixes a
        ``dense_vector``'s ``dims`` at index-creation time and refuses the create without it.
        """
        for base, mapping in index_mappings(self.settings.embedding.dims).items():
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


def build_client(settings: Settings) -> ElasticsearchLike:
    """The in-memory double for local runs, the live adapter when configured.

    Ingestion and the service share one live adapter (:class:`~excel_rag.live.LiveElasticsearch`):
    its bulk path goes through ``elasticsearch.helpers.bulk``, which chunks a large workbook and
    raises on any rejected document instead of reporting it as written.
    """
    if not settings.use_live_elasticsearch:
        from ..fake_es import in_memory_client

        return in_memory_client(settings)
    from ..live import LiveElasticsearch

    return LiveElasticsearch(settings.elasticsearch)


__all__ = ["IndexResult", "Indexer", "build_client", "document_body"]
