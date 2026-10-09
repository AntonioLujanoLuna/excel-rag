"""The indexer: index creation, budget refusal, and the atomic version flip."""

from __future__ import annotations

import pytest

from excel_rag.es import INDEX_CHUNKS, INDEX_MAPPINGS, INDEX_STRUCTURE, INDEX_VERSIONS
from excel_rag.fake_es import InMemoryElasticsearch, in_memory_client
from excel_rag.ingest import IngestError, ingest_workbook
from excel_rag.ingest.indexer import Indexer, build_client, document_body
from excel_rag.settings import Settings


def _indexer(settings: Settings | None = None) -> tuple[InMemoryElasticsearch, Indexer]:
    settings = settings or Settings()
    client = in_memory_client(settings)
    return client, Indexer(client, settings)


def test_build_client_defaults_to_the_in_memory_double() -> None:
    client = build_client(Settings())
    assert client.indices_exists(INDEX_CHUNKS)


def test_ensure_indices_is_idempotent() -> None:
    client = InMemoryElasticsearch()
    indexer = Indexer(client, Settings())
    indexer.ensure_indices()
    indexer.ensure_indices()
    for base in INDEX_MAPPINGS:
        assert client.indices_exists(base)
    assert client.mapping(INDEX_STRUCTURE)["properties"]["row_span"]["type"] == "integer_range"


def test_index_workbook_writes_documents_and_manifest(build) -> None:
    client, indexer = _indexer()
    ingested = ingest_workbook(build.path("two_tables_one_sheet"), workbook_id="wb", version=1)
    result = indexer.index_workbook(ingested)
    assert result.chunks_written == len(ingested.chunks)
    assert result.structure_written == len(ingested.structure)
    assert result.replaced_version is None
    assert indexer.active_version("wb") == 1
    assert client.count(INDEX_CHUNKS, {"term": {"workbook_id": "wb"}}) == len(ingested.chunks)


def test_document_body_drops_empty_embedding_fields(build) -> None:
    ingested = ingest_workbook(build.path("two_tables_one_sheet"), workbook_id="wb", version=1)
    body = document_body(ingested.chunks[0])
    assert "embedding" not in body
    assert "embedding_model" not in body
    assert "colbert" not in body


def test_version_replacement_leaves_only_the_new_version(build) -> None:
    client, indexer = _indexer()
    first = ingest_workbook(build.path("two_tables_one_sheet"), workbook_id="wb", version=1)
    indexer.index_workbook(first)
    second = ingest_workbook(build.path("merged_multilevel_header"), workbook_id="wb", version=2)
    result = indexer.index_workbook(second)

    assert result.replaced_version == 1
    assert result.deleted_documents > 0
    assert indexer.active_version("wb") == 2
    assert client.count(INDEX_CHUNKS, {"term": {"version": 1}}) == 0
    assert client.count(INDEX_STRUCTURE, {"term": {"version": 1}}) == 0


def test_budget_refusal(build) -> None:
    from fixtures import make_fixtures as mk

    settings = Settings(budgets={"max_documents_per_workbook": 100})
    _, indexer = _indexer(settings)
    # A wide, tall region produces well over 100 documents.
    path = mk.large_region(build.directory, rows=400, columns=40)
    ingested = ingest_workbook(path, workbook_id="wb", version=1)
    assert len(ingested.chunks) + len(ingested.structure) > 100
    with pytest.raises(IngestError, match="above the 100-document ceiling"):
        indexer.index_workbook(ingested)


def test_manifest_index_document_is_readable(build) -> None:
    client, indexer = _indexer()
    ingested = ingest_workbook(build.path("table_object"), workbook_id="wb", version=1)
    indexer.index_workbook(ingested)
    stored = client.get_document(INDEX_VERSIONS, "wb")
    assert stored is not None
    assert stored["active_version"] == 1


def test_index_prefix_is_honoured(build) -> None:
    settings = Settings(elasticsearch={"index_prefix": "dev-"})
    client = in_memory_client(settings)
    indexer = Indexer(client, settings)
    ingested = ingest_workbook(build.path("table_object"), workbook_id="wb", version=1)
    indexer.index_workbook(ingested)
    assert client.indices_exists("dev-excel_chunks")
    assert indexer.active_version("wb") == 1
