"""The whole stack against a real Elasticsearch cluster.

Everything else in the suite runs against the in-memory double, which accepts any mapping and any
document and refreshes instantly -- so it cannot notice a mapping a cluster refuses, a bulk item a
cluster rejects, or a write that is not yet visible. These tests run the same seam as
``tests/test_integration.py`` against a cluster, and are skipped unless one is named:

    EXCEL_RAG_TEST_ES_URL=http://127.0.0.1:9200 uv run pytest tests/live

Each test writes under its own index prefix and deletes its indices afterwards.
"""

from __future__ import annotations

import os
import sys
import uuid
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from excel_rag import create_app
from excel_rag.es import INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS
from excel_rag.ingest import Indexer, build_client, ingest_workbook
from excel_rag.settings import ElasticsearchSettings, Settings

sys.path.insert(0, str(Path(__file__).parent.parent / "fixtures"))

import make_fixtures

ES_URL = os.environ.get("EXCEL_RAG_TEST_ES_URL")
pytestmark = pytest.mark.skipif(not ES_URL, reason="set EXCEL_RAG_TEST_ES_URL to run")

WORKBOOK_ID = "wb-live"
FINANCE = "finance-team"


@pytest.fixture
def settings() -> Iterator[Settings]:
    prefix = f"test-{uuid.uuid4().hex[:8]}-"
    configured = Settings(
        use_live_elasticsearch=True,
        elasticsearch=ElasticsearchSettings(urls=(str(ES_URL),), index_prefix=prefix),
    )
    yield configured
    client = build_client(configured)
    for base in (INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS):
        name = configured.elasticsearch.index_name(base)
        if client.indices_exists(name):
            client.delete_index(name)


def _index(settings: Settings, path: Path, *, version: int, acl: str = FINANCE) -> Any:
    ingested = ingest_workbook(path, workbook_id=WORKBOOK_ID, version=version, acl_scope=(acl,))
    Indexer(build_client(settings), settings).index_workbook(ingested)
    return ingested


def _api(settings: Settings) -> TestClient:
    return TestClient(create_app(settings))


def test_the_cluster_accepts_the_mappings_at_the_configured_dims(settings: Settings) -> None:
    client = build_client(settings)
    Indexer(client, settings).ensure_indices()
    mapping = client.client.indices.get_mapping(index=settings.elasticsearch.chunks_index)
    properties = next(iter(mapping.body.values()))["mappings"]["properties"]
    assert properties["embedding"]["dims"] == settings.embedding.dims
    assert properties["version_key"]["type"] == "keyword"


def test_ingested_documents_are_searchable_and_expand(settings: Settings, tmp_path: Path) -> None:
    _index(settings, make_fixtures.cross_sheet_formula(tmp_path), version=1)
    response = _api(settings).post(
        "/api/v1/search/excel",
        json={
            "query": "Forecast B2 compute",
            "filters": {"acl_scopes": [FINANCE]},
            "expand_references": True,
            "reference_depth": 1,
        },
    )
    assert response.status_code == 200, response.text
    payload = response.json()
    assert payload["hits"]
    ranges = {(node["sheet"], node["a1_range"]) for node in payload["nodes"].values()}
    assert {("Actuals", "D2:D500"), ("Assumptions", "C7")} <= ranges


def test_range_and_dependents_use_span_intersection(settings: Settings, tmp_path: Path) -> None:
    _index(settings, make_fixtures.cross_sheet_formula(tmp_path), version=1)
    api = _api(settings)
    body = {"workbook_id": WORKBOOK_ID, "sheet_name": "Actuals", "a1": "D100"}
    overlapping = api.post("/api/v1/excel/range", json=body)
    assert overlapping.status_code == 200, overlapping.text
    assert overlapping.json()["nodes"]
    dependents = api.post("/api/v1/excel/dependents", json=body)
    assert dependents.status_code == 200, dependents.text
    assert ("Forecast", "B2") in {
        (node["sheet"], node["a1_range"]) for node in dependents.json()["nodes"]
    }


def test_a_scope_the_documents_lack_returns_nothing(settings: Settings, tmp_path: Path) -> None:
    _index(settings, make_fixtures.cross_sheet_formula(tmp_path), version=1)
    response = _api(settings).post(
        "/api/v1/search/excel", json={"query": "revenue", "filters": {"acl_scopes": ["hr-team"]}}
    )
    assert response.status_code == 200
    assert response.json()["hits"] == []


def test_a_replacement_version_garbage_collects_the_old_one(
    settings: Settings, tmp_path: Path
) -> None:
    _index(settings, make_fixtures.cross_sheet_formula(tmp_path), version=1)
    second = _index(settings, make_fixtures.two_tables_one_sheet(tmp_path), version=2)
    client = build_client(settings)
    assert client.count(settings.elasticsearch.chunks_index, {"term": {"version": 1}}) == 0
    assert client.count(settings.elasticsearch.chunks_index) == len(second.chunks)


def test_reindexing_the_active_version_leaves_no_stale_documents(
    settings: Settings, tmp_path: Path
) -> None:
    _index(settings, make_fixtures.cross_sheet_formula(tmp_path), version=1)
    second = _index(settings, make_fixtures.two_tables_one_sheet(tmp_path), version=1)
    client = build_client(settings)
    assert client.count(settings.elasticsearch.chunks_index) == len(second.chunks)
    assert client.count(settings.elasticsearch.structure_index) == len(second.structure)


def test_a_large_region_is_written_in_full(settings: Settings, tmp_path: Path) -> None:
    """Bulk writes go through `helpers.bulk`, which chunks and raises on any rejected item."""
    ingested = _index(settings, make_fixtures.large_region(tmp_path, rows=5_000), version=1)
    client = build_client(settings)
    assert client.count(settings.elasticsearch.chunks_index) == len(ingested.chunks)
    assert client.count(settings.elasticsearch.structure_index) == len(ingested.structure)
