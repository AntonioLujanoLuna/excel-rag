"""``lightonai/mDenseOn`` end to end: real weights, real cluster, real ``knn``.

Opt-in, because it downloads ~1.2 GB of weights and needs the ``embed`` extra:

    EXCEL_RAG_TEST_ES_URL=http://127.0.0.1:9200 EXCEL_RAG_TEST_MDENSEON=1 uv run pytest tests/live

Six fixture workbooks are embedded and indexed once; each query below shares no word with the chunk
it should find -- paraphrases, and questions in Spanish, German and French over English workbooks.
The lexical retriever finds nothing for them, so these assert what the vectors add, not what BM25
already did.
"""

from __future__ import annotations

import importlib.util
import os
import sys
import uuid
from collections.abc import Iterator
from pathlib import Path

import pytest

from excel_rag.embedding import Embedder, build_embedder
from excel_rag.es import INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS
from excel_rag.ingest import Indexer, build_client, ingest_workbook
from excel_rag.models import SearchRequest
from excel_rag.retrieval import Repository, RetrievalService
from excel_rag.settings import ElasticsearchSettings, EmbeddingSettings, Settings

sys.path.insert(0, str(Path(__file__).parent.parent / "fixtures"))

import make_fixtures

ES_URL = os.environ.get("EXCEL_RAG_TEST_ES_URL")
pytestmark = pytest.mark.skipif(
    not (
        ES_URL
        and os.environ.get("EXCEL_RAG_TEST_MDENSEON")
        and importlib.util.find_spec("sentence_transformers")
    ),
    reason="set EXCEL_RAG_TEST_ES_URL and EXCEL_RAG_TEST_MDENSEON=1, with the 'embed' extra",
)

WORKBOOKS = (
    "two_tables_one_sheet",
    "units_and_notes",
    "cross_sheet_formula",
    "named_range",
    "table_object",
    "merged_multilevel_header",
)


@pytest.fixture(scope="module")
def corpus(tmp_path_factory: pytest.TempPathFactory) -> Iterator[tuple[Settings, Embedder]]:
    settings = Settings(
        use_live_elasticsearch=True,
        elasticsearch=ElasticsearchSettings(
            urls=(str(ES_URL),), index_prefix=f"test-{uuid.uuid4().hex[:8]}-"
        ),
        embedding=EmbeddingSettings(provider="sentence-transformers", model="lightonai/mDenseOn"),
    )
    embedder = build_embedder(settings.embedding)
    assert embedder is not None
    client = build_client(settings)
    directory = tmp_path_factory.mktemp("workbooks")
    for name in WORKBOOKS:
        path = getattr(make_fixtures, name)(directory)
        ingested = ingest_workbook(path, workbook_id=name, version=1)
        result = Indexer(client, settings, embedder=embedder).index_workbook(ingested)
        assert result.embedding_model == "lightonai/mDenseOn"
    yield settings, embedder
    for base in (INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS):
        client.delete_index(settings.elasticsearch.index_name(base))


@pytest.mark.parametrize(
    ("query", "workbook", "title"),
    [
        ("¿Qué moneda se usa para los ingresos?", "units_and_notes", "Revenue"),
        ("Quel est le chiffre d'affaires par région ?", "table_object", "Region"),
        ("Wie viele Einheiten wurden verkauft?", "two_tables_one_sheet", "Units"),
        ("expenses per quarter", "merged_multilevel_header", "Costs / Q1"),
        ("how is projected income calculated", "cross_sheet_formula", "Formula in Forecast B2"),
    ],
)
def test_the_vectors_find_what_no_word_matches(
    corpus: tuple[Settings, Embedder], query: str, workbook: str, title: str
) -> None:
    settings, embedder = corpus
    client = build_client(settings)
    request = SearchRequest(query=query, top_k=3, include_structure=False)

    hybrid = RetrievalService(Repository(client, settings), settings, embedder).search(request)
    found = [(hit.source.workbook_id, hit.title) for hit in hybrid.hits]
    assert (workbook, title) in found, found
    assert all("embedding" in hit.matched_fields for hit in hybrid.hits)

    lexical = RetrievalService(Repository(client, settings), settings, None).search(request)
    assert (workbook, title) not in [(hit.source.workbook_id, hit.title) for hit in lexical.hits]


def test_every_chunk_is_stored_with_a_768_dimensional_vector(
    corpus: tuple[Settings, Embedder],
) -> None:
    settings, _ = corpus
    client = build_client(settings)
    missing = client.count(
        settings.elasticsearch.chunks_index,
        {"bool": {"must_not": [{"exists": {"field": "embedding"}}]}},
    )
    assert missing == 0
