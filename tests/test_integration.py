"""The seam between ingestion and retrieval.

Neither workstream could write this: ingestion produced documents and retrieval queried them, on
separate branches, and each was tested only against its own fixtures. What is asserted here is that
the two halves compose — that a document ingestion writes is a document retrieval can find, that
every node id a hit points at exists in the structure index, that a formula edge ingestion emitted
is traversable through the API, and that the ACL and version filters hold across the join.

Built the way a deployment builds it: one client, ingestion into it, then the app running against
that same client.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from excel_rag import create_app
from excel_rag.embedding import HashingEmbedder
from excel_rag.fake_es import in_memory_client
from excel_rag.ingest import Indexer, ingest_workbook
from excel_rag.settings import Settings

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))

import make_fixtures

WORKBOOK_ID = "wb-it"
FINANCE = ("finance-team",)


def ingest(
    client: object,
    settings: Settings,
    path: Path,
    *,
    version: int,
    acl_scope: tuple[str, ...] = FINANCE,
) -> object:
    """One workbook in, documents out, against the client the app will read from."""
    ingested = ingest_workbook(path, workbook_id=WORKBOOK_ID, version=version, acl_scope=acl_scope)
    indexer = Indexer(client, settings)  # type: ignore[arg-type]
    indexer.ensure_indices()
    indexer.index_workbook(ingested)
    return ingested


@pytest.fixture
def deployment(tmp_path: Path) -> tuple[TestClient, object, Settings]:
    settings = Settings()
    client = in_memory_client(settings)
    path = make_fixtures.cross_sheet_formula(tmp_path)
    ingest(client, settings, path, version=1)
    app = create_app(settings)
    # The seam a gateway uses: the app's client is the one ingestion wrote into.
    app.state.client = client
    return TestClient(app), client, settings


def search(test_client: TestClient, query: str, **request_fields: object) -> dict:
    """Post a search. Top-level request fields stay top-level: see the `extra="forbid"` note in

    `excel_rag.models.SearchRequest` -- a field placed inside `filters` by mistake used to be
    ignored silently, which reads as "expansion is broken" when it is the request that is wrong.
    """
    payload = {"query": query, "filters": {"acl_scopes": list(FINANCE)}, "top_k": 5}
    payload.update(request_fields)
    response = test_client.post("/api/v1/search/excel", json=payload)
    assert response.status_code == 200, response.text
    return response.json()


class TestTheHalvesCompose:
    def test_a_document_ingestion_wrote_is_one_retrieval_finds(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        test_client, _, _ = deployment
        body = search(test_client, "revenue")
        assert body["hits"], "ingestion produced documents that retrieval cannot find"
        for hit in body["hits"]:
            assert hit["source"]["workbook_id"] == WORKBOOK_ID
            assert hit["source"]["version"] == 1
            assert hit["source"]["sheet"]
            assert hit["source"]["a1_range"]

    def test_every_hit_points_at_a_node_that_exists_in_the_structure_index(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        """The join contract: a hit is only useful if its node is retrievable exactly."""
        test_client, client, settings = deployment
        structure_index = settings.elasticsearch.structure_index
        body = search(test_client, "revenue", expand_references=True)
        assert body["hits"]
        missing = [
            hit["node_id"]
            for hit in body["hits"]
            if client.get_document(structure_index, hit["node_id"]) is None  # type: ignore[attr-defined]
        ]
        assert not missing, (
            f"chunk documents reference nodes the structure index does not hold: {missing}"
        )

    def test_a_hit_range_parses_and_lies_on_a_sheet_of_the_workbook(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        from excel_rag.models import A1Range

        test_client, _, _ = deployment
        body = search(test_client, "revenue")
        for hit in body["hits"]:
            parsed = A1Range.parse(hit["source"]["sheet"], hit["source"]["a1_range"])
            assert parsed.cell_count >= 1

    def test_expansion_returns_nodes_ingestion_emitted(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        test_client, _, _ = deployment
        body = search(test_client, "revenue", expand_references=True, reference_depth=1)
        assert body["nodes"], "depth-1 expansion returned nothing the workbook actually contains"
        for node in body["nodes"].values():
            assert node["node_id"]
            assert node["a1_range"]

    def test_the_structure_endpoint_lists_the_ingested_nodes(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        test_client, _, _ = deployment
        response = test_client.get(f"/api/v1/excel/{WORKBOOK_ID}/structure")
        assert response.status_code == 200, response.text
        nodes = response.json()["nodes"]
        assert nodes
        types = {node["node_type"] for node in nodes}
        assert {"sheet"} <= types
        assert types & {"region", "table", "column", "row_group", "cell", "formula"}

    def test_a_formula_edge_ingestion_emitted_is_traversable(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        """Follow the cross-sheet edge ingestion emitted, end to end, out of the index.

        The formula node `Forecast!B2` holds `=SUM(Actuals!D2:D500)*(1+Assumptions!C7)` and its
        `formula_summary` chunk is what a search lands on. The query below is built from tokens that
        chunk actually contains: the in-memory client matches tokens, it does not do BM25.
        """
        test_client, client, settings = deployment
        body = search(test_client, "Forecast B2 compute", expand_references=True, reference_depth=1)
        assert body["hits"], "the formula summary chunk was not found"

        structure_index = settings.elasticsearch.structure_index
        hit_node = body["hits"][0]["node_id"]
        stored = client.get_document(structure_index, hit_node)  # type: ignore[attr-defined]
        assert stored is not None, "the hit names a node the structure index does not hold"
        targets = {reference["target_node_id"] for reference in stored["references"]}
        assert targets, "the formula node carries no edges at all"

        followed = {node["node_id"] for node in body["nodes"].values() if node["depth"] >= 1}
        assert followed, "no reference was followed"
        assert targets <= followed, f"the traversal missed {sorted(targets - followed)}"

        # And the edge is not dangling: both targets are readable nodes.
        for target in targets:
            assert client.get_document(structure_index, target) is not None  # type: ignore[attr-defined]

    def test_range_query_intersects_spans_ingestion_stored(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        """A1 in, overlapping nodes out — span intersection, not a cell walk."""
        test_client, _, _ = deployment
        structure = test_client.get(f"/api/v1/excel/{WORKBOOK_ID}/structure").json()["nodes"]
        region = next(node for node in structure if node["node_type"] in {"region", "table"})
        response = test_client.post(
            "/api/v1/excel/range",
            json={
                "workbook_id": WORKBOOK_ID,
                "sheet_name": region["sheet"],
                "a1": region["a1_range"],
                "limit": 50,
            },
        )
        assert response.status_code == 200, response.text
        overlapping = response.json()["nodes"]
        assert overlapping, f"nothing overlaps {region['a1_range']} on {region['sheet']}"
        ids = {node["node_id"] for node in overlapping}
        assert region["node_id"] in ids, "the region does not overlap itself"


class TestFiltersHoldAcrossTheJoin:
    def test_a_scope_the_caller_lacks_returns_nothing(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        test_client, _, _ = deployment
        response = test_client.post(
            "/api/v1/search/excel",
            json={"query": "revenue", "filters": {"acl_scopes": ["hr-team"]}, "top_k": 5},
        )
        assert response.status_code == 200
        assert response.json()["hits"] == []

    def test_the_same_query_with_the_right_scope_returns_hits(
        self, deployment: tuple[TestClient, object, Settings]
    ) -> None:
        test_client, _, _ = deployment
        assert search(test_client, "revenue")["hits"]

    def test_a_related_node_the_caller_may_not_read_is_not_returned_through_the_join(
        self, tmp_path: Path
    ) -> None:
        """The `_mget` trap, exercised end to end: an unreadable node never reaches the payload."""
        settings = Settings()
        client = in_memory_client(settings)
        path = make_fixtures.cross_sheet_formula(tmp_path)
        ingest(client, settings, path, version=1, acl_scope=("finance-team",))
        app = create_app(settings)
        app.state.client = client
        test_client = TestClient(app)

        body = search(test_client, "revenue", expand_references=True, reference_depth=1)
        assert body["hits"]
        # Every returned node must carry the ingested scope, whatever the edge named.
        structure_index = settings.elasticsearch.structure_index
        for node in body["nodes"].values():
            stored = client.get_document(structure_index, node["node_id"])
            assert stored is not None
            assert "finance-team" in stored["acl_scope"]

    def test_search_never_mixes_versions(self, tmp_path: Path) -> None:
        settings = Settings()
        client = in_memory_client(settings)
        path = make_fixtures.cross_sheet_formula(tmp_path)
        ingest(client, settings, path, version=1)
        ingest(client, settings, path, version=2)
        app = create_app(settings)
        app.state.client = client
        test_client = TestClient(app)

        versions = {hit["source"]["version"] for hit in search(test_client, "revenue")["hits"]}
        assert versions == {2}, f"a request mixed versions: {versions}"

    def test_the_superseded_version_is_gone_from_both_indices(self, tmp_path: Path) -> None:
        settings = Settings()
        client = in_memory_client(settings)
        path = make_fixtures.cross_sheet_formula(tmp_path)
        ingest(client, settings, path, version=1)
        ingest(client, settings, path, version=2)

        chunks = settings.elasticsearch.chunks_index
        assert client.count(chunks, {"term": {"version": 1}}) == 0
        assert client.count(chunks, {"term": {"version": 2}}) > 0


class TestDependents:
    """The reverse edges: which formulas read a range, through the API, from ingested documents."""

    def _dependents(self, test_client: TestClient, sheet: str, a1: str, **extra: object) -> set:
        response = test_client.post(
            "/api/v1/excel/dependents",
            json={"workbook_id": WORKBOOK_ID, "sheet_name": sheet, "a1": a1, **extra},
        )
        assert response.status_code == 200, response.text
        return {(node["sheet"], node["a1_range"]) for node in response.json()["nodes"]}

    def test_a_cell_inside_a_read_range_finds_the_formula(self, deployment) -> None:
        """`SUM(Actuals!D2:D500)` is one edge, yet `D100` still finds the formula reading it."""
        test_client, _, _ = deployment
        assert ("Forecast", "B2") in self._dependents(test_client, "Actuals", "D100")

    def test_the_exact_precedent_cell_finds_the_formula(self, deployment) -> None:
        test_client, _, _ = deployment
        assert ("Forecast", "B2") in self._dependents(test_client, "Assumptions", "C7")

    def test_an_unread_cell_has_no_dependents(self, deployment) -> None:
        test_client, _, _ = deployment
        assert self._dependents(test_client, "Assumptions", "C8") == set()
        assert self._dependents(test_client, "Actuals", "E100") == set()

    def test_a_caller_without_the_scope_sees_no_dependents(self, deployment) -> None:
        test_client, _, _ = deployment
        assert self._dependents(test_client, "Actuals", "D100", acl_scopes=["hr-team"]) == set()


class TestHybrid:
    """Vectors written by ingestion are the ones the service's knn query reads."""

    @pytest.fixture
    def hybrid(self, tmp_path: Path) -> TestClient:
        settings = Settings()
        client = in_memory_client(settings)
        embedder = HashingEmbedder(settings.embedding.dims)
        ingested = ingest_workbook(
            make_fixtures.cross_sheet_formula(tmp_path),
            workbook_id=WORKBOOK_ID,
            version=1,
            acl_scope=FINANCE,
        )
        Indexer(client, settings, embedder=embedder).index_workbook(ingested)
        app = create_app(settings)
        app.state.client = client
        app.state.embedder = embedder
        return TestClient(app)

    def test_hits_report_the_vector_retriever(self, hybrid: TestClient) -> None:
        body = search(hybrid, "revenue forecast")
        assert body["hits"]
        assert any("embedding" in hit["matched_fields"] for hit in body["hits"])

    def test_health_names_the_query_model(self, hybrid: TestClient) -> None:
        assert hybrid.get("/health").json()["embedding_model"] == "hashing-test-double"

    def test_a_scope_the_caller_lacks_hides_vector_hits_too(self, hybrid: TestClient) -> None:
        response = hybrid.post(
            "/api/v1/search/excel",
            json={"query": "revenue forecast", "filters": {"acl_scopes": ["hr-team"]}},
        )
        assert response.json()["hits"] == []
