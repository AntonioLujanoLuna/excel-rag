"""Tests for the in-memory Elasticsearch double.

The double is the floor under every other test in this repo, so its query semantics get tested
directly: a clause it silently mis-handles would make a passing test meaningless.
"""

from __future__ import annotations

import pytest

from excel_rag.es import INDEX_CHUNKS, INDEX_MAPPINGS, INDEX_STRUCTURE
from excel_rag.fake_es import InMemoryElasticsearch, UnsupportedQueryError, in_memory_client
from excel_rag.settings import Settings

CHUNKS = [
    (
        "c1",
        {
            "id": "c1",
            "workbook_id": "wb42",
            "version": 3,
            "content": "Projected revenue by quarter",
            "chunk_type": "row_group",
            "acl_scope": ["finance-team"],
        },
    ),
    (
        "c2",
        {
            "id": "c2",
            "workbook_id": "wb42",
            "version": 3,
            "content": "Headcount plan and hiring assumptions",
            "chunk_type": "region",
            "acl_scope": ["hr-team"],
        },
    ),
    (
        "c3",
        {
            "id": "c3",
            "workbook_id": "wb77",
            "version": 1,
            "content": "Projected revenue sensitivity",
            "chunk_type": "table",
            "acl_scope": ["finance-team", "board"],
        },
    ),
]
STRUCTURE = [
    (
        "s1",
        {
            "node_id": "s1",
            "node_type": "range",
            "row_span": {"gte": 2, "lte": 500},
            "column_span": {"gte": 4, "lte": 4},
            "references": [
                {"target_node_id": "t1", "kind": "range", "resolved": True},
                {"target_node_id": "t2", "kind": "cell", "resolved": False},
            ],
        },
    ),
    ("s2", {"node_id": "s2", "node_type": "cell", "row_span": {"gte": 7, "lte": 7}}),
]


@pytest.fixture
def client() -> InMemoryElasticsearch:
    fake = InMemoryElasticsearch()
    fake.create_index(INDEX_CHUNKS, INDEX_MAPPINGS[INDEX_CHUNKS])
    fake.create_index(INDEX_STRUCTURE, INDEX_MAPPINGS[INDEX_STRUCTURE])
    fake.bulk_index(INDEX_CHUNKS, CHUNKS)
    fake.bulk_index(INDEX_STRUCTURE, STRUCTURE)
    return fake


class TestIndexManagement:
    def test_creating_twice_is_an_error(self) -> None:
        fake = InMemoryElasticsearch()
        fake.create_index("x", {})
        with pytest.raises(ValueError, match="already exists"):
            fake.create_index("x", {})

    def test_operating_on_a_missing_index_is_an_error(self) -> None:
        with pytest.raises(ValueError, match="no such index"):
            InMemoryElasticsearch().count("nope")

    def test_delete_index_removes_documents(self) -> None:
        fake = in_memory_client(Settings())
        assert fake.indices_exists("excel_chunks")
        fake.delete_index("excel_chunks")
        assert not fake.indices_exists("excel_chunks")

    def test_accepts_bare_elasticsearch_settings(self) -> None:
        fake = in_memory_client(Settings().elasticsearch)
        assert fake.indices_exists("excel_structure")

    def test_rejects_something_that_is_not_settings(self) -> None:
        with pytest.raises(TypeError, match="Settings"):
            in_memory_client(object())

    def test_index_prefix_is_applied(self) -> None:
        fake = in_memory_client(Settings(elasticsearch={"index_prefix": "dev-"}))
        assert fake.indices_exists("dev-excel_chunks")


class TestSearch:
    def test_match_all_counts_everything(self, client: InMemoryElasticsearch) -> None:
        result = client.search(INDEX_CHUNKS, {"match_all": {}}, size=10)
        assert result["hits"]["total"]["value"] == 3

    def test_size_windows_the_results(self, client: InMemoryElasticsearch) -> None:
        result = client.search(INDEX_CHUNKS, {"match_all": {}}, size=2)
        assert len(result["hits"]["hits"]) == 2
        assert result["hits"]["total"]["value"] == 3

    def test_term_filters_on_keyword(self, client: InMemoryElasticsearch) -> None:
        result = client.search(INDEX_CHUNKS, {"term": {"workbook_id": "wb42"}}, size=10)
        assert {hit["_id"] for hit in result["hits"]["hits"]} == {"c1", "c2"}

    def test_terms_matches_membership_in_a_list(self, client: InMemoryElasticsearch) -> None:
        result = client.search(
            INDEX_CHUNKS, {"terms": {"acl_scope": ["board", "hr-team"]}}, size=10
        )
        assert {hit["_id"] for hit in result["hits"]["hits"]} == {"c2", "c3"}

    def test_match_is_token_based(self, client: InMemoryElasticsearch) -> None:
        result = client.search(INDEX_CHUNKS, {"match": {"content": "projected revenue"}}, size=10)
        assert {hit["_id"] for hit in result["hits"]["hits"]} == {"c1", "c3"}

    def test_bool_must_and_must_not(self, client: InMemoryElasticsearch) -> None:
        query = {
            "bool": {
                "must": [{"match": {"content": "projected revenue"}}],
                "must_not": [{"term": {"workbook_id": "wb77"}}],
            }
        }
        result = client.search(INDEX_CHUNKS, query, size=10)
        assert {hit["_id"] for hit in result["hits"]["hits"]} == {"c1"}

    def test_bool_should_requires_one_by_default(self, client: InMemoryElasticsearch) -> None:
        query = {"bool": {"should": [{"term": {"workbook_id": "wb77"}}]}}
        result = client.search(INDEX_CHUNKS, query, size=10)
        assert {hit["_id"] for hit in result["hits"]["hits"]} == {"c3"}

    def test_minimum_should_match_is_honoured(self, client: InMemoryElasticsearch) -> None:
        query = {
            "bool": {
                "should": [
                    {"term": {"workbook_id": "wb42"}},
                    {"term": {"chunk_type": "table"}},
                ],
                "minimum_should_match": 2,
            }
        }
        assert client.count(INDEX_CHUNKS, query) == 0

    def test_hits_are_ordered_by_score_then_id(self, client: InMemoryElasticsearch) -> None:
        query = {
            "bool": {"should": [{"match": {"content": "projected"}}, {"term": {"version": 3}}]}
        }
        result = client.search(INDEX_CHUNKS, query, size=10)
        hits = result["hits"]["hits"]
        # c1 matches both clauses, c2 and c3 one each, so the tie breaks on id ascending.
        assert [hit["_id"] for hit in hits] == ["c1", "c2", "c3"]
        assert [hit["_score"] for hit in hits] == [2.0, 1.0, 1.0]

    def test_source_includes_projects_fields(self, client: InMemoryElasticsearch) -> None:
        result = client.search(
            INDEX_CHUNKS, {"match_all": {}}, size=1, source_includes=["id", "acl_scope"]
        )
        assert set(result["hits"]["hits"][0]["_source"]) == {"id", "acl_scope"}


class TestRanges:
    def test_integer_range_intersects(self, client: InMemoryElasticsearch) -> None:
        query = {"range": {"row_span": {"gte": 300, "lte": 400}}}
        assert {
            hit["_id"] for hit in client.search(INDEX_STRUCTURE, query, size=10)["hits"]["hits"]
        } == {"s1"}

    def test_integer_range_misses_when_disjoint(self, client: InMemoryElasticsearch) -> None:
        query = {"range": {"row_span": {"gte": 600, "lte": 700}}}
        assert client.count(INDEX_STRUCTURE, query) == 0

    def test_integer_range_intersection_is_inclusive(self, client: InMemoryElasticsearch) -> None:
        query = {"range": {"row_span": {"gte": 500, "lte": 900}}}
        assert client.count(INDEX_STRUCTURE, query) == 1

    def test_scalar_range_comparisons(self, client: InMemoryElasticsearch) -> None:
        assert client.count(INDEX_CHUNKS, {"range": {"version": {"gte": 2}}}) == 2
        assert client.count(INDEX_CHUNKS, {"range": {"version": {"lt": 3}}}) == 1


class TestNested:
    def test_nested_matches_any_element(self, client: InMemoryElasticsearch) -> None:
        query = {
            "nested": {
                "path": "references",
                "query": {"term": {"references.resolved": False}},
            }
        }
        assert {
            hit["_id"] for hit in client.search(INDEX_STRUCTURE, query, size=10)["hits"]["hits"]
        } == {"s1"}

    def test_nested_requires_path_and_query(self, client: InMemoryElasticsearch) -> None:
        with pytest.raises(UnsupportedQueryError, match="nested requires"):
            client.search(INDEX_STRUCTURE, {"nested": {"path": "references"}}, size=1)

    def test_nested_inner_clause_may_use_the_full_path(self, client: InMemoryElasticsearch) -> None:
        """`references.resolved` inside a nested query resolves to `resolved` in the element."""
        query = {
            "nested": {
                "path": "references",
                "query": {"bool": {"must": [{"term": {"references.target_node_id": "t1"}}]}},
            }
        }
        assert client.count(INDEX_STRUCTURE, query) == 1

    def test_nested_misses_when_no_element_matches(self, client: InMemoryElasticsearch) -> None:
        query = {
            "nested": {"path": "references", "query": {"term": {"references.target_node_id": "zz"}}}
        }
        assert client.count(INDEX_STRUCTURE, query) == 0

    def test_nested_on_a_missing_field_is_no_match(self, client: InMemoryElasticsearch) -> None:
        query = {"nested": {"path": "references", "query": {"term": {"references.kind": "cell"}}}}
        assert client.count(INDEX_CHUNKS, query) == 0


class TestKnn:
    """Exact cosine search, scored as a cluster scores a cosine `dense_vector`: (1 + cos) / 2."""

    def _knn(self, vector: list[float], **extra: object) -> dict:
        return {"knn": {"field": "embedding", "query_vector": vector, "k": 3, **extra}}

    def test_orders_by_cosine_and_honours_k(self) -> None:
        client = _vectors({"a": [1.0, 0.0], "b": [0.6, 0.8], "c": [0.0, 1.0], "d": [-1.0, 0.0]})
        result = client.search(INDEX_CHUNKS, self._knn([1.0, 0.0]), size=10)
        hits = result["hits"]["hits"]
        assert [hit["_id"] for hit in hits] == ["a", "b", "c"]
        assert hits[0]["_score"] == pytest.approx(1.0)
        assert hits[2]["_score"] == pytest.approx(0.5)

    def test_filters_apply_before_ranking(self) -> None:
        client = _vectors({"a": [1.0, 0.0], "b": [0.6, 0.8]}, scopes={"a": "hr"})
        query = self._knn([1.0, 0.0], filter=[{"term": {"acl_scope": "finance"}}])
        hits = client.search(INDEX_CHUNKS, query, size=10)["hits"]["hits"]
        assert [hit["_id"] for hit in hits] == ["b"]

    def test_missing_or_mismatched_vectors_are_not_candidates(self) -> None:
        client = _vectors({"a": [1.0, 0.0], "b": [1.0, 0.0, 0.0], "c": None})
        hits = client.search(INDEX_CHUNKS, self._knn([1.0, 0.0]), size=10)["hits"]["hits"]
        assert [hit["_id"] for hit in hits] == ["a"]

    def test_knn_with_a_query_in_one_request_is_refused(self) -> None:
        client = _vectors({"a": [1.0, 0.0]})
        query = {**self._knn([1.0, 0.0]), "match_all": {}}
        with pytest.raises(UnsupportedQueryError, match="separate"):
            client.search(INDEX_CHUNKS, query, size=1)


def _vectors(
    vectors: dict[str, list[float] | None], scopes: dict[str, str] | None = None
) -> InMemoryElasticsearch:
    client = InMemoryElasticsearch()
    client.create_index(INDEX_CHUNKS, {})
    client.bulk_index(
        INDEX_CHUNKS,
        [
            (
                identifier,
                {
                    "id": identifier,
                    "acl_scope": [(scopes or {}).get(identifier, "finance")],
                    **({"embedding": vector} if vector is not None else {}),
                },
            )
            for identifier, vector in vectors.items()
        ],
    )
    return client


class TestRefusals:
    @pytest.mark.parametrize(
        "query",
        [
            {"knn": {"field": "embedding", "k": 10}},
            {"script_score": {"query": {"match_all": {}}, "script": "1"}},
            {"function_score": {"query": {"match_all": {}}}},
            {"wildcard": {"content": "rev*"}},
            {"aggs": {"by_type": {"terms": {"field": "chunk_type"}}}},
        ],
    )
    def test_unsupported_clauses_raise(self, client: InMemoryElasticsearch, query: dict) -> None:
        with pytest.raises(UnsupportedQueryError):
            client.search(INDEX_CHUNKS, query, size=1)

    def test_the_error_names_the_clause_and_the_supported_set(
        self, client: InMemoryElasticsearch
    ) -> None:
        with pytest.raises(UnsupportedQueryError) as excinfo:
            client.search(INDEX_CHUNKS, {"script_score": {}}, size=1)
        assert "script_score" in str(excinfo.value)
        assert "match_all" in str(excinfo.value)

    def test_terms_requires_a_list(self, client: InMemoryElasticsearch) -> None:
        with pytest.raises(UnsupportedQueryError, match="terms expects a list"):
            client.search(INDEX_CHUNKS, {"terms": {"workbook_id": "wb42"}}, size=1)


class TestDocumentAccess:
    def test_mget_keeps_requested_order_and_skips_missing(
        self, client: InMemoryElasticsearch
    ) -> None:
        found = client.mget_documents(INDEX_CHUNKS, ["c3", "absent", "c1"])
        assert [document["id"] for document in found] == ["c3", "c1"]

    def test_mget_applies_no_acl_filter(self, client: InMemoryElasticsearch) -> None:
        """Documents the caller may not read come back anyway: filtering is the caller's job."""
        found = client.mget_documents(INDEX_CHUNKS, ["c1", "c2"])
        assert len(found) == 2

    def test_get_returns_none_for_a_missing_id(self, client: InMemoryElasticsearch) -> None:
        assert client.get_document(INDEX_CHUNKS, "absent") is None

    def test_documents_are_copied_on_the_way_in_and_out(
        self, client: InMemoryElasticsearch
    ) -> None:
        document = {"id": "c9", "acl_scope": ["a"]}
        client.index_document(INDEX_CHUNKS, "c9", document)
        document["acl_scope"].append("mutated")
        assert client.get_document(INDEX_CHUNKS, "c9")["acl_scope"] == ["a"]

    def test_delete_by_query_returns_the_count(self, client: InMemoryElasticsearch) -> None:
        assert client.delete_by_query(INDEX_CHUNKS, {"term": {"workbook_id": "wb42"}}) == 2
        assert client.count(INDEX_CHUNKS) == 1

    def test_calls_are_recorded_for_assertions(self, client: InMemoryElasticsearch) -> None:
        client.search(INDEX_CHUNKS, {"match_all": {}}, size=1)
        assert ("search", INDEX_CHUNKS) in client.calls
