"""The contract tests: the invariants two independent workstreams are built on.

These exist before either side lands, because a mismatch between a Pydantic field and an
Elasticsearch mapping is silent -- Elasticsearch accepts the extra key and the field simply never
matches. Everything here is deterministic and needs no cluster.
"""

from __future__ import annotations

import pytest

from excel_rag import __version__
from excel_rag.es import (
    ACL_FIELD,
    EMBEDDING_FIELD,
    EMBEDDING_MODEL_FIELD,
    INDEX_CHUNKS,
    INDEX_MAPPINGS,
    INDEX_STRUCTURE,
)
from excel_rag.models import (
    A1Range,
    ActiveVersionManifest,
    ChunkDocument,
    ChunkType,
    NodeType,
    SearchFilters,
    SearchRequest,
    StructureDocument,
)
from excel_rag.settings import Settings


class TestA1Range:
    @pytest.mark.parametrize(
        ("a1", "bounds"),
        [
            ("C7", (7, 7, 3, 3)),
            ("A1:F25", (1, 25, 1, 6)),
            ("$A$1:$F$25", (1, 25, 1, 6)),
            ("AA100:AB200", (100, 200, 27, 28)),
        ],
    )
    def test_parses_both_forms(self, a1: str, bounds: tuple[int, int, int, int]) -> None:
        parsed = A1Range.parse("Sheet1", a1)
        assert (parsed.min_row, parsed.max_row, parsed.min_col, parsed.max_col) == bounds

    def test_keeps_the_a1_string_a_caller_can_open(self) -> None:
        assert A1Range.parse("Sheet1", "$a$1:$f$25").a1 == "A1:F25"
        assert A1Range.parse("Sheet1", " c7 ").a1 == "C7"

    @pytest.mark.parametrize("a1", ["", "7C", "A0", "Sheet1!A1", "A1:B"])
    def test_rejects_non_ranges(self, a1: str) -> None:
        with pytest.raises(ValueError, match="not an A1 range"):
            A1Range.parse("Sheet1", a1)

    def test_intersection_is_sheet_scoped(self) -> None:
        target = A1Range.parse("Assumptions", "A1:F25")
        assert target.intersects(A1Range.parse("Assumptions", "F20:F30"))
        assert target.intersects(A1Range.parse("Assumptions", "A1"))
        assert not target.intersects(A1Range.parse("Assumptions", "G1:G9"))
        assert not target.intersects(A1Range.parse("Actuals", "A1:B2"))

    def test_spans_are_elasticsearch_integer_ranges(self) -> None:
        parsed = A1Range.parse("Sheet1", "B2:D9")
        assert parsed.row_span == {"gte": 2, "lte": 9}
        assert parsed.column_span == {"gte": 2, "lte": 4}
        assert parsed.cell_count == 24


class TestNodeId:
    def test_is_deterministic_and_readable(self) -> None:
        assert _node_id("wb42", 3, "cell", "Assumptions!C7") == _node_id(
            "wb42", 3, "cell", "Assumptions!C7"
        )

    def test_carries_workbook_version_kind_and_key(self) -> None:
        assert _node_id("wb42", 3, "cell", "Assumptions!C7") == "wb42:v3:cell:assumptions!c7"

    def test_different_versions_never_collide(self) -> None:
        assert _node_id("wb42", 1, "range", "A1:B2") != _node_id("wb42", 2, "range", "A1:B2")


def _node_id(workbook_id: str, version: int, kind: str, key: str) -> str:
    from excel_rag.models import node_id

    return node_id(workbook_id, version, kind, key)


class TestMappingsMatchTheModels:
    """Every document field the code writes must be a field the mapping can match on."""

    def test_chunk_document_fields_are_mapped(self) -> None:
        mapped = set(INDEX_MAPPINGS[INDEX_CHUNKS]["properties"])
        declared = set(ChunkDocument.model_fields)
        # `colbert` is a deliberate exception while the late-interaction stage is off; it is
        # carried so the pipeline can be switched on without a reindex.
        unmapped = declared - mapped - {"colbert"}
        assert not unmapped, f"chunk fields with no mapping: {sorted(unmapped)}"

    def test_structure_document_fields_are_mapped(self) -> None:
        mapped = set(INDEX_MAPPINGS[INDEX_STRUCTURE]["properties"])
        parent_fields = {"children"}  # hierarchy edges live on the child as `parent_id`
        unmapped = set(StructureDocument.model_fields) - mapped - parent_fields - {"node_id"}
        assert not unmapped, f"structure fields with no mapping: {sorted(unmapped)}"

    def test_embedding_fields_are_declared_together(self) -> None:
        properties = INDEX_MAPPINGS[INDEX_CHUNKS]["properties"]
        assert properties[EMBEDDING_FIELD]["type"] == "dense_vector"
        assert EMBEDDING_MODEL_FIELD in properties, "a vector without its model is not comparable"

    def test_acl_is_keyword_everywhere_it_appears(self) -> None:
        for index in (INDEX_CHUNKS, INDEX_STRUCTURE):
            assert INDEX_MAPPINGS[index]["properties"][ACL_FIELD]["type"] == "keyword"

    def test_spans_are_integer_ranges_in_both_indices(self) -> None:
        assert INDEX_MAPPINGS[INDEX_STRUCTURE]["properties"]["row_span"]["type"] == "integer_range"
        assert (
            INDEX_MAPPINGS[INDEX_STRUCTURE]["properties"]["column_span"]["type"] == "integer_range"
        )

    def test_reference_edges_are_nested(self) -> None:
        references = INDEX_MAPPINGS[INDEX_STRUCTURE]["properties"]["references"]
        assert references["type"] == "nested", "a flat array cannot be queried edge by edge"
        assert set(references["properties"]) >= {"target_node_id", "sheet_name", "a1_range", "kind"}

    def test_active_version_manifest_is_mapped(self) -> None:
        from excel_rag.es import INDEX_VERSIONS

        assert set(ActiveVersionManifest.model_fields) <= set(
            INDEX_MAPPINGS[INDEX_VERSIONS]["properties"]
        )


class TestSearchRequest:
    def test_defaults_are_bounded(self) -> None:
        request = SearchRequest(query="revenue")
        assert request.top_k == 10
        assert request.reference_depth == 1
        assert request.max_related_nodes == 20
        assert request.expand_references is False

    @pytest.mark.parametrize("query", ["", "   "])
    def test_blank_queries_are_rejected(self, query: str) -> None:
        with pytest.raises(ValueError, match="must not be blank"):
            SearchRequest(query=query)

    def test_top_k_is_capped(self) -> None:
        with pytest.raises(ValueError):
            SearchRequest(query="x", top_k=10_000)

    def test_empty_acl_filter_means_unfiltered(self) -> None:
        assert SearchFilters().acl_scopes == ()


class TestSettings:
    def test_defaults_are_safe_for_a_local_run(self) -> None:
        settings = Settings()
        assert settings.use_live_elasticsearch is False
        assert settings.budgets.reference_depth == 1
        assert settings.budgets.max_documents_per_workbook == 200_000

    def test_index_prefix_applies_to_every_index(self) -> None:
        settings = Settings(elasticsearch={"index_prefix": "dev-"})
        assert settings.elasticsearch.chunks_index == "dev-excel_chunks"
        assert settings.elasticsearch.structure_index == "dev-excel_structure"
        assert settings.elasticsearch.versions_index == "dev-excel_versions"

    def test_unknown_keys_are_refused(self) -> None:
        with pytest.raises(ValueError):
            Settings(nonsense=True)  # type: ignore[call-arg]

    def test_enums_cover_the_document_types(self) -> None:
        assert {ChunkType.ROW_GROUP, ChunkType.COLUMN, ChunkType.FORMULA_SUMMARY} <= set(ChunkType)
        assert {NodeType.CELL, NodeType.NAMED_RANGE, NodeType.FORMULA} <= set(NodeType)


def test_version_is_declared() -> None:
    assert __version__ == "0.1.0"
