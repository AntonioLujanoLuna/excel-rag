"""Repository tests: the one place ACL and version filtering is enforced.

The point of these is that no query path can forget the filters. They exercise the server-side
clauses (primary search, structure query, range intersection) and, crucially, the caller-side
re-check that ``_mget`` requires.
"""

from __future__ import annotations

import pytest
from conftest import (
    C_HR,
    C_OTHER,
    C_REVENUE,
    C_SECRET,
    C_STALE,
    EXEC,
    FINANCE,
    N_CELL,
    N_FORECAST_CELL,
    N_GHOST,
    N_OTHER_TABLE,
    N_REGION,
    N_SECRET,
    OTHER_WB,
    WB,
    range_of,
)

from excel_rag.es import INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS
from excel_rag.fake_es import InMemoryElasticsearch
from excel_rag.models import ChunkType, SearchFilters
from excel_rag.retrieval import Repository, Scope, UnknownWorkbook
from excel_rag.settings import Settings


class TestIndexManagement:
    def test_ensure_indices_creates_the_three_indices(self, settings: Settings) -> None:
        client = InMemoryElasticsearch()
        Repository(client, settings).ensure_indices()
        assert client.indices_exists(INDEX_CHUNKS)
        assert client.indices_exists(INDEX_STRUCTURE)
        assert client.indices_exists(INDEX_VERSIONS)

    def test_ensure_indices_is_idempotent(self, settings: Settings, empty_client) -> None:
        Repository(empty_client, settings).ensure_indices()
        assert empty_client.count(INDEX_CHUNKS) == 0

    def test_counts_reports_every_index(self, settings: Settings, client) -> None:
        counts = Repository(client, settings).counts()
        assert set(counts) == {INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS}
        assert counts[INDEX_CHUNKS] > 0


class TestScope:
    def test_resolves_active_version_and_scope(self, settings: Settings, client) -> None:
        scope = Repository(client, settings).resolve_scope(
            workbook_ids=(WB,), acl_scopes=(FINANCE,)
        )
        assert scope.versions == {WB: 1}
        assert scope.acl_scopes == (FINANCE,)

    def test_unknown_workbook_is_a_lookup_error(self, settings: Settings, client) -> None:
        with pytest.raises(UnknownWorkbook) as excinfo:
            Repository(client, settings).resolve_scope(workbook_ids=("nope",))
        assert excinfo.value.workbook_id == "nope"

    def test_no_filter_pins_every_manifested_workbook(self, settings: Settings, client) -> None:
        scope = Repository(client, settings).resolve_scope()
        assert scope.versions == {WB: 1, OTHER_WB: 1}

    def test_empty_acl_scopes_fall_back_to_the_default(self, client) -> None:
        settings = Settings(default_acl_scope=(FINANCE,))
        scope = Repository(client, settings).resolve_scope()
        assert scope.acl_scopes == (FINANCE,)

    def test_visible_rejects_wrong_scope_and_wrong_version(self) -> None:
        scope = Scope(acl_scopes=(FINANCE,), versions={WB: 1})
        assert scope.visible({"workbook_id": WB, "version": 1, "acl_scope": [FINANCE]})
        assert not scope.visible({"workbook_id": WB, "version": 1, "acl_scope": [EXEC]})
        assert not scope.visible({"workbook_id": WB, "version": 2, "acl_scope": [FINANCE]})
        assert not scope.visible({"workbook_id": OTHER_WB, "version": 1, "acl_scope": [FINANCE]})


class TestPrimarySearch:
    def test_matches_and_excludes_the_stale_version(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,))
        scored = repository.search_chunks("revenue", filters=SearchFilters(), scope=scope, size=50)
        ids = {chunk.id for chunk, _ in scored}
        assert C_REVENUE in ids
        assert C_STALE not in ids, "the active-version pin must drop version 2"

    def test_acl_filter_hides_other_scopes(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        ids = {
            chunk.id
            for chunk, _ in repository.search_chunks(
                "revenue", filters=SearchFilters(), scope=scope, size=50
            )
        }
        assert C_SECRET not in ids
        assert C_HR not in ids

    def test_filters_by_workbook_sheet_and_type(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(acl_scopes=(FINANCE,))
        by_workbook = repository.search_chunks(
            "revenue", filters=SearchFilters(workbook_ids=(OTHER_WB,)), scope=scope, size=50
        )
        assert {chunk.id for chunk, _ in by_workbook} == {C_OTHER}

        by_sheet = repository.search_chunks(
            "revenue", filters=SearchFilters(sheet_names=("Forecast",)), scope=scope, size=50
        )
        assert C_OTHER not in {chunk.id for chunk, _ in by_sheet}

        by_type = repository.search_chunks(
            "revenue",
            filters=SearchFilters(chunk_types=(ChunkType.FORMULA_SUMMARY,)),
            scope=scope,
            size=50,
        )
        assert {chunk.chunk_type for chunk, _ in by_type} == {ChunkType.FORMULA_SUMMARY}

    def test_blank_query_falls_back_to_match_all_over_the_scope(
        self, settings: Settings, client
    ) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        scored = repository.search_chunks("   ", filters=SearchFilters(), scope=scope, size=50)
        assert scored, "an all-blank query still returns the scope, it is the route that refuses it"


class TestGetNodes:
    def test_acl_and_version_are_rechecked_after_mget(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = Scope(acl_scopes=(FINANCE,), versions={WB: 1})
        fetch = repository.get_nodes([N_REGION, N_CELL, N_SECRET, N_GHOST], scope=scope)
        assert {node.node_id for node in fetch.nodes} == {N_REGION, N_CELL}
        assert fetch.denied == (N_SECRET,), "a node the caller may not read must not be returned"
        assert fetch.missing == (N_GHOST,)

    def test_visible_scope_returns_the_secret(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = Scope(acl_scopes=(EXEC,), versions={WB: 1})
        fetch = repository.get_nodes([N_SECRET], scope=scope)
        assert [node.node_id for node in fetch.nodes] == [N_SECRET]

    def test_empty_request_is_empty(self, settings: Settings, client) -> None:
        fetch = Repository(client, settings).get_nodes([], scope=Scope())
        assert fetch.nodes == () and fetch.denied == () and fetch.missing == ()


class TestStructureQueries:
    def test_query_structure_filters_by_workbook_and_sheet(
        self, settings: Settings, client
    ) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        nodes = repository.query_structure(
            scope=scope, workbook_ids=(WB,), sheet_names=("Forecast",)
        )
        assert {node.node_id for node in nodes} >= {N_REGION, N_FORECAST_CELL}
        assert N_OTHER_TABLE not in {node.node_id for node in nodes}

    def test_query_structure_applies_acl(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = Scope(acl_scopes=(EXEC,), versions={WB: 1})
        nodes = repository.query_structure(scope=scope, workbook_ids=(WB,))
        assert {node.node_id for node in nodes} == {N_SECRET}

    def test_range_intersection_uses_spans(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        nodes = repository.query_range(
            scope=scope,
            workbook_id=WB,
            sheet_name="Forecast",
            region=range_of("Forecast", "A12:C20"),
        )
        assert {node.node_id for node in nodes} == {N_REGION, N_FORECAST_CELL}

    def test_range_intersection_honours_node_type_and_acl(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        only_cells = repository.query_range(
            scope=scope,
            workbook_id=WB,
            sheet_name="Forecast",
            region=range_of("Forecast", "A12:C20"),
            node_types=("cell",),
        )
        assert {node.node_id for node in only_cells} == {N_FORECAST_CELL}

        exec_scope = Scope(acl_scopes=(EXEC,), versions={WB: 1})
        denied = repository.query_range(
            scope=exec_scope,
            workbook_id=WB,
            sheet_name="Forecast",
            region=range_of("Forecast", "A12:C20"),
        )
        assert denied == []

    def test_disjoint_range_returns_nothing(self, settings: Settings, client) -> None:
        repository = Repository(client, settings)
        scope = repository.resolve_scope(workbook_ids=(WB,), acl_scopes=(FINANCE,))
        nodes = repository.query_range(
            scope=scope, workbook_id=WB, sheet_name="Actuals", region=range_of("Actuals", "A1:B2")
        )
        assert nodes == []
