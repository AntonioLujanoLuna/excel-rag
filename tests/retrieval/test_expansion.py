"""Expansion tests: bounded BFS, ACL re-check per level, cycle termination, explicit truncation.

The graph is small and fixed, so the assertions are about *what the traversal did*: which nodes came
back, which were denied, which could not be resolved, and exactly how a budget stop is reported.
"""

from __future__ import annotations

from conftest import (
    FINANCE,
    N_CELL,
    N_FORMULA,
    N_FORMULA2,
    N_GHOST,
    N_REGION,
    N_SECRET,
    WB,
)

from excel_rag.models import UnresolvedReason
from excel_rag.retrieval import Repository, Scope, expand
from excel_rag.settings import Settings

FINANCE_SCOPE = Scope(acl_scopes=(FINANCE,), versions={WB: 1})


def _expand(
    repository: Repository,
    seeds: list[str],
    *,
    max_depth: int = 1,
    max_nodes: int = 100,
    max_bytes: int = 1_000_000,
    deadline: float | None = None,
):
    return expand(
        seed_node_ids=seeds,
        repository=repository,
        scope=FINANCE_SCOPE,
        max_depth=max_depth,
        max_nodes=max_nodes,
        max_bytes=max_bytes,
        deadline=deadline,
    )


class TestTraversal:
    def test_depth_zero_returns_only_the_seeds(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=0)
        assert set(result.nodes) == {N_REGION}
        assert result.nodes[N_REGION].depth == 0

    def test_depth_one_follows_references(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=1)
        assert {N_REGION, N_CELL, N_FORMULA} <= set(result.nodes)
        assert result.nodes[N_CELL].depth == 1
        assert result.nodes[N_CELL].cached_value == 0.05

    def test_depth_two_reaches_the_second_hop(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=3)
        assert N_FORMULA2 in result.nodes

    def test_acl_applies_at_every_level(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=1)
        assert N_SECRET not in result.nodes
        assert result.denied >= 1

    def test_a_dangling_reference_is_surfaced(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=1)
        texts = {ref.reference_text for ref in result.unresolved}
        assert N_GHOST in texts
        assert any(ref.reason is UnresolvedReason.MALFORMED for ref in result.unresolved)

    def test_the_documents_own_unresolved_references_are_returned(
        self, settings: Settings, client
    ) -> None:
        result = _expand(Repository(client, settings), [N_FORMULA], max_depth=0)
        reasons = {ref.reason for ref in result.unresolved}
        assert UnresolvedReason.INDIRECT in reasons


class TestCycles:
    def test_an_a_to_b_to_a_cycle_terminates(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_FORMULA], max_depth=3)
        assert set(result.nodes) == {N_FORMULA, N_FORMULA2}
        assert result.nodes[N_FORMULA].depth == 0
        assert result.nodes[N_FORMULA2].depth == 1

    def test_a_cycle_of_one_seed_still_terminates(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_FORMULA2], max_depth=3)
        assert set(result.nodes) == {N_FORMULA2, N_FORMULA}


class TestBudgets:
    def test_node_budget_truncates_and_counts(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=2, max_nodes=2)
        assert N_REGION in result.nodes
        assert len(result.nodes) == 1 + 2, "the seed plus two related nodes"
        assert result.truncation.truncated
        assert "max_related_nodes" in (result.truncation.reason or "")
        assert result.truncation.dropped_nodes >= 1

    def test_the_node_budget_never_drops_a_seed(self, settings: Settings, client) -> None:
        """Seeds are the hits' own nodes: a zero related-node budget still returns them."""
        result = _expand(
            Repository(client, settings), [N_REGION, N_FORMULA2], max_depth=1, max_nodes=0
        )
        assert set(result.nodes) == {N_REGION, N_FORMULA2}
        assert "max_related_nodes" in (result.truncation.reason or "")

    def test_byte_budget_truncates_and_counts(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=0, max_bytes=10)
        assert result.nodes == {}
        assert result.truncation.truncated
        assert "max_payload_bytes" in (result.truncation.reason or "")
        assert result.truncation.dropped_nodes >= 1

    def test_depth_bound_is_reported(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=1)
        assert result.truncation.truncated
        assert "depth_limit" in (result.truncation.reason or "")
        assert result.truncation.depth_limit == 1

    def test_timeout_is_reported(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_REGION], max_depth=2, deadline=0.0)
        assert "timeout" in (result.truncation.reason or "")
        assert result.truncation.dropped_nodes >= 1

    def test_no_truncation_when_the_graph_is_exhausted(self, settings: Settings, client) -> None:
        result = _expand(Repository(client, settings), [N_FORMULA], max_depth=3)
        assert not result.truncation.truncated
        assert result.truncation.bytes_returned > 0


def test_expansion_of_nothing_is_empty(settings: Settings, client) -> None:
    result = _expand(Repository(client, settings), [], max_depth=2)
    assert result.nodes == {}
    assert not result.truncation.truncated
