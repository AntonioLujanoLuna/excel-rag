"""Bounded breadth-first expansion over the ``references`` edges of ``excel_structure``.

A formula graph has cycles -- ``A1`` reads ``B1`` reads ``A1`` -- so the traversal keeps a visited
set and terminates, returning each node exactly once; the design's dependency graph is a graph, not
a tree, and a naive walk would loop forever. Because the response schema carries no cycle field,
the termination is a tested behaviour rather than a payload field: a cycle is de-duplicated, and
``tests/retrieval`` asserts an ``A -> B -> A`` pair yields exactly two nodes.

Every level re-applies the ACL and version filters, because a related node is fetched by ``_mget``
(see :mod:`excel_rag.retrieval.repository`), which enforces neither. **Silent omission is a bug**:
anything the budgets drop is counted in :class:`~excel_rag.models.TruncationInfo`, and a reference
that resolves to no node is surfaced in ``unresolved_references`` rather than dropped. The one
deliberate exception is a node the caller may not read: it is excluded without being named, because
surfacing it would leak its existence.
"""

from __future__ import annotations

import json
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

from ..models import (
    NodePayload,
    StructureDocument,
    TruncationInfo,
    UnresolvedReason,
    UnresolvedReference,
)
from .repository import Repository, Scope


@dataclass(frozen=True)
class ExpansionResult:
    """What the traversal returned, what it could not resolve, and what the budgets dropped."""

    nodes: dict[str, NodePayload] = field(default_factory=dict)
    unresolved: tuple[UnresolvedReference, ...] = ()
    truncation: TruncationInfo = field(default_factory=TruncationInfo)
    denied: int = 0
    """Count of ids that exist but were filtered by ACL or version. Never surfaced, only counted."""


def expand(
    *,
    seed_node_ids: Sequence[str],
    repository: Repository,
    scope: Scope,
    max_depth: int,
    max_nodes: int,
    max_bytes: int,
    deadline: float | None = None,
    clock: Callable[[], float] = time.monotonic,
) -> ExpansionResult:
    """Expand the reference graph from ``seed_node_ids`` under the request budgets.

    ``max_depth`` is the number of reference hops to follow; ``0`` returns the seed nodes only.
    ``deadline`` is an absolute :func:`time.monotonic` instant, checked between levels.
    """
    nodes: dict[str, NodePayload] = {}
    unresolved: list[UnresolvedReference] = []
    seen_unresolved: set[tuple[str, str]] = set()
    reasons: list[str] = []
    dropped = 0
    denied = 0
    bytes_used = 0
    depth_limit: int | None = None

    seeds = list(dict.fromkeys(seed_node_ids))
    attempted: set[str] = set(seeds)
    current: list[str] = seeds
    level = 0

    def note_unresolved(reference: UnresolvedReference) -> None:
        key = (reference.reference_text, str(reference.reason))
        if key not in seen_unresolved:
            seen_unresolved.add(key)
            unresolved.append(reference)

    while current:
        if deadline is not None and clock() >= deadline:
            reasons.append("timeout")
            dropped += len(current)
            break

        fetch = repository.get_nodes(current, scope=scope)
        denied += len(fetch.denied)
        present = fetch.by_id()
        for missing_id in fetch.missing:
            note_unresolved(
                UnresolvedReference(
                    reference_text=missing_id,
                    reason=UnresolvedReason.MALFORMED,
                    detail="referenced node is not present in the structure index",
                )
            )

        budget_hit = False
        for document in fetch.nodes:
            if len(nodes) >= max_nodes:
                reasons.append("max_related_nodes")
                dropped += 1
                budget_hit = True
                continue
            payload = node_payload(document, level)
            size = _payload_bytes(payload)
            if bytes_used + size > max_bytes:
                reasons.append("max_payload_bytes")
                dropped += 1
                budget_hit = True
                continue
            nodes[document.node_id] = payload
            bytes_used += size
            for unresolved_reference in document.unresolved_references:
                note_unresolved(unresolved_reference)
        if budget_hit:
            break

        if level >= max_depth:
            if max_depth > 0:
                remaining = _unexpanded(present, attempted)
                if remaining:
                    reasons.append("depth_limit")
                    dropped += len(remaining)
                    depth_limit = max_depth
            break

        next_frontier: list[str] = []
        for document in fetch.nodes:
            for reference in document.references:
                if not reference.resolved:
                    continue
                target = reference.target_node_id
                if target in attempted:
                    # Already scheduled or visited: this is a duplicate edge or a cycle.
                    continue
                attempted.add(target)
                next_frontier.append(target)
        current = next_frontier
        level += 1

    reason = "; ".join(dict.fromkeys(reasons)) if reasons else None
    truncation = TruncationInfo(
        truncated=bool(reasons),
        reason=reason,
        dropped_nodes=dropped,
        depth_limit=depth_limit,
        bytes_returned=bytes_used,
    )
    return ExpansionResult(
        nodes=nodes,
        unresolved=tuple(unresolved),
        truncation=truncation,
        denied=denied,
    )


def _unexpanded(present: dict[str, StructureDocument], attempted: set[str]) -> list[str]:
    """Reference targets at the deepest processed level that the depth bound left unvisited."""
    remaining: list[str] = []
    for document in present.values():
        for reference in document.references:
            if reference.resolved and reference.target_node_id not in attempted:
                remaining.append(reference.target_node_id)
    return remaining


def node_payload(document: StructureDocument, depth: int = 0) -> NodePayload:
    """Project a stored structural node onto the response payload, tagging its expansion depth."""
    return NodePayload(
        node_id=document.node_id,
        node_type=document.node_type,
        sheet=document.sheet_name,
        a1_range=document.a1_range,
        value=document.cached_value,
        display_value=document.display_value,
        formula=document.formula,
        cached_value=document.cached_value,
        table_name=document.table_name,
        column_name=document.column_name,
        named_range=document.named_range,
        references=document.references,
        depth=depth,
    )


def _payload_bytes(payload: NodePayload) -> int:
    return len(json.dumps(payload.model_dump(mode="json"), ensure_ascii=False).encode("utf-8"))


__all__ = ["ExpansionResult", "expand", "node_payload"]
