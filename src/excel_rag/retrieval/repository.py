"""The one place Elasticsearch queries are built, and the one place they are filtered.

Every read the service performs -- the primary BM25 search over ``excel_chunks``, a direct
structure query, a span-intersection query, and the ``_mget`` that fetches related nodes -- passes
through this module. That is deliberate: the design's hard invariant is that the ACL filter and the
active-version pin apply to primary hits, to every related-node lookup and to every direct
structure query, and the one query builder that cannot forget is a single one.

The split between :meth:`Scope.filters` and :meth:`Scope.visible` is the crux. Server-side clauses
filter what the index returns. But ``_mget`` does **not** apply document-level ACLs (`_mget` is a
document fetch, not a query), so the caller has to re-check every document it hands back. The
repository does that in :meth:`Repository.get_nodes` and classifies what it dropped.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any

from ..es import ACL_FIELD, INDEX_MAPPINGS, ElasticsearchLike
from ..models import A1Range, ChunkDocument, SearchFilters, StructureDocument
from ..settings import Settings

#: How many active-version manifest rows a single unfiltered scope resolution reads. A deployment
#: has one row per workbook; this is an operational ceiling, not a per-request user bound.
MANIFEST_PAGE_SIZE = 1000


class UnknownWorkbook(LookupError):
    """Raised when a request names a workbook that has no active-version manifest."""

    def __init__(self, workbook_id: str) -> None:
        super().__init__(workbook_id)
        self.workbook_id = workbook_id


@dataclass(frozen=True)
class Scope:
    """The ACL scopes and active-version pins a query must honour.

    ``acl_scopes`` empty means "no scope filter", which is correct only where the deployment has no
    ACLs; the route layer decides whether that is the case. ``versions`` maps ``workbook_id`` to the
    single active version that may be read; a document from any other version is invisible.
    """

    acl_scopes: tuple[str, ...] = ()
    versions: Mapping[str, int] = field(default_factory=dict)

    def filters(self) -> list[dict[str, Any]]:
        """The server-side filter clauses: ACL scope first, then the version pin."""
        clauses: list[dict[str, Any]] = []
        if self.acl_scopes:
            clauses.append({"terms": {ACL_FIELD: list(self.acl_scopes)}})
        if self.versions:
            should = [
                {
                    "bool": {
                        "filter": [
                            {"term": {"workbook_id": workbook_id}},
                            {"term": {"version": version}},
                        ]
                    }
                }
                for workbook_id, version in sorted(self.versions.items())
            ]
            clauses.append({"bool": {"should": should, "minimum_should_match": 1}})
        return clauses

    def visible(self, document: Mapping[str, Any]) -> bool:
        """Re-check a document the index handed over without applying a filter (``_mget``).

        Returns ``False`` when the document's ACL scope does not intersect the caller's, or when its
        ``(workbook_id, version)`` is not the active pair. This is the check that keeps ``_mget``
        from leaking a related node the caller may not read.
        """
        if self.acl_scopes:
            document_scopes = document.get(ACL_FIELD) or ()
            if not set(document_scopes) & set(self.acl_scopes):
                return False
        if self.versions:
            workbook_id = document.get("workbook_id")
            version = document.get("version")
            if self.versions.get(str(workbook_id)) != version:
                return False
        return True


@dataclass(frozen=True)
class NodeFetch:
    """The result of an ``_mget`` for related nodes, classified by why a node is absent.

    ``nodes`` is what the caller may see. ``denied`` are ids that exist but were filtered by ACL or
    version -- they are never surfaced, because revealing them would leak existence. ``missing`` are
    ids the structure index does not hold at all; those are a data-integrity signal and are safe to
    report as unresolved.
    """

    nodes: tuple[StructureDocument, ...] = ()
    denied: tuple[str, ...] = ()
    missing: tuple[str, ...] = ()

    def by_id(self) -> dict[str, StructureDocument]:
        return {node.node_id: node for node in self.nodes}


class Repository:
    """Every Elasticsearch query the retrieval service makes, in one object."""

    def __init__(self, client: ElasticsearchLike, settings: Settings) -> None:
        self._client = client
        self._settings = settings

    # -- index management -----------------------------------------------------------------------
    def ensure_indices(self) -> None:
        """Create any of the three indices that does not yet exist. Idempotent."""
        for base, mapping in INDEX_MAPPINGS.items():
            name = self._settings.elasticsearch.index_name(base)
            if not self._client.indices_exists(name):
                self._client.create_index(name, mapping)

    @property
    def chunks_index(self) -> str:
        return self._settings.elasticsearch.chunks_index

    @property
    def structure_index(self) -> str:
        return self._settings.elasticsearch.structure_index

    @property
    def versions_index(self) -> str:
        return self._settings.elasticsearch.versions_index

    def es_requests(self) -> int:
        """How many Elasticsearch calls this client has served, for per-request accounting.

        Both the in-memory double and the live adapter record ``calls``; a client that does not is
        reported as zero rather than guessed at.
        """
        calls = getattr(self._client, "calls", None)
        if isinstance(calls, list):
            return len(calls)
        return 0

    def counts(self) -> dict[str, int]:
        """Document count per index, missing indices reported as zero (for ``/health``)."""
        result: dict[str, int] = {}
        for name in (self.chunks_index, self.structure_index, self.versions_index):
            if self._client.indices_exists(name):
                result[name] = int(self._client.count(name))
            else:
                result[name] = 0
        return result

    # -- scope ----------------------------------------------------------------------------------
    def resolve_scope(
        self, *, workbook_ids: Sequence[str] = (), acl_scopes: Sequence[str] = ()
    ) -> Scope:
        """Resolve the ACL scopes and active versions one request may read.

        A request that names workbooks must name ones that exist, or the answer would be a silent
        empty result for a typo: :class:`UnknownWorkbook` is raised instead. With no workbook filter
        every manifested workbook is pinned to its active version.
        """
        scopes = tuple(acl_scopes) or tuple(self._settings.default_acl_scope)
        versions = self._active_versions(tuple(workbook_ids))
        if workbook_ids:
            for workbook_id in workbook_ids:
                if workbook_id not in versions:
                    raise UnknownWorkbook(workbook_id)
        return Scope(acl_scopes=scopes, versions=versions)

    def _active_versions(self, workbook_ids: tuple[str, ...]) -> dict[str, int]:
        query: dict[str, Any]
        if workbook_ids:
            query = {"terms": {"workbook_id": list(workbook_ids)}}
            size = max(len(workbook_ids) * 2, 10)
        else:
            query = {"match_all": {}}
            size = MANIFEST_PAGE_SIZE
        result = self._client.search(self.versions_index, query, size=size)
        versions: dict[str, int] = {}
        for hit in _hits(result):
            source = _source(hit)
            versions[str(source["workbook_id"])] = int(source["active_version"])
        return versions

    def active_version(self, workbook_id: str) -> int | None:
        """The single active version for a workbook, or ``None`` when it is unknown."""
        return self._active_versions((workbook_id,)).get(workbook_id)

    # -- primary search -------------------------------------------------------------------------
    def search_chunks(
        self,
        query: str,
        *,
        filters: SearchFilters,
        scope: Scope,
        size: int,
    ) -> list[tuple[ChunkDocument, float]]:
        """BM25-style match over title/content/headers, filtered by scope, version and facets.

        The match clauses are ``should`` (a chunk need only match one field), the ACL and version
        pins are ``filter`` (they must hold and do not contribute to the score), and the facet
        filters are ``filter`` too.
        """
        body: dict[str, Any] = {}
        should: list[dict[str, Any]] = []
        if query.strip():
            should = [
                {"match": {"title": query}},
                {"match": {"content": query}},
                {"match": {"headers": query}},
            ]
        if should:
            body["should"] = should
            body["minimum_should_match"] = 1
        filter_clauses = scope.filters() + self._facet_filters(filters)
        if filter_clauses:
            body["filter"] = filter_clauses
        es_query: dict[str, Any] = {"bool": body} if body else {"match_all": {}}
        result = self._client.search(self.chunks_index, es_query, size=size)
        scored: list[tuple[ChunkDocument, float]] = []
        for hit in _hits(result):
            chunk = ChunkDocument.model_validate(_source(hit))
            scored.append((chunk, float(hit.get("_score") or 0.0)))
        return scored

    @staticmethod
    def _facet_filters(filters: SearchFilters) -> list[dict[str, Any]]:
        clauses: list[dict[str, Any]] = []
        if filters.workbook_ids:
            clauses.append({"terms": {"workbook_id": list(filters.workbook_ids)}})
        if filters.sheet_names:
            clauses.append({"terms": {"sheet_name": list(filters.sheet_names)}})
        if filters.chunk_types:
            clauses.append({"terms": {"chunk_type": [str(kind) for kind in filters.chunk_types]}})
        return clauses

    # -- structure ------------------------------------------------------------------------------
    def query_structure(
        self,
        *,
        scope: Scope,
        workbook_ids: Sequence[str] = (),
        sheet_names: Sequence[str] = (),
        node_types: Sequence[str] = (),
        limit: int = 200,
    ) -> list[StructureDocument]:
        """Structural nodes for a workbook, filtered by scope, version and facets."""
        clauses = scope.filters()
        if workbook_ids:
            clauses.append({"terms": {"workbook_id": list(workbook_ids)}})
        if sheet_names:
            clauses.append({"terms": {"sheet_name": list(sheet_names)}})
        if node_types:
            clauses.append({"terms": {"node_type": list(node_types)}})
        es_query: dict[str, Any] = {"bool": {"filter": clauses}} if clauses else {"match_all": {}}
        return self._structure_search(es_query, limit=limit)

    def query_range(
        self,
        *,
        scope: Scope,
        workbook_id: str,
        sheet_name: str,
        region: A1Range,
        node_types: Sequence[str] = (),
        limit: int = 50,
    ) -> list[StructureDocument]:
        """Nodes whose ``row_span``/``column_span`` intersect ``region``.

        Answered by ``integer_range`` intersection on both axes -- never by walking cells. The
        rectangle arrives as two ``range`` clauses; a node matches only when both spans overlap.
        """
        clauses = scope.filters()
        clauses.extend(
            [
                {"term": {"workbook_id": workbook_id}},
                {"term": {"sheet_name": sheet_name}},
                {"range": {"row_span": {"gte": region.min_row, "lte": region.max_row}}},
                {"range": {"column_span": {"gte": region.min_col, "lte": region.max_col}}},
            ]
        )
        if node_types:
            clauses.append({"terms": {"node_type": list(node_types)}})
        return self._structure_search({"bool": {"filter": clauses}}, limit=limit)

    def _structure_search(self, es_query: dict[str, Any], *, limit: int) -> list[StructureDocument]:
        result = self._client.search(self.structure_index, es_query, size=limit)
        return [StructureDocument.model_validate(_source(hit)) for hit in _hits(result)]

    # -- related nodes --------------------------------------------------------------------------
    def get_nodes(self, node_ids: Sequence[str], *, scope: Scope) -> NodeFetch:
        """Fetch structural nodes by id, applying the ACL and version filters ``_mget`` skips.

        Ids are de-duplicated while preserving order. Every returned document is re-checked with
        :meth:`Scope.visible`: a document the caller may not read is dropped into ``denied`` and
        never appears in ``nodes``.
        """
        unique = list(dict.fromkeys(node_ids))
        if not unique:
            return NodeFetch()
        found = self._client.mget_documents(self.structure_index, unique)
        by_id: dict[str, Mapping[str, Any]] = {}
        for stored in found:
            by_id[str(stored["node_id"])] = stored
        visible: list[StructureDocument] = []
        denied: list[str] = []
        missing: list[str] = []
        for node_id in unique:
            document = by_id.get(node_id)
            if document is None:
                missing.append(node_id)
            elif scope.visible(document):
                visible.append(StructureDocument.model_validate(document))
            else:
                denied.append(node_id)
        return NodeFetch(nodes=tuple(visible), denied=tuple(denied), missing=tuple(missing))


# -------------------------------------------------------------------------------------------------
# Small helpers for the Elasticsearch response envelope
# -------------------------------------------------------------------------------------------------
def _hits(result: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    hits = result.get("hits", {})
    inner = hits.get("hits", []) if isinstance(hits, Mapping) else []
    return [hit for hit in inner if isinstance(hit, Mapping)]


def _source(hit: Mapping[str, Any]) -> Mapping[str, Any]:
    source = hit.get("_source")
    if not isinstance(source, Mapping):  # pragma: no cover - defensive against a malformed response
        raise ValueError("Elasticsearch hit carried no _source")
    return source


__all__ = ["MANIFEST_PAGE_SIZE", "NodeFetch", "Repository", "Scope", "UnknownWorkbook"]
