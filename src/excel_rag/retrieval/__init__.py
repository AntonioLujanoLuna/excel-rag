"""Retrieval internals: the one place Elasticsearch queries are built and filtered.

The package is layered so the ACL and version filters cannot be forgotten in a corner:

* :mod:`excel_rag.retrieval.repository` wraps an :class:`~excel_rag.es.ElasticsearchLike` client and
  owns **every** query. Primary chunk search, direct structure queries, range-intersection queries
  and related-node ``_mget`` lookups all go through a single
  :class:`~excel_rag.retrieval.repository.Scope`, which builds the server-side filters and re-checks
  ``_mget`` results the index does not filter.
* :mod:`excel_rag.retrieval.fusion` re-ranks the candidate window the index returned.
* :mod:`excel_rag.retrieval.expansion` walks the ``references`` edges under the request budgets.
* :mod:`excel_rag.retrieval.service` orchestrates the three into a
  :class:`~excel_rag.models.SearchResponse`.
"""

from __future__ import annotations

from .expansion import ExpansionResult, expand, node_payload
from .fusion import Candidate, RankedCandidate, cosine_similarity, rank
from .repository import NodeFetch, Repository, Scope, UnknownWorkbook
from .service import RetrievalService

__all__ = [
    "Candidate",
    "ExpansionResult",
    "NodeFetch",
    "RankedCandidate",
    "Repository",
    "RetrievalService",
    "Scope",
    "UnknownWorkbook",
    "cosine_similarity",
    "expand",
    "node_payload",
    "rank",
]
