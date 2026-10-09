"""The live Elasticsearch adapter, and the production ``knn`` query builder.

This is the path a real cluster runs; it is written so the production shape exists even though no
cluster is reachable from the machine the rest of this repo was developed on. **The live path is
UNMEASURED here** -- the benchmarks run against the in-memory double, and no number in this repo
comes from an Elasticsearch cluster.

``elasticsearch`` is an optional extra, so the import is lazy: this module imports cleanly with the
package absent, and only raises when a connection is actually attempted.
"""

from __future__ import annotations

import importlib
from collections.abc import Iterable, Mapping, Sequence
from typing import Any

from .settings import ElasticsearchSettings


def build_knn_query(
    *,
    field: str,
    query_vector: Sequence[float],
    k: int,
    num_candidates: int,
    filter: Sequence[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """The production ``knn`` clause for a dense-vector search over ``excel_chunks``.

    ``filter`` carries the same ACL and version clauses the lexical query uses, so a vector search
    cannot cross an access scope or a workbook version. The in-memory double refuses ``knn`` on
    purpose; this builder is the live counterpart.
    """
    body: dict[str, Any] = {
        "field": field,
        "query_vector": list(query_vector),
        "k": k,
        "num_candidates": num_candidates,
    }
    if filter:
        body["filter"] = [dict(clause) for clause in filter]
    return {"knn": body}


class LiveElasticsearch:
    """An :class:`~excel_rag.es.ElasticsearchLike` over a real ``elasticsearch.Elasticsearch``.

    ``client`` may be injected (tests do this to exercise the adapter without a cluster); otherwise
    the real client is built lazily on first use.
    """

    def __init__(self, settings: ElasticsearchSettings, client: Any | None = None) -> None:
        self._settings = settings
        self._client = client
        #: Mirrors the in-memory double so per-request Elasticsearch call accounting is identical.
        self.calls: list[tuple[str, str]] = []

    @property
    def client(self) -> Any:
        if self._client is None:
            self._client = connect(self._settings)
        return self._client

    # -- index management -------------------------------------------------------------------
    def indices_exists(self, index: str) -> bool:
        self.calls.append(("indices_exists", index))
        return bool(self.client.indices.exists(index=index))

    def create_index(self, index: str, mappings: Mapping[str, Any]) -> None:
        self.calls.append(("create_index", index))
        self.client.indices.create(index=index, mappings=dict(mappings))

    def delete_index(self, index: str) -> None:
        self.calls.append(("delete_index", index))
        self.client.indices.delete(index=index)

    # -- documents --------------------------------------------------------------------------
    def index_document(
        self, index: str, document_id: str, document: Mapping[str, Any], *, refresh: bool = False
    ) -> None:
        self.calls.append(("index_document", index))
        self.client.index(index=index, id=document_id, document=dict(document), refresh=refresh)

    def bulk_index(
        self,
        index: str,
        documents: Iterable[tuple[str, Mapping[str, Any]]],
        *,
        refresh: bool = False,
    ) -> int:
        self.calls.append(("bulk_index", index))
        actions = [
            {"_index": index, "_id": document_id, "_source": dict(document)}
            for document_id, document in documents
        ]
        if not actions:
            return 0
        helpers = _helpers_module()
        if helpers is not None:
            result = helpers.bulk(self.client, actions, refresh=refresh, raise_on_error=True)
            return int(result[0]) if isinstance(result, tuple) else int(result)
        for action in actions:
            self.client.index(
                index=action["_index"],
                id=action["_id"],
                document=action["_source"],
                refresh=refresh,
            )
        return len(actions)

    def get_document(self, index: str, document_id: str) -> Mapping[str, Any] | None:
        self.calls.append(("get_document", index))
        try:
            response = self.client.get(index=index, id=document_id)
        except Exception as exc:
            if type(exc).__name__ == "NotFoundError":
                return None
            raise
        return _source_of(_as_mapping(response))

    def mget_documents(
        self, index: str, document_ids: Sequence[str]
    ) -> Sequence[Mapping[str, Any]]:
        self.calls.append(("mget_documents", index))
        if not document_ids:
            return []
        response = _as_mapping(self.client.mget(index=index, ids=list(document_ids)))
        documents: list[Mapping[str, Any]] = []
        for document in response.get("docs", []) or []:
            if isinstance(document, Mapping) and document.get("found"):
                source = document.get("_source")
                if isinstance(source, Mapping):
                    documents.append(source)
        return documents

    # -- search -----------------------------------------------------------------------------
    def search(
        self,
        index: str,
        query: Mapping[str, Any],
        *,
        size: int = 10,
        source_includes: Sequence[str] | None = None,
    ) -> Mapping[str, Any]:
        self.calls.append(("search", index))
        kwargs: dict[str, Any] = {"index": index, "size": size}
        body = dict(query)
        knn = body.pop("knn", None)
        if knn is not None:
            kwargs["knn"] = knn
        if body:
            kwargs["query"] = body
        if source_includes is not None:
            kwargs["source_includes"] = list(source_includes)
        return _as_mapping(self.client.search(**kwargs))

    def delete_by_query(self, index: str, query: Mapping[str, Any]) -> int:
        self.calls.append(("delete_by_query", index))
        response = _as_mapping(self.client.delete_by_query(index=index, query=dict(query)))
        return int(response.get("deleted", 0))

    def count(self, index: str, query: Mapping[str, Any] | None = None) -> int:
        self.calls.append(("count", index))
        kwargs: dict[str, Any] = {"index": index}
        if query is not None:
            kwargs["query"] = query
        response = _as_mapping(self.client.count(**kwargs))
        return int(response.get("count", 0))


def connect(settings: ElasticsearchSettings) -> Any:
    """Build a real ``elasticsearch.Elasticsearch``, importing the optional extra lazily."""
    try:
        module = importlib.import_module("elasticsearch")
    except ModuleNotFoundError as exc:  # pragma: no cover - exercised only without the extra
        raise RuntimeError("the live client needs the 'es' extra: install excel-rag[es]") from exc
    kwargs: dict[str, Any] = {
        "hosts": list(settings.urls),
        "request_timeout": settings.request_timeout_seconds,
    }
    if settings.username and settings.password is not None:
        kwargs["basic_auth"] = (settings.username, settings.password.get_secret_value())
    client: Any = module.Elasticsearch(**kwargs)
    return client


def _helpers_module() -> Any | None:
    try:
        return importlib.import_module("elasticsearch.helpers")
    except ModuleNotFoundError:
        return None


def _as_mapping(response: Any) -> Mapping[str, Any]:
    body = getattr(response, "body", None)
    if body is None and isinstance(response, Mapping):
        body = response
    if isinstance(body, Mapping):
        return body
    converted: Mapping[str, Any] = dict(response)
    return converted


def _source_of(response: Mapping[str, Any]) -> Mapping[str, Any] | None:
    source = response.get("_source")
    return source if isinstance(source, Mapping) else None


__all__ = ["LiveElasticsearch", "build_knn_query", "connect"]
