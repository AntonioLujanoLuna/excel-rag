"""Live-adapter tests against a stand-in client.

No cluster is reachable here, so the adapter is exercised with a fake that mimics the
``elasticsearch.Elasticsearch`` surface. This proves the delegation shape; it does **not** prove the
adapter works against a real cluster, and the docs say so.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_rag import live as live_module
from excel_rag.live import LiveElasticsearch, build_knn_query, connect
from excel_rag.settings import ElasticsearchSettings, Settings


class _NotFound(Exception):
    pass


class _NotFoundByModule(Exception):
    pass


# Mirror the library's class name so the adapter's name-based catch matches.
_NotFoundByModule.__name__ = "NotFoundError"


class _IndicesClient:
    def __init__(self) -> None:
        self.existing: set[str] = set()

    def exists(self, index: str) -> bool:
        return index in self.existing

    def create(self, index: str, mappings: dict[str, Any]) -> dict[str, Any]:
        self.existing.add(index)
        return {"acknowledged": True}

    def delete(self, index: str) -> dict[str, Any]:
        self.existing.discard(index)
        return {"acknowledged": True}


class _FakeReal:
    def __init__(self) -> None:
        self.indices = _IndicesClient()
        self.docs: dict[tuple[str, str], dict[str, Any]] = {}
        self.calls: list[dict[str, Any]] = []

    def index(self, *, index: str, id: str, document: dict[str, Any], refresh: bool = False) -> Any:
        self.docs[(index, id)] = dict(document)
        return {"result": "created"}

    def get(self, *, index: str, id: str) -> Any:
        if (index, id) not in self.docs:
            raise _NotFoundByModule(f"missing {id}")
        return {"_source": self.docs[(index, id)]}

    def mget(self, *, index: str, ids: list[str]) -> Any:
        docs = []
        for document_id in ids:
            if (index, document_id) in self.docs:
                docs.append(
                    {"_id": document_id, "found": True, "_source": self.docs[(index, document_id)]}
                )
            else:
                docs.append({"_id": document_id, "found": False})
        return {"docs": docs}

    def search(self, **kwargs: Any) -> Any:
        self.calls.append(kwargs)
        return {
            "hits": {
                "total": {"value": 1},
                "hits": [{"_id": "x", "_score": 1.0, "_source": {"a": 1}}],
            }
        }

    def count(self, *, index: str, query: Any = None) -> Any:
        return {"count": 7}

    def delete_by_query(self, *, index: str, query: Any, refresh: bool = False) -> Any:
        self.last_delete_refresh = refresh
        return {"deleted": 2}


@pytest.fixture
def fake() -> _FakeReal:
    return _FakeReal()


@pytest.fixture
def live(fake: _FakeReal) -> LiveElasticsearch:
    return LiveElasticsearch(ElasticsearchSettings(), client=fake)


class TestBuildKnnQuery:
    def test_shape_carries_vector_and_candidates(self) -> None:
        query = build_knn_query(
            field="embedding", query_vector=(1.0, 2.0), k=10, num_candidates=100
        )
        assert query == {
            "knn": {
                "field": "embedding",
                "query_vector": [1.0, 2.0],
                "k": 10,
                "num_candidates": 100,
            }
        }

    def test_filter_is_carried_through(self) -> None:
        query = build_knn_query(
            field="embedding",
            query_vector=(1.0,),
            k=5,
            num_candidates=50,
            filter=[{"terms": {"acl_scope": ["finance-team"]}}],
        )
        assert query["knn"]["filter"] == [{"terms": {"acl_scope": ["finance-team"]}}]


class TestIndexManagement:
    def test_exists_create_delete(self, live: LiveElasticsearch) -> None:
        assert live.indices_exists("i") is False
        live.create_index("i", {"properties": {}})
        assert live.indices_exists("i") is True
        live.delete_index("i")
        assert live.indices_exists("i") is False


class TestDocuments:
    def test_index_get_and_missing(self, live: LiveElasticsearch) -> None:
        live.index_document("i", "d1", {"a": 1})
        assert live.get_document("i", "d1") == {"a": 1}
        assert live.get_document("i", "absent") is None

    def test_get_reraises_an_unexpected_error(
        self, live: LiveElasticsearch, fake: _FakeReal
    ) -> None:
        def boom(*, index: str, id: str) -> Any:
            raise _NotFound("genuine failure")

        fake.get = boom  # type: ignore[method-assign]
        with pytest.raises(_NotFound):
            live.get_document("i", "d1")

    def test_mget_keeps_order_and_skips_missing(self, live: LiveElasticsearch) -> None:
        live.index_document("i", "d1", {"a": 1})
        live.index_document("i", "d2", {"a": 2})
        assert live.mget_documents("i", ["d2", "absent", "d1"]) == [{"a": 2}, {"a": 1}]

    def test_mget_of_nothing_is_empty(self, live: LiveElasticsearch) -> None:
        assert live.mget_documents("i", []) == []

    def test_bulk_index_without_helpers_falls_back_to_single_index(
        self, live: LiveElasticsearch, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(live_module, "_helpers_module", lambda: None)
        written = live.bulk_index("i", [("d1", {"a": 1}), ("d2", {"a": 2})])
        assert written == 2
        assert live.get_document("i", "d2") == {"a": 2}

    def test_bulk_index_uses_helpers_when_available(
        self, live: LiveElasticsearch, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        class _Helpers:
            @staticmethod
            def bulk(
                client: Any, actions: list[Any], refresh: bool = False, raise_on_error: bool = True
            ) -> tuple[int, list[Any]]:
                return (len(actions), [])

        monkeypatch.setattr(live_module, "_helpers_module", lambda: _Helpers)
        assert live.bulk_index("i", [("d1", {"a": 1})]) == 1

    def test_bulk_index_of_nothing_is_zero(self, live: LiveElasticsearch) -> None:
        assert live.bulk_index("i", []) == 0


class TestSearch:
    def test_plain_query_is_passed_as_query(self, live: LiveElasticsearch, fake: _FakeReal) -> None:
        result = live.search("i", {"match_all": {}}, size=5)
        assert result["hits"]["hits"][0]["_id"] == "x"
        assert fake.calls[-1]["query"] == {"match_all": {}}

    def test_a_knn_query_is_passed_as_knn(self, live: LiveElasticsearch, fake: _FakeReal) -> None:
        knn = build_knn_query(field="embedding", query_vector=(1.0,), k=3, num_candidates=30)
        live.search("i", {**knn, "query": {"match_all": {}}}, size=3, source_includes=["id"])
        assert fake.calls[-1]["knn"]["k"] == 3
        assert fake.calls[-1]["source_includes"] == ["id"]

    def test_count_and_delete_by_query(self, live: LiveElasticsearch) -> None:
        assert live.count("i", {"match_all": {}}) == 7
        assert live.delete_by_query("i", {"match_all": {}}) == 2

    def test_delete_by_query_refreshes(self, live: LiveElasticsearch, fake: _FakeReal) -> None:
        """Unrefreshed, a cluster keeps counting and returning deleted documents."""
        live.delete_by_query("i", {"match_all": {}})
        assert fake.last_delete_refresh is True

    def test_calls_are_recorded(self, live: LiveElasticsearch) -> None:
        live.count("i")
        assert ("count", "i") in live.calls


class TestConnect:
    def test_connection_is_lazy_until_first_use(self, fake: _FakeReal) -> None:
        live = LiveElasticsearch(ElasticsearchSettings(), client=fake)
        assert live.client is fake

    def test_missing_extra_raises_a_helpful_error(self, monkeypatch: pytest.MonkeyPatch) -> None:
        def boom(name: str, package: str | None = None) -> Any:
            raise ModuleNotFoundError(name)

        monkeypatch.setattr(live_module.importlib, "import_module", boom)
        with pytest.raises(RuntimeError, match="es"):
            connect(Settings().elasticsearch)

    def test_builds_a_client_from_settings(self, monkeypatch: pytest.MonkeyPatch) -> None:
        captured: dict[str, Any] = {}

        class _Module:
            @staticmethod
            def Elasticsearch(**kwargs: Any) -> Any:
                captured.update(kwargs)
                return object()

        monkeypatch.setattr(live_module.importlib, "import_module", lambda name: _Module)
        settings = ElasticsearchSettings(
            urls=("http://es:9200",), username="u", password="p", request_timeout_seconds=2.0
        )
        connect(settings)
        assert captured["hosts"] == ["http://es:9200"]
        assert captured["basic_auth"] == ("u", "p")
        assert captured["request_timeout"] == 2.0
