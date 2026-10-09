"""Application-wiring tests: the token gate, the request-size guard, and the client factory."""

from __future__ import annotations

from typing import Any

from fastapi.testclient import TestClient

from excel_rag import __version__
from excel_rag.api.deps import get_client
from excel_rag.app import MAX_REQUEST_BYTES, create_app, create_client
from excel_rag.fake_es import InMemoryElasticsearch
from excel_rag.live import LiveElasticsearch
from excel_rag.settings import Settings


class TestHealth:
    def test_reports_version_and_document_counts(self, api: Any) -> None:
        payload = api.get("/health").json()
        assert payload["status"] == "ok"
        assert payload["version"] == __version__
        assert payload["use_live_elasticsearch"] is False
        assert payload["indices"]["excel_chunks"] > 0


class TestLazyPackageAttribute:
    def test_create_app_is_exposed_lazily(self) -> None:
        import excel_rag

        assert callable(excel_rag.create_app)

    def test_an_unknown_attribute_is_an_attribute_error(self) -> None:
        import excel_rag

        try:
            _ = excel_rag.nonsense
        except AttributeError as exc:
            assert "nonsense" in str(exc)
        else:  # pragma: no cover - defensive
            raise AssertionError("expected AttributeError")


class TestTokenGate:
    def _app(self, client: InMemoryElasticsearch) -> TestClient:
        settings = Settings(server={"service_token": "s3cret"})
        app = create_app(settings)
        app.dependency_overrides[get_client] = lambda: client
        return TestClient(app)

    def test_health_stays_open(self, client: InMemoryElasticsearch) -> None:
        assert self._app(client).get("/health").status_code == 200

    def test_a_missing_token_is_401(self, client: InMemoryElasticsearch) -> None:
        response = self._app(client).post(
            "/api/v1/search/excel", json={"query": "revenue", "include_structure": False}
        )
        assert response.status_code == 401
        assert response.json()["error"]["type"] == "unauthorized"

    def test_a_wrong_token_is_401(self, client: InMemoryElasticsearch) -> None:
        response = self._app(client).post(
            "/api/v1/search/excel",
            json={"query": "revenue", "include_structure": False},
            headers={"X-Service-Token": "nope"},
        )
        assert response.status_code == 401

    def test_the_right_token_in_either_header_inlet_is_200(
        self, client: InMemoryElasticsearch
    ) -> None:
        instance = self._app(client)
        body = {"query": "revenue", "include_structure": False}
        assert (
            instance.post(
                "/api/v1/search/excel", json=body, headers={"X-Service-Token": "s3cret"}
            ).status_code
            == 200
        )
        assert (
            instance.post(
                "/api/v1/search/excel", json=body, headers={"Authorization": "Bearer s3cret"}
            ).status_code
            == 200
        )


class TestSizeGuard:
    def test_an_oversized_body_is_413(self, api: Any) -> None:
        response = api.post(
            "/api/v1/search/excel",
            json={"query": "x" * (MAX_REQUEST_BYTES + 1000), "include_structure": False},
        )
        assert response.status_code == 413
        assert response.json()["error"]["type"] == "request_too_large"


class TestClientFactory:
    def test_default_client_is_the_in_memory_double(self) -> None:
        assert isinstance(create_client(Settings()), InMemoryElasticsearch)

    def test_live_settings_build_the_adapter_without_connecting(self) -> None:
        client = create_client(Settings(use_live_elasticsearch=True))
        assert isinstance(client, LiveElasticsearch)
        # No connection is attempted until the first call, so no cluster is needed here.
        assert client.calls == []

    def test_create_app_defaults_to_safe_settings(self) -> None:
        app = create_app()
        assert isinstance(app.state.client, InMemoryElasticsearch)
        assert app.state.settings.use_live_elasticsearch is False
