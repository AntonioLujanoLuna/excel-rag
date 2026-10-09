"""ACL scopes come from the authenticated caller, not from the request body.

Before principals, `filters.acl_scopes` was whatever the caller wrote: a caller could name a scope
it did not hold, or name none and read everything. These tests pin the replacement: a restricted
principal is filtered by its own scopes, may narrow them, and is refused (403) for naming one it
does not hold -- on search and on both direct-inspection routes.
"""

from __future__ import annotations

import pytest
from conftest import C_HR, C_REVENUE, C_SECRET, EXEC, FINANCE, HR, N_SECRET, WB
from fastapi.testclient import TestClient
from pydantic import ValidationError

from excel_rag.api.deps import get_client
from excel_rag.app import create_app
from excel_rag.fake_es import InMemoryElasticsearch
from excel_rag.settings import Principal, Settings

FIN_TOKEN = "fin-token"
BOTH_TOKEN = "both-token"
ADMIN_TOKEN = "admin-token"


@pytest.fixture
def secured(client: InMemoryElasticsearch) -> TestClient:
    settings = Settings(
        server={
            "principals": [
                {"name": "finance", "token": FIN_TOKEN, "acl_scopes": [FINANCE]},
                {"name": "fin-hr", "token": BOTH_TOKEN, "acl_scopes": [FINANCE, HR]},
                {"name": "admin", "token": ADMIN_TOKEN, "unrestricted": True},
            ]
        }
    )
    app = create_app(settings)
    app.dependency_overrides[get_client] = lambda: client
    return TestClient(app)


def _search(api: TestClient, token: str, scopes: list[str] | None = None) -> set[str]:
    filters: dict[str, object] = {"workbook_ids": [WB]}
    if scopes is not None:
        filters["acl_scopes"] = scopes
    response = api.post(
        "/api/v1/search/excel",
        json={"query": "revenue headcount executive", "filters": filters, "top_k": 20},
        headers={"Authorization": f"Bearer {token}"},
    )
    assert response.status_code == 200, response.text
    return {hit["chunk_id"] for hit in response.json()["hits"]}


class TestSearch:
    def test_a_principal_naming_no_scope_is_filtered_by_all_of_its_own(
        self, secured: TestClient
    ) -> None:
        found = _search(secured, BOTH_TOKEN)
        assert {C_REVENUE, C_HR} <= found
        assert C_SECRET not in found

    def test_a_principal_may_narrow_its_scopes(self, secured: TestClient) -> None:
        found = _search(secured, BOTH_TOKEN, [HR])
        assert C_HR in found
        assert C_REVENUE not in found

    def test_naming_a_scope_the_principal_does_not_hold_is_403(self, secured: TestClient) -> None:
        response = secured.post(
            "/api/v1/search/excel",
            json={"query": "revenue", "filters": {"acl_scopes": [EXEC]}},
            headers={"X-Service-Token": FIN_TOKEN},
        )
        assert response.status_code == 403
        assert response.json()["error"]["type"] == "forbidden_scope"

    def test_an_unrestricted_principal_passes_its_requested_scopes_through(
        self, secured: TestClient
    ) -> None:
        assert C_SECRET in _search(secured, ADMIN_TOKEN, [EXEC])

    def test_an_unknown_token_is_401(self, secured: TestClient) -> None:
        response = secured.post(
            "/api/v1/search/excel",
            json={"query": "revenue"},
            headers={"Authorization": "Bearer nope"},
        )
        assert response.status_code == 401


class TestDirectInspection:
    def test_structure_is_filtered_by_the_principal(self, secured: TestClient) -> None:
        response = secured.get(
            f"/api/v1/excel/{WB}/structure", headers={"X-Service-Token": FIN_TOKEN}
        )
        assert response.status_code == 200
        assert N_SECRET not in {node["node_id"] for node in response.json()["nodes"]}

    def test_structure_refuses_a_foreign_scope(self, secured: TestClient) -> None:
        response = secured.get(
            f"/api/v1/excel/{WB}/structure",
            params={"acl_scopes": [EXEC]},
            headers={"X-Service-Token": FIN_TOKEN},
        )
        assert response.status_code == 403

    def test_range_is_filtered_by_the_principal(self, secured: TestClient) -> None:
        body = {"workbook_id": WB, "sheet_name": "Assumptions", "a1": "A1:Z50"}
        restricted = secured.post(
            "/api/v1/excel/range", json=body, headers={"X-Service-Token": FIN_TOKEN}
        )
        admin = secured.post(
            "/api/v1/excel/range",
            json={**body, "acl_scopes": [EXEC]},
            headers={"X-Service-Token": ADMIN_TOKEN},
        )
        assert N_SECRET not in {node["node_id"] for node in restricted.json()["nodes"]}
        assert N_SECRET in {node["node_id"] for node in admin.json()["nodes"]}


def test_a_principal_without_scopes_must_say_it_is_unrestricted() -> None:
    with pytest.raises(ValidationError, match="unrestricted"):
        Principal(name="oops", token="t")
