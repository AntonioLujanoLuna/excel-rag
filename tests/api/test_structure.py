"""API tests for the direct-inspection endpoints: workbook structure and A1 range intersection."""

from __future__ import annotations

from typing import Any

from conftest import N_FORECAST_CELL, N_REGION, N_SECRET, OTHER_WB, WB


class TestWorkbookStructure:
    def test_returns_nodes_for_the_active_version(self, api: Any) -> None:
        response = api.get(f"/api/v1/excel/{WB}/structure")
        assert response.status_code == 200
        payload = response.json()
        assert payload["workbook_id"] == WB
        assert payload["version"] == 1
        ids = {node["node_id"] for node in payload["nodes"]}
        assert N_REGION in ids
        assert payload["es_requests"] >= 1

    def test_filters_by_sheet_and_node_type(self, api: Any) -> None:
        by_type = api.get(f"/api/v1/excel/{WB}/structure", params={"node_types": "cell"}).json()
        assert by_type["nodes"]
        assert {node["node_type"] for node in by_type["nodes"]} == {"cell"}

        by_sheet = api.get(
            f"/api/v1/excel/{WB}/structure", params={"sheet_names": "Forecast"}
        ).json()
        assert all(node["sheet"] == "Forecast" for node in by_sheet["nodes"])

    def test_unscoped_deployment_sees_every_scope(self, api: Any) -> None:
        payload = api.get(f"/api/v1/excel/{WB}/structure").json()
        # With no default ACL scope the deployment is unscoped, so the exec node is visible here.
        assert N_SECRET in {node["node_id"] for node in payload["nodes"]}

    def test_unknown_workbook_is_a_404(self, api: Any) -> None:
        response = api.get("/api/v1/excel/ghost/structure")
        assert response.status_code == 404
        assert response.json()["error"]["type"] == "unknown_workbook"

    def test_a_bad_limit_is_rejected(self, api: Any) -> None:
        assert api.get(f"/api/v1/excel/{WB}/structure", params={"limit": 0}).status_code == 400


class TestRange:
    def test_intersection_returns_overlapping_nodes(self, api: Any) -> None:
        response = api.post(
            "/api/v1/excel/range",
            json={"workbook_id": WB, "sheet_name": "Forecast", "a1": "A12:C20"},
        )
        assert response.status_code == 200
        payload = response.json()
        ids = {node["node_id"] for node in payload["nodes"]}
        assert ids == {N_REGION, N_FORECAST_CELL}
        assert payload["a1"] == "A12:C20"
        assert payload["version"] == 1

    def test_node_type_filter_narrows_the_result(self, api: Any) -> None:
        payload = api.post(
            "/api/v1/excel/range",
            json={
                "workbook_id": WB,
                "sheet_name": "Forecast",
                "a1": "A12:C20",
                "node_types": ["cell"],
            },
        ).json()
        assert {node["node_id"] for node in payload["nodes"]} == {N_FORECAST_CELL}

    def test_disjoint_rectangle_is_empty(self, api: Any) -> None:
        payload = api.post(
            "/api/v1/excel/range",
            json={"workbook_id": WB, "sheet_name": "Actuals", "a1": "A1:B2"},
        ).json()
        assert payload["nodes"] == []

    def test_a_malformed_rectangle_is_a_400(self, api: Any) -> None:
        response = api.post(
            "/api/v1/excel/range",
            json={"workbook_id": WB, "sheet_name": "Forecast", "a1": "not-a-range"},
        )
        assert response.status_code == 400
        assert response.json()["error"]["type"] == "invalid_range"

    def test_unknown_workbook_is_a_404(self, api: Any) -> None:
        response = api.post(
            "/api/v1/excel/range",
            json={"workbook_id": "ghost", "sheet_name": "Forecast", "a1": "A1"},
        )
        assert response.status_code == 404
        assert response.json()["error"]["type"] == "unknown_workbook"

    def test_other_workbook_has_its_own_nodes(self, api: Any) -> None:
        payload = api.post(
            "/api/v1/excel/range",
            json={"workbook_id": OTHER_WB, "sheet_name": "Sensitivity", "a1": "A1:B2"},
        ).json()
        assert [node["sheet"] for node in payload["nodes"]] == ["Sensitivity"]
