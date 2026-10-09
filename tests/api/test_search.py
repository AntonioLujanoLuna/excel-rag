"""API tests for ``POST /api/v1/search/excel`` -- the frozen request/response contract.

The ACL test here is the trap from the brief: a related node the caller may not read is one
``_mget`` would happily return, so the endpoint must not put it in the response at all.
"""

from __future__ import annotations

import json
from typing import Any

from conftest import C_FORMULA, C_REVENUE, C_STALE, N_CELL, N_REGION, N_SECRET, WB


def _post(api: Any, **overrides: Any) -> Any:
    body: dict[str, Any] = {
        "query": "revenue",
        "filters": {"workbook_ids": [WB]},
        "top_k": 10,
        "include_structure": False,
    }
    body.update(overrides)
    return api.post("/api/v1/search/excel", json=body)


class TestSuccess:
    def test_returns_scored_hits_with_coordinates(self, api: Any) -> None:
        response = _post(api)
        assert response.status_code == 200
        payload = response.json()
        hit = next(hit for hit in payload["hits"] if hit["chunk_id"] == C_REVENUE)
        assert hit["source"] == {
            "workbook_id": WB,
            "version": 1,
            "sheet": "Forecast",
            "a1_range": "A12:D25",
        }
        assert payload["es_requests"] >= 1
        assert "took_ms" in payload

    def test_empty_result_is_well_formed(self, api: Any) -> None:
        payload = _post(api, query="zzzzzzzz").json()
        assert payload["hits"] == []
        assert payload["nodes"] == {}
        assert payload["truncation"]["truncated"] is False

    def test_filters_by_workbook_sheet_and_type(self, api: Any) -> None:
        by_sheet = _post(api, filters={"workbook_ids": [WB], "sheet_names": ["Forecast"]}).json()
        assert by_sheet["hits"]
        assert {hit["source"]["sheet"] for hit in by_sheet["hits"]} == {"Forecast"}

        no_match = _post(api, filters={"workbook_ids": [WB], "sheet_names": ["HR"]}).json()
        assert no_match["hits"] == [], "the HR chunk has no 'revenue' token"

        by_type = _post(
            api, filters={"workbook_ids": [WB], "chunk_types": ["formula_summary"]}
        ).json()
        assert by_type["hits"], "the formula summary should match"
        assert {hit["chunk_id"] for hit in by_type["hits"]} == {C_FORMULA}

    def test_version_pinning_never_mixes_versions(self, api: Any) -> None:
        payload = _post(api, filters={}, include_structure=False).json()
        assert C_STALE not in {hit["chunk_id"] for hit in payload["hits"]}
        assert {hit["source"]["version"] for hit in payload["hits"]} == {1}


class TestStructureAndAcl:
    def test_across_sheet_related_node_is_returned(self, api: Any) -> None:
        payload = _post(
            api,
            filters={"workbook_ids": [WB], "acl_scopes": ["finance-team"]},
            include_structure=True,
            expand_references=True,
            reference_depth=1,
        ).json()
        assert N_REGION in payload["nodes"]
        assert N_CELL in payload["nodes"]
        assert payload["nodes"][N_CELL]["value"] == 0.05

    def test_a_related_node_the_caller_may_not_read_is_never_returned(self, api: Any) -> None:
        response = _post(
            api,
            filters={"workbook_ids": [WB], "acl_scopes": ["finance-team"]},
            include_structure=True,
            expand_references=True,
            reference_depth=1,
        )
        payload = response.json()
        # The denied node must not be returned as a node ...
        assert N_SECRET not in payload["nodes"]
        # ... and its value must not leak through any other part of the envelope.
        assert "999.0" not in json.dumps(payload)

    def test_expansion_truncation_is_reported(self, api: Any) -> None:
        payload = _post(
            api,
            filters={"workbook_ids": [WB], "acl_scopes": ["finance-team"]},
            include_structure=True,
            expand_references=True,
            reference_depth=1,
            max_related_nodes=1,
        ).json()
        truncation = payload["truncation"]
        assert truncation["truncated"] is True
        assert truncation["reason"]
        assert truncation["dropped_nodes"] >= 1

    def test_unresolved_references_are_surfaced(self, api: Any) -> None:
        payload = _post(
            api,
            filters={"workbook_ids": [WB], "acl_scopes": ["finance-team"]},
            include_structure=True,
            expand_references=True,
            reference_depth=0,
        ).json()
        assert any(ref["reason"] == "indirect" for ref in payload["unresolved_references"])


class TestErrors:
    def test_blank_query_is_a_400_with_an_error_body(self, api: Any) -> None:
        response = _post(api, query="   ")
        assert response.status_code == 400
        assert response.json()["error"]["type"] == "validation_error"

    def test_unknown_workbook_is_a_404(self, api: Any) -> None:
        response = _post(api, filters={"workbook_ids": ["ghost-wb"]})
        assert response.status_code == 404
        assert response.json()["error"]["type"] == "unknown_workbook"

    def test_top_k_over_the_cap_is_rejected(self, api: Any) -> None:
        response = _post(api, top_k=10_000)
        assert response.status_code == 400

    def test_unknown_route_gets_the_error_envelope(self, api: Any) -> None:
        response = api.get("/api/v1/nope")
        assert response.status_code == 404
        assert "error" in response.json()
