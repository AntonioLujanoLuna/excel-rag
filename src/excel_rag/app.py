"""The FastAPI application factory.

``create_app(settings)`` builds a stateless service: one Elasticsearch client (the in-memory double
unless ``use_live_elasticsearch`` says otherwise), three routers under ``/api/v1`` plus ``/health``,
a request-size guard and one error envelope. The client lives on ``app.state`` and is reached
through :func:`~excel_rag.api.deps.get_client`, so a test injects its own client by overriding that
dependency rather than by rebuilding the app.
"""

from __future__ import annotations

from typing import Any

from fastapi import FastAPI, Request
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from starlette.exceptions import HTTPException as StarletteHTTPException

from . import __version__
from .api import health, search, structure
from .api.schemas import error_body
from .es import ElasticsearchLike
from .fake_es import in_memory_client
from .retrieval import UnknownWorkbook
from .settings import Settings

#: Transport-level ceiling on a request body. This is an HTTP guard, not a retrieval budget -- the
#: computed bounds a query may spend live in :class:`~excel_rag.settings.BudgetSettings`.
MAX_REQUEST_BYTES = 1_000_000


def create_client(settings: Settings) -> ElasticsearchLike:
    """The client the app runs against: the live adapter, or the in-memory double.

    The live adapter is imported lazily so importing this module never requires the optional ``es``
    extra.
    """
    if settings.use_live_elasticsearch:
        from .live import LiveElasticsearch

        return LiveElasticsearch(settings.elasticsearch)
    return in_memory_client(settings)


def create_app(settings: Settings | None = None) -> FastAPI:
    resolved = settings or Settings()
    app = FastAPI(
        title="excel-rag",
        version=__version__,
        description="Structure-aware retrieval over Excel workbooks. Evidence, not answers.",
    )
    app.state.settings = resolved
    app.state.client = create_client(resolved)

    app.include_router(health.router)
    app.include_router(search.router)
    app.include_router(structure.router)

    _register_handlers(app)
    _install_size_guard(app)
    return app


def _install_size_guard(app: FastAPI) -> None:
    @app.middleware("http")
    async def _limit_body(request: Request, call_next: Any) -> Any:
        content_length = request.headers.get("content-length")
        if (
            content_length is not None
            and content_length.isdigit()
            and int(content_length) > MAX_REQUEST_BYTES
        ):
            return JSONResponse(
                status_code=413,
                content=error_body(
                    "request_too_large",
                    f"request body exceeds {MAX_REQUEST_BYTES} bytes",
                ),
            )
        return await call_next(request)


def _register_handlers(app: FastAPI) -> None:
    @app.exception_handler(RequestValidationError)
    async def _validation(request: Request, exc: RequestValidationError) -> JSONResponse:
        """A blank or malformed request is a 400 with the standard envelope, not a bare 422."""
        return JSONResponse(
            status_code=400,
            content=error_body(
                "validation_error",
                "request failed validation",
                tuple(jsonable_encoder(exc.errors())),
            ),
        )

    @app.exception_handler(UnknownWorkbook)
    async def _unknown_workbook(request: Request, exc: UnknownWorkbook) -> JSONResponse:
        return JSONResponse(
            status_code=404,
            content=error_body(
                "unknown_workbook",
                f"no active version for workbook {exc.workbook_id!r}",
            ),
        )

    @app.exception_handler(StarletteHTTPException)
    async def _http_error(request: Request, exc: StarletteHTTPException) -> JSONResponse:
        raw: object = exc.detail
        if isinstance(raw, dict) and "type" in raw and "message" in raw:
            normalized = {str(key): value for key, value in raw.items()}
            body = {
                "error": {
                    "type": normalized["type"],
                    "message": normalized["message"],
                    "details": normalized.get("details"),
                }
            }
        else:
            body = error_body("http_error", str(raw))
        return JSONResponse(status_code=exc.status_code, content=body, headers=exc.headers)


__all__ = ["MAX_REQUEST_BYTES", "create_app", "create_client"]
