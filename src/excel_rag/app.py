"""The FastAPI application factory.

``create_app(settings)`` builds a stateless service: one Elasticsearch client (the in-memory double
unless ``use_live_elasticsearch`` says otherwise), three routers under ``/api/v1`` plus ``/health``,
a request-size guard and one error envelope. The client lives on ``app.state`` and is reached
through :func:`~excel_rag.api.deps.get_client`, so a test injects its own client by overriding that
dependency rather than by rebuilding the app.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from weakref import WeakSet

from fastapi import FastAPI, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from . import __version__
from .api import health, search, structure
from .api.schemas import error_body
from .embedding import build_embedder
from .es import ElasticsearchLike
from .fake_es import in_memory_client
from .retrieval import TooManyWorkbooks, UnknownWorkbook
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
        lifespan=_warm_embedder,
        title="excel-rag",
        version=__version__,
        description="Structure-aware retrieval over Excel workbooks. Evidence, not answers.",
    )
    app.state.settings = resolved
    app.state.client = create_client(resolved)
    #: One embedder per app: the model loads once (on the first query, or at startup through the
    #: lifespan) and is shared by every request thread.
    app.state.embedder = build_embedder(resolved.embedding)
    #: Clients whose indices are known to exist, so a request does not re-check them.
    app.state.ensured_clients = WeakSet()

    app.include_router(health.router)
    app.include_router(search.router)
    app.include_router(structure.router)

    _register_handlers(app)
    _install_size_guard(app)
    return app


@asynccontextmanager
async def _warm_embedder(app: FastAPI) -> AsyncIterator[None]:
    """Load the query model before serving, so the first search does not pay for it.

    A server started by uvicorn runs this; a test client used without a ``with`` block does not,
    and its embedder (if any) loads on first use instead.
    """
    embedder = getattr(app.state, "embedder", None)
    if embedder is not None:
        await run_in_threadpool(embedder.embed_query, "warm-up")
    yield


def _install_size_guard(app: FastAPI) -> None:
    app.add_middleware(BodySizeLimit, limit=MAX_REQUEST_BYTES)


def _too_large(limit: int) -> HTTPException:
    return HTTPException(
        status_code=413,
        detail={"type": "request_too_large", "message": f"request body exceeds {limit} bytes"},
    )


class BodySizeLimit:
    """Refuse a request body over ``limit`` bytes, whether or not it declares its length.

    A declared ``Content-Length`` over the limit is refused before the body is read. A chunked
    body declares nothing, so the bytes are counted as the app reads them and the read fails with a
    413 once the count passes the limit -- FastAPI re-raises an ``HTTPException`` from a body read,
    so the refusal reaches the standard error envelope instead of becoming a generic 400.
    """

    def __init__(self, app: ASGIApp, limit: int) -> None:
        self.app = app
        self.limit = limit

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        for name, value in scope.get("headers", ()):
            if name == b"content-length" and value.isdigit() and int(value) > self.limit:
                response = JSONResponse(
                    status_code=413,
                    content=error_body(
                        "request_too_large", f"request body exceeds {self.limit} bytes"
                    ),
                )
                await response(scope, receive, send)
                return

        received = 0

        async def counted_receive() -> Message:
            nonlocal received
            message = await receive()
            if message["type"] == "http.request":
                received += len(message.get("body", b""))
                if received > self.limit:
                    raise _too_large(self.limit)
            return message

        await self.app(scope, counted_receive, send)


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

    @app.exception_handler(TooManyWorkbooks)
    async def _too_many_workbooks(request: Request, exc: TooManyWorkbooks) -> JSONResponse:
        return JSONResponse(
            status_code=400,
            content=error_body(
                "workbook_filter_required",
                f"{exc.count} workbooks are active, more than an unfiltered search pins "
                f"({exc.limit}); filter by workbook_ids",
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


__all__ = ["MAX_REQUEST_BYTES", "BodySizeLimit", "create_app", "create_client"]
