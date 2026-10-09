"""FastAPI dependencies: the client, the settings, and the objects built over them.

Everything the routes need arrives through a dependency so a test can override any of them. The
client in particular lives on ``app.state`` and is reached through :func:`get_client`; a test that
wants a hand-built in-memory client sets ``app.dependency_overrides[get_client]`` and the whole
stack above it -- repository, service -- is rebuilt over that client.
"""

from __future__ import annotations

import secrets
from typing import Annotated, cast

from fastapi import Depends, HTTPException, Request

from ..es import ElasticsearchLike
from ..retrieval import Repository, RetrievalService
from ..settings import Settings


def get_settings(request: Request) -> Settings:
    return cast(Settings, request.app.state.settings)


def get_client(request: Request) -> ElasticsearchLike:
    return cast(ElasticsearchLike, request.app.state.client)


def get_repository(
    client: Annotated[ElasticsearchLike, Depends(get_client)],
    settings: Annotated[Settings, Depends(get_settings)],
) -> Repository:
    """A repository over the app's client. Ensuring the indices exist is idempotent."""
    repository = Repository(client, settings)
    repository.ensure_indices()
    return repository


def get_service(
    repository: Annotated[Repository, Depends(get_repository)],
    settings: Annotated[Settings, Depends(get_settings)],
) -> RetrievalService:
    return RetrievalService(repository, settings)


def _bearer(header: str | None) -> str | None:
    if not header:
        return None
    prefix = "bearer "
    if header.lower().startswith(prefix):
        return header[len(prefix) :].strip()
    return None


def require_token(
    request: Request,
    settings: Annotated[Settings, Depends(get_settings)],
) -> None:
    """Enforce the optional service token on the ``/api/v1`` routes.

    A deployment with no token configured is open (the default, and what the tests use). With one
    configured, the caller must present it as ``X-Service-Token`` or ``Authorization: Bearer``.
    The comparison is constant-time.
    """
    token = settings.server.service_token
    if token is None:
        return
    supplied = request.headers.get("x-service-token") or _bearer(
        request.headers.get("authorization")
    )
    if not supplied or not secrets.compare_digest(supplied, token.get_secret_value()):
        raise HTTPException(
            status_code=401,
            detail={"type": "unauthorized", "message": "a valid service token is required"},
        )


__all__ = [
    "get_client",
    "get_repository",
    "get_service",
    "get_settings",
    "require_token",
]
