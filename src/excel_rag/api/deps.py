"""FastAPI dependencies: the client, the settings, and the objects built over them.

Everything the routes need arrives through a dependency so a test can override any of them. The
client in particular lives on ``app.state`` and is reached through :func:`get_client`; a test that
wants a hand-built in-memory client sets ``app.dependency_overrides[get_client]`` and the whole
stack above it -- repository, service -- is rebuilt over that client.
"""

from __future__ import annotations

import secrets
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Annotated, Any, cast
from weakref import WeakSet

from fastapi import Depends, HTTPException, Request

from ..es import ElasticsearchLike
from ..retrieval import Repository, RetrievalService
from ..settings import Settings


def get_settings(request: Request) -> Settings:
    return cast(Settings, request.app.state.settings)


def get_client(request: Request) -> ElasticsearchLike:
    return cast(ElasticsearchLike, request.app.state.client)


def get_repository(
    request: Request,
    client: Annotated[ElasticsearchLike, Depends(get_client)],
    settings: Annotated[Settings, Depends(get_settings)],
) -> Repository:
    """A repository over the app's client, with its indices ensured once per client.

    Ensuring is idempotent but not free -- three existence checks, which on a live cluster are three
    round trips -- so it runs the first time a client is seen rather than on every request. A
    client swapped in through ``dependency_overrides`` is a new client and is ensured in turn.
    """
    repository = Repository(client, settings)
    ensured: WeakSet[Any] | None = getattr(request.app.state, "ensured_clients", None)
    if ensured is None or client not in ensured:
        repository.ensure_indices()
        if ensured is not None:
            ensured.add(client)
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


@dataclass(frozen=True)
class Caller:
    """Who is calling, and which ACL scopes the request may be filtered by."""

    name: str
    acl_scopes: tuple[str, ...] = ()
    unrestricted: bool = False

    def scopes_for(self, requested: Sequence[str]) -> tuple[str, ...]:
        """The scopes one request is filtered by.

        An unrestricted caller passes its requested scopes through (an empty tuple then falls back
        to ``default_acl_scope`` in the repository). A restricted caller gets the scopes it holds,
        or the subset it asked for; asking for a scope it does not hold is a 403, not a silent
        widening and not a silent drop.
        """
        if self.unrestricted:
            return tuple(requested)
        if not requested:
            return self.acl_scopes
        foreign = sorted(set(requested) - set(self.acl_scopes))
        if foreign:
            raise HTTPException(
                status_code=403,
                detail={
                    "type": "forbidden_scope",
                    "message": f"caller {self.name!r} does not hold scope(s) {foreign}",
                },
            )
        return tuple(dict.fromkeys(requested))


#: An open deployment (no token configured): callers assert their own scopes, as before.
ANONYMOUS = Caller(name="anonymous", unrestricted=True)


def get_caller(
    request: Request,
    settings: Annotated[Settings, Depends(get_settings)],
) -> Caller:
    """Identify the caller on the ``/api/v1`` routes.

    With no token configured the deployment is open (the default, and what most tests use). With a
    ``service_token`` or ``principals`` configured, the caller must present a token as
    ``X-Service-Token`` or ``Authorization: Bearer``. Every configured token is compared, in
    constant time, so the match does not leak which one was close.
    """
    server = settings.server
    if not server.authenticated:
        return ANONYMOUS
    supplied = request.headers.get("x-service-token") or _bearer(
        request.headers.get("authorization")
    )
    matched: Caller | None = None
    if supplied:
        if server.service_token is not None and secrets.compare_digest(
            supplied.encode(), server.service_token.get_secret_value().encode()
        ):
            matched = Caller(name="service", unrestricted=True)
        for principal in server.principals:
            if (
                secrets.compare_digest(
                    supplied.encode(), principal.token.get_secret_value().encode()
                )
                and matched is None
            ):
                matched = Caller(
                    name=principal.name,
                    acl_scopes=principal.acl_scopes,
                    unrestricted=principal.unrestricted,
                )
    if matched is None:
        raise HTTPException(
            status_code=401,
            detail={"type": "unauthorized", "message": "a valid service token is required"},
        )
    return matched


__all__ = [
    "ANONYMOUS",
    "Caller",
    "get_caller",
    "get_client",
    "get_repository",
    "get_service",
    "get_settings",
]
