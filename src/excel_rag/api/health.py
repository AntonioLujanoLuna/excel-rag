"""``GET /health`` -- liveness plus what the service is bound to.

It reports the version, how the service is wired (live versus in-memory client) and the document
count per index, so an operator can see at a glance which store answered and whether it holds data.
"""

from __future__ import annotations

from typing import Annotated

from fastapi import APIRouter, Depends

from .. import __version__
from ..retrieval import Repository
from ..settings import Settings
from .deps import get_repository, get_settings
from .schemas import HealthResponse

router = APIRouter(tags=["health"])


@router.get("/health", response_model=HealthResponse)
def health(
    repository: Annotated[Repository, Depends(get_repository)],
    settings: Annotated[Settings, Depends(get_settings)],
) -> HealthResponse:
    return HealthResponse(
        status="ok",
        version=__version__,
        use_live_elasticsearch=settings.use_live_elasticsearch,
        indices=repository.counts(),
    )


__all__ = ["router"]
