"""Configuration: every bound on retrieval is declared here, not hardcoded in a branch."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, SecretStr, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

from .es import INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS


class BudgetSettings(BaseModel):
    """What one request may spend. The design bounds reference traversal on four axes."""

    reference_depth: int = Field(
        default=1, ge=0, le=3, description="Maximum breadth-first expansion depth."
    )
    max_related_nodes: int = Field(default=20, ge=0, le=200)
    max_payload_bytes: int = Field(
        default=1_000_000,
        ge=1_000,
        description="Byte ceiling on the node payload returned with a response.",
    )
    timeout_seconds: float = Field(default=5.0, gt=0.0)
    max_top_k: int = Field(default=100, ge=1, le=1000)
    #: A workbook that would index more than this many documents is refused rather than silently
    #: truncated; the design forbids a cell-per-document explosion.
    max_documents_per_workbook: int = Field(default=200_000, ge=100)


class EmbeddingSettings(BaseModel):
    """The dense-vector model, used at ingestion for chunks and at search time for the query.

    The default is ``lightonai/mDenseOn`` run in-process through sentence-transformers (the
    ``embed`` extra). ``hashing`` is a deterministic test double with no semantics; ``none`` indexes
    and searches without vectors (lexical only). Changing ``model`` or ``dims`` means re-indexing:
    a query vector is only ever compared with chunks embedded by the same model.
    """

    provider: Literal["sentence-transformers", "hashing", "none"] = "sentence-transformers"
    model: str = Field(
        default="lightonai/mDenseOn", description="Recorded on every chunk document."
    )
    dims: int = Field(default=768, ge=8)
    batch_size: int = Field(default=32, ge=1)
    #: Token cap per text. mDenseOn accepts 8,192, but a chunk is a description of a region, not a
    #: document, and on CPU the cost grows with the length; long chunks are truncated, not refused.
    max_seq_length: int | None = Field(default=1024, ge=16)
    #: ``cpu``, ``cuda``, ``mps``... ``None`` lets sentence-transformers pick.
    device: str | None = None


class RerankSettings(BaseModel):
    """An optional cross-encoder that rescores the top fused candidates (off by default).

    ``window`` is how many fused candidates are rescored; it is at least ``top_k`` for any request,
    so reranking reorders the window and never drops a hit fusion would have returned. ``overlap``
    is a deterministic test double with no semantics.
    """

    provider: Literal["cross-encoder", "overlap", "none"] = "none"
    model: str = "cross-encoder/mmarco-mMiniLMv2-L12-H384-v1"
    window: int = Field(default=50, ge=1, le=500)
    batch_size: int = Field(default=32, ge=1)
    max_length: int | None = Field(default=512, ge=16)
    device: str | None = None


class ElasticsearchSettings(BaseModel):
    urls: tuple[str, ...] = ("http://127.0.0.1:9200",)
    username: str | None = None
    password: SecretStr | None = None
    index_prefix: str = ""
    request_timeout_seconds: float = Field(default=5.0, gt=0.0)

    def index_name(self, base: str) -> str:
        return f"{self.index_prefix}{base}"

    @property
    def chunks_index(self) -> str:
        return self.index_name(INDEX_CHUNKS)

    @property
    def structure_index(self) -> str:
        return self.index_name(INDEX_STRUCTURE)

    @property
    def versions_index(self) -> str:
        return self.index_name(INDEX_VERSIONS)


class Principal(BaseModel):
    """A caller identified by its token, and the ACL scopes that caller holds.

    The scopes a request is filtered by come from here, not from the request body: a restricted
    principal may *narrow* its scopes per request but never name one it does not hold. A principal
    with no scopes must say so explicitly with ``unrestricted`` -- an empty scope list means "no
    filter" downstream, which is too dangerous to be a default.
    """

    name: str
    token: SecretStr
    acl_scopes: tuple[str, ...] = ()
    unrestricted: bool = False

    @model_validator(mode="after")
    def _scoped_or_explicitly_unrestricted(self) -> Principal:
        if not self.unrestricted and not self.acl_scopes:
            raise ValueError(
                f"principal {self.name!r} needs acl_scopes, or unrestricted=true to read everything"
            )
        return self


class ServerSettings(BaseModel):
    host: str = "127.0.0.1"
    port: int = Field(default=8080, ge=1, le=65535)
    service_token: SecretStr | None = Field(
        default=None,
        description="Optional shared token for a trusted, unrestricted caller (one that applies "
        "its own authorisation and passes the scopes it allows); required on /api/v1 when set.",
    )
    principals: tuple[Principal, ...] = Field(
        default=(),
        description="Per-caller tokens with the ACL scopes each holds. When any token is "
        "configured the /api/v1 routes require one.",
    )

    @property
    def authenticated(self) -> bool:
        return self.service_token is not None or bool(self.principals)


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix="EXCEL_RAG_", env_nested_delimiter="__", extra="forbid", frozen=True
    )

    #: When False the service runs against the in-memory client, which is what the tests and a
    #: laptop demo use; nothing else changes.
    use_live_elasticsearch: bool = False
    default_acl_scope: tuple[str, ...] = ()
    budgets: BudgetSettings = Field(default_factory=BudgetSettings)
    embedding: EmbeddingSettings = Field(default_factory=EmbeddingSettings)
    rerank: RerankSettings = Field(default_factory=RerankSettings)
    elasticsearch: ElasticsearchSettings = Field(default_factory=ElasticsearchSettings)
    server: ServerSettings = Field(default_factory=ServerSettings)
