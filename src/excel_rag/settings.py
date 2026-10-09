"""Configuration: every bound on retrieval is declared here, not hardcoded in a branch."""

from __future__ import annotations

from pydantic import BaseModel, Field, SecretStr
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
    dims: int = Field(default=768, ge=8)
    model: str = Field(default="unspecified", description="Recorded on every chunk document.")
    #: Whether the query embedding is produced in-process (False) or supplied by the caller.
    embed_queries_locally: bool = False


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


class ServerSettings(BaseModel):
    host: str = "127.0.0.1"
    port: int = Field(default=8080, ge=1, le=65535)
    service_token: SecretStr | None = Field(
        default=None, description="Optional shared token, required on /api/v1 routes when set."
    )


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
    elasticsearch: ElasticsearchSettings = Field(default_factory=ElasticsearchSettings)
    server: ServerSettings = Field(default_factory=ServerSettings)
