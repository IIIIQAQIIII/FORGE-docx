"""Mission 05 — Batch Assemble Engine models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class AssemblyItemResult:
    index: int
    source_path: str
    source_sha256: str
    status: str                      # success | success_with_warnings | failed
    normalized_status: str = ""
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    source_fingerprints: dict[str, str] = field(default_factory=dict)
    normalized_fingerprints: dict[str, str] = field(default_factory=dict)

    imported_block_count: int = 0
    normalized_path: Optional[str] = None
    per_item_preservation: Optional[dict[str, Any]] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "source_path": self.source_path,
            "source_sha256": self.source_sha256,
            "status": self.status,
            "normalized_status": self.normalized_status,
            "warnings": list(self.warnings),
            "errors": list(self.errors),
            "source_fingerprints": dict(self.source_fingerprints),
            "normalized_fingerprints": dict(self.normalized_fingerprints),
            "imported_block_count": self.imported_block_count,
            "normalized_path": self.normalized_path,
            "per_item_preservation": self.per_item_preservation,
        }


@dataclass
class AssemblyResult:
    status: str = "ok"
    total: int = 0
    processed: int = 0
    failed: int = 0

    target_profile_id: Optional[str] = None
    resolution: Optional[dict[str, Any]] = None
    assembly_profile: Optional[dict[str, Any]] = None

    output: Optional[str] = None
    normalized_outputs: list[str] = field(default_factory=list)

    items: list[AssemblyItemResult] = field(default_factory=list)

    content_preservation: Optional[dict[str, Any]] = None
    assembly_payload_sha256: str = ""
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "total": self.total,
            "processed": self.processed,
            "failed": self.failed,
            "target_profile_id": self.target_profile_id,
            "resolution": self.resolution,
            "assembly_profile": self.assembly_profile,
            "output": self.output,
            "normalized_outputs": list(self.normalized_outputs),
            "items": [item.to_dict() for item in self.items],
            "content_preservation": self.content_preservation,
            "assembly_payload_sha256": self.assembly_payload_sha256,
            "warnings": list(self.warnings),
            "errors": list(self.errors),
        }
