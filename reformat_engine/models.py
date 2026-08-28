"""Mission 04-A — Reformat Plan models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class Operation:
    block_id: str
    block_type: str
    semantic_role: Optional[str]
    role_confidence: Optional[float]
    action: str
    style_slot: Optional[str]
    reason: str
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ProfileCoverage:
    complete: bool
    missing_slots: list[str] = field(default_factory=list)


@dataclass
class ReformatPlan:
    target_profile_id: str
    source_fingerprint: dict[str, Any] = field(default_factory=dict)
    source_file_sha256: str = ""
    ready: bool = True
    operations: list[Operation] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    profile_coverage: Optional[ProfileCoverage] = None
