"""Mission 07 — Edit Engine models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class EditOperation:
    op: str
    target_block_id: Optional[str] = None
    source_locator: Optional[str] = None
    old_text: Optional[str] = None
    new_text: Optional[str] = None
    expected_occurrences: int = 1
    text: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "op": self.op,
            "target_block_id": self.target_block_id,
            "source_locator": self.source_locator,
            "old_text": self.old_text,
            "new_text": self.new_text,
            "expected_occurrences": self.expected_occurrences,
            "text": self.text,
        }


@dataclass
class BoundEditOperation:
    op: str
    element: Any
    block_id: Optional[str]
    locator: Optional[str]
    old_text: Optional[str] = None
    new_text: Optional[str] = None
    expected_occurrences: int = 1
    text: Optional[str] = None
    match_plan: Optional[list] = None

    def to_dict(self) -> dict:
        return {
            "op": self.op,
            "block_id": self.block_id,
            "locator": self.locator,
            "old_text": self.old_text,
            "new_text": self.new_text,
            "expected_occurrences": self.expected_occurrences,
            "text": self.text,
            "matches": len(self.match_plan or []),
        }


@dataclass
class EditPlan:
    source_sha256: str = ""
    ready: bool = True
    operations: list = field(default_factory=list)  # BoundEditOperation
    blockers: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
