"""FORGE_HOME user asset directory."""

from __future__ import annotations

import os
from pathlib import Path


def get_forge_home() -> Path:
    env = os.environ.get("FORGE_HOME")
    if env:
        return Path(env).expanduser().resolve()
    return Path.home() / ".forge-docx-mcp"


def ensure_forge_home() -> Path:
    home = get_forge_home()
    for sub in ("profiles", "templates", "drafts"):
        (home / sub).mkdir(parents=True, exist_ok=True)
    return home


def profiles_dir() -> Path:
    ensure_forge_home()
    return get_forge_home() / "profiles"


def templates_dir() -> Path:
    ensure_forge_home()
    return get_forge_home() / "templates"


def drafts_dir() -> Path:
    ensure_forge_home()
    return get_forge_home() / "drafts"
