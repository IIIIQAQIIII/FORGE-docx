"""Build a clean end-user ZIP from the public FORGE repository.

Usage:
    python3 build_release.py

Output:
    dist/forge-docx-vX.Y.Z.zip
    dist/forge-docx-vX.Y.Z.zip.sha256

The archive is whitelist-based: only runtime files, license, docs and templates
are included. Development scripts, tests, Git metadata, virtual environments,
local config and previous outputs are deliberately excluded.
"""

from __future__ import annotations

import hashlib
import shutil
import tempfile
import zipfile
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
DIST_DIR = PROJECT_DIR / "dist"
PACKAGE_DIR_NAME = "forge-docx"

RUNTIME_FILES = [
    "server.py",
    "requirements.txt",
    "VERSION",
    "README.md",
    "INSTALL.md",
    "LICENSE",
    "install_mcp.sh",
    "install_mcp.ps1",
    # Backward-compatible client-specific installers.
    "install_codex_mcp.sh",
    "install_codex_mcp.ps1",
]


def read_version() -> str:
    version_path = PROJECT_DIR / "VERSION"
    if not version_path.is_file():
        raise SystemExit("VERSION file is missing")
    version = version_path.read_text(encoding="utf-8").strip()
    if not version:
        raise SystemExit("VERSION is empty")
    return version if version.startswith("v") else f"v{version}"


def validate_release_inputs() -> list[Path]:
    missing = [name for name in RUNTIME_FILES if not (PROJECT_DIR / name).is_file()]
    templates_dir = PROJECT_DIR / "templates"
    if not templates_dir.is_dir():
        missing.append("templates/")
    if missing:
        raise SystemExit("Missing release files: " + ", ".join(missing))

    templates = sorted(templates_dir.rglob("*.docx"))
    if not templates:
        raise SystemExit("No .docx templates found in templates/")
    return templates


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_release() -> Path:
    version = read_version()
    templates = validate_release_inputs()
    DIST_DIR.mkdir(exist_ok=True)

    zip_path = DIST_DIR / f"forge-docx-{version}.zip"
    checksum_path = zip_path.with_suffix(zip_path.suffix + ".sha256")

    with tempfile.TemporaryDirectory(prefix="forge-docx-release-") as tmp:
        staging_root = Path(tmp) / PACKAGE_DIR_NAME
        staging_root.mkdir(parents=True)

        for name in RUNTIME_FILES:
            shutil.copy2(PROJECT_DIR / name, staging_root / name)

        staging_templates = staging_root / "templates"
        staging_templates.mkdir()
        for source in templates:
            relative = source.relative_to(PROJECT_DIR / "templates")
            destination = staging_templates / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)

        # Include an empty outputs directory without shipping users' generated files.
        (staging_root / "outputs").mkdir()
        (staging_root / "outputs" / ".gitkeep").write_text("", encoding="utf-8")

        manifest_lines = [
            f"FORGE {version}",
            "",
            "Release contents:",
        ]
        for path in sorted(p for p in staging_root.rglob("*") if p.is_file()):
            manifest_lines.append(str(path.relative_to(staging_root)))
        (staging_root / "MANIFEST.txt").write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")

        if zip_path.exists():
            zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for path in sorted(p for p in staging_root.rglob("*") if p.is_file()):
                arcname = Path(PACKAGE_DIR_NAME) / path.relative_to(staging_root)
                archive.write(path, arcname)

    checksum = sha256_file(zip_path)
    checksum_path.write_text(f"{checksum}  {zip_path.name}\n", encoding="utf-8")

    print(f"Built: {zip_path}")
    print(f"SHA256: {checksum}")
    print(f"Checksum file: {checksum_path}")
    return zip_path


if __name__ == "__main__":
    build_release()
