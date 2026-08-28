"""Build and verify the end-user FORGE runtime ZIP.

This is intentionally an explicit runtime allowlist. It does not package
tests, private snapshot tooling, Git metadata, user data or generated output.
Before reporting success, the archive is extracted into a clean virtual
environment, dependencies are installed, ``server`` is imported and the MCP
smoke test runs against the extracted runtime.
"""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
DIST_DIR = PROJECT_DIR / "dist"
PACKAGE_DIR_NAME = "forge-word-docx-mcp"

RUNTIME_FILES = [
    "server.py",
    "reformat.py",
    "document_ir.py",
    "format_model.py",
    "forge_errors.py",
    "forge_utils.py",
    "forge_version.py",
    "requirements.txt",
    "VERSION",
    "README.md",
    "INSTALL.md",
    "LICENSE",
    "install_mcp.sh",
    "install_mcp.ps1",
    "install_codex_mcp.sh",
    "install_codex_mcp.ps1",
    "mcp-config.example.json",
]

RUNTIME_DIRS = [
    "assembly_engine",
    "edit_engine",
    "intelligence",
    "open_format",
    "profiles",
    "reformat_engine",
    "semantics",
]


def read_version() -> str:
    path = PROJECT_DIR / "VERSION"
    if not path.is_file():
        raise SystemExit("VERSION is missing")
    version = path.read_text(encoding="utf-8").strip().lstrip("v")
    if not version:
        raise SystemExit("VERSION is empty")
    return version


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def validate_inputs() -> list[Path]:
    missing = [name for name in RUNTIME_FILES if not (PROJECT_DIR / name).is_file()]
    missing.extend(name + "/" for name in RUNTIME_DIRS if not (PROJECT_DIR / name).is_dir())
    templates_dir = PROJECT_DIR / "templates"
    if not templates_dir.is_dir():
        missing.append("templates/")
    if missing:
        raise SystemExit("Missing release inputs: " + ", ".join(missing))
    templates = sorted(templates_dir.rglob("*.docx"))
    if not templates:
        raise SystemExit("No DOCX templates found")
    return templates


def stage_runtime(root: Path, templates: list[Path], version: str) -> None:
    for name in RUNTIME_FILES:
        _copy_file(PROJECT_DIR / name, root / name)

    for name in RUNTIME_DIRS:
        shutil.copytree(
            PROJECT_DIR / name,
            root / name,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
        )

    templates_root = PROJECT_DIR / "templates"
    for source in templates:
        _copy_file(source, root / "templates" / source.relative_to(templates_root))

    (root / "outputs").mkdir()
    (root / "outputs" / ".gitkeep").write_text("", encoding="utf-8")

    manifest = [f"FORGE v{version} runtime package", "", "Files:"]
    manifest.extend(str(path.relative_to(root)) for path in sorted(root.rglob("*")) if path.is_file())
    (root / "MANIFEST.txt").write_text("\n".join(manifest) + "\n", encoding="utf-8")


def create_archive(staging_root: Path, zip_path: Path) -> None:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(staging_root.rglob("*")):
            if path.is_file():
                archive.write(path, Path(PACKAGE_DIR_NAME) / path.relative_to(staging_root))


def _venv_python(venv_dir: Path) -> Path:
    return venv_dir / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")


def smoke_test_release(zip_path: Path) -> None:
    """Install and smoke-test the archive outside the source tree."""
    source_smoke = PROJECT_DIR / "scripts" / "mcp_smoke.py"
    if not source_smoke.is_file():
        raise SystemExit("scripts/mcp_smoke.py is required for release validation")

    with tempfile.TemporaryDirectory(prefix="forge_release_smoke_") as tmp:
        tmp_root = Path(tmp)
        with zipfile.ZipFile(zip_path) as archive:
            archive.extractall(tmp_root)
        runtime_root = tmp_root / PACKAGE_DIR_NAME
        venv_dir = tmp_root / "venv"
        subprocess.run([sys.executable, "-m", "venv", str(venv_dir)], check=True)
        python = _venv_python(venv_dir)
        subprocess.run([str(python), "-m", "pip", "install", "-r", "requirements.txt"], cwd=runtime_root, check=True)
        subprocess.run([str(python), "-c", "import server; print(server.FORGE_VERSION)"], cwd=runtime_root, check=True)

        # The smoke test is copied only into the temporary extraction; it is
        # deliberately absent from the end-user archive.
        smoke_path = runtime_root / "scripts" / "mcp_smoke.py"
        smoke_path.parent.mkdir()
        shutil.copy2(source_smoke, smoke_path)
        subprocess.run([str(python), str(smoke_path)], cwd=runtime_root, check=True)


def build_release() -> Path:
    version = read_version()
    templates = validate_inputs()
    zip_path = DIST_DIR / f"{PACKAGE_DIR_NAME}-v{version}.zip"
    checksum_path = zip_path.with_suffix(".zip.sha256")

    with tempfile.TemporaryDirectory(prefix="forge_release_stage_") as tmp:
        staging_root = Path(tmp) / PACKAGE_DIR_NAME
        staging_root.mkdir()
        stage_runtime(staging_root, templates, version)
        create_archive(staging_root, zip_path)

    checksum = sha256_file(zip_path)
    checksum_path.write_text(f"{checksum}  {zip_path.name}\n", encoding="utf-8")
    smoke_test_release(zip_path)
    print(f"Built and smoke-tested: {zip_path}")
    print(f"SHA256: {checksum}")
    return zip_path


if __name__ == "__main__":
    build_release()
