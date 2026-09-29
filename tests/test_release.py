"""Release ZIP must contain a clean, working Windows source distribution."""

from pathlib import Path
import subprocess
import sys
import tomllib
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]


def test_version_matches_pyproject() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert (ROOT / "VERSION").read_text(encoding="utf-8").strip() == project["project"]["version"]
    assert project["project"]["version"] == "0.2.0"


def test_release_archive_contains_launchers_and_docs() -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / "build_release.py")],
                   cwd=ROOT, check=True, capture_output=True, text=True)
    archive_path = ROOT / "dist" / "doc-to-md-converter-v0.2.0-windows.zip"
    assert archive_path.is_file()
    with ZipFile(archive_path) as archive:
        names = set(archive.namelist())
        top = "doc-to-md-converter-v0.2.0/"
        required = {
            "start.bat", "install-context-menu.bat", "uninstall-context-menu.bat",
            "README.md", "pyproject.toml", "VERSION", "CHANGELOG.md",
            "RELEASE_NOTES.md", "doc_to_md_converter/__main__.py",
            "doc_to_md_converter/converter.py", "doc_to_md_converter/pdf_converter.py",
            "docs/INSTALL_WINDOWS.md", "docs/TROUBLESHOOTING.md", "docs/FAQ.md",
        }
        assert {top + entry for entry in required} <= names
        assert all(name.startswith(top) for name in names)
        assert not any("/.venv" in name or "__pycache__" in name or "/tests/" in name
                       or "/.git/" in name for name in names)
        assert archive.testzip() is None
        assert archive.read(top + "VERSION").decode().strip() == "0.2.0"
