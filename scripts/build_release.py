"""Create a clean, versioned Windows source distribution using the standard library."""

from __future__ import annotations

from pathlib import Path
import tomllib
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]


def build_archive() -> Path:
    metadata = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    version = metadata["project"]["version"]
    recorded = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
    if version != recorded:
        raise ValueError(f"VERSION {recorded!r} and project.version {version!r} disagree")

    files = [
        ROOT / name for name in (
            "start.bat",
            "install-context-menu.bat",
            "uninstall-context-menu.bat",
            "README.md",
            "CHANGELOG.md",
            "RELEASE_NOTES.md",
            "VERSION",
            "pyproject.toml",
        )
    ]
    files += sorted((ROOT / "doc_to_md_converter").glob("*.py"))
    files += sorted((ROOT / "docs").glob("*.md"))
    for file in files:
        if not file.is_file():
            raise FileNotFoundError(f"Release file not found: {file}")

    destination = ROOT / "dist" / f"doc-to-md-converter-v{version}-windows.zip"
    destination.parent.mkdir(parents=True, exist_ok=True)
    top = f"doc-to-md-converter-v{version}"
    with ZipFile(destination, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for file in files:
            archive.write(file, f"{top}/{file.relative_to(ROOT).as_posix()}")
    with ZipFile(destination) as archive:
        damaged = archive.testzip()
        if damaged:
            raise RuntimeError(f"Broken file inside ZIP: {damaged}")
    return destination


if __name__ == "__main__":
    print(build_archive())
