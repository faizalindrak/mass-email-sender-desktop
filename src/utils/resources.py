from __future__ import annotations

import sys
from pathlib import Path


_base_path: Path | None = None


def get_base_path() -> Path:
    """Return the resolved base path for bundled resources."""
    global _base_path

    if _base_path is not None:
        return _base_path

    if hasattr(sys, "_MEIPASS"):
        _base_path = Path(getattr(sys, "_MEIPASS"))
    else:
        _base_path = Path(__file__).resolve().parent.parent

    return _base_path


def get_resource_path(relative_path: str) -> str:
    """Resolve a resource path that works in both dev and PyInstaller builds."""
    base_path = get_base_path()
    candidate = base_path / relative_path

    if candidate.exists():
        return str(candidate)

    # Fallback to current working directory to help during development
    return str((Path.cwd() / relative_path).resolve())