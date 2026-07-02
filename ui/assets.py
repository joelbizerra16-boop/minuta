from __future__ import annotations

from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_LOGO_PNG = _PROJECT_ROOT / "baixados.png"
_LOGO_SVG = _PROJECT_ROOT / "assets" / "logo.svg"


def get_brand_logo_path() -> Path | None:
    """Logo oficial BRIDA: PNG do projeto ou SVG em assets/."""
    if _LOGO_PNG.is_file():
        return _LOGO_PNG
    if _LOGO_SVG.is_file():
        return _LOGO_SVG
    return None
