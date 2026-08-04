"""Registro centralizado de fontes Unicode para geração de PDFs (ReportLab)."""

from __future__ import annotations

from pathlib import Path

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

_UTILS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _UTILS_DIR.parent
BUNDLED_FONTS_DIR = _PROJECT_ROOT / "assets" / "fonts"
WINDOWS_FONT_DIR = Path("C:/Windows/Fonts")

_PDF_FONTS_REGISTERED = False
_CACHED_FONT_PAIR: tuple[str, str] | None = None


def resolve_pdf_font_candidates() -> list[tuple[str, Path, Path]]:
    """Ordem: fontes embutidas no projeto → fontes do sistema Windows."""
    return [
        (
            "DejaVuSans",
            BUNDLED_FONTS_DIR / "DejaVuSans.ttf",
            BUNDLED_FONTS_DIR / "DejaVuSans-Bold.ttf",
        ),
        ("Calibri", WINDOWS_FONT_DIR / "calibri.ttf", WINDOWS_FONT_DIR / "calibrib.ttf"),
        ("Segoe UI", WINDOWS_FONT_DIR / "segoeui.ttf", WINDOWS_FONT_DIR / "segoeuib.ttf"),
        ("Arial", WINDOWS_FONT_DIR / "arial.ttf", WINDOWS_FONT_DIR / "arialbd.ttf"),
    ]


def register_pdf_fonts() -> tuple[str, str]:
    """Registra e devolve (regular, bold) com suporte a acentos portugueses."""
    global _PDF_FONTS_REGISTERED, _CACHED_FONT_PAIR

    if _PDF_FONTS_REGISTERED and _CACHED_FONT_PAIR is not None:
        return _CACHED_FONT_PAIR

    for font_name, regular_path, bold_path in resolve_pdf_font_candidates():
        if not regular_path.is_file() or not bold_path.is_file():
            continue
        bold_name = f"{font_name}-Bold"
        registered = set(pdfmetrics.getRegisteredFontNames())
        if font_name not in registered:
            pdfmetrics.registerFont(TTFont(font_name, str(regular_path)))
        if bold_name not in registered:
            pdfmetrics.registerFont(TTFont(bold_name, str(bold_path)))
        _CACHED_FONT_PAIR = (font_name, bold_name)
        _PDF_FONTS_REGISTERED = True
        return _CACHED_FONT_PAIR

    _CACHED_FONT_PAIR = ("Helvetica", "Helvetica-Bold")
    _PDF_FONTS_REGISTERED = True
    return _CACHED_FONT_PAIR


def normalize_pdf_text(value: object) -> str:
    """Garante texto Unicode limpo para desenho no PDF (sem remoção de acentos)."""
    return str(value if value is not None else "").replace("\r\n", "\n").replace("\r", "\n")
