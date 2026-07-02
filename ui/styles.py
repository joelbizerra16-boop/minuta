from __future__ import annotations

from pathlib import Path

import streamlit as st

_THEME_DIR = Path(__file__).resolve().parent / "theme"


def _read_css(filename: str) -> str:
    path = _THEME_DIR / filename
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8")


def inject_global_styles() -> None:
    css = (
        _read_css("global.css")
        + "\n"
        + _read_css("enhancements.css")
        + "\n"
        + _read_css("shell.css")
    )
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def inject_login_styles() -> None:
    css = _read_css("login.css")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
