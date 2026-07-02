from __future__ import annotations

import html
import re
from collections.abc import Callable

import streamlit as st


def _user_initials(name: str) -> str:
    parts = [part for part in re.split(r"\s+", name.strip()) if part]
    if not parts:
        return "OP"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return f"{parts[0][0]}{parts[-1][0]}".upper()


def render_app_topbar(
    title: str,
    subtitle: str,
    *,
    user_name: str,
    user_role: str,
    show_panel_toggle: bool = False,
    panel_open: bool = True,
    on_toggle_panel: Callable[[], None] | None = None,
) -> None:
    """Unico cabecalho da area principal (sem logo — logo fica no sidebar)."""
    initials = _user_initials(user_name)
    st.markdown(
        f"""
        <div class="brida-topbar">
            <div class="brida-topbar-left">
                <span class="brida-topbar-menu-icon">☰</span>
                <h1 class="brida-topbar-title">{html.escape(title)}</h1>
            </div>
            <div class="brida-topbar-right">
                <span class="brida-topbar-bell">🔔</span>
                <span class="brida-user-avatar">{html.escape(initials)}</span>
                <div class="brida-user-meta-block">
                    <span class="brida-user-meta-name">{html.escape(user_name)}</span>
                    <span class="brida-user-meta-role">{html.escape(user_role)}</span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if show_panel_toggle and on_toggle_panel is not None:
        st.markdown('<div class="brida-panel-toggle">', unsafe_allow_html=True)
        panel_label = "Ocultar arquivos" if panel_open else "Exibir arquivos"
        st.button(panel_label, on_click=on_toggle_panel, key="brida_toggle_panel")
        st.markdown("</div>", unsafe_allow_html=True)

    if subtitle:
        st.markdown(f'<p class="brida-page-subtitle">{html.escape(subtitle)}</p>', unsafe_allow_html=True)


def open_content_section(extra_class: str = "") -> None:
    classes = "brida-content-area"
    if extra_class:
        classes += f" {extra_class}"
    st.markdown(f'<div class="{classes}">', unsafe_allow_html=True)


def close_content_section() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


# Compatibilidade — navegacao agora e render_app_sidebar() em app.py
def render_app_navigation(*args, **kwargs) -> None:
    raise RuntimeError("Use render_app_sidebar() em app.py — navegacao no st.sidebar nativo.")


def render_fixed_sidebar(*args, **kwargs) -> None:
    raise RuntimeError("render_fixed_sidebar removido — use render_app_sidebar() em app.py.")


def open_main_content() -> None:
    return


def open_main_stage() -> None:
    return


def close_main_stage() -> None:
    return


def open_upload_panel() -> None:
    return


def close_upload_panel() -> None:
    return
