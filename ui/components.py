from __future__ import annotations

import html

import streamlit as st

from ui.assets import get_brand_logo_path
from ui.tokens import LOGO_WIDTH_HEADER


def render_brida_header(
    title: str,
    subtitle: str,
    *,
    on_menu,
    on_panel,
    on_logout,
    key_prefix: str,
) -> None:
    st.markdown('<div class="brida-app-header">', unsafe_allow_html=True)
    col_logo, col_header, col_home, col_menu_toggle, col_action = st.columns(
        [1.1, 4.1, 1.0, 1.0, 1.0],
        vertical_alignment="center",
    )
    with col_logo:
        logo_path = get_brand_logo_path()
        if logo_path is not None:
            st.image(str(logo_path), width=LOGO_WIDTH_HEADER)
    with col_header:
        st.markdown(f'<h1 class="brida-page-title">{html.escape(title)}</h1>', unsafe_allow_html=True)
        st.markdown(f'<p class="brida-page-subtitle">{html.escape(subtitle)}</p>', unsafe_allow_html=True)
    with col_home:
        st.button("Menu", use_container_width=True, on_click=on_menu, key=f"home_button_{key_prefix}")
    with col_menu_toggle:
        st.button("Painel", use_container_width=True, on_click=on_panel, key=f"toggle_sidebar_{key_prefix}")
    with col_action:
        st.button("Sair", use_container_width=True, on_click=on_logout, key=f"logout_{key_prefix}")
    st.markdown("</div>", unsafe_allow_html=True)


def open_form_shell(extra_class: str = "") -> None:
    classes = "brida-form-shell"
    if extra_class:
        classes += f" {extra_class}"
    st.markdown(f'<div class="{classes}">', unsafe_allow_html=True)


def close_form_shell() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def open_table_shell(title: str = "", caption: str = "") -> None:
    st.markdown('<div class="table-shell brida-table-shell">', unsafe_allow_html=True)
    if title:
        st.markdown(f'<p class="brida-section-title">{html.escape(title)}</p>', unsafe_allow_html=True)
    if caption:
        st.markdown(f'<p class="brida-page-subtitle">{html.escape(caption)}</p>', unsafe_allow_html=True)


def close_table_shell() -> None:
    st.markdown("</div>", unsafe_allow_html=True)
