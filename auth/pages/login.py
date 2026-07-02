from __future__ import annotations

import html

import streamlit as st

from auth.bootstrap import get_auth_service
from auth.security.session import create_session


def render_login_page(
    logo_path,
    on_success_screen: str,
    navigate_callback,
) -> None:
    st.markdown(
        """
    <style>
    .login-stage {
        width: 100%;
        max-width: 1040px;
        margin: 8vh auto 0;
        padding: 0 1.5rem;
        box-sizing: border-box;
    }
    .login-logo-wrap {
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
        padding: 1rem 0;
    }
    .login-shell {
        max-width: 400px;
        margin: 0 auto;
        display: flex;
        flex-direction: column;
        text-align: left;
    }
    .login-intro h2 {
        margin: 0 0 10px;
        color: #1F3A5F;
        font-size: 1.65rem;
        font-weight: 700;
    }
    .login-intro p {
        margin: 0;
        color: #607085;
        line-height: 1.5;
        font-size: 0.95rem;
    }
    .login-feedback-error {
        color: #B42318;
        margin-bottom: 0.75rem;
        font-size: 0.92rem;
    }
    section.main .block-container {
        max-width: 1040px;
        margin-top: 8vh;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    logo_col, login_col = st.columns([1, 1], gap="large", vertical_alignment="center")

    with logo_col:
        if logo_path is not None:
            st.image(str(logo_path), width=320)

    with login_col:
        st.markdown(
            """
            <div class="login-shell">
                <div class="login-intro">
                    <h2>Acesse sua conta</h2>
                    <p>Informe suas credenciais para entrar no sistema.</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        show_password = st.checkbox("Mostrar senha")
        login_error = st.session_state.get("login_error", "")
        if login_error:
            st.markdown(
                f'<div class="login-feedback-error">{html.escape(login_error)}</div>',
                unsafe_allow_html=True,
            )

        with st.form("login_form"):
            username = st.text_input("Usuario", placeholder="Digite seu usuario")
            password = st.text_input(
                "Senha",
                type="default" if show_password else "password",
                placeholder="Digite sua senha",
            )
            submitted = st.form_submit_button("Entrar", use_container_width=True)

    if submitted:
        auth_result = get_auth_service().authenticate(username, password)
        if auth_result.success and auth_result.user is not None:
            create_session(auth_result.user)
            st.session_state["login_error"] = ""
            st.session_state["login_success"] = "Acesso validado com sucesso."
            navigate_callback(on_success_screen)
            st.rerun()
        st.session_state["login_error"] = auth_result.error_message
        st.session_state["login_success"] = ""
