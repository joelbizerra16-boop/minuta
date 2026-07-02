from __future__ import annotations

import streamlit as st

from auth.models.usuario import PERFIL_ADMIN, UsuarioPublico

SESSION_LOGGED_IN = "logado"
SESSION_USER = "auth_user"


def create_session(user: UsuarioPublico) -> None:
    st.session_state[SESSION_LOGGED_IN] = True
    st.session_state[SESSION_USER] = user.to_dict()


def clear_session() -> None:
    st.session_state[SESSION_LOGGED_IN] = False
    st.session_state.pop(SESSION_USER, None)
    st.session_state.pop("auth_access_error", None)


def clear_session_on_logout() -> None:
    clear_session()
    st.session_state.pop("login_error", None)
    st.session_state.pop("login_success", None)
    st.session_state.pop("usuarios_action", None)
    st.session_state.pop("usuarios_selected_id", None)
    st.session_state.pop("usuarios_feedback", None)


def is_logged_in() -> bool:
    return bool(st.session_state.get(SESSION_LOGGED_IN, False) and st.session_state.get(SESSION_USER))


def get_current_user() -> UsuarioPublico | None:
    payload = st.session_state.get(SESSION_USER)
    if not isinstance(payload, dict):
        return None
    return UsuarioPublico(
        id=int(payload.get("id", 0)),
        nome=str(payload.get("nome", "")),
        usuario=str(payload.get("usuario", "")),
        perfil=str(payload.get("perfil", "")),
        ativo=bool(payload.get("ativo", True)),
        bloqueado=bool(payload.get("bloqueado", False)),
        criado_em=str(payload.get("criado_em", "")),
        ultimo_login=payload.get("ultimo_login"),
    )


def is_admin() -> bool:
    current_user = get_current_user()
    return current_user is not None and current_user.perfil == PERFIL_ADMIN


def require_admin() -> bool:
    return is_admin()
