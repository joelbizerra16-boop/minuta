from __future__ import annotations

import html
from datetime import datetime

import streamlit as st

from auth.bootstrap import get_usuario_service
from auth.models.usuario import PERFIL_ADMIN, PERFIL_OPERADOR
from auth.security.session import get_current_user


def _format_datetime(value: str | None) -> str:
    if not value:
        return "--"
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed.strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(value)


def _clear_user_form_state() -> None:
    st.session_state.pop("usuarios_action", None)
    st.session_state.pop("usuarios_selected_id", None)


def _perfil_badge(perfil: str) -> str:
    css = "brida-badge-operador"
    if perfil == PERFIL_ADMIN:
        css = "brida-badge-admin"
    return f'<span class="brida-badge {css}">{html.escape(perfil)}</span>'


def _status_badge(user) -> str:
    if not user.ativo:
        return '<span class="brida-badge brida-badge-inativo">Inativo</span>'
    if user.bloqueado:
        return '<span class="brida-badge brida-badge-bloqueado">Bloqueado</span>'
    return '<span class="brida-badge brida-badge-ativo">Ativo</span>'


def _render_user_form(mode: str, user_id: int | None = None) -> None:
    service = get_usuario_service()
    existing = service.get_user(user_id) if user_id is not None else None
    title = "Novo usuario" if mode == "create" else "Editar usuario"
    st.markdown(f'<div class="section-title">{html.escape(title)}</div>', unsafe_allow_html=True)

    with st.form(f"usuario_form_{mode}_{user_id or 'new'}"):
        nome = st.text_input("Nome completo", value=existing.nome if existing else "")
        usuario = st.text_input(
            "Usuario",
            value=existing.usuario if existing else "",
            disabled=mode == "edit",
        )
        perfil = st.selectbox(
            "Perfil",
            options=[PERFIL_ADMIN, PERFIL_OPERADOR],
            index=0 if (existing and existing.perfil == PERFIL_ADMIN) else 1,
        )
        senha = None
        if mode == "create":
            senha = st.text_input("Senha", type="password", placeholder="Minimo 6 caracteres")
        action_col_1, action_col_2 = st.columns(2)
        with action_col_1:
            submitted = st.form_submit_button("Salvar", use_container_width=True)
        with action_col_2:
            cancelled = st.form_submit_button("Cancelar", use_container_width=True)

    if cancelled:
        _clear_user_form_state()
        st.rerun()

    if submitted:
        try:
            if mode == "create":
                service.create_user(nome, usuario, str(senha or ""), perfil)
                st.session_state["usuarios_feedback"] = ("success", "Usuario cadastrado com sucesso.")
            else:
                service.update_user(int(user_id), nome, perfil)
                st.session_state["usuarios_feedback"] = ("success", "Usuario atualizado com sucesso.")
            _clear_user_form_state()
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))


def _render_password_form(user_id: int) -> None:
    service = get_usuario_service()
    existing = service.get_user(user_id)
    if existing is None:
        st.error("Usuario nao encontrado.")
        return

    st.markdown(
        f'<div class="section-title">Alterar senha — {html.escape(existing.nome)}</div>',
        unsafe_allow_html=True,
    )
    with st.form(f"usuario_password_form_{user_id}"):
        nova_senha = st.text_input("Nova senha", type="password", placeholder="Minimo 6 caracteres")
        confirmar_senha = st.text_input("Confirmar senha", type="password")
        action_col_1, action_col_2 = st.columns(2)
        with action_col_1:
            submitted = st.form_submit_button("Atualizar senha", use_container_width=True)
        with action_col_2:
            cancelled = st.form_submit_button("Cancelar", use_container_width=True)

    if cancelled:
        _clear_user_form_state()
        st.rerun()

    if submitted:
        if nova_senha != confirmar_senha:
            st.error("As senhas informadas nao conferem.")
            return
        try:
            service.change_password(user_id, nova_senha)
            st.session_state["usuarios_feedback"] = ("success", "Senha alterada com sucesso.")
            _clear_user_form_state()
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))


def render_usuarios_page(
    render_header_callback,
    navigate_callback,
    menu_screen: str,
) -> None:
    render_header_callback("Gerenciamento de Usuarios", "Cadastro, edicao e controle de acesso")

    current_user = get_current_user()
    feedback = st.session_state.pop("usuarios_feedback", None)
    if feedback:
        level, message = feedback
        if level == "success":
            st.success(message)
        else:
            st.error(message)

    action = st.session_state.get("usuarios_action")
    selected_id = st.session_state.get("usuarios_selected_id")

    if action == "create":
        _render_user_form("create")
        return

    if action == "edit" and selected_id is not None:
        _render_user_form("edit", int(selected_id))
        return

    if action == "password" and selected_id is not None:
        _render_password_form(int(selected_id))
        return

    top_col_1, top_col_2 = st.columns([3, 1])
    with top_col_1:
        st.markdown('<div class="section-title">Lista de Usuarios</div>', unsafe_allow_html=True)
    with top_col_2:
        if st.button("+ Novo usuario", use_container_width=True, key="new_user_button", type="primary"):
            st.session_state["usuarios_action"] = "create"
            st.rerun()

    service = get_usuario_service()
    users = service.list_users(include_inactive=True)
    if not users:
        st.info("Nenhum usuario cadastrado.")
        return

    table_rows = []
    for user in users:
        table_rows.append(
            "<tr>"
            f"<td>@{html.escape(user.usuario)}</td>"
            f"<td><strong>{html.escape(user.nome)}</strong></td>"
            f"<td>{_perfil_badge(user.perfil)}</td>"
            f"<td>{_status_badge(user)}</td>"
            f"<td>{html.escape(_format_datetime(user.ultimo_login))}</td>"
            "</tr>"
        )

    with st.container(border=True):
        st.markdown(
            """
        <table class="brida-users-table">
            <thead>
                <tr>
                    <th>Usuario</th>
                    <th>Nome</th>
                    <th>Perfil</th>
                    <th>Status</th>
                    <th>Ultimo acesso</th>
                </tr>
            </thead>
            <tbody>
        """
            + "".join(table_rows)
            + """
            </tbody>
        </table>
        """,
            unsafe_allow_html=True,
        )

    for user in users:
        label_col, action_cols = st.columns([2.2, 3.8])
        with label_col:
            st.caption(f"Acoes — {user.nome}")
        with action_cols:
            buttons = st.columns(5)
            with buttons[0]:
                if st.button("Editar", key=f"edit_user_{user.id}", disabled=not user.ativo):
                    st.session_state["usuarios_action"] = "edit"
                    st.session_state["usuarios_selected_id"] = user.id
                    st.rerun()
            with buttons[1]:
                if st.button("Senha", key=f"password_user_{user.id}", disabled=not user.ativo):
                    st.session_state["usuarios_action"] = "password"
                    st.session_state["usuarios_selected_id"] = user.id
                    st.rerun()
            with buttons[2]:
                if user.ativo and not user.bloqueado and user.id != (current_user.id if current_user else -1):
                    if st.button("Bloquear", key=f"block_user_{user.id}"):
                        try:
                            service.block_user(user.id)
                            st.session_state["usuarios_feedback"] = ("success", f"Usuario {user.usuario} bloqueado.")
                            st.rerun()
                        except ValueError as exc:
                            st.error(str(exc))
            with buttons[3]:
                if user.ativo and user.bloqueado:
                    if st.button("Desbloquear", key=f"unblock_user_{user.id}"):
                        try:
                            service.unblock_user(user.id)
                            st.session_state["usuarios_feedback"] = ("success", f"Usuario {user.usuario} desbloqueado.")
                            st.rerun()
                        except ValueError as exc:
                            st.error(str(exc))
            with buttons[4]:
                if user.ativo and user.id != (current_user.id if current_user else -1):
                    if st.button("Excluir", key=f"delete_user_{user.id}"):
                        try:
                            service.delete_user(user.id)
                            st.session_state["usuarios_feedback"] = ("success", f"Usuario {user.usuario} inativado.")
                            st.rerun()
                        except ValueError as exc:
                            st.error(str(exc))
