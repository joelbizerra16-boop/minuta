from __future__ import annotations

from dataclasses import dataclass, field

from auth.models.usuario import Usuario


@dataclass
class UsuarioComparisonDiff:
    user_id: int
    username: str
    field: str
    json_value: str
    sql_value: str


@dataclass
class UsuarioComparisonReport:
    json_total: int = 0
    sql_total: int = 0
    matched: int = 0
    divergent: int = 0
    json_only: list[str] = field(default_factory=list)
    sql_only: list[str] = field(default_factory=list)
    diffs: list[UsuarioComparisonDiff] = field(default_factory=list)
    invalid_users: list[str] = field(default_factory=list)
    duplicate_usernames: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        return (
            self.divergent == 0
            and not self.json_only
            and not self.sql_only
            and not self.invalid_users
            and not self.duplicate_usernames
            and self.json_total == self.sql_total
            and self.json_total > 0
        )


def validate_usuario(usuario: Usuario) -> list[str]:
    errors: list[str] = []
    if usuario.id <= 0:
        errors.append("id invalido")
    if not usuario.nome.strip():
        errors.append("nome obrigatorio")
    if not usuario.usuario.strip():
        errors.append("login obrigatorio")
    if not usuario.senha_hash.strip():
        errors.append("senha_hash obrigatorio")
    if usuario.perfil not in {"ADMIN", "OPERADOR"}:
        errors.append(f"perfil invalido: {usuario.perfil}")
    if not usuario.senha_hash.startswith("pbkdf2_sha256$"):
        errors.append("hash de senha fora do padrao")
    return errors


def compare_usuarios(json_users: list[Usuario], sql_users: list[Usuario]) -> UsuarioComparisonReport:
    report = UsuarioComparisonReport()
    report.json_total = len(json_users)
    report.sql_total = len(sql_users)

    seen_usernames: dict[str, int] = {}
    for user in json_users:
        errors = validate_usuario(user)
        if errors:
            report.invalid_users.append(f"{user.usuario} ({', '.join(errors)})")
        count = seen_usernames.get(user.usuario, 0) + 1
        seen_usernames[user.usuario] = count
    report.duplicate_usernames = [name for name, count in seen_usernames.items() if count > 1]

    json_map = {user.id: user for user in json_users}
    sql_map = {user.id: user for user in sql_users}

    for user_id, json_user in json_map.items():
        sql_user = sql_map.get(user_id)
        if sql_user is None:
            report.json_only.append(json_user.usuario)
            continue
        user_diffs = _compare_pair(json_user, sql_user)
        if user_diffs:
            report.divergent += 1
            report.diffs.extend(user_diffs)
        else:
            report.matched += 1

    for user_id, sql_user in sql_map.items():
        if user_id not in json_map:
            report.sql_only.append(sql_user.usuario)

    return report


def _compare_pair(json_user: Usuario, sql_user: Usuario) -> list[UsuarioComparisonDiff]:
    diffs: list[UsuarioComparisonDiff] = []
    fields = (
        "nome",
        "usuario",
        "senha_hash",
        "perfil",
        "ativo",
        "bloqueado",
        "criado_em",
        "ultimo_login",
    )
    for field_name in fields:
        json_value = _normalize_field(field_name, getattr(json_user, field_name))
        sql_value = _normalize_field(field_name, getattr(sql_user, field_name))
        if json_value != sql_value:
            diffs.append(
                UsuarioComparisonDiff(
                    user_id=json_user.id,
                    username=json_user.usuario,
                    field=field_name,
                    json_value=json_value,
                    sql_value=sql_value,
                )
            )
    return diffs


def _normalize_field(field_name: str, value: object) -> str:
    if value is None:
        return ""
    if field_name in {"ativo", "bloqueado"}:
        return "1" if bool(value) else "0"
    if field_name in {"criado_em", "ultimo_login"}:
        from auth.repository.usuario_mapper import format_iso_datetime, parse_iso_datetime

        parsed = parse_iso_datetime(str(value) if value else None)
        return format_iso_datetime(parsed) or ""
    return str(value).strip()
