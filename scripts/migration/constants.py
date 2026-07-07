from __future__ import annotations

DOMAIN_TABLES: tuple[str, ...] = (
    "perfil",
    "usuario",
    "configuracao",
    "motorista",
    "veiculo",
    "destinatario",
    "rota",
    "nota_fiscal",
    "item_nota_fiscal",
    "carregamento",
    "item_carregamento",
    "documento",
    "documento_xml",
    "historico_operacional",
    "evento_auditoria",
)

# Ordem de carga respeitando FK (Extract -> Validate -> Transform -> Load).
MIGRATION_TABLE_ORDER: tuple[str, ...] = (
    "perfil",
    "motorista",
    "veiculo",
    "destinatario",
    "rota",
    "usuario",
    "configuracao",
    "nota_fiscal",
    "item_nota_fiscal",
    "carregamento",
    "item_carregamento",
    "documento",
    "documento_xml",
    "historico_operacional",
    "evento_auditoria",
)

TRUNCATE_TABLE_ORDER: tuple[str, ...] = tuple(reversed(MIGRATION_TABLE_ORDER))

BOOLEAN_COLUMNS: dict[str, frozenset[str]] = {
    "usuario": frozenset({"bloqueado", "ativo"}),
    "motorista": frozenset({"ativo"}),
    "veiculo": frozenset({"ativo"}),
    "destinatario": frozenset({"ativo"}),
    "rota": frozenset({"ativo"}),
    "carregamento": frozenset({"reentrega"}),
    "documento_xml": frozenset({"ativo"}),
}

# FK para validacao pre-carga: tabela -> [(coluna, tabela_pai, coluna_pai)]
FOREIGN_KEY_CHECKS: dict[str, tuple[tuple[str, str, str], ...]] = {
    "usuario": (
        ("perfil_id", "perfil", "id"),
        ("criado_por_id", "usuario", "id"),
        ("atualizado_por_id", "usuario", "id"),
    ),
    "configuracao": (("atualizado_por_usuario_id", "usuario", "id"),),
    "nota_fiscal": (
        ("destinatario_id", "destinatario", "id"),
        ("rota_id", "rota", "id"),
    ),
    "item_nota_fiscal": (("nota_fiscal_id", "nota_fiscal", "id"),),
    "carregamento": (
        ("usuario_id", "usuario", "id"),
        ("motorista_id", "motorista", "id"),
        ("veiculo_id", "veiculo", "id"),
        ("ultima_impressao_usuario_id", "usuario", "id"),
    ),
    "item_carregamento": (
        ("carregamento_id", "carregamento", "id"),
        ("nota_fiscal_id", "nota_fiscal", "id"),
    ),
    "documento": (
        ("carregamento_id", "carregamento", "id"),
        ("usuario_id", "usuario", "id"),
    ),
    "documento_xml": (("usuario_id", "usuario", "id"),),
    "historico_operacional": (
        ("carregamento_id", "carregamento", "id"),
        ("usuario_id", "usuario", "id"),
        ("item_carregamento_id", "item_carregamento", "id"),
    ),
    "evento_auditoria": (
        ("usuario_id", "usuario", "id"),
    ),
}

DEFAULT_BATCH_SIZE = 500
