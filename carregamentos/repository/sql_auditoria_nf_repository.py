from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone

from sqlalchemy import bindparam, text

from infrastructure.persistence.engine_info import get_dialect_name
from infrastructure.persistence.sql_compat import trim_both_zeros
from infrastructure.unit_of_work import UnitOfWork

_EVENTOS_CARREGAMENTOS_SQL = """
SELECT
    'historico' AS fonte,
    ho.carregamento_id AS carregamento_id,
    ho.evento AS evento,
    ho.criado_em AS criado_em,
    ho.usuario_id AS usuario_id,
    ho.descricao AS descricao,
    NULL AS metadados_json
FROM historico_operacional ho
WHERE ho.carregamento_id IN :carregamento_ids
UNION ALL
SELECT
    'auditoria' AS fonte,
    ea.entidade_id AS carregamento_id,
    ea.evento AS evento,
    ea.criado_em AS criado_em,
    ea.usuario_id AS usuario_id,
    ea.descricao AS descricao,
    ea.metadados_json AS metadados_json
FROM evento_auditoria ea
WHERE ea.entidade_tipo = 'carregamento'
  AND ea.entidade_id IN :carregamento_ids
ORDER BY criado_em DESC, carregamento_id DESC
"""


def _build_extrato_nf_sql(dialect: str) -> str:
    trim_ic = trim_both_zeros("ic.numero_nf", dialect=dialect)
    trim_nr = trim_both_zeros("nr.numero_nf", dialect=dialect)
    nf_match = f"{trim_ic} = {trim_nr}"
    return f"""
WITH nf_ref AS (
    SELECT :numero_nf AS numero_nf, :chave_nfe AS chave_nfe
),
carregamentos_nf AS (
    SELECT DISTINCT c.id AS carregamento_id
    FROM carregamento c
    INNER JOIN item_carregamento ic ON ic.carregamento_id = c.id
    CROSS JOIN nf_ref nr
    WHERE (
        nr.chave_nfe <> ''
        AND ic.chave_nfe = nr.chave_nfe
    ) OR (
        {nf_match}
    )
),
rota_nf AS (
    SELECT
        ic.carregamento_id AS carregamento_id,
        MIN(COALESCE(ic.rota, '')) AS rota,
        MIN(COALESCE(ic.destinatario, '')) AS destinatario
    FROM item_carregamento ic
    INNER JOIN carregamentos_nf cn ON cn.carregamento_id = ic.carregamento_id
    CROSS JOIN nf_ref nr
    WHERE (
        nr.chave_nfe <> ''
        AND ic.chave_nfe = nr.chave_nfe
    ) OR (
        {nf_match}
    )
    GROUP BY ic.carregamento_id
),
base_eventos AS (
    SELECT
        'historico' AS fonte,
        ho.id AS evento_id,
        ho.carregamento_id AS carregamento_id,
        ho.evento AS evento,
        ho.criado_em AS criado_em,
        ho.usuario_id AS usuario_id,
        COALESCE(ho.descricao, '') AS descricao,
        NULL AS metadados_json
    FROM historico_operacional ho
    INNER JOIN carregamentos_nf cn ON cn.carregamento_id = ho.carregamento_id
    UNION ALL
    SELECT
        'auditoria' AS fonte,
        ea.id AS evento_id,
        ea.entidade_id AS carregamento_id,
        ea.evento AS evento,
        ea.criado_em AS criado_em,
        ea.usuario_id AS usuario_id,
        COALESCE(ea.descricao, '') AS descricao,
        ea.metadados_json AS metadados_json
    FROM evento_auditoria ea
    INNER JOIN carregamentos_nf cn ON cn.carregamento_id = ea.entidade_id
    WHERE ea.entidade_tipo = 'carregamento'
)
SELECT
    be.fonte,
    be.evento_id,
    be.carregamento_id,
    be.evento,
    be.criado_em,
    be.descricao,
    be.metadados_json,
    c.numero_carregamento,
    COALESCE(c.motorista, '') AS motorista,
    COALESCE(c.placa, '') AS placa,
    COALESCE(c.modalidade, '') AS modalidade,
    COALESCE(c.status, '') AS status,
    COALESCE(u.usuario, '') AS usuario,
    COALESCE(rn.rota, '') AS rota,
    COALESCE(rn.destinatario, '') AS destinatario
FROM base_eventos be
INNER JOIN carregamento c ON c.id = be.carregamento_id
INNER JOIN usuario u ON u.id = be.usuario_id
LEFT JOIN rota_nf rn ON rn.carregamento_id = be.carregamento_id
ORDER BY be.criado_em ASC, be.evento_id ASC
"""


@dataclass(frozen=True)
class EventoCarregamentoRegistro:
    fonte: str
    carregamento_id: int
    evento: str
    criado_em: datetime | None
    usuario_id: int | None
    descricao: str
    metadados_json: str | None


@dataclass(frozen=True)
class MovimentacaoNfExtratoRegistro:
    fonte: str
    evento_id: int
    carregamento_id: int
    evento: str
    criado_em: datetime | None
    descricao: str
    metadados_json: str | None
    numero_carregamento: str
    motorista: str
    placa: str
    modalidade: str
    status: str
    usuario: str
    rota: str
    destinatario: str


class SqlAuditoriaNfRepository:
    @staticmethod
    def _parse_criado_em(value: object) -> datetime | None:
        """Normaliza criado_em vindo de SQL bruto (SQLite retorna str, nao datetime)."""
        if value is None:
            return None
        if isinstance(value, datetime):
            if value.tzinfo is None:
                return value.replace(tzinfo=timezone.utc)
            return value
        text_value = str(value).strip()
        if not text_value:
            return None
        normalized = text_value.replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError:
            return None
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed

    def buscar_eventos_por_carregamentos(self, carregamento_ids: list[int]) -> list[EventoCarregamentoRegistro]:
        if not carregamento_ids:
            return []

        stmt = text(_EVENTOS_CARREGAMENTOS_SQL).bindparams(
            bindparam("carregamento_ids", expanding=True),
        )

        with UnitOfWork() as uow:
            rows = uow.session.execute(
                stmt,
                {"carregamento_ids": sorted({int(item) for item in carregamento_ids})},
            ).mappings().all()

        return [
            EventoCarregamentoRegistro(
                fonte=str(row.get("fonte", "") or ""),
                carregamento_id=int(row.get("carregamento_id", 0) or 0),
                evento=str(row.get("evento", "") or ""),
                criado_em=self._parse_criado_em(row.get("criado_em")),
                usuario_id=int(row["usuario_id"]) if row.get("usuario_id") is not None else None,
                descricao=str(row.get("descricao", "") or ""),
                metadados_json=row.get("metadados_json"),
            )
            for row in rows
        ]

    def buscar_extrato_movimentacoes_nf(
        self,
        *,
        numero_nf: str,
        chave_nfe: str = "",
    ) -> list[MovimentacaoNfExtratoRegistro]:
        numero = str(numero_nf or "").strip()
        chave = str(chave_nfe or "").strip()
        if not numero and not chave:
            return []

        with UnitOfWork() as uow:
            dialect = get_dialect_name(uow.session)
            rows = uow.session.execute(
                text(_build_extrato_nf_sql(dialect)),
                {"numero_nf": numero, "chave_nfe": chave},
            ).mappings().all()

        return [
            MovimentacaoNfExtratoRegistro(
                fonte=str(row.get("fonte", "") or ""),
                evento_id=int(row.get("evento_id", 0) or 0),
                carregamento_id=int(row.get("carregamento_id", 0) or 0),
                evento=str(row.get("evento", "") or ""),
                criado_em=self._parse_criado_em(row.get("criado_em")),
                descricao=str(row.get("descricao", "") or ""),
                metadados_json=row.get("metadados_json"),
                numero_carregamento=str(row.get("numero_carregamento", "") or ""),
                motorista=str(row.get("motorista", "") or ""),
                placa=str(row.get("placa", "") or ""),
                modalidade=str(row.get("modalidade", "") or ""),
                status=str(row.get("status", "") or ""),
                usuario=str(row.get("usuario", "") or ""),
                rota=str(row.get("rota", "") or ""),
                destinatario=str(row.get("destinatario", "") or ""),
            )
            for row in rows
        ]
