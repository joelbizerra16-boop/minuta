from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class TipoOperacaoNf(str, Enum):
    IMPRESSAO_ORIGINAL = "IMPRESSAO_ORIGINAL"
    REIMPRESSAO = "REIMPRESSAO"
    COMPLEMENTACAO = "COMPLEMENTACAO"
    REENTREGA = "REENTREGA"
    CANCELAMENTO = "CANCELAMENTO"


TIPO_OPERACAO_LABELS: dict[TipoOperacaoNf, str] = {
    TipoOperacaoNf.IMPRESSAO_ORIGINAL: "IMPRESSAO ORIGINAL",
    TipoOperacaoNf.REIMPRESSAO: "REIMPRESSAO",
    TipoOperacaoNf.COMPLEMENTACAO: "COMPLEMENTACAO",
    TipoOperacaoNf.REENTREGA: "REENTREGA",
    TipoOperacaoNf.CANCELAMENTO: "CANCELAMENTO",
}

TIPO_OPERACAO_BADGE_CLASS: dict[TipoOperacaoNf, str] = {
    TipoOperacaoNf.IMPRESSAO_ORIGINAL: "nf-badge-impressao",
    TipoOperacaoNf.REIMPRESSAO: "nf-badge-reimpressao",
    TipoOperacaoNf.COMPLEMENTACAO: "nf-badge-complementacao",
    TipoOperacaoNf.REENTREGA: "nf-badge-reentrega",
    TipoOperacaoNf.CANCELAMENTO: "nf-badge-cancelamento",
}


@dataclass(frozen=True)
class NfAuditoriaResumo:
    primeira_utilizacao: str
    ultima_utilizacao: str
    total_impressoes: int
    total_reimpressoes: int
    total_complementacoes: int
    total_reentregas: int
    carregamentos_distintos: int


@dataclass(frozen=True)
class NfAuditoriaEvento:
    data: str
    hora: str
    usuario: str
    operacao: TipoOperacaoNf
    operacao_label: str
    numero_carregamento: str
    motorista: str
    rota: str
    placa: str
    status_carregamento: str
    filial: str
    tipo_operacao: str
    ordenacao: float


@dataclass
class NfAuditoriaCard:
    token: str
    nf: str
    cliente: str
    situacao_atual: str
    quantidade_utilizacoes: int
    ultima_utilizacao: str
    resumo: NfAuditoriaResumo
    eventos: list[NfAuditoriaEvento] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "token": self.token,
            "nf": self.nf,
            "cliente": self.cliente,
            "situacao_atual": self.situacao_atual,
            "quantidade_utilizacoes": self.quantidade_utilizacoes,
            "ultima_utilizacao": self.ultima_utilizacao,
            "resumo": {
                "primeira_utilizacao": self.resumo.primeira_utilizacao,
                "ultima_utilizacao": self.resumo.ultima_utilizacao,
                "total_impressoes": self.resumo.total_impressoes,
                "total_reimpressoes": self.resumo.total_reimpressoes,
                "total_complementacoes": self.resumo.total_complementacoes,
                "total_reentregas": self.resumo.total_reentregas,
                "carregamentos_distintos": self.resumo.carregamentos_distintos,
            },
            "eventos": [
                {
                    "data": evento.data,
                    "hora": evento.hora,
                    "usuario": evento.usuario,
                    "operacao": evento.operacao.value,
                    "operacao_label": evento.operacao_label,
                    "numero_carregamento": evento.numero_carregamento,
                    "motorista": evento.motorista,
                    "rota": evento.rota,
                    "placa": evento.placa,
                    "status_carregamento": evento.status_carregamento,
                    "filial": evento.filial,
                    "tipo_operacao": evento.tipo_operacao,
                    "ordenacao": evento.ordenacao,
                }
                for evento in self.eventos
            ],
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> NfAuditoriaCard:
        resumo_payload = payload.get("resumo", {}) or {}
        eventos = []
        for item in payload.get("eventos", []) or []:
            operacao = TipoOperacaoNf(str(item.get("operacao", TipoOperacaoNf.IMPRESSAO_ORIGINAL.value)))
            eventos.append(
                NfAuditoriaEvento(
                    data=str(item.get("data", "") or ""),
                    hora=str(item.get("hora", "") or ""),
                    usuario=str(item.get("usuario", "") or ""),
                    operacao=operacao,
                    operacao_label=str(item.get("operacao_label", "") or TIPO_OPERACAO_LABELS[operacao]),
                    numero_carregamento=str(item.get("numero_carregamento", "") or ""),
                    motorista=str(item.get("motorista", "") or ""),
                    rota=str(item.get("rota", "") or ""),
                    placa=str(item.get("placa", "") or ""),
                    status_carregamento=str(item.get("status_carregamento", "") or ""),
                    filial=str(item.get("filial", "") or ""),
                    tipo_operacao=str(item.get("tipo_operacao", "") or ""),
                    ordenacao=float(item.get("ordenacao", 0) or 0),
                )
            )
        return cls(
            token=str(payload.get("token", "") or ""),
            nf=str(payload.get("nf", "") or ""),
            cliente=str(payload.get("cliente", "") or ""),
            situacao_atual=str(payload.get("situacao_atual", "") or ""),
            quantidade_utilizacoes=int(payload.get("quantidade_utilizacoes", 0) or 0),
            ultima_utilizacao=str(payload.get("ultima_utilizacao", "") or ""),
            resumo=NfAuditoriaResumo(
                primeira_utilizacao=str(resumo_payload.get("primeira_utilizacao", "") or ""),
                ultima_utilizacao=str(resumo_payload.get("ultima_utilizacao", "") or ""),
                total_impressoes=int(resumo_payload.get("total_impressoes", 0) or 0),
                total_reimpressoes=int(resumo_payload.get("total_reimpressoes", 0) or 0),
                total_complementacoes=int(resumo_payload.get("total_complementacoes", 0) or 0),
                total_reentregas=int(resumo_payload.get("total_reentregas", 0) or 0),
                carregamentos_distintos=int(resumo_payload.get("carregamentos_distintos", 0) or 0),
            ),
            eventos=eventos,
        )


@dataclass
class AuditoriaNfLote:
    data_consulta: str
    cards: list[NfAuditoriaCard] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "data_consulta": self.data_consulta,
            "cards": [card.to_dict() for card in self.cards],
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> AuditoriaNfLote:
        return cls(
            data_consulta=str(payload.get("data_consulta", "") or ""),
            cards=[NfAuditoriaCard.from_dict(item) for item in payload.get("cards", []) or []],
        )
