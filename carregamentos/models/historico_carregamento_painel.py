from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class HistoricoCarregamentoEstatisticas:
    numero_carregamento: str
    total_nfs: int
    peso_total_kg: float
    valor_total: float
    primeira_impressao_data: str
    primeira_impressao_hora: str
    ultima_impressao_data: str
    ultima_impressao_hora: str
    quantidade_reimpressoes: int
    quantidade_complementacoes: int
    quantidade_reentregas: int


@dataclass(frozen=True)
class HistoricoNfLinha:
    nf: str
    cliente: str
    cidade: str
    uf: str
    peso_kg: float
    valor_nf: float
    primeira_utilizacao_data: str
    primeira_utilizacao_hora: str
    usuario_carregamento: str
    numero_carregamento: str
    quantidade_reimpressoes: int
    ultima_reimpressao_data: str
    ultima_reimpressao_hora: str
    ultimo_usuario_impressao: str
    status_atual: str
    origem: str
    excel_utilizado: str
    xml_status: str
    xml_arquivo: str


@dataclass(frozen=True)
class HistoricoImpressaoLinha:
    data: str
    hora: str
    usuario: str
    tipo: str
    resultado: str


@dataclass(frozen=True)
class HistoricoComplementacaoLinha:
    data: str
    hora: str
    usuario: str
    nfs_adicionadas: int
    observacao: str


@dataclass(frozen=True)
class HistoricoReentregaLinha:
    data: str
    hora: str
    usuario: str
    motivo: str
    status: str


@dataclass
class HistoricoCarregamentoPainel:
    carregamento_id: int
    numero_carregamento: str
    excel_contexto: str
    data_analise: str
    estatisticas: HistoricoCarregamentoEstatisticas
    nfs: list[HistoricoNfLinha] = field(default_factory=list)
    impressoes: list[HistoricoImpressaoLinha] = field(default_factory=list)
    complementacoes: list[HistoricoComplementacaoLinha] = field(default_factory=list)
    reentregas: list[HistoricoReentregaLinha] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "carregamento_id": self.carregamento_id,
            "numero_carregamento": self.numero_carregamento,
            "excel_contexto": self.excel_contexto,
            "data_analise": self.data_analise,
            "estatisticas": {
                "numero_carregamento": self.estatisticas.numero_carregamento,
                "total_nfs": self.estatisticas.total_nfs,
                "peso_total_kg": self.estatisticas.peso_total_kg,
                "valor_total": self.estatisticas.valor_total,
                "primeira_impressao_data": self.estatisticas.primeira_impressao_data,
                "primeira_impressao_hora": self.estatisticas.primeira_impressao_hora,
                "ultima_impressao_data": self.estatisticas.ultima_impressao_data,
                "ultima_impressao_hora": self.estatisticas.ultima_impressao_hora,
                "quantidade_reimpressoes": self.estatisticas.quantidade_reimpressoes,
                "quantidade_complementacoes": self.estatisticas.quantidade_complementacoes,
                "quantidade_reentregas": self.estatisticas.quantidade_reentregas,
            },
            "nfs": [item.__dict__ for item in self.nfs],
            "impressoes": [item.__dict__ for item in self.impressoes],
            "complementacoes": [item.__dict__ for item in self.complementacoes],
            "reentregas": [item.__dict__ for item in self.reentregas],
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> HistoricoCarregamentoPainel:
        stats_payload = payload.get("estatisticas", {}) or {}
        return cls(
            carregamento_id=int(payload.get("carregamento_id", 0) or 0),
            numero_carregamento=str(payload.get("numero_carregamento", "") or ""),
            excel_contexto=str(payload.get("excel_contexto", "") or ""),
            data_analise=str(payload.get("data_analise", "") or ""),
            estatisticas=HistoricoCarregamentoEstatisticas(
                numero_carregamento=str(stats_payload.get("numero_carregamento", "") or ""),
                total_nfs=int(stats_payload.get("total_nfs", 0) or 0),
                peso_total_kg=float(stats_payload.get("peso_total_kg", 0) or 0),
                valor_total=float(stats_payload.get("valor_total", 0) or 0),
                primeira_impressao_data=str(stats_payload.get("primeira_impressao_data", "") or ""),
                primeira_impressao_hora=str(stats_payload.get("primeira_impressao_hora", "") or ""),
                ultima_impressao_data=str(stats_payload.get("ultima_impressao_data", "") or ""),
                ultima_impressao_hora=str(stats_payload.get("ultima_impressao_hora", "") or ""),
                quantidade_reimpressoes=int(stats_payload.get("quantidade_reimpressoes", 0) or 0),
                quantidade_complementacoes=int(stats_payload.get("quantidade_complementacoes", 0) or 0),
                quantidade_reentregas=int(stats_payload.get("quantidade_reentregas", 0) or 0),
            ),
            nfs=[HistoricoNfLinha(**item) for item in payload.get("nfs", []) or []],
            impressoes=[HistoricoImpressaoLinha(**item) for item in payload.get("impressoes", []) or []],
            complementacoes=[HistoricoComplementacaoLinha(**item) for item in payload.get("complementacoes", []) or []],
            reentregas=[HistoricoReentregaLinha(**item) for item in payload.get("reentregas", []) or []],
        )
