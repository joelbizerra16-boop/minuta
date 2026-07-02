from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass(frozen=True)
class RastreabilidadeNfResumo:
    numero_nf: str
    chave_nfe: str
    destinatario: str
    quantidade_itens: int
    peso_total: float
    quantidade_carregamentos: int
    quantidade_reentregas: int
    primeira_saida: str
    ultima_saida: str
    status_atual: str
    nf_criado_em: datetime | None = None


@dataclass(frozen=True)
class RastreabilidadeHistoricoLinha:
    data_hora: datetime
    numero_carregamento: str
    modalidade: str
    status: str
    usuario: str
    motorista: str
    veiculo: str
    placa: str
    rota: str
    pdf_gerado: str
    documento: str
    reentrega: bool = False
    balcao: bool = False


@dataclass(frozen=True)
class RastreabilidadeReentregaLinha:
    data: str
    carregamento: str
    usuario: str
    motivo: str
    status: str


@dataclass(frozen=True)
class RastreabilidadeVeiculoLinha:
    veiculo: str
    placa: str
    quantidade_viagens: int
    motorista: str


@dataclass(frozen=True)
class RastreabilidadeUsuarioLinha:
    usuario: str
    quantidade_operacoes: int
    primeira_operacao: str
    ultima_operacao: str


@dataclass(frozen=True)
class RastreabilidadeModalidadeLinha:
    modalidade: str
    quantidade: int


@dataclass(frozen=True)
class RastreabilidadeDocumentoLinha:
    minuta: str
    romaneio: str
    data: str
    usuario: str
    quantidade_impressoes: int
    ultima_impressao: str


@dataclass(frozen=True)
class RastreabilidadeEstatisticas:
    total_carregamentos: int
    total_itens_expedidos: int
    peso_expedido: float
    total_reentregas: int
    total_balcao: int
    veiculos_diferentes: int
    motoristas_diferentes: int
    usuarios_envolvidos: int


@dataclass(frozen=True)
class RastreabilidadeTimelineEvento:
    rotulo: str
    data_hora: datetime


@dataclass
class RastreabilidadeNfRelatorio:
    empresa: str
    emitido_em: datetime
    emitido_por: str
    resumo: RastreabilidadeNfResumo
    historico: list[RastreabilidadeHistoricoLinha] = field(default_factory=list)
    reentregas: list[RastreabilidadeReentregaLinha] = field(default_factory=list)
    veiculos: list[RastreabilidadeVeiculoLinha] = field(default_factory=list)
    usuarios: list[RastreabilidadeUsuarioLinha] = field(default_factory=list)
    modalidades: list[RastreabilidadeModalidadeLinha] = field(default_factory=list)
    documentos: list[RastreabilidadeDocumentoLinha] = field(default_factory=list)
    estatisticas: RastreabilidadeEstatisticas | None = None
    timeline: list[RastreabilidadeTimelineEvento] = field(default_factory=list)
