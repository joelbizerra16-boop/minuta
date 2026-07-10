from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class CenarioOperacional(str, Enum):
    NOVO = "NOVO"
    MESMO_CARREGAMENTO = "MESMO_CARREGAMENTO"
    REIMPRESSAO = "REIMPRESSAO"
    COMPLEMENTACAO = "COMPLEMENTACAO"
    REENTREGA = "REENTREGA"
    OUTRO_CARREGAMENTO = "OUTRO_CARREGAMENTO"
    CONFLITO_MULTIPLO = "CONFLITO_MULTIPLO"
    NF_CANCELADA = "NF_CANCELADA"
    NF_BLOQUEADA = "NF_BLOQUEADA"
    INCONSISTENTE = "INCONSISTENTE"


class ClassificacaoNfLote(str, Enum):
    NOVA = "NOVA"
    EXISTENTE = "EXISTENTE"
    CANCELADA = "CANCELADA"
    NAO_AUTORIZADA = "NAO_AUTORIZADA"


class DecisaoOperacional(str, Enum):
    NOVO = "NOVO"
    REIMPRIMIR = "REIMPRIMIR"
    COMPLEMENTAR = "COMPLEMENTAR"
    REENTREGA = "REENTREGA"
    CANCELAR = "CANCELAR"


class ClassificacaoOperacionalNf(str, Enum):
    NOVA = "NOVA"
    REENTREGA = "REENTREGA"
    REIMPRESSAO = "REIMPRESSAO"
    COMPLEMENTACAO = "COMPLEMENTACAO"
    BALCAO = "BALCAO"
    DUPLICIDADE = "DUPLICIDADE"
    INVALIDA = "INVALIDA"


class OperacaoProposta(str, Enum):
    INSERT_ITENS = "INSERT_ITENS"
    REUTILIZAR_REGISTROS = "REUTILIZAR_REGISTROS"
    REIMPRESSAO_PDF = "REIMPRESSAO_PDF"
    REGISTRAR_REENTREGA = "REGISTRAR_REENTREGA"
    IGNORAR = "IGNORAR"
    REGISTRAR_OCORRENCIA = "REGISTRAR_OCORRENCIA"


class SeveridadeDiagnostico(str, Enum):
    INFORMATIVO = "INFORMATIVO"
    REQUER_CONFIRMACAO = "REQUER_CONFIRMACAO"
    BLOQUEIO_ESTRUTURAL = "BLOQUEIO_ESTRUTURAL"


@dataclass(frozen=True)
class DiagnosticoNfOperacional:
    token: str
    nf: str
    chave_nfe: str
    classificacao: ClassificacaoOperacionalNf
    impactos: tuple[str, ...] = ()
    riscos: tuple[str, ...] = ()
    recomendacao: str = ""
    acao_proposta: OperacaoProposta = OperacaoProposta.IGNORAR
    requer_confirmacao: bool = False
    mensagens: tuple[str, ...] = ()
    vinculo_carregamento_id: int | None = None
    numero_carregamento: str = ""


@dataclass
class OcorrenciaProcessamentoNf:
    token: str
    nf: str
    classificacao: ClassificacaoOperacionalNf
    sucesso: bool
    mensagem: str = ""


@dataclass
class ResultadoProcessamentoLote:
    total_recebidas: int = 0
    processadas: int = 0
    reentregas: int = 0
    reimpressoes: int = 0
    duplicidades: int = 0
    invalidas: int = 0
    complementadas: int = 0
    ocorrencias: list[OcorrenciaProcessamentoNf] = field(default_factory=list)

    def resumo_texto(self) -> str:
        partes = [f"{self.processadas} processada(s)"]
        if self.reentregas:
            partes.append(f"{self.reentregas} reentrega(s)")
        if self.complementadas:
            partes.append(f"{self.complementadas} complementada(s)")
        if self.reimpressoes:
            partes.append(f"{self.reimpressoes} reimpressao(oes)")
        if self.duplicidades:
            partes.append(f"{self.duplicidades} duplicidade(s) reutilizada(s)")
        if self.invalidas:
            partes.append(f"{self.invalidas} invalida(s)")
        return " | ".join(partes)


@dataclass
class PlanoOperacionalLote:
    decisao_lote: DecisaoOperacional
    diagnostico_nf: list[DiagnosticoNfOperacional] = field(default_factory=list)
    severidade: SeveridadeDiagnostico = SeveridadeDiagnostico.INFORMATIVO
    bloqueio_estrutural: bool = False
    mensagem_bloqueio: str = ""
    carregamento_id: int | None = None
    itens_para_inserir: int = 0
    nfs_para_inserir: int = 0
    nfs_para_reutilizar: int = 0


@dataclass(frozen=True)
class VinculoNfHistorico:
    carregamento_id: int
    numero_carregamento: str
    data: str
    motorista: str
    placa: str
    status: str
    modalidade: str
    nf: str
    chave_nfe: str


@dataclass(frozen=True)
class NfLoteResumo:
    token: str
    nf: str
    chave_nfe: str
    status_nf: str
    classificacao: ClassificacaoNfLote
    vinculo: VinculoNfHistorico | None = None


@dataclass
class DiagnosticoCarregamento:
    cenario: CenarioOperacional
    nfs_total: int = 0
    nfs_novas: int = 0
    nfs_existentes: int = 0
    nfs_canceladas: int = 0
    nfs_nao_autorizadas: int = 0
    carregamento_id: int | None = None
    numero_carregamento: str = ""
    carregamento_data: str = ""
    carregamento_motorista: str = ""
    carregamento_placa: str = ""
    carregamento_status: str = ""
    carregamentos_distintos: int = 0
    requer_decisao: bool = False
    bloqueia_fechamento: bool = False
    mensagens: list[str] = field(default_factory=list)
    nfs: list[NfLoteResumo] = field(default_factory=list)
    opcoes_decisao: list[DecisaoOperacional] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "cenario": self.cenario.value,
            "nfs_total": self.nfs_total,
            "nfs_novas": self.nfs_novas,
            "nfs_existentes": self.nfs_existentes,
            "nfs_canceladas": self.nfs_canceladas,
            "nfs_nao_autorizadas": self.nfs_nao_autorizadas,
            "carregamento_id": self.carregamento_id,
            "numero_carregamento": self.numero_carregamento,
            "carregamento_data": self.carregamento_data,
            "carregamento_motorista": self.carregamento_motorista,
            "carregamento_placa": self.carregamento_placa,
            "carregamento_status": self.carregamento_status,
            "carregamentos_distintos": self.carregamentos_distintos,
            "requer_decisao": self.requer_decisao,
            "bloqueia_fechamento": self.bloqueia_fechamento,
            "mensagens": list(self.mensagens),
            "nfs": [
                {
                    "token": item.token,
                    "nf": item.nf,
                    "chave_nfe": item.chave_nfe,
                    "status_nf": item.status_nf,
                    "classificacao": item.classificacao.value,
                    "vinculo": (
                        {
                            "carregamento_id": item.vinculo.carregamento_id,
                            "numero_carregamento": item.vinculo.numero_carregamento,
                            "data": item.vinculo.data,
                            "motorista": item.vinculo.motorista,
                            "placa": item.vinculo.placa,
                            "status": item.vinculo.status,
                            "modalidade": item.vinculo.modalidade,
                            "nf": item.vinculo.nf,
                            "chave_nfe": item.vinculo.chave_nfe,
                        }
                        if item.vinculo
                        else None
                    ),
                }
                for item in self.nfs
            ],
            "opcoes_decisao": [item.value for item in self.opcoes_decisao],
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> DiagnosticoCarregamento:
        nfs: list[NfLoteResumo] = []
        for item in payload.get("nfs", []) or []:
            if not isinstance(item, dict):
                continue
            vinculo_payload = item.get("vinculo")
            vinculo = None
            if isinstance(vinculo_payload, dict):
                vinculo = VinculoNfHistorico(
                    carregamento_id=int(vinculo_payload.get("carregamento_id", 0) or 0),
                    numero_carregamento=str(vinculo_payload.get("numero_carregamento", "") or ""),
                    data=str(vinculo_payload.get("data", "") or ""),
                    motorista=str(vinculo_payload.get("motorista", "") or ""),
                    placa=str(vinculo_payload.get("placa", "") or ""),
                    status=str(vinculo_payload.get("status", "") or ""),
                    modalidade=str(vinculo_payload.get("modalidade", "") or ""),
                    nf=str(vinculo_payload.get("nf", "") or ""),
                    chave_nfe=str(vinculo_payload.get("chave_nfe", "") or ""),
                )
            nfs.append(
                NfLoteResumo(
                    token=str(item.get("token", "") or ""),
                    nf=str(item.get("nf", "") or ""),
                    chave_nfe=str(item.get("chave_nfe", "") or ""),
                    status_nf=str(item.get("status_nf", "") or ""),
                    classificacao=ClassificacaoNfLote(str(item.get("classificacao", ClassificacaoNfLote.NOVA.value))),
                    vinculo=vinculo,
                )
            )
        return cls(
            cenario=CenarioOperacional(str(payload.get("cenario", CenarioOperacional.NOVO.value))),
            nfs_total=int(payload.get("nfs_total", 0) or 0),
            nfs_novas=int(payload.get("nfs_novas", 0) or 0),
            nfs_existentes=int(payload.get("nfs_existentes", 0) or 0),
            nfs_canceladas=int(payload.get("nfs_canceladas", 0) or 0),
            nfs_nao_autorizadas=int(payload.get("nfs_nao_autorizadas", 0) or 0),
            carregamento_id=payload.get("carregamento_id"),
            numero_carregamento=str(payload.get("numero_carregamento", "") or ""),
            carregamento_data=str(payload.get("carregamento_data", "") or ""),
            carregamento_motorista=str(payload.get("carregamento_motorista", "") or ""),
            carregamento_placa=str(payload.get("carregamento_placa", "") or ""),
            carregamento_status=str(payload.get("carregamento_status", "") or ""),
            carregamentos_distintos=int(payload.get("carregamentos_distintos", 0) or 0),
            requer_decisao=bool(payload.get("requer_decisao", False)),
            bloqueia_fechamento=bool(payload.get("bloqueia_fechamento", False)),
            mensagens=[str(item) for item in payload.get("mensagens", []) or []],
            nfs=nfs,
            opcoes_decisao=[
                DecisaoOperacional(str(value))
                for value in payload.get("opcoes_decisao", []) or []
            ],
        )
