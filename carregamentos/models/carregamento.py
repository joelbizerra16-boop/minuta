from __future__ import annotations

import re
from dataclasses import dataclass

MODALIDADE_VEICULO = "VEÍCULO"
MODALIDADE_BALCAO = "BALCÃO"
STATUS_FINALIZADO = "FINALIZADO"
STATUS_FINALIZADO_R = "FINALIZADO R"


def utc_now_iso() -> str:
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_chave_nfe(value: object) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits if len(digits) == 44 else ""


def normalize_nf_number(value: object) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits.lstrip("0") or digits


@dataclass
class CarregamentoItem:
    nf: str
    cprod: str
    descricao: str
    quantidade: float
    unidade: str
    peso: float
    destinatario: str
    rota: str
    chave_nfe: str = ""
    status_nf: str = ""

    def to_dict(self) -> dict:
        return {
            "nf": self.nf,
            "cprod": self.cprod,
            "descricao": self.descricao,
            "quantidade": self.quantidade,
            "unidade": self.unidade,
            "peso": self.peso,
            "destinatario": self.destinatario,
            "rota": self.rota,
            "chave_nfe": self.chave_nfe,
            "status_nf": self.status_nf,
        }

    @classmethod
    def from_dict(cls, payload: dict) -> CarregamentoItem:
        return cls(
            nf=str(payload.get("nf", "")),
            cprod=str(payload.get("cprod", "")),
            descricao=str(payload.get("descricao", "")),
            quantidade=float(payload.get("quantidade", 0) or 0),
            unidade=str(payload.get("unidade", "")),
            peso=float(payload.get("peso", 0) or 0),
            destinatario=str(payload.get("destinatario", "")),
            rota=str(payload.get("rota", "")),
            chave_nfe=str(payload.get("chave_nfe", "")),
            status_nf=str(payload.get("status_nf", "")),
        )


@dataclass
class Carregamento:
    id: int
    numero_carregamento: str
    data: str
    hora: str
    usuario: str
    usuario_id: int | None
    motorista: str
    placa: str
    filial: str
    data_saida: str
    quantidade_nf: int
    quantidade_itens: int
    peso_total: float
    status: str
    modalidade: str
    reentrega: bool
    minuta_pdf_path: str | None
    romaneio_pdf_path: str | None
    itens: list[CarregamentoItem]
    criado_em: str
    quantidade_impressoes: int = 0
    ultima_impressao_em: str | None = None
    ultima_impressao_usuario: str | None = None

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "numero_carregamento": self.numero_carregamento,
            "data": self.data,
            "hora": self.hora,
            "usuario": self.usuario,
            "usuario_id": self.usuario_id,
            "motorista": self.motorista,
            "placa": self.placa,
            "filial": self.filial,
            "data_saida": self.data_saida,
            "quantidade_nf": self.quantidade_nf,
            "quantidade_itens": self.quantidade_itens,
            "peso_total": self.peso_total,
            "status": self.status,
            "modalidade": self.modalidade,
            "reentrega": self.reentrega,
            "minuta_pdf_path": self.minuta_pdf_path,
            "romaneio_pdf_path": self.romaneio_pdf_path,
            "itens": [item.to_dict() for item in self.itens],
            "criado_em": self.criado_em,
            "quantidade_impressoes": self.quantidade_impressoes,
            "ultima_impressao_em": self.ultima_impressao_em,
            "ultima_impressao_usuario": self.ultima_impressao_usuario,
        }

    @classmethod
    def from_dict(cls, payload: dict) -> Carregamento:
        raw_items = payload.get("itens", [])
        items = [CarregamentoItem.from_dict(item) for item in raw_items if isinstance(item, dict)]
        status = str(payload.get("status", STATUS_FINALIZADO))
        return cls(
            id=int(payload.get("id", 0)),
            numero_carregamento=str(payload.get("numero_carregamento", "")),
            data=str(payload.get("data", "")),
            hora=str(payload.get("hora", "")),
            usuario=str(payload.get("usuario", "")),
            usuario_id=payload.get("usuario_id"),
            motorista=str(payload.get("motorista", "")),
            placa=str(payload.get("placa", "")),
            filial=str(payload.get("filial", "")),
            data_saida=str(payload.get("data_saida", "")),
            quantidade_nf=int(payload.get("quantidade_nf", 0) or 0),
            quantidade_itens=int(payload.get("quantidade_itens", 0) or 0),
            peso_total=float(payload.get("peso_total", 0) or 0),
            status=status,
            modalidade=str(payload.get("modalidade", MODALIDADE_VEICULO)),
            reentrega=bool(payload.get("reentrega", status == STATUS_FINALIZADO_R)),
            minuta_pdf_path=payload.get("minuta_pdf_path"),
            romaneio_pdf_path=payload.get("romaneio_pdf_path"),
            itens=items,
            criado_em=str(payload.get("criado_em", utc_now_iso())),
            quantidade_impressoes=int(payload.get("quantidade_impressoes", 0) or 0),
            ultima_impressao_em=payload.get("ultima_impressao_em"),
            ultima_impressao_usuario=payload.get("ultima_impressao_usuario"),
        )


@dataclass
class CarregamentoFiltro:
    data_inicial: str | None = None
    data_final: str | None = None
    termo_pesquisa: str | None = None


@dataclass(frozen=True)
class NfHistoricoConflito:
    nf: str
    chave_nfe: str
    numero_carregamento: str
    data: str
    motorista: str
    placa: str
    modalidade: str
    status: str

    def formatar_mensagem(self) -> str:
        data_formatada = self._formatar_data(self.data)
        veiculo = self.placa if self.placa and self.placa != "--" else "nao informado"
        motorista = self.motorista if self.motorista and self.motorista != "--" else "nao informado"
        return (
            f"Esta Nota Fiscal ja esta vinculada ao carregamento nº {self.numero_carregamento}, "
            f"realizado em {data_formatada}, veiculo {veiculo} e motorista {motorista}."
        )

    def formatar_mensagem_balcao(self) -> str:
        return "Esta Nota Fiscal ja foi registrada anteriormente."

    @staticmethod
    def _formatar_data(value: str) -> str:
        parts = str(value or "").split("-")
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return value or "--"


@dataclass(frozen=True)
class HistoricoListagemResult:
    linhas: tuple[HistoricoItemLinha, ...]
    carregamentos_distintos: int


@dataclass(frozen=True)
class HistoricoItemLinha:
    carregamento_id: int
    item_index: int
    data: str
    carregamento: str
    nf: str
    chave_nfe: str
    produto: str
    descricao: str
    quantidade: float
    peso: float
    destinatario: str
    rota: str
    motorista: str
    placa: str
    usuario: str
    modalidade: str
    status: str
