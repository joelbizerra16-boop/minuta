from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from carregamentos.models.carregamento import NfHistoricoConflito, normalize_chave_nfe, normalize_nf_number
from carregamentos.repository.carregamento_repository import CarregamentoRepository


@dataclass(frozen=True)
class NfLoteIdentidade:
    nf: str
    chave_nfe: str


def extrair_identidades_nf_do_lote(processed_df: pd.DataFrame) -> list[NfLoteIdentidade]:
    if processed_df.empty:
        return []

    identidades: list[NfLoteIdentidade] = []
    vistos: set[str] = set()
    for _, row in processed_df.iterrows():
        chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
        nf = normalize_nf_number(row.get("NF", ""))
        token = f"chave:{chave}" if chave else f"nf:{nf}"
        if not chave and not nf:
            continue
        if token in vistos:
            continue
        vistos.add(token)
        identidades.append(NfLoteIdentidade(nf=str(row.get("NF", "") or ""), chave_nfe=chave))
    return identidades


def localizar_nf_no_lote(processed_df: pd.DataFrame, termo: str) -> pd.DataFrame:
    termo_limpo = str(termo or "").strip()
    if processed_df.empty or not termo_limpo:
        return processed_df.iloc[0:0]

    termo_chave = normalize_chave_nfe(termo_limpo)
    termo_nf = normalize_nf_number(termo_limpo)
    mascara = pd.Series(False, index=processed_df.index)

    if termo_chave:
        mascara = mascara | processed_df["ChaveNFe"].astype(str).map(normalize_chave_nfe).eq(termo_chave)
    if termo_nf:
        mascara = mascara | processed_df["NF"].astype(str).map(normalize_nf_number).eq(termo_nf)
    if not termo_chave and not termo_nf:
        termo_lower = termo_limpo.lower()
        mascara = (
            processed_df["NF"].astype(str).str.lower().str.contains(termo_lower, na=False)
            | processed_df["ChaveNFe"].astype(str).str.lower().str.contains(termo_lower, na=False)
        )

    return processed_df[mascara].copy()


class NfHistoricoValidator:
    def __init__(self, repository: CarregamentoRepository):
        self._repository = repository

    def validar_conflitos_do_lote(self, processed_df: pd.DataFrame) -> list[NfHistoricoConflito]:
        conflitos: list[NfHistoricoConflito] = []
        conflitos_vistos: set[str] = set()

        for identidade in extrair_identidades_nf_do_lote(processed_df):
            for conflito in self._buscar_conflitos(identidade.chave_nfe, identidade.nf):
                token = f"{conflito.chave_nfe}:{conflito.nf}:{conflito.numero_carregamento}"
                if token in conflitos_vistos:
                    continue
                conflitos_vistos.add(token)
                conflitos.append(conflito)
        return conflitos

    def _buscar_conflitos(self, chave_nfe: str, nf: str) -> list[NfHistoricoConflito]:
        encontrados: list[NfHistoricoConflito] = []
        chave_normalizada = normalize_chave_nfe(chave_nfe)
        nf_normalizada = normalize_nf_number(nf)

        for carregamento in self._repository.list_all():
            for item in carregamento.itens:
                item_chave = normalize_chave_nfe(item.chave_nfe)
                item_nf = normalize_nf_number(item.nf)
                chave_igual = bool(chave_normalizada and item_chave and chave_normalizada == item_chave)
                nf_igual = bool(nf_normalizada and item_nf and nf_normalizada == item_nf)
                if not chave_igual and not nf_igual:
                    continue
                encontrados.append(
                    NfHistoricoConflito(
                        nf=str(item.nf or nf),
                        chave_nfe=item_chave or chave_normalizada,
                        numero_carregamento=carregamento.numero_carregamento,
                        data=carregamento.data,
                        motorista=carregamento.motorista,
                        placa=carregamento.placa,
                        modalidade=carregamento.modalidade,
                        status=carregamento.status,
                    )
                )
        return encontrados
