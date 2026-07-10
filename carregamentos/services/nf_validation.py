from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

import pandas as pd

from carregamentos.models.carregamento import Carregamento, NfHistoricoConflito, normalize_chave_nfe, normalize_nf_number
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


def _montar_indice_conflitos(
    carregamentos: list[Carregamento],
) -> tuple[dict[str, list[NfHistoricoConflito]], dict[str, list[NfHistoricoConflito]]]:
    por_chave: dict[str, list[NfHistoricoConflito]] = defaultdict(list)
    por_nf: dict[str, list[NfHistoricoConflito]] = defaultdict(list)

    for carregamento in carregamentos:
        for item in carregamento.itens:
            item_chave = normalize_chave_nfe(item.chave_nfe)
            item_nf = normalize_nf_number(item.nf)
            conflito = NfHistoricoConflito(
                nf=str(item.nf or ""),
                chave_nfe=item_chave,
                numero_carregamento=carregamento.numero_carregamento,
                data=carregamento.data,
                motorista=carregamento.motorista,
                placa=carregamento.placa,
                modalidade=carregamento.modalidade,
                status=carregamento.status,
            )
            if item_chave:
                por_chave[item_chave].append(conflito)
            if item_nf:
                por_nf[item_nf].append(conflito)

    return por_chave, por_nf


class NfHistoricoValidator:
    def __init__(self, repository: CarregamentoRepository):
        self._repository = repository

    @staticmethod
    def _identidades_para_consulta(
        identidades: list[NfLoteIdentidade],
    ) -> tuple[set[str], set[str]]:
        chaves: set[str] = set()
        numeros: set[str] = set()
        for identidade in identidades:
            chave = normalize_chave_nfe(identidade.chave_nfe)
            nf_norm = normalize_nf_number(identidade.nf)
            nf_raw = str(identidade.nf or "").strip()
            if chave:
                chaves.add(chave)
            if nf_norm:
                numeros.add(nf_norm)
            if nf_raw:
                numeros.add(nf_raw)
        return chaves, numeros

    def validar_conflitos_do_lote(self, processed_df: pd.DataFrame) -> list[NfHistoricoConflito]:
        identidades = extrair_identidades_nf_do_lote(processed_df)
        if not identidades:
            return []

        chaves, numeros = self._identidades_para_consulta(identidades)
        carregamentos = self._repository.list_by_item_identidades(
            chaves_nfe=chaves,
            numeros_nf=numeros,
        )
        por_chave, por_nf = _montar_indice_conflitos(carregamentos)

        conflitos: list[NfHistoricoConflito] = []
        conflitos_vistos: set[str] = set()

        for identidade in identidades:
            chave_normalizada = normalize_chave_nfe(identidade.chave_nfe)
            nf_normalizada = normalize_nf_number(identidade.nf)
            candidatos: list[NfHistoricoConflito] = []
            if chave_normalizada:
                candidatos.extend(por_chave.get(chave_normalizada, []))
            if nf_normalizada:
                candidatos.extend(por_nf.get(nf_normalizada, []))

            for conflito in candidatos:
                token = f"{conflito.chave_nfe}:{conflito.nf}:{conflito.numero_carregamento}"
                if token in conflitos_vistos:
                    continue
                conflitos_vistos.add(token)
                conflitos.append(conflito)

        return conflitos

    def _buscar_conflitos(self, chave_nfe: str, nf: str) -> list[NfHistoricoConflito]:
        """Compatibilidade com chamadas unitárias legadas."""
        identidade = NfLoteIdentidade(nf=str(nf or ""), chave_nfe=normalize_chave_nfe(chave_nfe))
        chaves, numeros = self._identidades_para_consulta([identidade])
        carregamentos = self._repository.list_by_item_identidades(
            chaves_nfe=chaves,
            numeros_nf=numeros,
        )
        por_chave, por_nf = _montar_indice_conflitos(carregamentos)
        chave_normalizada = normalize_chave_nfe(chave_nfe)
        nf_normalizada = normalize_nf_number(nf)
        encontrados: list[NfHistoricoConflito] = []
        if chave_normalizada:
            encontrados.extend(por_chave.get(chave_normalizada, []))
        if nf_normalizada:
            encontrados.extend(por_nf.get(nf_normalizada, []))
        dedup: list[NfHistoricoConflito] = []
        vistos: set[str] = set()
        for conflito in encontrados:
            token = f"{conflito.chave_nfe}:{conflito.nf}:{conflito.numero_carregamento}"
            if token in vistos:
                continue
            vistos.add(token)
            dedup.append(conflito)
        return dedup
