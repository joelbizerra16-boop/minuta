from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from auth.models.usuario import UsuarioPublico
from carregamentos.models.carregamento import (
    MODALIDADE_BALCAO,
    MODALIDADE_VEICULO,
    STATUS_FINALIZADO,
    STATUS_FINALIZADO_R,
    Carregamento,
    CarregamentoFiltro,
    CarregamentoItem,
    HistoricoItemLinha,
    HistoricoListagemResult,
    NfHistoricoConflito,
    utc_now_iso,
)
from carregamentos.repository.carregamento_repository import CarregamentoRepository
from carregamentos.services.nf_validation import NfHistoricoValidator, localizar_nf_no_lote


class CarregamentoService:
    def __init__(self, repository: CarregamentoRepository, data_dir: Path):
        self._repository = repository
        self._data_dir = data_dir
        self._nf_validator = NfHistoricoValidator(repository)

    def list_carregamentos(self, filtro: CarregamentoFiltro | None = None) -> list[Carregamento]:
        if filtro is None:
            return sorted(self._repository.list_all(), key=lambda item: (item.data, item.hora, item.id), reverse=True)
        return self._repository.search(filtro)

    def validar_conflitos_nf(self, processed_df: pd.DataFrame) -> list[NfHistoricoConflito]:
        return self._nf_validator.validar_conflitos_do_lote(processed_df)

    def search_itens_listagem(self, filtro: CarregamentoFiltro) -> HistoricoListagemResult:
        carregamentos = self._repository.search(filtro)
        termo = self._normalize_search_term(filtro.termo_pesquisa)
        linhas: list[HistoricoItemLinha] = []
        carregamento_ids: set[int] = set()

        for carregamento in carregamentos:
            if termo and not self._carregamento_matches_term(carregamento, termo):
                item_matches = [
                    (index, item)
                    for index, item in enumerate(carregamento.itens)
                    if self._item_matches_term(item, termo)
                ]
            else:
                item_matches = list(enumerate(carregamento.itens))

            for item_index, item in item_matches:
                linhas.append(self._build_historico_linha(carregamento, item_index, item))
                carregamento_ids.add(int(carregamento.id))

        return HistoricoListagemResult(
            linhas=tuple(linhas),
            carregamentos_distintos=len(carregamento_ids),
        )

    def get_carregamento(self, carregamento_id: int) -> Carregamento | None:
        return self._repository.get_by_id(carregamento_id)

    def read_document(self, relative_path: str | None) -> bytes:
        if not relative_path:
            return b""
        candidates = [
            self._data_dir / relative_path,
            self._repository.storage_dir.parent / relative_path,
            self._repository.storage_dir / Path(relative_path).name,
        ]
        for document_path in candidates:
            if document_path.is_file():
                return document_path.read_bytes()
        return b""

    def build_processing_signature(self, summary: dict[str, Any], processed_df: pd.DataFrame) -> str:
        signature_parts = [
            str(summary.get("numero_carga", "")),
            str(summary.get("nf_count", 0)),
            str(summary.get("item_count", 0)),
            str(summary.get("peso_total", 0)),
            str(len(processed_df)),
        ]
        return hashlib.sha256("|".join(signature_parts).encode("utf-8")).hexdigest()

    def register_from_processing(
        self,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        carregamento_pdf: bytes,
        romaneio_pdf: bytes | None,
        *,
        modalidade: str = MODALIDADE_VEICULO,
        is_reentrega: bool = False,
    ) -> Carregamento | None:
        if processed_df.empty:
            return None
        if modalidade == MODALIDADE_VEICULO and not carregamento_pdf:
            return None

        numero_carregamento = str(summary.get("numero_carga", "") or "").strip()
        if not numero_carregamento or numero_carregamento == "--":
            return None

        if not is_reentrega:
            conflitos = self.validar_conflitos_nf(processed_df)
            if conflitos:
                raise ValueError("Existem NFs ja vinculadas a outro carregamento.")

        status = STATUS_FINALIZADO_R if is_reentrega else STATUS_FINALIZADO
        motorista = "--" if modalidade == MODALIDADE_BALCAO else str(summary.get("motorista", "--") or "--")
        placa = "--" if modalidade == MODALIDADE_BALCAO else str(summary.get("placa", "--") or "--")

        return self._persist_carregamento(
            numero_carregamento=numero_carregamento,
            summary=summary,
            processed_df=processed_df,
            current_user=current_user,
            carregamento_pdf=carregamento_pdf,
            romaneio_pdf=romaneio_pdf,
            modalidade=modalidade,
            status=status,
            reentrega=is_reentrega,
            motorista=motorista,
            placa=placa,
        )

    def register_entrega_balcao(
        self,
        termo_busca: str,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        carregamento_pdf: bytes | None = None,
        romaneio_pdf: bytes | None = None,
        *,
        is_reentrega: bool = False,
        standalone_balcao: bool = False,
    ) -> Carregamento | None:
        nf_df = localizar_nf_no_lote(processed_df, termo_busca)
        if nf_df.empty:
            raise ValueError("NF nao encontrada no lote atual.")

        if not is_reentrega:
            conflitos = self.validar_conflitos_nf(nf_df)
            if conflitos:
                raise ValueError("conflito_nf")

        nf_label = str(nf_df.iloc[0].get("NF", "") or termo_busca).strip()
        if standalone_balcao:
            numero_carregamento = f"BALCAO-{nf_label}"
            balcao_summary = {
                "filial": str(summary.get("filial", "BRIDA") or "BRIDA"),
                "data_saida": str(summary.get("data_saida", "--") or "--"),
                "nf_count": int(nf_df["NF"].nunique()),
                "item_count": int(len(nf_df)),
                "peso_total": float(nf_df["Peso"].sum()),
            }
        else:
            numero_base = str(summary.get("numero_carga", "--") or "--")
            numero_carregamento = f"{numero_base}-BALCAO-{nf_label}"
            balcao_summary = dict(summary)
            balcao_summary["nf_count"] = int(nf_df["NF"].nunique())
            balcao_summary["item_count"] = int(len(nf_df))
            balcao_summary["peso_total"] = float(nf_df["Peso"].sum())

        return self._persist_carregamento(
            numero_carregamento=numero_carregamento,
            summary=balcao_summary,
            processed_df=nf_df,
            current_user=current_user,
            carregamento_pdf=carregamento_pdf or b"",
            romaneio_pdf=romaneio_pdf,
            modalidade=MODALIDADE_BALCAO,
            status=STATUS_FINALIZADO_R if is_reentrega else STATUS_FINALIZADO,
            reentrega=is_reentrega,
            motorista="--",
            placa="--",
            allow_empty_pdf=True,
        )

    def _persist_carregamento(
        self,
        numero_carregamento: str,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        carregamento_pdf: bytes,
        romaneio_pdf: bytes | None,
        modalidade: str,
        status: str,
        reentrega: bool,
        motorista: str,
        placa: str,
        allow_empty_pdf: bool = False,
    ) -> Carregamento:
        if not carregamento_pdf and not allow_empty_pdf:
            raise ValueError("Documento da minuta nao disponivel.")

        now = datetime.now()
        itens = self._build_itens_from_dataframe(processed_df)
        carregamento = Carregamento(
            id=0,
            numero_carregamento=numero_carregamento,
            data=now.strftime("%Y-%m-%d"),
            hora=now.strftime("%H:%M:%S"),
            usuario=str(current_user.usuario if current_user else "sistema"),
            usuario_id=current_user.id if current_user else None,
            motorista=motorista,
            placa=placa,
            filial=str(summary.get("filial", "BRIDA") or "BRIDA"),
            data_saida=str(summary.get("data_saida", "--") or "--"),
            quantidade_nf=int(summary.get("nf_count", 0) or 0),
            quantidade_itens=int(summary.get("item_count", 0) or len(itens)),
            peso_total=float(summary.get("peso_total", 0) or 0),
            status=status,
            modalidade=modalidade,
            reentrega=reentrega,
            minuta_pdf_path=None,
            romaneio_pdf_path=None,
            itens=itens,
            criado_em=utc_now_iso(),
        )
        saved = self._repository.save(carregamento)
        storage_folder = self._repository.storage_dir / str(saved.id)
        storage_folder.mkdir(parents=True, exist_ok=True)

        if carregamento_pdf:
            minuta_relative = f"carregamentos/{saved.id}/minuta_carregamento.pdf"
            (storage_folder / "minuta_carregamento.pdf").write_bytes(carregamento_pdf)
            saved.minuta_pdf_path = minuta_relative

        if romaneio_pdf:
            romaneio_relative = f"carregamentos/{saved.id}/romaneio_entrega.pdf"
            (storage_folder / "romaneio_entrega.pdf").write_bytes(romaneio_pdf)
            saved.romaneio_pdf_path = romaneio_relative

        return self._repository.save(saved)

    def _build_itens_from_dataframe(self, processed_df: pd.DataFrame) -> list[CarregamentoItem]:
        itens: list[CarregamentoItem] = []
        for _, row in processed_df.iterrows():
            itens.append(
                CarregamentoItem(
                    nf=str(row.get("NF", "") or ""),
                    cprod=str(row.get("cProd", "") or ""),
                    descricao=str(row.get("Descricao", "") or ""),
                    quantidade=float(row.get("Qtd", 0) or 0),
                    unidade=str(row.get("Unidade", "") or ""),
                    peso=float(row.get("Peso", 0) or 0),
                    destinatario=str(row.get("Destinatario", "") or ""),
                    rota=str(row.get("ROTA", "") or ""),
                    chave_nfe=str(row.get("ChaveNFe", "") or ""),
                    status_nf=str(row.get("Status", "") or ""),
                )
            )
        return itens

    def _build_historico_linha(
        self,
        carregamento: Carregamento,
        item_index: int,
        item: CarregamentoItem,
    ) -> HistoricoItemLinha:
        return HistoricoItemLinha(
            carregamento_id=carregamento.id,
            item_index=item_index,
            data=carregamento.data,
            carregamento=carregamento.numero_carregamento,
            nf=item.nf,
            chave_nfe=item.chave_nfe,
            produto=item.cprod,
            descricao=item.descricao,
            quantidade=item.quantidade,
            peso=item.peso,
            destinatario=item.destinatario,
            rota=item.rota,
            motorista=carregamento.motorista,
            placa=carregamento.placa,
            usuario=carregamento.usuario,
            modalidade=carregamento.modalidade,
            status=carregamento.status,
        )

    def _normalize_search_term(self, value: str | None) -> str:
        return re.sub(r"\s+", " ", str(value or "").strip().lower())

    def _carregamento_matches_term(self, carregamento: Carregamento, termo: str) -> bool:
        values = [
            carregamento.numero_carregamento,
            carregamento.usuario,
            carregamento.motorista,
            carregamento.placa,
            carregamento.filial,
            carregamento.data_saida,
            carregamento.modalidade,
            carregamento.status,
        ]
        return any(termo in str(value).lower() for value in values if value)

    def _item_matches_term(self, item: CarregamentoItem, termo: str) -> bool:
        values = [
            item.nf,
            item.cprod,
            item.descricao,
            item.destinatario,
            item.rota,
            item.chave_nfe,
            item.status_nf,
        ]
        return any(termo in str(value).lower() for value in values if value)
