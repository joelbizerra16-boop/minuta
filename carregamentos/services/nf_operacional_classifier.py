from __future__ import annotations

from collections import defaultdict

import pandas as pd

from carregamentos.models.carregamento import (
    Carregamento,
    CarregamentoItem,
    normalize_chave_nfe,
    normalize_nf_number,
)
from carregamentos.models.operacional import (
    ClassificacaoNfLote,
    ClassificacaoOperacionalNf,
    DecisaoOperacional,
    DiagnosticoCarregamento,
    DiagnosticoNfOperacional,
    NfLoteResumo,
    OperacaoProposta,
)
from carregamentos.services.validacao_item_carregamento import item_ja_existe_no_carregamento


def _row_token(row: pd.Series) -> str:
    chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
    if chave:
        return chave
    nf_norm = normalize_nf_number(row.get("NF", ""))
    return f"nf:{nf_norm}" if nf_norm else ""


def _build_item_from_row(row: pd.Series) -> CarregamentoItem:
    return CarregamentoItem(
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


class NfOperacionalClassifier:
    """Classifica cada NF do lote antes da persistência."""

    def classificar_lote(
        self,
        processed_df: pd.DataFrame,
        diagnostico: DiagnosticoCarregamento,
        decisao_lote: DecisaoOperacional,
        carregamento_existente: Carregamento | None = None,
    ) -> list[DiagnosticoNfOperacional]:
        if processed_df.empty:
            return []

        nf_resumos = {item.token: item for item in diagnostico.nfs}
        rows_por_token: dict[str, list[pd.Series]] = defaultdict(list)
        for _, row in processed_df.iterrows():
            token = _row_token(row)
            if token:
                rows_por_token[token].append(row)

        resultados: list[DiagnosticoNfOperacional] = []
        for token, rows in rows_por_token.items():
            resumo = nf_resumos.get(token)
            nf = str(rows[0].get("NF", "") or "")
            chave = normalize_chave_nfe(rows[0].get("ChaveNFe", ""))
            if resumo is None:
                resultados.append(
                    self._diagnostico_invalida(
                        token=token,
                        nf=nf,
                        chave_nfe=chave,
                        mensagem="NF nao encontrada no diagnostico operacional.",
                    )
                )
                continue
            resultados.append(
                self._classificar_nf(
                    resumo=resumo,
                    rows=rows,
                    decisao_lote=decisao_lote,
                    carregamento_existente=carregamento_existente,
                    carregamento_alvo_id=diagnostico.carregamento_id,
                    numero_carregamento=diagnostico.numero_carregamento,
                )
            )
        return resultados

    def _classificar_nf(
        self,
        *,
        resumo: NfLoteResumo,
        rows: list[pd.Series],
        decisao_lote: DecisaoOperacional,
        carregamento_existente: Carregamento | None,
        carregamento_alvo_id: int | None,
        numero_carregamento: str,
    ) -> DiagnosticoNfOperacional:
        if resumo.classificacao == ClassificacaoNfLote.CANCELADA:
            return self._diagnostico_invalida(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                mensagem="NF cancelada — requer correcao fiscal antes do processamento.",
                impactos=("NF com status cancelado no lote.",),
                riscos=("Processar NF cancelada pode gerar inconsistencia fiscal.",),
                recomendacao="Remova a NF cancelada do lote ou substitua o XML.",
            )

        if resumo.classificacao == ClassificacaoNfLote.NAO_AUTORIZADA:
            return self._diagnostico_invalida(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                mensagem="NF nao autorizada — verifique o status no XML.",
                recomendacao="Confirme a autorizacao da SEFAZ antes de continuar.",
            )

        if decisao_lote == DecisaoOperacional.REIMPRIMIR:
            return DiagnosticoNfOperacional(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                classificacao=ClassificacaoOperacionalNf.REIMPRESSAO,
                impactos=("Nenhum item sera alterado.", "Nova impressao de documentos."),
                riscos=("Minuta/romaneio serao reemitidos.",),
                recomendacao="Confirme se deseja emitir nova impressao.",
                acao_proposta=OperacaoProposta.REIMPRESSAO_PDF,
                requer_confirmacao=True,
                vinculo_carregamento_id=carregamento_alvo_id,
                numero_carregamento=numero_carregamento,
            )

        if decisao_lote == DecisaoOperacional.REENTREGA:
            return DiagnosticoNfOperacional(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                classificacao=ClassificacaoOperacionalNf.REENTREGA,
                impactos=(
                    "Esta NF ja possui historico operacional.",
                    "Sera reutilizado o historico existente.",
                ),
                riscos=("Status do carregamento sera alterado para FINALIZADO R.",),
                recomendacao="Confirme somente se a reentrega e intencional.",
                acao_proposta=OperacaoProposta.REGISTRAR_REENTREGA,
                requer_confirmacao=True,
                vinculo_carregamento_id=carregamento_alvo_id,
                numero_carregamento=numero_carregamento,
            )

        if decisao_lote == DecisaoOperacional.NOVO:
            return DiagnosticoNfOperacional(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                classificacao=ClassificacaoOperacionalNf.NOVA,
                impactos=("NF sera registrada em novo carregamento.",),
                recomendacao="Prosseguir com o registro.",
                acao_proposta=OperacaoProposta.INSERT_ITENS,
                requer_confirmacao=False,
            )

        # COMPLEMENTAR — classificação por NF no lote misto
        if resumo.classificacao == ClassificacaoNfLote.NOVA:
            return DiagnosticoNfOperacional(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                classificacao=ClassificacaoOperacionalNf.NOVA,
                impactos=(f"Itens da NF {resumo.nf} serao adicionados ao carregamento {numero_carregamento}.",),
                recomendacao="Prosseguir com a complementacao desta NF.",
                acao_proposta=OperacaoProposta.INSERT_ITENS,
                requer_confirmacao=False,
                vinculo_carregamento_id=carregamento_alvo_id,
                numero_carregamento=numero_carregamento,
            )

        # NF existente no mesmo carregamento
        vinculo = resumo.vinculo
        if vinculo and carregamento_alvo_id and int(vinculo.carregamento_id) == int(carregamento_alvo_id):
            itens_nf = [_build_item_from_row(row) for row in rows]
            todos_duplicados = carregamento_existente is not None and all(
                item_ja_existe_no_carregamento(carregamento_existente, item) for item in itens_nf
            )
            if todos_duplicados:
                return DiagnosticoNfOperacional(
                    token=resumo.token,
                    nf=resumo.nf,
                    chave_nfe=resumo.chave_nfe,
                    classificacao=ClassificacaoOperacionalNf.DUPLICIDADE,
                    impactos=(
                        f"NF {resumo.nf} ja pertence ao carregamento {numero_carregamento}.",
                        "Itens existentes serao reutilizados.",
                    ),
                    riscos=("Nenhum INSERT sera realizado para esta NF.",),
                    recomendacao="Reutilizar registros existentes.",
                    acao_proposta=OperacaoProposta.REUTILIZAR_REGISTROS,
                    requer_confirmacao=False,
                    mensagens=("Duplicidade detectada — reutilizacao preventiva.",),
                    vinculo_carregamento_id=carregamento_alvo_id,
                    numero_carregamento=numero_carregamento,
                )
            return DiagnosticoNfOperacional(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                classificacao=ClassificacaoOperacionalNf.REENTREGA,
                impactos=(
                    f"NF {resumo.nf} possui historico no carregamento {numero_carregamento}.",
                    "Sera reutilizado o historico existente.",
                ),
                riscos=("Itens ja vinculados nao serao duplicados.",),
                recomendacao="Confirmar complementacao sem reinserir esta NF.",
                acao_proposta=OperacaoProposta.REUTILIZAR_REGISTROS,
                requer_confirmacao=False,
                vinculo_carregamento_id=carregamento_alvo_id,
                numero_carregamento=numero_carregamento,
            )

        if vinculo and carregamento_alvo_id and int(vinculo.carregamento_id) != int(carregamento_alvo_id):
            return self._diagnostico_invalida(
                token=resumo.token,
                nf=resumo.nf,
                chave_nfe=resumo.chave_nfe,
                mensagem=(
                    f"NF {resumo.nf} pertence ao carregamento {vinculo.numero_carregamento} "
                    f"e nao ao carregamento alvo {numero_carregamento}."
                ),
                recomendacao="Separe as NFs por carregamento ou autorize reentrega.",
            )

        return DiagnosticoNfOperacional(
            token=resumo.token,
            nf=resumo.nf,
            chave_nfe=resumo.chave_nfe,
            classificacao=ClassificacaoOperacionalNf.COMPLEMENTACAO,
            impactos=(f"NF {resumo.nf} sera tratada na complementacao.",),
            acao_proposta=OperacaoProposta.INSERT_ITENS,
            vinculo_carregamento_id=carregamento_alvo_id,
            numero_carregamento=numero_carregamento,
        )

    @staticmethod
    def _diagnostico_invalida(
        *,
        token: str,
        nf: str,
        chave_nfe: str,
        mensagem: str,
        impactos: tuple[str, ...] = (),
        riscos: tuple[str, ...] = (),
        recomendacao: str = "",
    ) -> DiagnosticoNfOperacional:
        return DiagnosticoNfOperacional(
            token=token,
            nf=nf,
            chave_nfe=chave_nfe,
            classificacao=ClassificacaoOperacionalNf.INVALIDA,
            impactos=impactos,
            riscos=riscos,
            recomendacao=recomendacao,
            acao_proposta=OperacaoProposta.REGISTRAR_OCORRENCIA,
            requer_confirmacao=False,
            mensagens=(mensagem,),
        )
