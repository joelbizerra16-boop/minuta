from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone

import pandas as pd

from carregamentos.models.auditoria_nf import (
    TIPO_OPERACAO_LABELS,
    AuditoriaNfLote,
    NfAuditoriaCard,
    NfAuditoriaEvento,
    NfAuditoriaResumo,
    TipoOperacaoNf,
)
from carregamentos.models.carregamento import Carregamento, normalize_chave_nfe, normalize_nf_number
from carregamentos.models.operacional import (
    CenarioOperacional,
    ClassificacaoNfLote,
    DecisaoOperacional,
    DiagnosticoCarregamento,
    NfLoteResumo,
    VinculoNfHistorico,
)
from carregamentos.repository.carregamento_repository import CarregamentoRepository
from carregamentos.repository.sql_auditoria_nf_repository import SqlAuditoriaNfRepository
from infrastructure.models.constants import (
    AUDIT_EVENTO_COMPLEMENTACAO,
    AUDIT_EVENTO_ENTREGA_BALCAO,
    AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
    AUDIT_EVENTO_REENTREGA,
    AUDIT_EVENTO_REIMPRESSAO,
    HISTORICO_EVENTO_COMPLEMENTACAO,
    HISTORICO_EVENTO_ENTREGA_BALCAO,
    HISTORICO_EVENTO_FINALIZACAO,
    HISTORICO_EVENTO_REENTREGA,
)


def _is_canceled_nf_status(value: object) -> bool:
    text = str(value or "").strip().lower()
    return "cancel" in text


def _is_authorized_nf_status(value: object) -> bool:
    text = str(value or "").strip().lower()
    return "autoriz" in text


class IndiceHistoricoCarregamento:
    def __init__(self, carregamentos: list[Carregamento]) -> None:
        self._por_chave: dict[str, VinculoNfHistorico] = {}
        self._por_nf: dict[str, VinculoNfHistorico] = {}
        for carregamento in carregamentos:
            nfs_vistas: set[str] = set()
            for item in carregamento.itens:
                chave = normalize_chave_nfe(item.chave_nfe)
                nf_norm = normalize_nf_number(item.nf)
                token = chave or (f"nf:{nf_norm}" if nf_norm else "")
                if not token or token in nfs_vistas:
                    continue
                nfs_vistas.add(token)
                vinculo = VinculoNfHistorico(
                    carregamento_id=int(carregamento.id),
                    numero_carregamento=str(carregamento.numero_carregamento or ""),
                    data=str(carregamento.data or ""),
                    motorista=str(carregamento.motorista or ""),
                    placa=str(carregamento.placa or ""),
                    status=str(carregamento.status or ""),
                    modalidade=str(carregamento.modalidade or ""),
                    nf=str(item.nf or ""),
                    chave_nfe=chave,
                )
                if chave:
                    self._por_chave[chave] = vinculo
                if nf_norm:
                    self._por_nf[nf_norm] = vinculo

    def buscar(self, chave_nfe: str, nf: str) -> VinculoNfHistorico | None:
        chave = normalize_chave_nfe(chave_nfe)
        if chave and chave in self._por_chave:
            return self._por_chave[chave]
        nf_norm = normalize_nf_number(nf)
        if nf_norm and nf_norm in self._por_nf:
            return self._por_nf[nf_norm]
        return None


class AnaliseOperacionalService:
    """Ponto unico de inteligencia operacional baseada em carregamento."""

    def __init__(self, repository: CarregamentoRepository) -> None:
        self._repository = repository
        self._indice_cache: IndiceHistoricoCarregamento | None = None
        self._carregamentos_cache: list[Carregamento] | None = None
        self._auditoria_nf_repo = SqlAuditoriaNfRepository()

    def invalidar_cache(self) -> None:
        self._indice_cache = None
        self._carregamentos_cache = None

    def _obter_carregamentos(self) -> list[Carregamento]:
        if self._carregamentos_cache is None:
            self._carregamentos_cache = self._repository.list_all()
        return self._carregamentos_cache

    def _obter_indice(self) -> IndiceHistoricoCarregamento:
        if self._indice_cache is None:
            self._indice_cache = IndiceHistoricoCarregamento(self._obter_carregamentos())
        return self._indice_cache

    def analisar_lote_processado(self, processed_df: pd.DataFrame) -> DiagnosticoCarregamento:
        if processed_df.empty:
            return DiagnosticoCarregamento(
                cenario=CenarioOperacional.INCONSISTENTE,
                bloqueia_fechamento=True,
                mensagens=["Nenhum dado processado para analise operacional."],
            )

        indice = self._obter_indice()
        nfs_lote = self._extrair_nfs_unicas(processed_df)
        diagnostico = DiagnosticoCarregamento(cenario=CenarioOperacional.NOVO, nfs=nfs_lote)
        diagnostico.nfs_total = len(nfs_lote)

        carregamentos_encontrados: dict[int, VinculoNfHistorico] = {}
        for nf_resumo in nfs_lote:
            if nf_resumo.classificacao == ClassificacaoNfLote.CANCELADA:
                diagnostico.nfs_canceladas += 1
                continue
            if nf_resumo.classificacao == ClassificacaoNfLote.NAO_AUTORIZADA:
                diagnostico.nfs_nao_autorizadas += 1
                continue
            if nf_resumo.vinculo is None:
                diagnostico.nfs_novas += 1
                continue

            diagnostico.nfs_existentes += 1
            carregamentos_encontrados[nf_resumo.vinculo.carregamento_id] = nf_resumo.vinculo

        diagnostico.carregamentos_distintos = len(carregamentos_encontrados)

        if diagnostico.nfs_canceladas > 0:
            diagnostico.cenario = CenarioOperacional.NF_CANCELADA
            diagnostico.requer_decisao = True
            diagnostico.mensagens.append(
                f"{diagnostico.nfs_canceladas} NF(s) cancelada(s) no lote. "
                "Revise o lote e confirme a operacao desejada."
            )
            diagnostico.opcoes_decisao = [DecisaoOperacional.CANCELAR]
            return diagnostico

        if diagnostico.nfs_novas == diagnostico.nfs_total:
            diagnostico.cenario = CenarioOperacional.NOVO
            diagnostico.opcoes_decisao = [DecisaoOperacional.NOVO]
            return diagnostico

        if diagnostico.carregamentos_distintos > 1:
            diagnostico.cenario = CenarioOperacional.CONFLITO_MULTIPLO
            diagnostico.requer_decisao = True
            diagnostico.opcoes_decisao = [DecisaoOperacional.CANCELAR]
            diagnostico.mensagens.append(
                "As NFs do lote pertencem a carregamentos diferentes. "
                "Revise o historico e confirme a operacao desejada."
            )
            return diagnostico

        if diagnostico.carregamentos_distintos == 1:
            vinculo = next(iter(carregamentos_encontrados.values()))
            diagnostico.carregamento_id = vinculo.carregamento_id
            diagnostico.numero_carregamento = vinculo.numero_carregamento
            diagnostico.carregamento_data = self._formatar_data(vinculo.data)
            diagnostico.carregamento_motorista = vinculo.motorista or "--"
            diagnostico.carregamento_placa = vinculo.placa or "--"
            diagnostico.carregamento_status = vinculo.status or "--"

            if diagnostico.nfs_existentes == diagnostico.nfs_total:
                diagnostico.cenario = CenarioOperacional.REIMPRESSAO
                diagnostico.requer_decisao = True
                diagnostico.opcoes_decisao = [
                    DecisaoOperacional.REIMPRIMIR,
                    DecisaoOperacional.REENTREGA,
                    DecisaoOperacional.CANCELAR,
                ]
                diagnostico.mensagens.append(
                    f"Todas as {diagnostico.nfs_existentes} NF(s) pertencem ao carregamento "
                    f"{vinculo.numero_carregamento}."
                )
                return diagnostico

            if diagnostico.nfs_novas > 0 and diagnostico.nfs_existentes > 0:
                diagnostico.cenario = CenarioOperacional.COMPLEMENTACAO
                diagnostico.requer_decisao = True
                diagnostico.opcoes_decisao = [
                    DecisaoOperacional.COMPLEMENTAR,
                    DecisaoOperacional.CANCELAR,
                ]
                diagnostico.mensagens.append(
                    f"{diagnostico.nfs_existentes} NF(s) ja pertencem ao carregamento "
                    f"{vinculo.numero_carregamento} e {diagnostico.nfs_novas} NF(s) sao novas."
                )
                return diagnostico

        diagnostico.cenario = CenarioOperacional.INCONSISTENTE
        diagnostico.bloqueia_fechamento = True
        diagnostico.requer_decisao = True
        diagnostico.opcoes_decisao = [DecisaoOperacional.CANCELAR]
        diagnostico.mensagens.append("Nao foi possivel classificar o carregamento do lote.")
        return diagnostico

    def analisar_xml_records(self, xml_records: list[dict[str, object]]) -> DiagnosticoCarregamento:
        if not xml_records:
            return DiagnosticoCarregamento(cenario=CenarioOperacional.NOVO)

        rows: list[dict[str, object]] = []
        for record in xml_records:
            if not isinstance(record, dict):
                continue
            rows.append(
                {
                    "NF": record.get("NF", record.get("nf_normalizada", "")),
                    "ChaveNFe": record.get("ChaveNFe", ""),
                    "Status": record.get("StatusNF", record.get("Status", "")),
                }
            )
        if not rows:
            return DiagnosticoCarregamento(cenario=CenarioOperacional.NOVO)
        return self.analisar_lote_processado(pd.DataFrame(rows))

    def montar_auditoria_nfs_lote(self, processed_df: pd.DataFrame) -> AuditoriaNfLote:
        nfs_lote = self._extrair_nfs_unicas(processed_df)
        clientes_por_token = self._mapear_clientes_lote(processed_df)
        if not nfs_lote:
            return AuditoriaNfLote(data_consulta=self._formatar_data_hora_consulta(datetime.now()))

        carregamentos = self._obter_carregamentos()
        nf_para_carregamentos, carregamento_map, rota_por_nf = self._mapear_nf_carregamentos(
            carregamentos,
            {item.token for item in nfs_lote},
        )
        carregamento_ids = sorted({cid for ids in nf_para_carregamentos.values() for cid in ids})
        eventos_brutos = self._auditoria_nf_repo.buscar_eventos_por_carregamentos(carregamento_ids)
        usuario_ids = {
            int(evento.usuario_id)
            for evento in eventos_brutos
            if evento.usuario_id is not None
        }
        for carregamento in carregamentos:
            if carregamento.usuario_id:
                usuario_ids.add(int(carregamento.usuario_id))
        usuarios_por_id = self._carregar_usuarios(usuario_ids)

        eventos_por_carregamento: dict[int, list] = defaultdict(list)
        for evento in eventos_brutos:
            eventos_por_carregamento[int(evento.carregamento_id)].append(evento)

        audit_primeira_por_carregamento = {
            int(evento.carregamento_id)
            for evento in eventos_brutos
            if evento.fonte == "auditoria" and str(evento.evento or "") == AUDIT_EVENTO_PRIMEIRA_IMPRESSAO
        }

        cards: list[NfAuditoriaCard] = []
        for nf_resumo in nfs_lote:
            cards.append(
                self._montar_card_nf(
                    nf_resumo=nf_resumo,
                    cliente=clientes_por_token.get(nf_resumo.token, "--"),
                    carregamento_ids=nf_para_carregamentos.get(nf_resumo.token, set()),
                    carregamento_map=carregamento_map,
                    rota_por_nf=rota_por_nf,
                    eventos_por_carregamento=eventos_por_carregamento,
                    usuarios_por_id=usuarios_por_id,
                    audit_primeira_por_carregamento=audit_primeira_por_carregamento,
                )
            )

        cards.sort(key=lambda item: item.nf)
        return AuditoriaNfLote(
            data_consulta=self._formatar_data_hora_consulta(datetime.now()),
            cards=cards,
        )

    def _montar_card_nf(
        self,
        *,
        nf_resumo: NfLoteResumo,
        cliente: str,
        carregamento_ids: set[int],
        carregamento_map: dict[int, Carregamento],
        rota_por_nf: dict[str, str],
        eventos_por_carregamento: dict[int, list],
        usuarios_por_id: dict[int, str],
        audit_primeira_por_carregamento: set[int],
    ) -> NfAuditoriaCard:
        situacao = (
            "Nunca utilizada anteriormente"
            if nf_resumo.classificacao == ClassificacaoNfLote.NOVA
            else "Ja utilizada anteriormente"
        )
        eventos_card: list[NfAuditoriaEvento] = []
        vistos: set[tuple[str, int, str, str]] = set()

        for carregamento_id in sorted(carregamento_ids):
            carregamento = carregamento_map.get(carregamento_id)
            if carregamento is None:
                continue
            for evento in eventos_por_carregamento.get(carregamento_id, []):
                if (
                    evento.fonte == "historico"
                    and str(evento.evento or "") == HISTORICO_EVENTO_FINALIZACAO
                    and carregamento_id in audit_primeira_por_carregamento
                ):
                    continue
                operacao = self._classificar_evento(evento.evento, evento.fonte)
                if operacao is None:
                    continue
                data, hora, ordenacao = self._formatar_evento_data_hora(evento.criado_em, carregamento)
                chave = (operacao.value, carregamento_id, data, hora)
                if chave in vistos:
                    continue
                vistos.add(chave)
                usuario = usuarios_por_id.get(int(evento.usuario_id or 0), carregamento.usuario or "--")
                eventos_card.append(
                    NfAuditoriaEvento(
                        data=data,
                        hora=hora,
                        usuario=str(usuario or "--"),
                        operacao=operacao,
                        operacao_label=TIPO_OPERACAO_LABELS[operacao],
                        numero_carregamento=str(carregamento.numero_carregamento or "--"),
                        motorista=str(carregamento.motorista or "--"),
                        rota=str(rota_por_nf.get(nf_resumo.token, "--")),
                        placa=str(carregamento.placa or "--"),
                        status_carregamento=str(carregamento.status or "--"),
                        filial=str(carregamento.filial or "--"),
                        tipo_operacao=str(carregamento.modalidade or "--"),
                        ordenacao=ordenacao,
                    )
                )

        eventos_card.sort(key=lambda item: item.ordenacao, reverse=True)
        resumo = self._montar_resumo_nf(eventos_card, len(carregamento_ids))
        ultima_utilizacao = resumo.ultima_utilizacao or "--"

        return NfAuditoriaCard(
            token=nf_resumo.token,
            nf=nf_resumo.nf,
            cliente=cliente or "--",
            situacao_atual=situacao,
            quantidade_utilizacoes=len(eventos_card),
            ultima_utilizacao=ultima_utilizacao,
            resumo=resumo,
            eventos=eventos_card,
        )

    @staticmethod
    def _mapear_clientes_lote(processed_df: pd.DataFrame) -> dict[str, str]:
        clientes: dict[str, str] = {}
        for _, row in processed_df.iterrows():
            chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
            nf_norm = normalize_nf_number(row.get("NF", ""))
            token = chave or (f"nf:{nf_norm}" if nf_norm else "")
            if not token or token in clientes:
                continue
            clientes[token] = str(row.get("Destinatario", "") or "--")
        return clientes

    @staticmethod
    def _mapear_nf_carregamentos(
        carregamentos: list[Carregamento],
        tokens_lote: set[str],
    ) -> tuple[dict[str, set[int]], dict[int, Carregamento], dict[str, str]]:
        nf_para_carregamentos: dict[str, set[int]] = defaultdict(set)
        carregamento_map: dict[int, Carregamento] = {}
        rota_por_nf: dict[str, str] = {}

        for carregamento in carregamentos:
            carregamento_map[int(carregamento.id)] = carregamento
            vistos_no_carregamento: set[str] = set()
            for item in carregamento.itens:
                chave = normalize_chave_nfe(item.chave_nfe)
                nf_norm = normalize_nf_number(item.nf)
                token = chave or (f"nf:{nf_norm}" if nf_norm else "")
                if not token or token not in tokens_lote or token in vistos_no_carregamento:
                    continue
                vistos_no_carregamento.add(token)
                nf_para_carregamentos[token].add(int(carregamento.id))
                if token not in rota_por_nf:
                    rota_por_nf[token] = str(item.rota or "--")

        return nf_para_carregamentos, carregamento_map, rota_por_nf

    @staticmethod
    def _carregar_usuarios(usuario_ids: set[int]) -> dict[int, str]:
        if not usuario_ids:
            return {}
        from sqlalchemy import select

        from infrastructure.models.usuario import UsuarioORM
        from infrastructure.unit_of_work import UnitOfWork

        with UnitOfWork() as uow:
            rows = uow.session.scalars(
                select(UsuarioORM).where(UsuarioORM.id.in_(sorted(usuario_ids)))
            ).all()
        return {int(row.id): str(row.usuario or "--") for row in rows}

    @staticmethod
    def _classificar_evento(evento: str, fonte: str) -> TipoOperacaoNf | None:
        evento_norm = str(evento or "").strip().upper()
        if evento_norm in {AUDIT_EVENTO_PRIMEIRA_IMPRESSAO, HISTORICO_EVENTO_FINALIZACAO, HISTORICO_EVENTO_ENTREGA_BALCAO, AUDIT_EVENTO_ENTREGA_BALCAO}:
            return TipoOperacaoNf.IMPRESSAO_ORIGINAL
        if evento_norm == AUDIT_EVENTO_REIMPRESSAO:
            return TipoOperacaoNf.REIMPRESSAO
        if evento_norm in {AUDIT_EVENTO_COMPLEMENTACAO, HISTORICO_EVENTO_COMPLEMENTACAO}:
            return TipoOperacaoNf.COMPLEMENTACAO
        if evento_norm in {AUDIT_EVENTO_REENTREGA, HISTORICO_EVENTO_REENTREGA}:
            return TipoOperacaoNf.REENTREGA
        return None

    @staticmethod
    def _formatar_evento_data_hora(criado_em, carregamento: Carregamento) -> tuple[str, str, float]:
        if criado_em is not None:
            dt = criado_em
            if isinstance(dt, datetime) and dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            if isinstance(dt, datetime):
                local = dt.astimezone()
                return local.strftime("%d/%m/%Y"), local.strftime("%H:%M"), local.timestamp()
        data, hora = AnaliseOperacionalService._split_data_hora_carregamento(
            str(carregamento.data or ""),
            str(carregamento.hora or ""),
        )
        try:
            ordenacao = datetime.strptime(f"{data} {hora}", "%d/%m/%Y %H:%M").timestamp()
        except ValueError:
            ordenacao = 0.0
        return data, hora, ordenacao

    @staticmethod
    def _split_data_hora_carregamento(data_value: str, hora_value: str) -> tuple[str, str]:
        data_text = str(data_value or "").strip()
        hora_text = str(hora_value or "").strip()
        if len(data_text.split("-")) == 3:
            ano, mes, dia = data_text.split("-")
            data_fmt = f"{dia}/{mes}/{ano}"
        else:
            data_fmt = data_text or "--"
        hora_fmt = hora_text[:5] if len(hora_text) >= 5 else (hora_text or "--")
        return data_fmt, hora_fmt

    @staticmethod
    def _montar_resumo_nf(eventos: list[NfAuditoriaEvento], carregamentos_distintos: int) -> NfAuditoriaResumo:
        if not eventos:
            return NfAuditoriaResumo(
                primeira_utilizacao="--",
                ultima_utilizacao="--",
                total_impressoes=0,
                total_reimpressoes=0,
                total_complementacoes=0,
                total_reentregas=0,
                carregamentos_distintos=carregamentos_distintos,
            )

        ordenados = sorted(eventos, key=lambda item: item.ordenacao)
        primeira = ordenados[0]
        ultima = ordenados[-1]
        return NfAuditoriaResumo(
            primeira_utilizacao=f"{primeira.data} {primeira.hora}".strip(),
            ultima_utilizacao=f"{ultima.data} {ultima.hora}".strip(),
            total_impressoes=sum(
                1 for item in eventos if item.operacao in {TipoOperacaoNf.IMPRESSAO_ORIGINAL, TipoOperacaoNf.REIMPRESSAO}
            ),
            total_reimpressoes=sum(1 for item in eventos if item.operacao == TipoOperacaoNf.REIMPRESSAO),
            total_complementacoes=sum(1 for item in eventos if item.operacao == TipoOperacaoNf.COMPLEMENTACAO),
            total_reentregas=sum(1 for item in eventos if item.operacao == TipoOperacaoNf.REENTREGA),
            carregamentos_distintos=carregamentos_distintos,
        )

    @staticmethod
    def _formatar_data_hora_consulta(value: datetime) -> str:
        local = value.astimezone() if value.tzinfo else value.replace(tzinfo=timezone.utc).astimezone()
        return local.strftime("%d/%m/%Y %H:%M")

    def filtrar_nfs_novas(self, processed_df: pd.DataFrame, diagnostico: DiagnosticoCarregamento) -> pd.DataFrame:
        if processed_df.empty:
            return processed_df.iloc[0:0]
        tokens_novos = {
            item.token
            for item in diagnostico.nfs
            if item.classificacao == ClassificacaoNfLote.NOVA
        }
        if not tokens_novos:
            return processed_df.iloc[0:0]

        def row_token(row: pd.Series) -> str:
            chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
            if chave:
                return chave
            nf_norm = normalize_nf_number(row.get("NF", ""))
            return f"nf:{nf_norm}" if nf_norm else ""

        mask = processed_df.apply(row_token, axis=1).isin(tokens_novos)
        return processed_df[mask].copy()

    def _extrair_nfs_unicas(self, processed_df: pd.DataFrame) -> list[NfLoteResumo]:
        indice = self._obter_indice()
        agrupado: dict[str, NfLoteResumo] = {}
        for _, row in processed_df.iterrows():
            chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
            nf_norm = normalize_nf_number(row.get("NF", ""))
            token = chave or (f"nf:{nf_norm}" if nf_norm else "")
            if not token:
                continue
            status_nf = str(row.get("Status", "") or "")
            if token in agrupado:
                continue

            if _is_canceled_nf_status(status_nf):
                classificacao = ClassificacaoNfLote.CANCELADA
                vinculo = None
            elif not _is_authorized_nf_status(status_nf):
                classificacao = ClassificacaoNfLote.NAO_AUTORIZADA
                vinculo = None
            else:
                vinculo = indice.buscar(chave, str(row.get("NF", "") or ""))
                classificacao = ClassificacaoNfLote.EXISTENTE if vinculo else ClassificacaoNfLote.NOVA

            agrupado[token] = NfLoteResumo(
                token=token,
                nf=str(row.get("NF", "") or ""),
                chave_nfe=chave,
                status_nf=status_nf,
                classificacao=classificacao,
                vinculo=vinculo,
            )
        return list(agrupado.values())

    @staticmethod
    def _formatar_data(value: str) -> str:
        parts = str(value or "").split("-")
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return value or "--"
