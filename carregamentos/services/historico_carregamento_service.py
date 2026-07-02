from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime, timezone

from sqlalchemy import or_, select

from carregamentos.models.carregamento import normalize_chave_nfe, normalize_nf_number
from carregamentos.models.historico_carregamento_painel import (
    HistoricoCarregamentoEstatisticas,
    HistoricoCarregamentoPainel,
    HistoricoComplementacaoLinha,
    HistoricoImpressaoLinha,
    HistoricoNfLinha,
    HistoricoReentregaLinha,
)
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from infrastructure.models.carregamento import CarregamentoORM
from infrastructure.models.constants import (
    AUDIT_CATEGORIA_CARREGAMENTO,
    AUDIT_EVENTO_COMPLEMENTACAO,
    AUDIT_EVENTO_ENTREGA_BALCAO,
    AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
    AUDIT_EVENTO_REENTREGA,
    AUDIT_EVENTO_REIMPRESSAO,
    HISTORICO_EVENTO_COMPLEMENTACAO,
    HISTORICO_EVENTO_REENTREGA,
)
from infrastructure.models.nota_fiscal import NotaFiscalORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.repositories.sql.historico_repository import SqlHistoricoRepository
from infrastructure.unit_of_work import UnitOfWork


class HistoricoCarregamentoService:
    """Monta o painel de auditoria operacional de um carregamento existente."""

    def __init__(self, repository: SqlCarregamentoRepository) -> None:
        self._repository = repository

    def montar_painel_auditoria(
        self,
        carregamento_id: int,
        *,
        excel_contexto: str = "",
        data_analise: datetime | None = None,
    ) -> HistoricoCarregamentoPainel | None:
        if carregamento_id <= 0:
            return None

        with UnitOfWork() as uow:
            session = uow.session
            row = session.scalars(
                self._repository._base_stmt().where(CarregamentoORM.id == carregamento_id)
            ).first()
            if row is None:
                return None

            carregamento = self._repository._to_domain(session, row)
            historico_repo = SqlHistoricoRepository(session)
            audit_repo = SqlEventoAuditoriaRepository(session)
            historicos = historico_repo.list_by_carregamento(carregamento_id)
            auditorias = audit_repo.list_by_entidade("carregamento", carregamento_id)

            usuario_ids = {
                int(item.usuario_id)
                for item in historicos
                if item.usuario_id is not None
            }
            usuario_ids.update(
                int(item.usuario_id)
                for item in auditorias
                if item.usuario_id is not None
            )
            usuarios_por_id = self._carregar_usuarios(session, usuario_ids)
            notas_por_chave, notas_por_numero = self._carregar_notas_fiscais(session, carregamento.itens)

        agora = data_analise or datetime.now()
        excel_nome = str(excel_contexto or "").strip() or "Nao registrado no historico"
        primeira_data, primeira_hora = self._split_data_hora(carregamento.data, carregamento.hora)
        ultima_data, ultima_hora = self._split_datetime_iso(carregamento.ultima_impressao_em)

        reimpressoes = max(int(carregamento.quantidade_impressoes or 0) - 1, 0)
        complementacoes = self._montar_complementacoes(historicos, auditorias, usuarios_por_id)
        reentregas = self._montar_reentregas(historicos, usuarios_por_id, carregamento.status)
        impressoes = self._montar_impressoes(auditorias, usuarios_por_id)

        nfs = self._montar_linhas_nf(
            carregamento=carregamento,
            notas_por_chave=notas_por_chave,
            notas_por_numero=notas_por_numero,
            excel_contexto=excel_nome,
            primeira_data=primeira_data,
            primeira_hora=primeira_hora,
            ultima_data=ultima_data,
            ultima_hora=ultima_hora,
            reimpressoes=reimpressoes,
        )

        valor_total = sum(item.valor_nf for item in nfs)
        estatisticas = HistoricoCarregamentoEstatisticas(
            numero_carregamento=carregamento.numero_carregamento,
            total_nfs=len(nfs),
            peso_total_kg=float(carregamento.peso_total or 0),
            valor_total=valor_total,
            primeira_impressao_data=primeira_data,
            primeira_impressao_hora=primeira_hora,
            ultima_impressao_data=ultima_data or primeira_data,
            ultima_impressao_hora=ultima_hora or primeira_hora,
            quantidade_reimpressoes=reimpressoes,
            quantidade_complementacoes=len(complementacoes),
            quantidade_reentregas=len(reentregas),
        )

        return HistoricoCarregamentoPainel(
            carregamento_id=carregamento_id,
            numero_carregamento=carregamento.numero_carregamento,
            excel_contexto=excel_nome,
            data_analise=self._formatar_data_hora_br(agora),
            estatisticas=estatisticas,
            nfs=nfs,
            impressoes=impressoes,
            complementacoes=complementacoes,
            reentregas=reentregas,
        )

    @staticmethod
    def _carregar_usuarios(session, usuario_ids: set[int]) -> dict[int, str]:
        if not usuario_ids:
            return {}
        rows = session.scalars(select(UsuarioORM).where(UsuarioORM.id.in_(usuario_ids))).all()
        return {int(row.id): str(row.usuario or "--") for row in rows}

    @staticmethod
    def _carregar_notas_fiscais(session, itens) -> tuple[dict[str, NotaFiscalORM], dict[str, NotaFiscalORM]]:
        chaves = {normalize_chave_nfe(item.chave_nfe) for item in itens if normalize_chave_nfe(item.chave_nfe)}
        numeros = {normalize_nf_number(item.nf) for item in itens if normalize_nf_number(item.nf)}
        if not chaves and not numeros:
            return {}, {}

        filtros = []
        if chaves:
            filtros.append(NotaFiscalORM.chave_nfe.in_(sorted(chaves)))
        if numeros:
            filtros.append(NotaFiscalORM.numero_nf.in_(sorted(numeros)))
        rows = session.scalars(select(NotaFiscalORM).where(or_(*filtros))).all()

        por_chave: dict[str, NotaFiscalORM] = {}
        por_numero: dict[str, NotaFiscalORM] = {}
        for row in rows:
            chave = normalize_chave_nfe(row.chave_nfe)
            numero = normalize_nf_number(row.numero_nf)
            if chave:
                por_chave[chave] = row
            if numero:
                por_numero[numero] = row
        return por_chave, por_numero

    def _montar_linhas_nf(
        self,
        *,
        carregamento,
        notas_por_chave: dict[str, NotaFiscalORM],
        notas_por_numero: dict[str, NotaFiscalORM],
        excel_contexto: str,
        primeira_data: str,
        primeira_hora: str,
        ultima_data: str,
        ultima_hora: str,
        reimpressoes: int,
    ) -> list[HistoricoNfLinha]:
        agrupado: dict[str, dict[str, object]] = defaultdict(
            lambda: {
                "peso": 0.0,
                "cliente": "",
                "rota": "",
                "chave": "",
                "status_nf": "",
            }
        )

        for item in carregamento.itens:
            nf_norm = normalize_nf_number(item.nf) or str(item.nf or "").strip()
            if not nf_norm:
                continue
            bucket = agrupado[nf_norm]
            bucket["peso"] = float(bucket["peso"]) + float(item.peso or 0)
            if not bucket["cliente"]:
                bucket["cliente"] = str(item.destinatario or "").strip()
            if not bucket["rota"]:
                bucket["rota"] = str(item.rota or "").strip()
            if not bucket["chave"]:
                bucket["chave"] = normalize_chave_nfe(item.chave_nfe)
            if not bucket["status_nf"]:
                bucket["status_nf"] = str(item.status_nf or "").strip()

        linhas: list[HistoricoNfLinha] = []
        for nf_norm, bucket in sorted(agrupado.items(), key=lambda item: item[0]):
            chave = str(bucket["chave"] or "")
            nota = notas_por_chave.get(chave) if chave else None
            if nota is None:
                nota = notas_por_numero.get(nf_norm)

            cidade = str(nota.municipio or bucket["rota"] or "--") if nota else str(bucket["rota"] or "--")
            uf = str(nota.uf or "--") if nota else "--"
            valor_nf = float(nota.valor_total or 0) if nota else 0.0
            xml_status = str(nota.status_nf or bucket["status_nf"] or "--") if nota else str(bucket["status_nf"] or "--")
            xml_arquivo = str(nota.arquivo_origem or "--") if nota else "--"

            linhas.append(
                HistoricoNfLinha(
                    nf=nf_norm,
                    cliente=str(bucket["cliente"] or "--"),
                    cidade=cidade,
                    uf=uf,
                    peso_kg=float(bucket["peso"]),
                    valor_nf=valor_nf,
                    primeira_utilizacao_data=primeira_data,
                    primeira_utilizacao_hora=primeira_hora,
                    usuario_carregamento=str(carregamento.usuario or "--"),
                    numero_carregamento=str(carregamento.numero_carregamento or "--"),
                    quantidade_reimpressoes=reimpressoes,
                    ultima_reimpressao_data=ultima_data or "--",
                    ultima_reimpressao_hora=ultima_hora or "--",
                    ultimo_usuario_impressao=str(carregamento.ultima_impressao_usuario or carregamento.usuario or "--"),
                    status_atual=str(carregamento.status or "--"),
                    origem="Excel",
                    excel_utilizado=excel_contexto,
                    xml_status=xml_status,
                    xml_arquivo=xml_arquivo,
                )
            )
        return linhas

    def _montar_impressoes(self, auditorias, usuarios_por_id: dict[int, str]) -> list[HistoricoImpressaoLinha]:
        eventos_impressao = {
            AUDIT_EVENTO_PRIMEIRA_IMPRESSAO: "Primeira impressao",
            AUDIT_EVENTO_REIMPRESSAO: "Reimpressao",
            AUDIT_EVENTO_ENTREGA_BALCAO: "Entrega no balcao",
        }
        linhas: list[HistoricoImpressaoLinha] = []
        for evento in auditorias:
            if evento.categoria != AUDIT_CATEGORIA_CARREGAMENTO:
                continue
            tipo = eventos_impressao.get(str(evento.evento or ""))
            if not tipo:
                continue
            data, hora = self._split_datetime_record(evento.criado_em)
            usuario = usuarios_por_id.get(int(evento.usuario_id or 0), "--")
            linhas.append(
                HistoricoImpressaoLinha(
                    data=data,
                    hora=hora,
                    usuario=usuario,
                    tipo=tipo,
                    resultado="OK",
                )
            )
        linhas.sort(key=lambda item: (item.data, item.hora))
        return linhas

    def _montar_complementacoes(
        self,
        historicos,
        auditorias,
        usuarios_por_id: dict[int, str],
    ) -> list[HistoricoComplementacaoLinha]:
        audit_por_data: dict[str, int] = {}
        for evento in auditorias:
            if str(evento.evento or "") != AUDIT_EVENTO_COMPLEMENTACAO:
                continue
            quantidade = self._extrair_itens_adicionados(evento.metadados_json)
            chave = self._formatar_data_hora_br(evento.criado_em)
            audit_por_data[chave] = max(audit_por_data.get(chave, 0), quantidade)

        linhas: list[HistoricoComplementacaoLinha] = []
        for item in historicos:
            if str(item.evento or "") != HISTORICO_EVENTO_COMPLEMENTACAO:
                continue
            data, hora = self._split_datetime_record(item.criado_em)
            chave = f"{data} {hora}"
            linhas.append(
                HistoricoComplementacaoLinha(
                    data=data,
                    hora=hora,
                    usuario=usuarios_por_id.get(int(item.usuario_id or 0), "--"),
                    nfs_adicionadas=audit_por_data.get(chave, 0),
                    observacao=str(item.descricao or "Complementacao do carregamento."),
                )
            )
        return linhas

    def _montar_reentregas(
        self,
        historicos,
        usuarios_por_id: dict[int, str],
        status: str,
    ) -> list[HistoricoReentregaLinha]:
        linhas: list[HistoricoReentregaLinha] = []
        for item in historicos:
            if str(item.evento or "") != HISTORICO_EVENTO_REENTREGA:
                continue
            data, hora = self._split_datetime_record(item.criado_em)
            linhas.append(
                HistoricoReentregaLinha(
                    data=data,
                    hora=hora,
                    usuario=usuarios_por_id.get(int(item.usuario_id or 0), "--"),
                    motivo=str(item.descricao or "Reentrega registrada."),
                    status=str(status or "--"),
                )
            )
        return linhas

    @staticmethod
    def _extrair_itens_adicionados(metadados_json: str | None) -> int:
        if not metadados_json:
            return 0
        try:
            payload = json.loads(metadados_json)
        except json.JSONDecodeError:
            return 0
        if not isinstance(payload, dict):
            return 0
        return int(payload.get("itens_adicionados", 0) or 0)

    @staticmethod
    def _split_data_hora(data_value: str, hora_value: str) -> tuple[str, str]:
        data_text = str(data_value or "").strip()
        hora_text = str(hora_value or "").strip()
        if len(data_text.split("-")) == 3:
            ano, mes, dia = data_text.split("-")
            data_fmt = f"{dia}/{mes}/{ano}"
        else:
            data_fmt = data_text or "--"
        hora_fmt = hora_text[:5] if len(hora_text) >= 5 else (hora_text or "--")
        return data_fmt, hora_fmt

    @classmethod
    def _split_datetime_iso(cls, value: str | None) -> tuple[str, str]:
        if not value:
            return "", ""
        parsed = cls._coerce_datetime(value)
        if parsed is None:
            return "", ""
        local = parsed.astimezone()
        return local.strftime("%d/%m/%Y"), local.strftime("%H:%M")

    @classmethod
    def _split_datetime_record(cls, value: datetime | None) -> tuple[str, str]:
        if value is None:
            return "--", "--"
        parsed = cls._coerce_datetime(value)
        if parsed is None:
            return "--", "--"
        local = parsed.astimezone()
        return local.strftime("%d/%m/%Y"), local.strftime("%H:%M")

    @staticmethod
    def _coerce_datetime(value: object) -> datetime | None:
        if isinstance(value, datetime):
            dt = value
        else:
            text = str(value or "").strip()
            if not text:
                return None
            try:
                dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
            except ValueError:
                return None
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt

    @classmethod
    def _formatar_data_hora_br(cls, value: datetime) -> str:
        local = cls._coerce_datetime(value)
        if local is None:
            return "--"
        local = local.astimezone()
        return local.strftime("%d/%m/%Y %H:%M")
