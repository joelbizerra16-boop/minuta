from __future__ import annotations

from datetime import date

from sqlalchemy import and_, exists, func, select, text
from sqlalchemy.orm import Session

from carregamentos.models.retencao import RetencaoContagensArvore
from carregamentos.repository.retencao_repository import RetencaoRepository
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.persistence.engine_info import get_dialect_name
from infrastructure.persistence.sql_compat import trim_both_zeros
from infrastructure.unit_of_work import UnitOfWork


def _carregamentos_elegiveis_clause(data_corte: date):
    return CarregamentoORM.data < data_corte


def _carregamentos_elegiveis_por_data_clause(data_corte: date, data_alvo: date):
    return and_(CarregamentoORM.data < data_corte, CarregamentoORM.data == data_alvo)


class SqlRetencaoRepository(RetencaoRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    def possui_carregamentos_elegiveis(self, data_corte: date) -> bool:
        with self._uow() as uow:
            stmt = select(
                exists().where(_carregamentos_elegiveis_clause(data_corte))
            )
            return bool(uow.session.scalar(stmt))

    def possui_carregamentos_expirados(self, data_corte: date) -> bool:
        return self.possui_carregamentos_elegiveis(data_corte)

    def coletar_contagens_arvore(self, data_corte: date) -> RetencaoContagensArvore:
        with self._uow() as uow:
            return self._coletar_contagens_arvore(uow.session, data_corte, data_alvo=None)

    def obter_data_mais_antiga_elegivel(self, data_corte: date) -> date | None:
        with self._uow() as uow:
            return uow.session.scalar(
                select(func.min(CarregamentoORM.data)).where(_carregamentos_elegiveis_clause(data_corte))
            )

    def listar_carregamento_ids_por_data(self, data_corte: date, data_alvo: date) -> tuple[int, ...]:
        with self._uow() as uow:
            rows = uow.session.scalars(
                select(CarregamentoORM.id)
                .where(_carregamentos_elegiveis_por_data_clause(data_corte, data_alvo))
                .order_by(CarregamentoORM.id)
            ).all()
            return tuple(int(row) for row in rows)

    def coletar_contagens_arvore_por_data(self, data_corte: date, data_alvo: date) -> RetencaoContagensArvore:
        with self._uow() as uow:
            return self._coletar_contagens_arvore(uow.session, data_corte, data_alvo=data_alvo)

    def _coletar_contagens_arvore(
        self,
        session: Session,
        data_corte: date,
        *,
        data_alvo: date | None,
    ) -> RetencaoContagensArvore:
        if data_alvo is not None:
            elegiveis_clause = _carregamentos_elegiveis_por_data_clause(data_corte, data_alvo)
        else:
            elegiveis_clause = _carregamentos_elegiveis_clause(data_corte)

        elegiveis = select(CarregamentoORM.id).where(elegiveis_clause)

        carregamentos_elegiveis = int(
            session.scalar(select(func.count()).select_from(CarregamentoORM).where(elegiveis_clause)) or 0
        )
        if carregamentos_elegiveis == 0:
            return RetencaoContagensArvore(
                carregamentos_elegiveis=0,
                notas_fiscais=0,
                itens_carregamento=0,
                itens_nota_fiscal=0,
                documentos_xml=0,
                documentos_pdf=0,
                historicos=0,
                eventos=0,
                caminhos_pdf=(),
                espaco_xmls_bytes=0,
            )

        itens_carregamento = int(
            session.scalar(
                select(func.count())
                .select_from(ItemCarregamentoORM)
                .where(ItemCarregamentoORM.carregamento_id.in_(elegiveis))
            )
            or 0
        )

        documentos_pdf = int(
            session.scalar(
                select(func.count()).select_from(DocumentoORM).where(DocumentoORM.carregamento_id.in_(elegiveis))
            )
            or 0
        )

        historicos = int(
            session.scalar(
                select(func.count())
                .select_from(HistoricoOperacionalORM)
                .where(HistoricoOperacionalORM.carregamento_id.in_(elegiveis))
            )
            or 0
        )

        eventos = int(
            session.scalar(
                select(func.count())
                .select_from(EventoAuditoriaORM)
                .where(
                    EventoAuditoriaORM.entidade_tipo == "carregamento",
                    EventoAuditoriaORM.entidade_id.in_(elegiveis),
                )
            )
            or 0
        )

        notas_fiscais = self._contar_notas_fiscais_distintas(session, data_corte, data_alvo)
        itens_nota_fiscal = self._contar_itens_nota_fiscal(session, data_corte, data_alvo)
        documentos_xml, espaco_xmls_bytes = self._contar_documentos_xml(session, data_corte, data_alvo)
        caminhos_pdf = self._listar_caminhos_pdf(session, data_corte, data_alvo)

        return RetencaoContagensArvore(
            carregamentos_elegiveis=carregamentos_elegiveis,
            notas_fiscais=notas_fiscais,
            itens_carregamento=itens_carregamento,
            itens_nota_fiscal=itens_nota_fiscal,
            documentos_xml=documentos_xml,
            documentos_pdf=documentos_pdf,
            historicos=historicos,
            eventos=eventos,
            caminhos_pdf=caminhos_pdf,
            espaco_xmls_bytes=espaco_xmls_bytes,
        )

    @staticmethod
    def _contar_notas_fiscais_distintas(session: Session, data_corte: date, data_alvo: date | None = None) -> int:
        filtro_data = "AND c.data = :data_alvo" if data_alvo is not None else ""
        stmt = text(
            f"""
            SELECT COUNT(*) FROM (
                SELECT DISTINCT
                    CASE
                        WHEN TRIM(COALESCE(ic.chave_nfe, '')) <> '' THEN TRIM(ic.chave_nfe)
                        ELSE TRIM(COALESCE(ic.numero_nf, ''))
                    END AS nf_identidade
                FROM item_carregamento ic
                INNER JOIN carregamento c ON c.id = ic.carregamento_id
                WHERE c.data < :data_corte
                  {filtro_data}
                  AND (
                    TRIM(COALESCE(ic.chave_nfe, '')) <> ''
                    OR TRIM(COALESCE(ic.numero_nf, '')) <> ''
                  )
            ) AS nfs_distintas
            """
        )
        params = {"data_corte": data_corte}
        if data_alvo is not None:
            params["data_alvo"] = data_alvo
        return int(session.scalar(stmt, params) or 0)

    @staticmethod
    def _contar_itens_nota_fiscal(session: Session, data_corte: date, data_alvo: date | None = None) -> int:
        filtro_data = "AND c.data = :data_alvo" if data_alvo is not None else ""
        dialect = get_dialect_name(session)
        trim_ic = trim_both_zeros("ic.numero_nf", dialect=dialect)
        trim_nf = trim_both_zeros("nf.numero_nf", dialect=dialect)
        nf_match = f"{trim_ic} = {trim_nf}"
        stmt = text(
            f"""
            SELECT COUNT(*)
            FROM item_nota_fiscal inf
            WHERE inf.nota_fiscal_id IN (
                SELECT DISTINCT ic.nota_fiscal_id
                FROM item_carregamento ic
                INNER JOIN carregamento c ON c.id = ic.carregamento_id
                WHERE c.data < :data_corte
                  {filtro_data}
                  AND ic.nota_fiscal_id IS NOT NULL
            )
            OR inf.nota_fiscal_id IN (
                SELECT nf.id
                FROM nota_fiscal nf
                INNER JOIN item_carregamento ic ON (
                    (
                        TRIM(COALESCE(ic.chave_nfe, '')) <> ''
                        AND nf.chave_nfe = TRIM(ic.chave_nfe)
                    )
                    OR (
                        TRIM(COALESCE(ic.numero_nf, '')) <> ''
                        AND {nf_match}
                    )
                )
                INNER JOIN carregamento c ON c.id = ic.carregamento_id
                WHERE c.data < :data_corte
                  {filtro_data}
                  AND ic.nota_fiscal_id IS NULL
            )
            """
        )
        params = {"data_corte": data_corte}
        if data_alvo is not None:
            params["data_alvo"] = data_alvo
        return int(session.scalar(stmt, params) or 0)

    @staticmethod
    def _contar_documentos_xml(
        session: Session,
        data_corte: date,
        data_alvo: date | None = None,
    ) -> tuple[int, int]:
        if data_alvo is not None:
            elegiveis_clause = _carregamentos_elegiveis_por_data_clause(data_corte, data_alvo)
        else:
            elegiveis_clause = _carregamentos_elegiveis_clause(data_corte)
        chaves_stmt = (
            select(ItemCarregamentoORM.chave_nfe)
            .join(CarregamentoORM, CarregamentoORM.id == ItemCarregamentoORM.carregamento_id)
            .where(
                elegiveis_clause,
                ItemCarregamentoORM.chave_nfe.is_not(None),
                ItemCarregamentoORM.chave_nfe != "",
            )
            .distinct()
        )
        chaves = [str(value).strip() for value in session.scalars(chaves_stmt).all() if str(value or "").strip()]
        if not chaves:
            return 0, 0

        stmt = (
            select(
                func.count(DocumentoXmlORM.id),
                func.coalesce(func.sum(DocumentoXmlORM.tamanho), 0),
            )
            .where(
                DocumentoXmlORM.ativo.is_(True),
                DocumentoXmlORM.chave_nfe.in_(chaves),
            )
        )
        row = session.execute(stmt).one()
        return int(row[0] or 0), int(row[1] or 0)

    @staticmethod
    def _listar_caminhos_pdf(
        session: Session,
        data_corte: date,
        data_alvo: date | None = None,
    ) -> tuple[str, ...]:
        if data_alvo is not None:
            elegiveis_clause = _carregamentos_elegiveis_por_data_clause(data_corte, data_alvo)
        else:
            elegiveis_clause = _carregamentos_elegiveis_clause(data_corte)
        stmt = (
            select(DocumentoORM.caminho_arquivo)
            .join(CarregamentoORM, CarregamentoORM.id == DocumentoORM.carregamento_id)
            .where(elegiveis_clause)
            .order_by(DocumentoORM.id)
        )
        return tuple(str(path).strip() for path in session.scalars(stmt).all() if str(path or "").strip())

    def _uow(self) -> UnitOfWork:
        return UnitOfWork(self._session)
