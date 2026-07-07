from __future__ import annotations

from datetime import date

from sqlalchemy import select, text
from sqlalchemy.orm import Session

from carregamentos.models.simulacao_retencao import CarregamentoElegivelRef, ProblemaIntegridade, SaudePacote
from carregamentos.repository.simulacao_retencao_repository import (
    ArvoreCarregamentoRaw,
    DocumentoXmlRaw,
    SimulacaoRetencaoRepository,
)
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM
from infrastructure.unit_of_work import UnitOfWork


def _elegiveis_clause(data_corte: date):
    return CarregamentoORM.data < data_corte


class SqlSimulacaoRetencaoRepository(SimulacaoRetencaoRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    def listar_carregamentos_elegiveis(self, data_corte: date) -> list[CarregamentoElegivelRef]:
        with self._uow() as uow:
            rows = uow.session.scalars(
                select(CarregamentoORM)
                .where(_elegiveis_clause(data_corte))
                .order_by(CarregamentoORM.data, CarregamentoORM.id)
            ).all()
            return [
                CarregamentoElegivelRef(
                    id=int(row.id),
                    numero_carregamento=str(row.numero_carregamento),
                    data=row.data,
                )
                for row in rows
            ]

    def carregar_arvores_elegiveis(self, data_corte: date) -> list[ArvoreCarregamentoRaw]:
        with self._uow() as uow:
            session = uow.session
            carregamentos = session.scalars(
                select(CarregamentoORM).where(_elegiveis_clause(data_corte)).order_by(CarregamentoORM.id)
            ).all()
            if not carregamentos:
                return []

            carregamento_ids = [int(row.id) for row in carregamentos]
            itens = session.scalars(
                select(ItemCarregamentoORM)
                .where(ItemCarregamentoORM.carregamento_id.in_(carregamento_ids))
                .order_by(ItemCarregamentoORM.carregamento_id, ItemCarregamentoORM.id)
            ).all()
            documentos = session.scalars(
                select(DocumentoORM)
                .where(DocumentoORM.carregamento_id.in_(carregamento_ids))
                .order_by(DocumentoORM.carregamento_id, DocumentoORM.id)
            ).all()
            historicos = session.scalars(
                select(HistoricoOperacionalORM)
                .where(HistoricoOperacionalORM.carregamento_id.in_(carregamento_ids))
                .order_by(HistoricoOperacionalORM.carregamento_id, HistoricoOperacionalORM.id)
            ).all()
            eventos = session.scalars(
                select(EventoAuditoriaORM)
                .where(
                    EventoAuditoriaORM.entidade_tipo == "carregamento",
                    EventoAuditoriaORM.entidade_id.in_(carregamento_ids),
                )
                .order_by(EventoAuditoriaORM.entidade_id, EventoAuditoriaORM.id)
            ).all()

            itens_por_carregamento: dict[int, list[ItemCarregamentoORM]] = {cid: [] for cid in carregamento_ids}
            for item in itens:
                itens_por_carregamento[int(item.carregamento_id)].append(item)

            docs_por_carregamento: dict[int, list[DocumentoORM]] = {cid: [] for cid in carregamento_ids}
            for doc in documentos:
                docs_por_carregamento[int(doc.carregamento_id)].append(doc)

            hist_por_carregamento: dict[int, list[HistoricoOperacionalORM]] = {cid: [] for cid in carregamento_ids}
            for hist in historicos:
                hist_por_carregamento[int(hist.carregamento_id)].append(hist)

            evt_por_carregamento: dict[int, list[EventoAuditoriaORM]] = {cid: [] for cid in carregamento_ids}
            for evt in eventos:
                if evt.entidade_id is not None:
                    evt_por_carregamento[int(evt.entidade_id)].append(evt)

            nota_fiscal_ids: set[int] = set()
            chaves: set[str] = set()
            for item in itens:
                if item.nota_fiscal_id is not None:
                    nota_fiscal_ids.add(int(item.nota_fiscal_id))
                chave = str(item.chave_nfe or "").strip()
                if chave:
                    chaves.add(chave)

            item_nf_ids_por_nf: dict[int, list[int]] = {}
            if nota_fiscal_ids:
                item_nf_rows = session.scalars(
                    select(ItemNotaFiscalORM).where(ItemNotaFiscalORM.nota_fiscal_id.in_(sorted(nota_fiscal_ids)))
                ).all()
                for row in item_nf_rows:
                    item_nf_ids_por_nf.setdefault(int(row.nota_fiscal_id), []).append(int(row.id))

            arvores: list[ArvoreCarregamentoRaw] = []
            for carregamento in carregamentos:
                cid = int(carregamento.id)
                grupo_itens = itens_por_carregamento.get(cid, [])
                grupo_docs = docs_por_carregamento.get(cid, [])
                grupo_hist = hist_por_carregamento.get(cid, [])
                grupo_evt = evt_por_carregamento.get(cid, [])

                nfs_ids: list[int] = []
                chaves_nf: list[str] = []
                numeros_nf: list[str] = []
                item_nf_ids: list[int] = []
                nfs_vistas: set[str] = set()

                for item in grupo_itens:
                    if item.nota_fiscal_id is not None:
                        nf_id = int(item.nota_fiscal_id)
                        if nf_id not in nfs_ids:
                            nfs_ids.append(nf_id)
                        item_nf_ids.extend(item_nf_ids_por_nf.get(nf_id, []))
                    chave = str(item.chave_nfe or "").strip()
                    numero = str(item.numero_nf or "").strip()
                    if chave:
                        chaves_nf.append(chave)
                    if numero:
                        numeros_nf.append(numero)
                    token = chave or numero
                    if token:
                        nfs_vistas.add(token)

                arvores.append(
                    ArvoreCarregamentoRaw(
                        carregamento_id=cid,
                        numero_carregamento=str(carregamento.numero_carregamento),
                        data=carregamento.data,
                        item_ids=tuple(int(i.id) for i in grupo_itens),
                        chaves_nfe=tuple(dict.fromkeys(chaves_nf)),
                        numeros_nf=tuple(dict.fromkeys(numeros_nf)),
                        nota_fiscal_ids=tuple(nfs_ids),
                        documento_ids=tuple(int(d.id) for d in grupo_docs),
                        documento_tipos=tuple(str(d.tipo) for d in grupo_docs),
                        documento_caminhos=tuple(str(d.caminho_arquivo) for d in grupo_docs),
                        documento_nomes=tuple(str(d.nome_arquivo) for d in grupo_docs),
                        documento_hashes=tuple(str(d.hash_sha256) for d in grupo_docs),
                        historico_ids=tuple(int(h.id) for h in grupo_hist),
                        evento_ids=tuple(int(e.id) for e in grupo_evt),
                        item_nota_fiscal_ids=tuple(dict.fromkeys(item_nf_ids)),
                    )
                )
            return arvores

    def carregar_documentos_xml_por_chaves(self, chaves: list[str]) -> dict[str, DocumentoXmlRaw]:
        normalized = [str(chave or "").strip() for chave in chaves if str(chave or "").strip()]
        if not normalized:
            return {}
        with self._uow() as uow:
            rows = uow.session.scalars(
                select(DocumentoXmlORM).where(DocumentoXmlORM.chave_nfe.in_(normalized))
            ).all()
            result: dict[str, DocumentoXmlRaw] = {}
            for row in rows:
                result[row.chave_nfe] = DocumentoXmlRaw(
                    id=int(row.id),
                    chave_nfe=row.chave_nfe,
                    numero_nf=row.numero_nf,
                    caminho_arquivo=row.caminho_arquivo,
                    hash_sha256=row.hash_sha256,
                    tamanho=int(row.tamanho),
                    ativo=bool(row.ativo),
                )
            return result

    def detectar_orfaos(self, data_corte: date, carregamento_ids: list[int]) -> list[ProblemaIntegridade]:
        if not carregamento_ids:
            return []
        problemas: list[ProblemaIntegridade] = []
        with self._uow() as uow:
            session = uow.session

            hist_orfaos = session.execute(
                text(
                    """
                    SELECT ho.id, ho.carregamento_id
                    FROM historico_operacional ho
                    LEFT JOIN carregamento c ON c.id = ho.carregamento_id
                    WHERE c.id IS NULL
                    LIMIT 50
                    """
                )
            ).all()
            for row in hist_orfaos:
                problemas.append(
                    ProblemaIntegridade(
                        severidade=SaudePacote.CRITICO,
                        categoria="historico_sem_carregamento",
                        descricao=f"Historico {row.id} sem carregamento vinculado.",
                        carregamento_id=int(row.carregamento_id) if row.carregamento_id else None,
                        referencia=str(row.id),
                    )
                )

            doc_orfaos = session.execute(
                text(
                    """
                    SELECT d.id, d.carregamento_id, d.caminho_arquivo
                    FROM documento d
                    LEFT JOIN carregamento c ON c.id = d.carregamento_id
                    WHERE c.id IS NULL
                    LIMIT 50
                    """
                )
            ).all()
            for row in doc_orfaos:
                problemas.append(
                    ProblemaIntegridade(
                        severidade=SaudePacote.CRITICO,
                        categoria="pdf_sem_carregamento",
                        descricao=f"Documento PDF {row.id} sem carregamento vinculado.",
                        carregamento_id=int(row.carregamento_id) if row.carregamento_id else None,
                        referencia=str(row.caminho_arquivo or row.id),
                    )
                )

            evt_orfaos = session.execute(
                text(
                    """
                    SELECT ea.id, ea.entidade_id
                    FROM evento_auditoria ea
                    LEFT JOIN carregamento c ON c.id = ea.entidade_id
                    WHERE ea.entidade_tipo = 'carregamento'
                      AND c.id IS NULL
                    LIMIT 50
                    """
                )
            ).all()
            for row in evt_orfaos:
                problemas.append(
                    ProblemaIntegridade(
                        severidade=SaudePacote.CRITICO,
                        categoria="evento_sem_carregamento",
                        descricao=f"Evento {row.id} referencia carregamento inexistente.",
                        carregamento_id=int(row.entidade_id) if row.entidade_id else None,
                        referencia=str(row.id),
                    )
                )

            item_orfaos = session.execute(
                text(
                    """
                    SELECT ic.id, ic.carregamento_id
                    FROM item_carregamento ic
                    LEFT JOIN carregamento c ON c.id = ic.carregamento_id
                    WHERE c.id IS NULL
                    LIMIT 50
                    """
                )
            ).all()
            for row in item_orfaos:
                problemas.append(
                    ProblemaIntegridade(
                        severidade=SaudePacote.CRITICO,
                        categoria="item_sem_carregamento",
                        descricao=f"Item de carregamento {row.id} sem carregamento vinculado.",
                        carregamento_id=int(row.carregamento_id) if row.carregamento_id else None,
                        referencia=str(row.id),
                    )
                )

            xml_sem_vinculo = session.execute(
                text(
                    f"""
                    SELECT dx.id, dx.chave_nfe, dx.caminho_arquivo
                    FROM documento_xml dx
                    WHERE dx.ativo = 1
                      AND dx.chave_nfe NOT IN (
                        SELECT DISTINCT TRIM(ic.chave_nfe)
                        FROM item_carregamento ic
                        INNER JOIN carregamento c ON c.id = ic.carregamento_id
                        WHERE c.data < :data_corte
                          AND TRIM(COALESCE(ic.chave_nfe, '')) <> ''
                      )
                    LIMIT 50
                    """
                ),
                {"data_corte": data_corte},
            ).all()
            for row in xml_sem_vinculo:
                problemas.append(
                    ProblemaIntegridade(
                        severidade=SaudePacote.ATENCAO,
                        categoria="documento_xml_sem_carregamento_elegivel",
                        descricao=(
                            f"Documento XML {row.id} (chave {row.chave_nfe}) "
                            "sem vinculo com carregamento elegivel."
                        ),
                        referencia=str(row.caminho_arquivo or row.chave_nfe),
                    )
                )

            for cid in carregamento_ids:
                doc_count = session.scalar(
                    select(DocumentoORM.id).where(DocumentoORM.carregamento_id == cid).limit(1)
                )
                item_count = session.scalar(
                    select(ItemCarregamentoORM.id).where(ItemCarregamentoORM.carregamento_id == cid).limit(1)
                )
                if doc_count is not None and item_count is None:
                    problemas.append(
                        ProblemaIntegridade(
                            severidade=SaudePacote.ATENCAO,
                            categoria="carregamento_incompleto",
                            descricao="Carregamento possui PDF sem itens vinculados.",
                            carregamento_id=cid,
                        )
                    )

        return problemas

    def _uow(self) -> UnitOfWork:
        return UnitOfWork(self._session)
