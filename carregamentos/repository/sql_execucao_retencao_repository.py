from __future__ import annotations

from datetime import date

from sqlalchemy import delete, exists, or_, select
from sqlalchemy.orm import Session

from carregamentos.repository.execucao_retencao_repository import ExecucaoRetencaoRepository, RecursosCompartilhadosRemovidos
from carregamentos.repository.simulacao_retencao_repository import ArvoreCarregamentoRaw
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM


def _elegiveis_clause(data_corte: date):
    return CarregamentoORM.data < data_corte


class SqlExecucaoRetencaoRepository(ExecucaoRetencaoRepository):
    def carregar_arvores_por_ids(self, session: Session, carregamento_ids: list[int]) -> list[ArvoreCarregamentoRaw]:
        if not carregamento_ids:
            return []

        ids = sorted({int(value) for value in carregamento_ids})
        carregamentos = session.scalars(
            select(CarregamentoORM).where(CarregamentoORM.id.in_(ids)).order_by(CarregamentoORM.id)
        ).all()
        if len(carregamentos) != len(ids):
            encontrados = {int(row.id) for row in carregamentos}
            faltantes = [cid for cid in ids if cid not in encontrados]
            raise ValueError(f"Carregamento(s) nao encontrado(s): {faltantes}")

        itens = session.scalars(
            select(ItemCarregamentoORM)
            .where(ItemCarregamentoORM.carregamento_id.in_(ids))
            .order_by(ItemCarregamentoORM.carregamento_id, ItemCarregamentoORM.id)
        ).all()
        documentos = session.scalars(
            select(DocumentoORM)
            .where(DocumentoORM.carregamento_id.in_(ids))
            .order_by(DocumentoORM.carregamento_id, DocumentoORM.id)
        ).all()
        historicos = session.scalars(
            select(HistoricoOperacionalORM)
            .where(HistoricoOperacionalORM.carregamento_id.in_(ids))
            .order_by(HistoricoOperacionalORM.carregamento_id, HistoricoOperacionalORM.id)
        ).all()
        eventos = session.scalars(
            select(EventoAuditoriaORM)
            .where(
                EventoAuditoriaORM.entidade_tipo == "carregamento",
                EventoAuditoriaORM.entidade_id.in_(ids),
            )
            .order_by(EventoAuditoriaORM.entidade_id, EventoAuditoriaORM.id)
        ).all()

        itens_por_carregamento: dict[int, list[ItemCarregamentoORM]] = {cid: [] for cid in ids}
        for item in itens:
            itens_por_carregamento[int(item.carregamento_id)].append(item)

        docs_por_carregamento: dict[int, list[DocumentoORM]] = {cid: [] for cid in ids}
        for doc in documentos:
            docs_por_carregamento[int(doc.carregamento_id)].append(doc)

        hist_por_carregamento: dict[int, list[HistoricoOperacionalORM]] = {cid: [] for cid in ids}
        for hist in historicos:
            hist_por_carregamento[int(hist.carregamento_id)].append(hist)

        evt_por_carregamento: dict[int, list[EventoAuditoriaORM]] = {cid: [] for cid in ids}
        for evt in eventos:
            if evt.entidade_id is not None:
                evt_por_carregamento[int(evt.entidade_id)].append(evt)

        nota_fiscal_ids: set[int] = set()
        for item in itens:
            if item.nota_fiscal_id is not None:
                nota_fiscal_ids.add(int(item.nota_fiscal_id))

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

    def validar_carregamento_elegivel(self, session: Session, carregamento_id: int, data_corte: date) -> bool:
        row = session.get(CarregamentoORM, int(carregamento_id))
        if row is None:
            return False
        return row.data < data_corte

    def excluir_arvore_carregamento(self, session: Session, arvore: ArvoreCarregamentoRaw) -> None:
        # Ordem obrigatoria: Eventos -> Historicos -> PDF -> Itens -> Carregamento
        if arvore.evento_ids:
            session.execute(delete(EventoAuditoriaORM).where(EventoAuditoriaORM.id.in_(arvore.evento_ids)))
        if arvore.historico_ids:
            session.execute(
                delete(HistoricoOperacionalORM).where(HistoricoOperacionalORM.id.in_(arvore.historico_ids))
            )
        if arvore.documento_ids:
            session.execute(delete(DocumentoORM).where(DocumentoORM.id.in_(arvore.documento_ids)))
        if arvore.item_ids:
            session.execute(delete(ItemCarregamentoORM).where(ItemCarregamentoORM.id.in_(arvore.item_ids)))
        session.execute(delete(CarregamentoORM).where(CarregamentoORM.id == arvore.carregamento_id))

    def excluir_recursos_compartilhados_orfos(
        self,
        session: Session,
        candidatos_nf_ids: set[int],
        candidatos_chaves_xml: set[str],
    ) -> RecursosCompartilhadosRemovidos:
        # Ordem: Itens NF -> Nota Fiscal -> Documento XML
        nfs_removidas: list[int] = []
        for nf_id in sorted(candidatos_nf_ids):
            if self._nota_fiscal_ainda_referenciada(session, nf_id):
                continue
            session.execute(delete(ItemNotaFiscalORM).where(ItemNotaFiscalORM.nota_fiscal_id == nf_id))
            deleted = session.execute(delete(NotaFiscalORM).where(NotaFiscalORM.id == nf_id))
            if deleted.rowcount:
                nfs_removidas.append(nf_id)

        chaves_removidas: list[str] = []
        caminhos_xml: list[str] = []
        for chave in sorted({str(value).strip() for value in candidatos_chaves_xml if str(value).strip()}):
            if self._chave_ainda_referenciada(session, chave):
                continue
            rows = session.scalars(
                select(DocumentoXmlORM).where(
                    DocumentoXmlORM.chave_nfe == chave,
                    DocumentoXmlORM.ativo.is_(True),
                )
            ).all()
            for row in rows:
                caminho = str(row.caminho_arquivo or "").strip()
                if caminho:
                    caminhos_xml.append(caminho)
            session.execute(
                delete(DocumentoXmlORM).where(
                    DocumentoXmlORM.chave_nfe == chave,
                    DocumentoXmlORM.ativo.is_(True),
                )
            )
            chaves_removidas.append(chave)

        return RecursosCompartilhadosRemovidos(
            nota_fiscal_ids=tuple(nfs_removidas),
            chaves_xml=tuple(chaves_removidas),
            caminhos_xml=tuple(caminhos_xml),
        )

    @staticmethod
    def _nota_fiscal_ainda_referenciada(session: Session, nf_id: int) -> bool:
        if session.scalar(select(exists().where(ItemCarregamentoORM.nota_fiscal_id == nf_id))):
            return True
        nf = session.get(NotaFiscalORM, nf_id)
        if nf is None:
            return False
        chave = str(nf.chave_nfe or "").strip()
        numero = str(nf.numero_nf or "").strip()
        conditions = []
        if chave:
            conditions.append(
                (ItemCarregamentoORM.chave_nfe == chave) & ItemCarregamentoORM.nota_fiscal_id.is_(None)
            )
        if numero:
            conditions.append(
                (ItemCarregamentoORM.numero_nf == numero) & ItemCarregamentoORM.nota_fiscal_id.is_(None)
            )
        if not conditions:
            return False
        return bool(session.scalar(select(exists().where(or_(*conditions)))))

    @staticmethod
    def _chave_ainda_referenciada(session: Session, chave: str) -> bool:
        chave = str(chave or "").strip()
        if not chave:
            return False
        return bool(session.scalar(select(exists().where(ItemCarregamentoORM.chave_nfe == chave))))
