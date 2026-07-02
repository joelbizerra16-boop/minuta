from __future__ import annotations

from datetime import date, datetime, time, timezone
from decimal import Decimal

from carregamentos.models.carregamento import Carregamento, CarregamentoItem
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.constants import DOC_TIPO_MINUTA, DOC_TIPO_ROMANEIO
from infrastructure.models.documento import DocumentoORM


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except ValueError:
        return datetime.now(timezone.utc).date()


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(str(value), "%H:%M:%S").time()
    except ValueError:
        return datetime.now(timezone.utc).time()


def domain_to_orm(
    carregamento: Carregamento,
    row: CarregamentoORM | None = None,
    *,
    usuario_id: int,
) -> CarregamentoORM:
    target = row or CarregamentoORM()
    if carregamento.id > 0:
        target.id = carregamento.id
    target.numero_carregamento = carregamento.numero_carregamento
    target.usuario_id = usuario_id
    target.data = _parse_date(carregamento.data)
    target.hora = _parse_time(carregamento.hora)
    target.motorista = carregamento.motorista
    target.placa = carregamento.placa
    target.filial = carregamento.filial
    target.data_saida = carregamento.data_saida
    target.modalidade = carregamento.modalidade
    target.status = carregamento.status
    target.reentrega = bool(carregamento.reentrega)
    target.quantidade_nf = int(carregamento.quantidade_nf)
    target.quantidade_itens = int(carregamento.quantidade_itens)
    target.peso_total = Decimal(str(carregamento.peso_total))
    return target


def item_domain_to_orm(item: CarregamentoItem, carregamento_id: int, sequencia: int) -> ItemCarregamentoORM:
    return ItemCarregamentoORM(
        carregamento_id=carregamento_id,
        numero_nf=item.nf,
        codigo_produto=item.cprod,
        descricao=item.descricao,
        quantidade=Decimal(str(item.quantidade)),
        unidade=item.unidade,
        peso=Decimal(str(item.peso)),
        destinatario=item.destinatario,
        rota=item.rota,
        chave_nfe=item.chave_nfe or None,
        status_nf=item.status_nf or None,
        sequencia=sequencia,
    )


def orm_to_domain(
    row: CarregamentoORM,
    itens: list[ItemCarregamentoORM],
    documentos: list[DocumentoORM],
    *,
    usuario_login: str,
    ultima_impressao_usuario: str | None = None,
) -> Carregamento:
    minuta_path = None
    romaneio_path = None
    for documento in documentos:
        if documento.tipo == DOC_TIPO_MINUTA:
            minuta_path = documento.caminho_arquivo
        elif documento.tipo == DOC_TIPO_ROMANEIO:
            romaneio_path = documento.caminho_arquivo
    return Carregamento(
        id=int(row.id),
        numero_carregamento=row.numero_carregamento,
        data=row.data.isoformat(),
        hora=row.hora.strftime("%H:%M:%S"),
        usuario=usuario_login,
        usuario_id=int(row.usuario_id),
        motorista=row.motorista or "--",
        placa=row.placa or "--",
        filial=row.filial or "",
        data_saida=row.data_saida or "--",
        quantidade_nf=int(row.quantidade_nf),
        quantidade_itens=int(row.quantidade_itens),
        peso_total=float(row.peso_total),
        status=row.status,
        modalidade=row.modalidade,
        reentrega=bool(row.reentrega),
        minuta_pdf_path=minuta_path,
        romaneio_pdf_path=romaneio_path,
        itens=[
            CarregamentoItem(
                nf=item.numero_nf,
                cprod=item.codigo_produto or "",
                descricao=item.descricao or "",
                quantidade=float(item.quantidade or 0),
                unidade=item.unidade or "",
                peso=float(item.peso or 0),
                destinatario=item.destinatario or "",
                rota=item.rota or "",
                chave_nfe=item.chave_nfe or "",
                status_nf=item.status_nf or "",
            )
            for item in itens
        ],
        criado_em=row.criado_em.isoformat() if row.criado_em else "",
        quantidade_impressoes=int(row.quantidade_impressoes or 0),
        ultima_impressao_em=row.ultima_impressao_em.isoformat() if row.ultima_impressao_em else None,
        ultima_impressao_usuario=ultima_impressao_usuario,
    )
