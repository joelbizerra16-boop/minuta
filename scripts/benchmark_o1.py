"""Benchmark O1 — mede queries e tempos dos hot paths apos otimizacao P0/P1."""

from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

from sqlalchemy import event

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _ms(start: float) -> float:
    return (time.perf_counter() - start) * 1000.0


def main() -> None:
    import pandas as pd
    from sqlalchemy import select

    from auth.bootstrap import configure_auth_storage
    from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
    from carregamentos.models.operacional import DecisaoOperacional
    from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
    from carregamentos.services.nf_validation import NfHistoricoValidator
    from core.settings import get_settings
    from infrastructure.database import configure_database, get_engine, reset_database_state
    from infrastructure.models.carregamento import ItemCarregamentoORM
    from infrastructure.schema import ensure_full_schema
    from infrastructure.storage.xml_storage import SqlXmlRecordRepository
    from infrastructure.unit_of_work import UnitOfWork

    tmp = Path(tempfile.mkdtemp(prefix="bench_o1_"))
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(tmp)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp / 'o1.db').as_posix()}"
    get_settings.cache_clear()
    reset_database_state()
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=tmp,
        pdf_storage_dir=tmp / "docs",
        xml_storage_dir=tmp / "xml",
    )
    ensure_full_schema()
    configure_auth_storage(tmp)
    configure_carregamentos_storage(tmp)

    engine = get_engine()
    sql_count = {"n": 0}

    def _count_sql(conn, cursor, statement, parameters, context, executemany):
        sql_count["n"] += 1

    event.listen(engine, "before_cursor_execute", _count_sql)

    xml_repo = SqlXmlRecordRepository()
    # Seed: 200 NFs no banco, lote consulta 5
    seed = []
    for index in range(200):
        nf = f"{index + 1:06d}"
        seed.append(
            {
                "NF": nf,
                "ChaveNFe": f"{index + 1:044d}",
                "Destinatario": "Cliente",
                "StatusNF": "Autorizado",
                "Items": [{"cProd": "P1", "Descricao": "X", "Qtd": 1, "Unidade": "UN", "Peso": 1}],
                "DataReferenciaISO": "2026-01-01T00:00:00+00:00",
            }
        )
    xml_repo.upsert_records(seed)

    lote = seed[:5]
    sql_count["n"] = 0
    start = time.perf_counter()
    by_id = xml_repo.list_records_by_identities({str(r["ChaveNFe"]) for r in lote})
    t_xml_delta = _ms(start)
    q_xml_delta = sql_count["n"]

    sql_count["n"] = 0
    start = time.perf_counter()
    all_recs = xml_repo.list_all_records()
    t_xml_all = _ms(start)
    q_xml_all = sql_count["n"]

    # Seed carregamentos
    fechamento = get_fechamento_service()
    analise = get_analise_operacional_service()

    def _row(nf: str, chave: str, cprod: str = "P001") -> dict:
        return {
            "NF": nf,
            "cProd": cprod,
            "Descricao": "Produto",
            "Qtd": 1,
            "Unidade": "UN",
            "Peso": 10.0,
            "Destinatario": "Cliente",
            "ROTA": "R1",
            "ChaveNFe": chave,
            "Status": "Autorizado o uso da NF-e",
        }

    summary = {
        "numero_carga": "BENCH",
        "motorista": "M",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "10/07/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 10.0,
    }

    for index in range(30):
        nf = f"{9000 + index}"
        chave = f"{9000 + index:044d}"
        df = pd.DataFrame([_row(nf, chave, f"C{index}")])
        fechamento.executar_fechamento_veiculo(
            summary=summary,
            processed_df=df,
            current_user=None,
            gerar_minuta=True,
            gerar_romaneio=False,
            diagnostico=analise.analisar_lote_processado(df),
            decisao=DecisaoOperacional.NOVO,
        )

    probe_df = pd.DataFrame([_row("9000", f"{9000:044d}", "C0"), _row("9999", f"{9999:044d}", "NEW")])
    analise.invalidar_cache()

    sql_count["n"] = 0
    start = time.perf_counter()
    analise.analisar_lote_processado(probe_df)
    t_analise = _ms(start)
    q_analise = sql_count["n"]

    sql_count["n"] = 0
    start = time.perf_counter()
    NfHistoricoValidator(fechamento._repository).validar_conflitos_do_lote(probe_df)
    t_nf = _ms(start)
    q_nf = sql_count["n"]

    sql_count["n"] = 0
    start = time.perf_counter()
    fechamento._repository.list_all()
    t_list_all = _ms(start)
    q_list_all = sql_count["n"]

    # Complementacao save query count
    existente = fechamento._repository.list_all()[0]
    lista = list(existente.itens)
    from carregamentos.models.carregamento import CarregamentoItem

    lista.append(
        CarregamentoItem(
            nf="8888",
            cprod="NEW",
            descricao="N",
            quantidade=1,
            unidade="UN",
            peso=1.0,
            destinatario="A",
            rota="R",
            chave_nfe=f"{8888:044d}",
            status_nf="OK",
        )
    )
    existente.itens = lista
    sql_count["n"] = 0
    start = time.perf_counter()
    with UnitOfWork() as uow:
        SqlCarregamentoRepository(uow.session)._save_in_session(uow.session, existente)
    t_save = _ms(start)
    q_save = sql_count["n"]

    with UnitOfWork() as uow:
        item_count = len(uow.session.scalars(select(ItemCarregamentoORM)).all())

    print("## Benchmark O1 (SQLite isolado)")
    print()
    print("| Operacao | Tempo ms | SQL | Registros lidos |")
    print("| --- | ---: | ---: | ---: |")
    print(f"| XML list_all_records (200) | {t_xml_all:.1f} | {q_xml_all} | {len(all_recs)} |")
    print(f"| XML list_records_by_identities (5/200) | {t_xml_delta:.1f} | {q_xml_delta} | {len(by_id)} |")
    print(f"| Analise lote (2 NFs / 30 cargas) | {t_analise:.1f} | {q_analise} | targeted |")
    print(f"| NF validation conflitos | {t_nf:.1f} | {q_nf} | targeted |")
    print(f"| list_all carregamentos | {t_list_all:.1f} | {q_list_all} | 30 |")
    print(f"| save complemento (preload itens) | {t_save:.1f} | {q_save} | 1 carga |")
    print()
    print(f"Itens totais no banco: {item_count}")
    print(f"Ganho XML registros: {len(all_recs)} -> {len(by_id)} ({100 * (1 - len(by_id) / max(len(all_recs), 1)):.0f}% menos)")
    event.remove(engine, "before_cursor_execute", _count_sql)


if __name__ == "__main__":
    main()
