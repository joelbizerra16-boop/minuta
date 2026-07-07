"""Benchmark P1 — operações críticas com medição antes/depois."""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.bootstrap import configure_application_storage
from core.performance import clear_timings, record_timing


def _ms(start: float) -> float:
    return (time.perf_counter() - start) * 1000.0


def main() -> None:
    configure_application_storage()
    clear_timings()

    from carregamentos.bootstrap import (
        configure_carregamentos_storage,
        get_analise_operacional_service,
        get_carregamento_repository,
        get_xml_export_service,
    )
    from carregamentos.models.carregamento import CarregamentoFiltro
    from carregamentos.services.nf_validation import NfHistoricoValidator
    from infrastructure.services.documento_xml_service import DocumentoXmlService, XmlDocumentalItem

    configure_carregamentos_storage(ROOT / "data")
    repo = get_carregamento_repository()
    analise = get_analise_operacional_service()
    xml_export = get_xml_export_service()

    import pandas as pd

    from app import (
        XMLS_PROCESSADOS_JSON_PATH,
        carregar_xmls_processados_json,
        create_empty_processed_df,
        integrate_excel_with_xml,
        load_excel_base,
        prepare_processed_search_dataframe,
    )

    timings: list[tuple[str, float]] = []

    start = time.perf_counter()
    xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    timings.append(("sql.carregar_xmls_runtime", _ms(start)))

    processed_df = create_empty_processed_df()
    sample_excel = ROOT / "data" / "sample_benchmark.xlsx"
    if sample_excel.is_file():

        class _Upload:
            def __init__(self, data: bytes):
                self._data = data

            def getvalue(self) -> bytes:
                return self._data

            def seek(self, _offset: int) -> None:
                return None

        start = time.perf_counter()
        base_df = load_excel_base(_Upload(sample_excel.read_bytes()))
        timings.append(("excel.load_base", _ms(start)))

        start = time.perf_counter()
        processed_df, _, _, _ = integrate_excel_with_xml(base_df, xml_records)
        timings.append(("process.integrate_excel_xml", _ms(start)))

    if not processed_df.empty:
        start = time.perf_counter()
        analise.analisar_lote_processado(processed_df)
        timings.append(("process.analise_operacional (1a)", _ms(start)))

        start = time.perf_counter()
        analise.montar_auditoria_nfs_lote(processed_df)
        timings.append(("auditoria.montar_nfs_lote", _ms(start)))

        start = time.perf_counter()
        analise.analisar_lote_processado(processed_df)
        timings.append(("process.analise_operacional (cache)", _ms(start)))

        validator = NfHistoricoValidator(repo)
        start = time.perf_counter()
        validator.validar_conflitos_do_lote(processed_df)
        timings.append(("nf_validation.conflitos_lote", _ms(start)))

    start = time.perf_counter()
    repo.list_all()
    timings.append(("sql.carregamento_list_all", _ms(start)))

    start = time.perf_counter()
    repo.search(CarregamentoFiltro(data_inicial="2000-01-01", data_final="2099-12-31"))
    timings.append(("sql.carregamento_search", _ms(start)))

    start = time.perf_counter()
    prepare_processed_search_dataframe(processed_df if not processed_df.empty else create_empty_processed_df())
    timings.append(("dataframe.prepare_processed", _ms(start)))

    carregamentos = repo.list_all()
    if carregamentos:
        carregamento_id = int(carregamentos[-1].id)
        start = time.perf_counter()
        xml_export.collect_xmls_for_carregamento(carregamento_id)
        timings.append((f"export.xml_carregamento_id={carregamento_id}", _ms(start)))

    service = DocumentoXmlService()
    items = [
        XmlDocumentalItem(
            file_bytes=b"<nfeProc><NFe><infNFe Id=\"NFe1\"/></NFe></nfeProc>",
            hash_sha256=f"bench-{index}",
            original_filename=f"bench_{index}.xml",
            chave_nfe=f"{'1' * 43}{index % 10}",
        )
        for index in range(20)
    ]
    start = time.perf_counter()
    service.persist_raw_xml_batch(items[:5])
    timings.append(("import.xml_documental_batch(5)", _ms(start)))

    for label, elapsed in timings:
        record_timing(label, elapsed)

    print("## Benchmark P1")
    print()
    print("| Operação | Tempo (ms) |")
    print("| --- | ---: |")
    for label, elapsed in sorted(timings, key=lambda item: item[1], reverse=True):
        print(f"| {label} | {elapsed:.1f} |")
    print()
    print(f"XMLs em memória: {len(xml_records)}")
    print(f"Carregamentos: {len(carregamentos)}")


if __name__ == "__main__":
    main()
