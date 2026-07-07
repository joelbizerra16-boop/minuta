"""Benchmark offline das operações críticas (sem Streamlit UI)."""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.bootstrap import configure_application_storage
from core.performance import build_performance_report, clear_timings, measure, record_timing


def _timed(label: str, func, *args, **kwargs):
    start = time.perf_counter()
    result = func(*args, **kwargs)
    record_timing(label, (time.perf_counter() - start) * 1000.0)
    return result


def main() -> None:
    configure_application_storage()

    from app import (
        XMLS_PROCESSADOS_JSON_PATH,
        carregar_classificacao_produtos_json,
        carregar_separacao_json,
        carregar_xmls_processados_json,
        create_empty_processed_df,
        integrate_excel_with_xml,
        load_excel_base,
        prepare_processed_search_dataframe,
        build_balcao_lookup_dataframe,
        sincronizar_base_separacao,
    )

    clear_timings()

    with measure("startup.carregar_xmls"):
        xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))

    with measure("startup.carregar_classificacao"):
        classificacao_records, _ = carregar_classificacao_produtos_json(str(ROOT / "data" / "classificacao_produtos.json"))

    with measure("startup.sincronizar_separacao"):
        sincronizar_base_separacao(xml_records, classificacao_records)

    with measure("startup.carregar_separacao"):
        separacao_records, _ = carregar_separacao_json(str(ROOT / "data" / "separacao.json"))

    with measure("dataframe.balcao_lookup"):
        build_balcao_lookup_dataframe(xml_records)

    sample_excel = ROOT / "data" / "sample_benchmark.xlsx"
    if sample_excel.is_file():
        with measure("excel.load_base"):
            excel_bytes = sample_excel.read_bytes()

            class _Upload:
                def __init__(self, data: bytes):
                    self._data = data

                def getvalue(self) -> bytes:
                    return self._data

                def seek(self, _offset: int) -> None:
                    return None

            base_df = load_excel_base(_Upload(excel_bytes))

        with measure("process.integrate_excel_xml"):
            processed_df, _, _, _ = integrate_excel_with_xml(base_df, xml_records)

        with measure("dataframe.prepare_processed"):
            prepare_processed_search_dataframe(processed_df)
    else:
        with measure("dataframe.prepare_processed_empty"):
            prepare_processed_search_dataframe(create_empty_processed_df())

    print(build_performance_report())
    print(f"\nXMLs carregados: {len(xml_records)}")
    print(f"Registros separação: {len(separacao_records)}")


if __name__ == "__main__":
    main()
