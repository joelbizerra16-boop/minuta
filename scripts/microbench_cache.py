"""Compara custo de operações repetidas (simula rerenders)."""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.bootstrap import configure_application_storage

configure_application_storage(ROOT / "data")

from app import (
    SEPARACAO_JSON_PATH,
    TABLE_COLUMNS,
    XMLS_PROCESSADOS_JSON_PATH,
    build_balcao_lookup_dataframe,
    build_display_table,
    carregar_separacao_json,
    carregar_xmls_processados_json,
    create_empty_processed_df,
    generate_lote_pdf,
    get_latest_closed_lote_summary,
    get_lote_records,
    prepare_processed_search_dataframe,
)


def _repeat(label: str, iterations: int, func) -> float:
    start = time.perf_counter()
    for _ in range(iterations):
        func()
    elapsed_ms = (time.perf_counter() - start) * 1000.0
    print(f"{label}: {elapsed_ms:.1f} ms ({iterations} execuções)")
    return elapsed_ms


def main() -> None:
    xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    separacao_records, _ = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    empty = create_empty_processed_df()

    balcao_total = _repeat("balcao_lookup_sem_cache", 5, lambda: build_balcao_lookup_dataframe(xml_records))
    prep_total = _repeat("prepare_processed_sem_cache", 5, lambda: prepare_processed_search_dataframe(empty))
    display_total = _repeat(
        "build_display_table",
        5,
        lambda: build_display_table(empty[TABLE_COLUMNS].copy()),
    )

    latest = get_latest_closed_lote_summary(separacao_records)
    pdf_total = 0.0
    if latest:
        lote_id = str(latest.get("Lote", ""))
        records = get_lote_records(separacao_records, lote_id)
        pdf_total = _repeat("generate_lote_pdf", 3, lambda: generate_lote_pdf(latest, records))

    print()
    print("Estimativa de ganho por rerender (com cache de sessão):")
    print(f"  balcao_lookup: ~{balcao_total - (balcao_total / 5):.0f} ms economizados")
    print(f"  prepare_processed: ~{prep_total - (prep_total / 5):.0f} ms economizados")
    print(f"  display_table: ~{display_total - (display_total / 5):.0f} ms economizados")
    if pdf_total:
        print(f"  lote_pdf (tela separação): ~{pdf_total - (pdf_total / 3):.0f} ms por rerender")


if __name__ == "__main__":
    main()
