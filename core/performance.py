"""Instrumentação de desempenho para medição sem alterar regras de negócio."""

from __future__ import annotations

import time
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Iterator

_PERF_STORE: dict[str, "_PerfAccumulator"] = {}


@dataclass
class _PerfAccumulator:
    label: str
    total_ms: float = 0.0
    count: int = 0
    max_ms: float = 0.0

    def add(self, elapsed_ms: float) -> None:
        self.total_ms += elapsed_ms
        self.count += 1
        if elapsed_ms > self.max_ms:
            self.max_ms = elapsed_ms

    @property
    def avg_ms(self) -> float:
        if self.count == 0:
            return 0.0
        return self.total_ms / self.count


def _session_store() -> dict[str, _PerfAccumulator] | None:
    try:
        import streamlit as st
    except ImportError:
        return None

    key = "_perf_timings"
    if key not in st.session_state:
        st.session_state[key] = {}
    return st.session_state[key]


def record_timing(label: str, elapsed_ms: float) -> None:
    session_store = _session_store()
    if session_store is not None:
        entry = session_store.get(label)
        if entry is None:
            session_store[label] = _PerfAccumulator(label=label)
            entry = session_store[label]
        entry.add(elapsed_ms)

    entry = _PERF_STORE.get(label)
    if entry is None:
        _PERF_STORE[label] = _PerfAccumulator(label=label)
        entry = _PERF_STORE[label]
    entry.add(elapsed_ms)


@contextmanager
def measure(label: str) -> Iterator[None]:
    start = time.perf_counter()
    try:
        yield
    finally:
        record_timing(label, (time.perf_counter() - start) * 1000.0)


def clear_timings() -> None:
    _PERF_STORE.clear()
    session_store = _session_store()
    if session_store is not None:
        session_store.clear()


def get_timing_entries() -> list[_PerfAccumulator]:
    session_store = _session_store()
    if session_store:
        merged: dict[str, _PerfAccumulator] = {}
        for label, entry in {**_PERF_STORE, **session_store}.items():
            acc = merged.get(label)
            if acc is None:
                merged[label] = _PerfAccumulator(
                    label=label,
                    total_ms=entry.total_ms,
                    count=entry.count,
                    max_ms=entry.max_ms,
                )
            else:
                acc.total_ms += entry.total_ms
                acc.count += entry.count
                acc.max_ms = max(acc.max_ms, entry.max_ms)
        return sorted(merged.values(), key=lambda item: item.total_ms, reverse=True)
    return sorted(_PERF_STORE.values(), key=lambda item: item.total_ms, reverse=True)


def build_performance_report() -> str:
    entries = get_timing_entries()
    if not entries:
        return "Nenhuma medição registrada nesta sessão."

    lines = [
        "## Relatório de Performance",
        "",
        "| Operação | Total (ms) | Chamadas | Média (ms) | Máx (ms) |",
        "| --- | ---: | ---: | ---: | ---: |",
    ]
    for entry in entries:
        lines.append(
            f"| {entry.label} | {entry.total_ms:.1f} | {entry.count} | {entry.avg_ms:.1f} | {entry.max_ms:.1f} |"
        )

    lines.extend(["", "### Maiores gargalos", ""])
    for index, entry in enumerate(entries[:8], start=1):
        lines.append(f"{index}. **{entry.label}** — {entry.total_ms:.1f} ms total ({entry.count} chamada(s))")

    return "\n".join(lines)


def bump_processed_data_version() -> int:
    try:
        import streamlit as st
    except ImportError:
        return 0

    version = int(st.session_state.get("_processed_data_version", 0)) + 1
    st.session_state["_processed_data_version"] = version
    st.session_state.pop("_prepared_processed_df", None)
    st.session_state.pop("_display_table_df", None)
    st.session_state.pop("_display_table_version", None)
    return version


def invalidate_balcao_lookup_cache() -> None:
    try:
        import streamlit as st
    except ImportError:
        return

    st.session_state.pop("balcao_lookup_df", None)
    st.session_state.pop("_balcao_lookup_signature", None)


def invalidate_latest_closed_lote_pdf_cache() -> None:
    try:
        import streamlit as st
    except ImportError:
        return

    st.session_state.pop("_latest_closed_lote_pdf", None)
    st.session_state.pop("_latest_closed_lote_pdf_sig", None)
