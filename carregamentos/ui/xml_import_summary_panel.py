from __future__ import annotations

import html
import re
from dataclasses import dataclass, field

import streamlit as st

_UPLOAD_DUPLICATE_RE = re.compile(r"^XML duplicado ignorado: (.+)$")
_XML_ERROR_RE = re.compile(r"^Erro no XML (.+?): (.+)$")
_XML_NO_IDENTITY_RE = re.compile(r"^XML sem chave/NF identificavel: (.+)$")
_BATCH_DUPLICATE_RE = re.compile(r"^XML duplicado no lote ignorado: (.+)$")
_STORAGE_DUPLICATE_RE = re.compile(r"^XML duplicado ou desatualizado ignorado: (.+)$")
_SEPARATED_NF_RE = re.compile(r"^NF (.+?) ignorada no upload porque ja esta separada\.$")
_BATCH_KEPT_NEWER_RE = re.compile(
    r"^NF (.+?) duplicada no lote\. Foi mantido o arquivo mais recente: (.+)$"
)
_NF_UPDATED_RE = re.compile(r"^NF (.+?) atualizada pelo evento/XML mais recente\.$")
_DOCUMENTAL_FAILURE_RE = re.compile(r"^Persistencia documental dos XMLs nao concluida: (.+)$")
_IMPORT_ABORTED_RE = re.compile(r"^Importacao de XMLs abortada: (.+)$")


@dataclass
class XmlImportDetailItem:
    filename: str
    reason: str = ""


@dataclass
class XmlImportReport:
    selected: int = 0
    imported: int = 0
    duplicated: int = 0
    invalid: int = 0
    errors: int = 0
    rejected: int = 0
    elapsed_seconds: float = 0.0
    imported_files: list[str] = field(default_factory=list)
    duplicated_files: list[XmlImportDetailItem] = field(default_factory=list)
    invalid_files: list[XmlImportDetailItem] = field(default_factory=list)
    rejected_files: list[XmlImportDetailItem] = field(default_factory=list)
    error_files: list[XmlImportDetailItem] = field(default_factory=list)
    general_errors: list[str] = field(default_factory=list)

    def has_content(self) -> bool:
        return (
            self.selected > 0
            or self.imported > 0
            or self.duplicated > 0
            or self.invalid > 0
            or self.errors > 0
            or self.rejected > 0
            or bool(self.general_errors)
        )


def _append_unique_file(items: list[XmlImportDetailItem], filename: str, reason: str = "") -> None:
    normalized = filename.strip()
    if not normalized:
        return
    if any(item.filename == normalized for item in items):
        return
    items.append(XmlImportDetailItem(filename=normalized, reason=reason.strip()))


def _append_unique_name(names: list[str], filename: str) -> None:
    normalized = filename.strip()
    if normalized and normalized not in names:
        names.append(normalized)


def classify_xml_import_issues(
    issues: list[str],
    *,
    nf_to_arquivo: dict[str, str] | None = None,
) -> tuple[
    list[XmlImportDetailItem],
    list[XmlImportDetailItem],
    list[XmlImportDetailItem],
    list[XmlImportDetailItem],
    list[str],
    set[str],
]:
    duplicated: list[XmlImportDetailItem] = []
    invalid: list[XmlImportDetailItem] = []
    rejected: list[XmlImportDetailItem] = []
    errors: list[XmlImportDetailItem] = []
    general_errors: list[str] = []
    excluded_filenames: set[str] = set()
    nf_lookup = nf_to_arquivo or {}

    for issue in issues:
        text = str(issue or "").strip()
        if not text:
            continue

        match = _UPLOAD_DUPLICATE_RE.match(text)
        if match:
            filename = match.group(1).strip()
            _append_unique_file(duplicated, filename)
            excluded_filenames.add(filename)
            continue

        match = _XML_ERROR_RE.match(text)
        if match:
            filename = match.group(1).strip()
            reason = match.group(2).strip()
            _append_unique_file(errors, filename, reason)
            excluded_filenames.add(filename)
            continue

        match = _XML_NO_IDENTITY_RE.match(text)
        if match:
            filename = match.group(1).strip()
            _append_unique_file(invalid, filename)
            excluded_filenames.add(filename)
            continue

        match = _BATCH_DUPLICATE_RE.match(text)
        if match:
            filename = match.group(1).strip()
            _append_unique_file(duplicated, filename)
            excluded_filenames.add(filename)
            continue

        match = _STORAGE_DUPLICATE_RE.match(text)
        if match:
            filename = match.group(1).strip()
            _append_unique_file(duplicated, filename)
            excluded_filenames.add(filename)
            continue

        match = _SEPARATED_NF_RE.match(text)
        if match:
            nf = match.group(1).strip()
            filename = nf_lookup.get(nf, nf)
            _append_unique_file(rejected, filename, "NF ja separada")
            excluded_filenames.add(filename)
            continue

        match = _BATCH_KEPT_NEWER_RE.match(text)
        if match:
            continue

        match = _NF_UPDATED_RE.match(text)
        if match:
            continue

        match = _DOCUMENTAL_FAILURE_RE.match(text)
        if match:
            general_errors.append(match.group(1).strip())
            continue

        match = _IMPORT_ABORTED_RE.match(text)
        if match:
            general_errors.append(match.group(1).strip())
            continue

        general_errors.append(text)

    return duplicated, invalid, rejected, errors, general_errors, excluded_filenames


def build_xml_import_report(
    *,
    summary: dict[str, int],
    issues: list[str],
    upload_duplicate_messages: list[str] | None = None,
    selected_count: int = 0,
    pending_filenames: list[str] | None = None,
    elapsed_seconds: float = 0.0,
    nf_to_arquivo: dict[str, str] | None = None,
) -> XmlImportReport:
    all_issues = list(issues or []) + list(upload_duplicate_messages or [])
    duplicated, invalid, rejected, errors, general_errors, excluded = classify_xml_import_issues(
        all_issues,
        nf_to_arquivo=nf_to_arquivo,
    )

    imported_files: list[str] = []
    for filename in pending_filenames or []:
        normalized = str(filename or "").strip()
        if normalized and normalized not in excluded:
            _append_unique_name(imported_files, normalized)

    imported_count = int(summary.get("processados", 0) or 0)
    if imported_count <= 0 and imported_files:
        imported_count = len(imported_files)

    duplicated_count = int(summary.get("duplicados", 0) or 0) + len(
        [message for message in (upload_duplicate_messages or []) if message.strip()]
    )
    if duplicated_count < len(duplicated):
        duplicated_count = len(duplicated)

    return XmlImportReport(
        selected=max(int(selected_count or 0), int(summary.get("total_arquivos", 0) or 0)),
        imported=imported_count,
        duplicated=duplicated_count,
        invalid=len(invalid),
        errors=int(summary.get("erros", 0) or 0) or len(errors),
        rejected=len(rejected),
        elapsed_seconds=max(float(elapsed_seconds or 0.0), 0.0),
        imported_files=imported_files,
        duplicated_files=duplicated,
        invalid_files=invalid,
        rejected_files=rejected,
        error_files=errors,
        general_errors=general_errors,
    )


def merge_xml_import_reports(existing: XmlImportReport | None, incoming: XmlImportReport) -> XmlImportReport:
    if existing is None or not existing.has_content():
        return incoming

    merged = XmlImportReport(
        selected=existing.selected + incoming.selected,
        imported=existing.imported + incoming.imported,
        duplicated=existing.duplicated + incoming.duplicated,
        invalid=existing.invalid + incoming.invalid,
        errors=existing.errors + incoming.errors,
        rejected=existing.rejected + incoming.rejected,
        elapsed_seconds=existing.elapsed_seconds + incoming.elapsed_seconds,
    )

    for source in (existing, incoming):
        for filename in source.imported_files:
            _append_unique_name(merged.imported_files, filename)
        for item in source.duplicated_files:
            _append_unique_file(merged.duplicated_files, item.filename, item.reason)
        for item in source.invalid_files:
            _append_unique_file(merged.invalid_files, item.filename, item.reason)
        for item in source.rejected_files:
            _append_unique_file(merged.rejected_files, item.filename, item.reason)
        for item in source.error_files:
            _append_unique_file(merged.error_files, item.filename, item.reason)
        for message in source.general_errors:
            if message not in merged.general_errors:
                merged.general_errors.append(message)

    return merged


def _render_detail_section(title: str, items: list[XmlImportDetailItem]) -> None:
    if not items:
        return
    st.markdown(f"**{title}**")
    for item in items:
        if item.reason:
            st.markdown(
                f"- `{html.escape(item.filename)}`  \n"
                f"  Motivo: {html.escape(item.reason)}"
            )
        else:
            st.markdown(f"- `{html.escape(item.filename)}`")
    st.markdown("")


def _render_summary_block(report: XmlImportReport) -> None:
    rows = [
        ("XMLs selecionados", str(report.selected)),
        ("Importados", str(report.imported)),
        ("Duplicados", str(report.duplicated)),
        ("Invalidos", str(report.invalid)),
        ("Com erro", str(report.errors)),
        ("Tempo processamento", f"{report.elapsed_seconds:.1f} segundos"),
    ]
    label_width = max(len(label) for label, _ in rows)
    lines = ["Importacao de XML concluida", ""]
    for label, value in rows:
        dots = "." * max(1, label_width - len(label) + 4)
        lines.append(f"{label}{dots}: {value}")
    st.markdown(
        f'<pre style="margin:0; white-space:pre-wrap; font-family:inherit; font-size:0.92rem;">'
        f"{html.escape(chr(10).join(lines))}"
        f"</pre>",
        unsafe_allow_html=True,
    )


def render_xml_import_summary_panel(
    report: XmlImportReport | None,
    *,
    error_message: str = "",
) -> None:
    if error_message:
        st.error(error_message)

    if report is None or not report.has_content():
        return

    _render_summary_block(report)

    has_details = any(
        [
            report.imported_files,
            report.duplicated_files,
            report.invalid_files,
            report.rejected_files,
            report.error_files,
            report.general_errors,
        ]
    )
    if not has_details:
        return

    with st.expander("Visualizar detalhes", expanded=False):
        st.markdown("**Detalhes da importacao**")
        if report.imported_files:
            st.markdown("**XML importados**")
            for filename in report.imported_files:
                st.markdown(f"- `{html.escape(filename)}`")
            st.markdown("")

        _render_detail_section("XML duplicados", report.duplicated_files)
        _render_detail_section("XML invalidos", report.invalid_files)
        _render_detail_section("XML rejeitados", report.rejected_files)
        _render_detail_section("XML com erro", report.error_files)

        if report.general_errors:
            st.markdown("**Outros avisos**")
            for message in report.general_errors:
                st.markdown(f"- {html.escape(message)}")
