from __future__ import annotations

from carregamentos.ui.xml_import_summary_panel import (
    build_xml_import_report,
    classify_xml_import_issues,
    merge_xml_import_reports,
)


def test_classify_upload_and_parse_issues() -> None:
    issues = [
        "XML duplicado ignorado: dup_upload.xml",
        "Erro no XML invalido.xml: Estrutura XML invalida",
        "XML sem chave/NF identificavel: sem_chave.xml",
        "XML duplicado no lote ignorado: dup_lote.xml",
        "XML duplicado ou desatualizado ignorado: dup_storage.xml",
        "NF 123 ignorada no upload porque ja esta separada.",
    ]
    duplicated, invalid, rejected, errors, general_errors, excluded = classify_xml_import_issues(
        issues,
        nf_to_arquivo={"123": "nf_separada.xml"},
    )

    assert [item.filename for item in duplicated] == ["dup_upload.xml", "dup_lote.xml", "dup_storage.xml"]
    assert [item.filename for item in invalid] == ["sem_chave.xml"]
    assert [item.filename for item in rejected] == ["nf_separada.xml"]
    assert errors[0].filename == "invalido.xml"
    assert errors[0].reason == "Estrutura XML invalida"
    assert "invalido.xml" in excluded
    assert not general_errors


def test_build_xml_import_report_counts_and_imported_files() -> None:
    report = build_xml_import_report(
        summary={
            "total_arquivos": 3,
            "processados": 2,
            "erros": 1,
            "duplicados": 1,
        },
        issues=[
            "XML duplicado ignorado: dup.xml",
            "Erro no XML bad.xml: Assinatura invalida",
        ],
        upload_duplicate_messages=["XML duplicado ignorado: dup.xml"],
        selected_count=4,
        pending_filenames=["ok1.xml", "ok2.xml", "bad.xml"],
        elapsed_seconds=6.8,
    )

    assert report.selected == 4
    assert report.imported == 2
    assert report.duplicated == 2
    assert report.errors == 1
    assert report.elapsed_seconds == 6.8
    assert report.imported_files == ["ok1.xml", "ok2.xml"]


def test_merge_xml_import_reports_accumulates_values() -> None:
    first = build_xml_import_report(
        summary={"total_arquivos": 2, "processados": 1, "duplicados": 1},
        issues=["XML duplicado ignorado: a.xml"],
        selected_count=2,
        pending_filenames=["ok.xml"],
        elapsed_seconds=2.0,
    )
    second = build_xml_import_report(
        summary={"total_arquivos": 1, "processados": 1},
        issues=[],
        selected_count=1,
        pending_filenames=["ok2.xml"],
        elapsed_seconds=3.5,
    )

    merged = merge_xml_import_reports(first, second)

    assert merged.selected == 3
    assert merged.imported == 2
    assert merged.duplicated == 1
    assert merged.elapsed_seconds == 5.5
    assert merged.imported_files == ["ok.xml", "ok2.xml"]
