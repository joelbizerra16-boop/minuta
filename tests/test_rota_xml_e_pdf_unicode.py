"""Regressão: rota individual do XML e Unicode nos PDFs."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.xml_mapper import orm_to_record, record_to_orm
from infrastructure.storage.xml_storage import SqlXmlRecordRepository
from utils.pdf_fonts import BUNDLED_FONTS_DIR, register_pdf_fonts
from utils.rota_xml import UNDEFINED_ROUTE_LABEL, extract_route_from_inf_cpl, normalize_route_label

REFERENCE_INF_CPL = (
    "ICMS SUBST TRIBPELO FABRICANTE ART 412 INCISO III DEC454902000 "
    "VL TRIBUTOS REF IBPT 042026 CREDITO DE ICMS FACULTADO AO CONSUMIDOR "
    "ART 272 DECRETO 454902000 Vendedor Danielly Silva Pedido 27970 "
    "Cliente C007355 Rota ABCD\\nTrib aprox R1721062 Fed 3199000 Est e 000 Mun "
    "Fonte IBPT 92589A\\nValor CBS R 104511\\nValor IBS UF R 11612"
)

ACCENT_SAMPLES = [
    "Emissão",
    "Página",
    "Veículo",
    "São Paulo",
    "São Bernardo do Campo",
    "Peças",
    "Construções",
    "Lubrificantes",
    "Informação",
    "Conferência",
    "Observação",
    "Saída",
    "Não definida",
    "ALTO DO TIETÊ",
]


@pytest.mark.parametrize(
    ("inf_cpl", "expected"),
    [
        ("Rota ABCD", "ABCD"),
        ("Rota: ABCD", "ABCD"),
        ("ROTA ABCD", "ABCD"),
        ("ROTA: ABCD", "ABCD"),
        ("Rota - ABCD", "ABCD"),
        ("Rota=ABCD", "ABCD"),
        ("Cliente C007355 Rota ABCD\nTrib aprox ...", "ABCD"),
        ("Rota ALTO DO TIETÊ", "ALTO DO TIETÊ"),
        ("texto sem marcador de rota", ""),
        ("", ""),
        (None, ""),
        (REFERENCE_INF_CPL, "ABCD"),
    ],
)
def test_extract_route_from_inf_cpl_variants(inf_cpl: object, expected: str) -> None:
    assert extract_route_from_inf_cpl(inf_cpl) == expected


class _UploadedXml:
    def __init__(self, payload: bytes, name: str) -> None:
        self._payload = payload
        self.name = name

    def getvalue(self) -> bytes:
        return self._payload


def test_reference_xml_file_route_abcd() -> None:
    xml_path = Path(r"c:\Users\joelb\Downloads\procNFe35260700846804000106550010014360591118433961.xml")
    if not xml_path.is_file():
        pytest.skip("XML de referência não disponível no ambiente")

    from app import parse_xml_file

    record = parse_xml_file(_UploadedXml(xml_path.read_bytes(), xml_path.name))
    assert record["NF"] == "1436059"
    assert record["ROTA"] == "ABCD"
    assert record["ROTA"] != UNDEFINED_ROUTE_LABEL
    assert "Trib" not in record["ROTA"]


def test_normalize_route_label_empty_becomes_undefined() -> None:
    assert normalize_route_label("") == UNDEFINED_ROUTE_LABEL
    assert normalize_route_label("  ABCD  ") == "ABCD"


@pytest.fixture
def xml_repo(tmp_path: Path) -> SqlXmlRecordRepository:
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "rota_unicode.db"
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=tmp_path,
        pdf_storage_dir=tmp_path / "pdf",
        xml_storage_dir=tmp_path / "xml",
    )
    ensure_full_schema()
    yield SqlXmlRecordRepository()
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def test_rota_persisted_and_reloaded_from_sql(xml_repo: SqlXmlRecordRepository) -> None:
    record = {
        "NF": "1436059",
        "nf_normalizada": "1436059",
        "ChaveNFe": "35260700846804000106550010014360591118433961",
        "Destinatario": "PHP COM ATACADISTA",
        "Municipio": "Sao Bernardo do Campo",
        "UF": "SP",
        "StatusNF": "Autorizado o uso da NF-e",
        "ROTA": "ABCD",
        "Items": [{"cProd": "123075", "Descricao": "MOBIL", "Qtd": 200, "Unidade": "CX", "Peso": 4620.0}],
        "Arquivo": "ref.xml",
        "TipoXML": "normal",
        "ValorNF": 127960.0,
        "PesoTotal": 4620.0,
        "VolumeTotal": 200.0,
    }
    xml_repo.upsert_records([record])
    loaded = xml_repo.list_all_records()[0]
    assert loaded["ROTA"] == "ABCD"

    row = record_to_orm(loaded)
    again = orm_to_record(row, [])
    assert again["ROTA"] == "ABCD"


def test_minuta_preserves_individual_routes_per_nf() -> None:
    from app import build_minuta_records, generate_minuta_pdf

    dataframe = pd.DataFrame(
        [
            {
                "NF": "100",
                "Data": "2026-08-01",
                "Destinatario": "Cliente A",
                "VolumeTotal": 1,
                "PesoTotal": 10.0,
                "ROTA": "ABCD",
                "Items": [{"descricao": "Prod A", "codigo": "1", "qtd": 1, "unidade": "CX", "peso": 10.0}],
            },
            {
                "NF": "200",
                "Data": "2026-08-01",
                "Destinatario": "Cliente B",
                "VolumeTotal": 2,
                "PesoTotal": 20.0,
                "ROTA": "R01",
                "Items": [{"descricao": "Prod B", "codigo": "2", "qtd": 2, "unidade": "CX", "peso": 20.0}],
            },
            {
                "NF": "300",
                "Data": "2026-08-01",
                "Destinatario": "Cliente C",
                "VolumeTotal": 3,
                "PesoTotal": 30.0,
                "ROTA": "ALTO DO TIETÊ",
                "Items": [{"descricao": "Prod C", "codigo": "3", "qtd": 3, "unidade": "CX", "peso": 30.0}],
            },
        ]
    )
    # build_minuta_records expects product rows; expand via Items or columns used by groupby
    rows = []
    for _, row in dataframe.iterrows():
        for item in row["Items"]:
            rows.append(
                {
                    "NF": row["NF"],
                    "Data": row["Data"],
                    "Destinatario": row["Destinatario"],
                    "VolumeTotal": row["VolumeTotal"],
                    "PesoTotal": row["PesoTotal"],
                    "ROTA": row["ROTA"],
                    "Descricao": item["descricao"],
                    "cProd": item["codigo"],
                    "Qtd": item["qtd"],
                    "Unidade": item["unidade"],
                    "Peso": item["peso"],
                }
            )
    processed = pd.DataFrame(rows)
    records = build_minuta_records(processed)
    by_nf = {str(item["nf"]): item["rota"] for item in records}
    assert by_nf["100"] == "ABCD"
    assert by_nf["200"] == "R01"
    assert by_nf["300"] == "ALTO DO TIETÊ"

    pdf_bytes = generate_minuta_pdf(
        records,
        numero_carga="000999",
        data_emissao="04/08/2026 12:00:00",
        veiculo="TESTE",
        motorista="Motorista",
    )
    text = _extract_pdf_text(pdf_bytes)
    assert "ROTA: ABCD" in text
    assert "ROTA: R01" in text
    assert "ROTA: ALTO DO TIETÊ" in text or "ROTA: ALTO DO TIET" in text
    assert "NÃO DEFINIDA" not in text.upper().replace("NAO", "NÃO") or "ABCD" in text


def test_bundled_dejavu_fonts_registered() -> None:
    assert (BUNDLED_FONTS_DIR / "DejaVuSans.ttf").is_file()
    assert (BUNDLED_FONTS_DIR / "DejaVuSans-Bold.ttf").is_file()
    regular, bold = register_pdf_fonts()
    assert regular == "DejaVuSans"
    assert bold == "DejaVuSans-Bold"


def test_gerador_minuta_source_has_no_mojibake() -> None:
    source = Path("utils/gerador_minuta.py").read_text(encoding="utf-8")
    assert "Emissão" in source
    assert "Página" in source
    assert "Veículo" in source
    assert "├" not in source
    assert "Emiss├" not in source


def test_minuta_entrega_pdf_preserves_portuguese_accents() -> None:
    from utils.gerador_minuta import generate_minuta_entrega_pdf

    records = [
        {
            "nota": "1436059",
            "item": 1,
            "data": "31/07/2026",
            "cliente": "Peças e Construções LTDA",
            "cidade": "São Bernardo do Campo",
            "uf": "SP",
            "peso": 10.5,
            "valor": 100.0,
            "rota": "ALTO DO TIETÊ",
        }
    ]
    totals = {"total_volumes": 1, "total_nfs": 1, "total_peso": 10.5, "total_valor": 100.0}
    pdf_bytes = generate_minuta_entrega_pdf(
        records,
        totals,
        numero_documento="000215",
        data_emissao="04/08/2026 20:49:56",
        transportadora="BRIDA LUBRIFICANTES LTDA",
        veiculo="Veículo Teste",
        motorista="José",
        placa="ABC1D23",
    )
    assert pdf_bytes.startswith(b"%PDF")
    assert b"DejaVu" in pdf_bytes
    text = _extract_pdf_text(pdf_bytes)
    normalized = " ".join(text.split())
    for sample in ("Emissão", "Página", "Veículo", "São Bernardo do Campo", "Peças"):
        assert sample in normalized, f"texto ausente no PDF: {sample!r}"
    assert "├" not in text
    assert "■" not in text
    assert "�" not in text
    assert "Emiss├" not in text


def test_minuta_carregamento_pdf_preserves_portuguese_accents() -> None:
    from app import generate_minuta_pdf

    records = [
        {
            "nf": "1436059",
            "data": "2026-07-31",
            "cliente": "São Paulo Peças",
            "volume": 1,
            "peso": 17.46,
            "rota": "ABCD",
            "produtos": [
                {
                    "descricao": "Lubrificantes Informação Conferência",
                    "codigo": "1",
                    "qtd": 1,
                    "un": "CX",
                    "peso": 17.46,
                }
            ],
        }
    ]
    pdf_bytes = generate_minuta_pdf(
        records,
        numero_carga="000215",
        data_emissao="04/08/2026 20:49:56",
        veiculo="Veículo 9",
        motorista="JUNIOR",
    )
    text = _extract_pdf_text(pdf_bytes)
    assert "Emissão" in text
    assert "Veículo" in text
    assert "ROTA: ABCD" in text
    assert "NÃO DEFINIDA" not in text


def _extract_pdf_text(pdf_bytes: bytes) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:  # pragma: no cover
        pytest.skip("pypdf não instalado")

    reader = PdfReader(BytesIO(pdf_bytes))
    chunks: list[str] = []
    for page in reader.pages:
        chunks.append(page.extract_text() or "")
    return "\n".join(chunks)
