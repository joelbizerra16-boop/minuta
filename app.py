from pathlib import Path
from datetime import datetime
import base64
from io import BytesIO
import json
import re
import hashlib
import textwrap
import unicodedata
import xml.etree.ElementTree as ET

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
FIXED_LOGO_PATH = BASE_DIR / "baixados.png"
WINDOWS_FONT_DIR = Path("C:/Windows/Fonts")
DATA_DIR = BASE_DIR / "data"
PRACAS_JSON_PATH = DATA_DIR / "pracas.json"
XMLS_PROCESSADOS_JSON_PATH = DATA_DIR / "xmls_processados.json"
NFE_NAMESPACE = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
DISPLAY_PROCESSING_WARNINGS = False
NF_DEBUG_COLUMNS = ["NF Planilha", "NF XML", "Tipo XML", "Arquivo XML", "Correspondencia"]
TABLE_COLUMNS = [
    "Seq",
    "NF",
    "cProd",
    "Descricao",
    "Qtd",
    "Unidade",
    "Peso",
    "Destinatario",
    "ROTA",
    "Status",
]
UNDEFINED_ROUTE_LABEL = "NÃO DEFINIDA"
PDF_FONT_REGULAR = "Helvetica"
PDF_FONT_BOLD = "Helvetica-Bold"
PDF_FONT_MONO = "Courier"
PDF_FONT_MONO_BOLD = "Courier-Bold"
LOGIN_USERNAME = "minuta"
LOGIN_PASSWORD = "minuta123"
AUTH_QUERY_PARAM = "auth"
AUTH_QUERY_VALUE = "1"
ICON_MAP = {
    "dados_gerais": "folder",
    "filial": "building",
    "carregamento": "truck",
    "data_saida": "calendar",
    "motorista": "user_badge",
    "placa": "car",
    "resumo_carga": "chart",
    "nf": "receipt",
    "peso": "scale",
    "itens": "box",
    "erros": "alert",
    "xml": "file",
    "excel": "sheet",
    "processar": "play",
    "print": "printer",
}

ICON_SVG = {
    "folder": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M2.5 5.5a1.5 1.5 0 0 1 1.5-1.5h3l1.4 1.8H16a1.5 1.5 0 0 1 1.5 1.5v6.7A1.5 1.5 0 0 1 16 15.5H4A1.5 1.5 0 0 1 2.5 14z"/></svg>',
    "building": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4 16V4.5A1.5 1.5 0 0 1 5.5 3h7A1.5 1.5 0 0 1 14 4.5V16M7 6.5h1.5M10.5 6.5H12M7 9.5h1.5M10.5 9.5H12M7 12.5h1.5M10.5 12.5H12M3 16.5h14"/></svg>',
    "truck": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M2.5 6.5h8v5h3l1.7-2.2h2.3v2.2h-1M5.5 14.5a1.5 1.5 0 1 0 0 .01M14.5 14.5a1.5 1.5 0 1 0 0 .01M2.5 8.5v4h1.5M17.5 11.5v1h-1.5"/></svg>',
    "calendar": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M5 3.5v2M15 3.5v2M3.5 7h13M4.5 5h11A1.5 1.5 0 0 1 17 6.5v8A1.5 1.5 0 0 1 15.5 16h-11A1.5 1.5 0 0 1 3 14.5v-8A1.5 1.5 0 0 1 4.5 5zM6.5 9.5h2M10.5 9.5h2M6.5 12.5h2M10.5 12.5h2"/></svg>',
    "user_badge": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 10a2.75 2.75 0 1 0 0-5.5A2.75 2.75 0 0 0 10 10zm-4.5 5a4.5 4.5 0 0 1 9 0M14.5 5.5h2.5M15.75 4.25v2.5"/></svg>',
    "car": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M5.2 7.5 6.5 5h7l1.3 2.5M4.5 8.5h11A1.5 1.5 0 0 1 17 10v3h-1M5 13h10M6 14.5a1.25 1.25 0 1 0 0 .01M14 14.5a1.25 1.25 0 1 0 0 .01M3 13H2.5v-2A2.5 2.5 0 0 1 5 8.5"/></svg>',
    "chart": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4 15.5V9.5M10 15.5V5.5M16 15.5V11.5M3 16.5h14"/></svg>',
    "receipt": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M6 3.5h8A1.5 1.5 0 0 1 15.5 5v11l-1.75-1-1.75 1-1.75-1-1.75 1-1.75-1-1.75 1V5A1.5 1.5 0 0 1 6 3.5zM7 7h6M7 10h6M7 13h4"/></svg>',
    "scale": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 4v10.5M6.5 6h7M4 8.5h5l-2.5 4.5L4 8.5zm7 0h5l-2.5 4.5L11 8.5zM6 16.5h8"/></svg>',
    "box": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 3.5 16 6.5 10 9.5 4 6.5 10 3.5zm6 3v7L10 16.5l-6-3v-7M10 9.5v7"/></svg>',
    "alert": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 4.5 16 15.5H4L10 4.5zm0 4v3.5M10 13.75h.01"/></svg>',
    "file": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M6 3.5h5l3 3V15a1.5 1.5 0 0 1-1.5 1.5h-6A1.5 1.5 0 0 1 5 15V5A1.5 1.5 0 0 1 6.5 3.5zM11 3.5V7h3"/></svg>',
    "sheet": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M5 4.5h10A1.5 1.5 0 0 1 16.5 6v8A1.5 1.5 0 0 1 15 15.5H5A1.5 1.5 0 0 1 3.5 14V6A1.5 1.5 0 0 1 5 4.5zm0 3h10M8 4.5v11M12 4.5v11"/></svg>',
    "play": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M7 5.5v9l7-4.5-7-4.5z"/></svg>',
    "printer": '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M6 7V4.5h8V7M6.5 15.5h7A1.5 1.5 0 0 0 15 14v-3H5v3a1.5 1.5 0 0 0 1.5 1.5zM5 8h10a1.5 1.5 0 0 1 1.5 1.5V12H15M6.5 13h7"/></svg>',
}


def register_pdf_fonts() -> tuple[str, str]:
    regular_font = WINDOWS_FONT_DIR / "arial.ttf"
    bold_font = WINDOWS_FONT_DIR / "arialbd.ttf"

    if regular_font.is_file() and bold_font.is_file():
        if "Arial" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("Arial", str(regular_font)))
        if "Arial-Bold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("Arial-Bold", str(bold_font)))
        return "Arial", "Arial-Bold"

    return PDF_FONT_REGULAR, PDF_FONT_BOLD


def get_logo_path() -> Path | None:
    return FIXED_LOGO_PATH if FIXED_LOGO_PATH.is_file() else None


def get_logo_data_uri() -> str:
    logo_path = get_logo_path()
    if logo_path is None:
        return ""

    suffix = logo_path.suffix.lower()
    mime_type = "image/png"
    if suffix == ".jpg" or suffix == ".jpeg":
        mime_type = "image/jpeg"
    elif suffix == ".webp":
        mime_type = "image/webp"

    encoded_logo = base64.b64encode(logo_path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded_logo}"


def render_floating_logo() -> None:
    logo_data_uri = get_logo_data_uri()
    if not logo_data_uri:
        return

    st.markdown(
        f"""
    <style>
    .floating-company-logo {{
        position: fixed;
        top: 14px;
        left: 18px;
        z-index: 1000;
        pointer-events: none;
    }}
    .floating-company-logo img {{
        width: 120px;
        max-width: 14vw;
        height: auto;
        display: block;
    }}
    @media (max-width: 768px) {{
        .floating-company-logo {{
            top: 10px;
            left: 12px;
        }}
        .floating-company-logo img {{
            width: 100px;
            max-width: 28vw;
        }}
    }}
    </style>
    <div class="floating-company-logo">
        <img src="{logo_data_uri}" alt="Logo da empresa">
    </div>
    """,
        unsafe_allow_html=True,
    )


def update_pracas_json(uploaded_file) -> str:
    try:
        uploaded_file.seek(0)
        pracas_df = pd.read_excel(uploaded_file)
        uploaded_file.seek(0)
    except Exception as exc:
        raise ValueError(f"Erro ao ler o arquivo de pracas: {exc}") from exc

    json_content = pracas_df.to_json(orient="records", force_ascii=False, date_format="iso")
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if PRACAS_JSON_PATH.exists():
        PRACAS_JSON_PATH.unlink()

    PRACAS_JSON_PATH.write_text(json_content, encoding="utf-8")
    load_pracas_lookup.clear()
    return "Praças atualizadas com sucesso"


def normalize_praca_name(value: object) -> str:
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text).strip().upper()
    return text


@st.cache_data(show_spinner=False)
def load_pracas_lookup(json_path: str) -> dict[str, str]:
    path = Path(json_path)
    if not path.is_file():
        return {}

    try:
        pracas_df = pd.read_json(path)
    except ValueError:
        return {}

    if pracas_df.empty or "PRACA" not in pracas_df.columns or "ROTA" not in pracas_df.columns:
        return {}

    normalized_df = pracas_df[["PRACA", "ROTA"]].copy()
    normalized_df["PRACA"] = normalized_df["PRACA"].map(normalize_praca_name)
    normalized_df["ROTA"] = normalized_df["ROTA"].fillna(UNDEFINED_ROUTE_LABEL).astype(str).str.strip()
    normalized_df = normalized_df[normalized_df["PRACA"] != ""]
    normalized_df["ROTA"] = normalized_df["ROTA"].replace("", UNDEFINED_ROUTE_LABEL)

    return dict(zip(normalized_df["PRACA"], normalized_df["ROTA"]))


def apply_routes_to_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    updated_df = dataframe.copy()

    if "ROTA" not in updated_df.columns:
        updated_df["ROTA"] = UNDEFINED_ROUTE_LABEL

    if updated_df.empty:
        return updated_df

    if "Municipio" not in updated_df.columns:
        updated_df["ROTA"] = UNDEFINED_ROUTE_LABEL
        return updated_df

    route_lookup = load_pracas_lookup(str(PRACAS_JSON_PATH))
    if not route_lookup:
        updated_df["ROTA"] = UNDEFINED_ROUTE_LABEL
        return updated_df

    normalized_municipios = updated_df["Municipio"].map(normalize_praca_name)
    updated_df["ROTA"] = normalized_municipios.map(route_lookup).fillna(UNDEFINED_ROUTE_LABEL)
    return updated_df


def get_route_for_municipio(value: object) -> str:
    normalized = normalize_praca_name(value)
    if not normalized:
        return UNDEFINED_ROUTE_LABEL
    route_lookup = load_pracas_lookup(str(PRACAS_JSON_PATH))
    if not route_lookup:
        return UNDEFINED_ROUTE_LABEL
    return route_lookup.get(normalized, UNDEFINED_ROUTE_LABEL)


def render_label_icon(icon_name: str) -> str:
    return f'<span class="ui-icon" aria-hidden="true">{ICON_SVG[icon_name]}</span>'


def normalize_label(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^a-zA-Z0-9]+", "", text).lower()
    return text


def normalize_nf(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+[\.,]\d+", text):
        text = re.split(r"[\.,]", text, maxsplit=1)[0]
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    digits = digits.lstrip("0")
    return digits or "0"


def xml_local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def normalize_chave_nfe(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value).strip())
    return digits if len(digits) == 44 else ""


def extract_nf_from_chave(chave_nfe: str) -> str:
    if len(chave_nfe) != 44 or not chave_nfe.isdigit():
        return ""
    return normalize_nf(chave_nfe[25:34])


def detect_xml_type(root: ET.Element) -> str:
    root_name = xml_local_name(root.tag).lower()
    if "evento" in root_name:
        return "evento"
    if find_xml_text_by_localname(root, ["tpEvento", "descEvento", "chNFe"]) and not find_xml_text_by_localname(root, ["nNF"]):
        return "evento"
    return "normal"


def should_replace_xml(current_xml: dict[str, object], new_xml: dict[str, object]) -> bool:
    current_type = str(current_xml.get("TipoXML", "normal"))
    new_type = str(new_xml.get("TipoXML", "normal"))
    return current_type == "evento" and new_type == "normal"


def parse_float(value: object) -> float:
    if value is None or pd.isna(value):
        return 0.0
    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def find_column(columns: list[object], aliases: list[str]) -> str | None:
    normalized_columns = {normalize_label(column): str(column) for column in columns}
    normalized_aliases = [normalize_label(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in normalized_columns:
            return normalized_columns[alias]

    for normalized_column, original_column in normalized_columns.items():
        if any(alias in normalized_column or normalized_column in alias for alias in normalized_aliases):
            return original_column

    return None


def format_date_series(series: pd.Series) -> pd.Series:
    original_values = series.fillna("").astype(str).str.strip()
    parsed_dates = pd.to_datetime(series, errors="coerce", dayfirst=True)
    formatted = parsed_dates.dt.strftime("%d/%m/%Y")
    return formatted.where(parsed_dates.notna(), original_values)


def format_single_date(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""

    iso_match = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
    if iso_match:
        year, month, day = iso_match.groups()
        return f"{day}/{month}/{year}"

    br_match = re.search(r"(\d{2})/(\d{2})/(\d{4})", text)
    if br_match:
        return br_match.group(0)

    return format_date_series(pd.Series([text])).iloc[0]


def find_xml_text_by_localname(node: ET.Element, local_names: list[str]) -> str:
    for local_name in local_names:
        found = node.find(f".//{{*}}{local_name}")
        if found is not None and found.text:
            text = found.text.strip()
            if text:
                return text
    return ""


def format_product_description(descricao: object, codigo: object, markdown_bold: bool = False) -> str:
    descricao_text = str(descricao or "").strip()
    codigo_text = str(codigo or "").strip()

    if not codigo_text:
        return descricao_text

    codigo_display = f"**{codigo_text}**" if markdown_bold else codigo_text
    if descricao_text:
        return f"{descricao_text} - ({codigo_display})"
    return f"({codigo_display})"


def has_formatted_product_code(value: object) -> bool:
    text = str(value or "").strip()
    if not text:
        return False
    normalized_text = " ".join(text.split())
    return bool(re.search(r" - \([^()]+\)$", normalized_text))


def format_datetime_display(value: datetime | None = None) -> str:
    return (value or datetime.now()).strftime("%d/%m/%Y %H:%M:%S")


def format_decimal_br(value: object, decimals: int = 2) -> str:
    number = parse_float(value)
    formatted = f"{number:,.{decimals}f}"
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def format_quantity_display(value: object) -> str:
    number = parse_float(value)
    if number.is_integer():
        return str(int(number))
    formatted = f"{number:,.3f}".rstrip("0").rstrip(".")
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def sanitize_filename_part(value: object, default: str) -> str:
    text = str(value or "").strip()
    sanitized = re.sub(r"[^a-zA-Z0-9_-]+", "_", text).strip("_")
    return sanitized or default


def summarize_metadata(base_df: pd.DataFrame) -> dict[str, str]:
    data_saida_values = [value for value in base_df["Data Saida"].astype(str).tolist() if value]
    motorista_values = [value for value in base_df["Motorista"].astype(str).tolist() if value]
    placa_values = [value for value in base_df["Placa"].astype(str).tolist() if value]
    carga_values = [value for value in base_df["Numero Carga"].astype(str).tolist() if value]

    unique_dates = sorted(set(data_saida_values))
    unique_motoristas = sorted(set(motorista_values))
    unique_placas = sorted(set(placa_values))
    unique_cargas = sorted(set(carga_values))

    return {
        "numero_carga": unique_cargas[0] if len(unique_cargas) == 1 else ("Multiplos" if unique_cargas else "--"),
        "data_saida": unique_dates[0] if len(unique_dates) == 1 else ("Multiplas" if unique_dates else "--"),
        "motorista": unique_motoristas[0] if len(unique_motoristas) == 1 else ("Multiplos" if unique_motoristas else "--"),
        "placa": unique_placas[0] if len(unique_placas) == 1 else ("Multiplas" if unique_placas else "--"),
    }


def detect_excel_structure(uploaded_excel) -> tuple[str, int | None, int | None]:
    workbook = pd.ExcelFile(uploaded_excel)
    uploaded_excel.seek(0)

    overview_tokens = {"filial", "dtsaida", "data", "carga", "carregamento", "numerocarga", "motorista", "veiculo", "placa"}
    detail_tokens = {"seqent", "numeronota", "carregamento", "numeropedido", "pesokg"}
    best_sheet = workbook.sheet_names[0]
    best_overview_row = None
    best_detail_row = None
    best_overview_score = -1
    best_detail_score = -1

    for sheet_name in workbook.sheet_names:
        preview_df = pd.read_excel(workbook, sheet_name=sheet_name, header=None, nrows=20)
        for row_index in range(len(preview_df.index)):
            normalized_values = {
                normalize_label(value)
                for value in preview_df.iloc[row_index].tolist()
                if str(value).strip() and str(value).strip().lower() != "nan"
            }

            overview_score = len(overview_tokens.intersection(normalized_values))
            detail_score = len(detail_tokens.intersection(normalized_values))

            if detail_score > best_detail_score:
                best_detail_score = detail_score
                best_sheet = sheet_name
                best_detail_row = row_index

            if overview_score > best_overview_score:
                best_overview_score = overview_score
                best_overview_row = row_index

    return best_sheet, best_detail_row, best_overview_row


def extract_summary_metadata(preview_df: pd.DataFrame, overview_row: int | None) -> dict[str, str]:
    default_metadata = {"Filial": "BRIDA", "Numero Carga": "", "Data Saida": "", "Motorista": "", "Placa": ""}
    if overview_row is None or overview_row + 1 >= len(preview_df.index):
        return default_metadata

    header_values = [normalize_label(value) for value in preview_df.iloc[overview_row].tolist()]
    data_values = preview_df.iloc[overview_row + 1].tolist()
    mapping = dict(zip(header_values, data_values))

    filial = str(mapping.get("filial", "BRIDA") or "BRIDA").strip()
    numero_carga = str(mapping.get("carregamento", mapping.get("carga", mapping.get("numerocarga", ""))) or "").strip()
    data_saida = format_date_series(pd.Series([mapping.get("dtsaida", mapping.get("data", ""))])).iloc[0]
    motorista = str(mapping.get("motorista", "") or "").strip()
    placa = str(mapping.get("veiculo", mapping.get("placa", "")) or "").strip()

    return {
        "Filial": filial,
        "Numero Carga": numero_carga,
        "Data Saida": data_saida,
        "Motorista": motorista,
        "Placa": placa,
    }


def build_metadata_df(base_df: pd.DataFrame, summary_metadata: dict[str, str]) -> pd.DataFrame:
    metadata_df = pd.DataFrame(index=base_df.index)

    seq_column = find_column(
        list(base_df.columns),
        ["Seq. Ent", "Seq Ent", "Sequencia Entrega", "Seq", "Sequencia", "Carga"],
    )

    if seq_column:
        metadata_df["Seq"] = base_df[seq_column]
    else:
        metadata_df["Seq"] = range(1, len(base_df.index) + 1)

    metadata_df["Data Saida"] = summary_metadata.get("Data Saida", "")
    metadata_df["Motorista"] = summary_metadata.get("Motorista", "")
    metadata_df["Placa"] = summary_metadata.get("Placa", "")
    metadata_df["Numero Carga"] = summary_metadata.get("Numero Carga", "")
    metadata_df["Filial"] = summary_metadata.get("Filial", "BRIDA")
    metadata_df["Seq"] = metadata_df["Seq"].fillna("")
    metadata_df["Seq_sort"] = pd.to_numeric(metadata_df["Seq"], errors="coerce")
    return metadata_df


def summarize_filial(base_df: pd.DataFrame) -> str:
    if "Filial" not in base_df.columns:
        return "BRIDA"

    filiais = [value for value in base_df["Filial"].astype(str).tolist() if value]
    unique_filiais = sorted(set(filiais))
    if len(unique_filiais) == 1:
        return unique_filiais[0]
    if unique_filiais:
        return "Multiplas"
    return "BRIDA"


def load_excel_base(uploaded_excel) -> pd.DataFrame:
    try:
        sheet_name, detail_header_row, overview_row = detect_excel_structure(uploaded_excel)
        preview_df = pd.read_excel(uploaded_excel, sheet_name=sheet_name, header=None, nrows=20)
        uploaded_excel.seek(0)

        if detail_header_row is None:
            raise ValueError("Nao foi possivel localizar a tabela detalhada de notas no Excel.")

        base_df = pd.read_excel(uploaded_excel, sheet_name=sheet_name, header=detail_header_row)
        uploaded_excel.seek(0)
    except Exception as exc:
        raise ValueError(f"Nao foi possivel ler o Excel enviado: {exc}") from exc

    if base_df.empty:
        raise ValueError("O Excel enviado esta vazio.")

    summary_metadata = extract_summary_metadata(preview_df, overview_row)
    metadata_df = build_metadata_df(base_df, summary_metadata)
    nf_column = find_column(list(base_df.columns), ["NF", "Nota Fiscal", "Numero NF", "Numero Nota", "N. NF", "Nf"])

    if nf_column:
        filtered_df = metadata_df.copy()
        filtered_df["NF"] = base_df[nf_column].apply(normalize_nf)
        filtered_df["nf_normalizada"] = filtered_df["NF"]
        filtered_df = filtered_df[filtered_df["NF"] != ""].copy()

        if filtered_df.empty:
            raise ValueError("Nenhuma NF valida foi encontrada no Excel.")

        filtered_df.attrs["integration_mode"] = "excel_nf"
        filtered_df.attrs["issues"] = []
        return filtered_df

    available = ", ".join(str(column) for column in base_df.columns)
    metadata_df = metadata_df.dropna(how="all").copy()
    metadata_df.attrs["integration_mode"] = "xml_base"
    metadata_df.attrs["issues"] = [
        "A planilha enviada nao possui coluna NF. Foram considerados todos os XMLs enviados como base.",
        f"Colunas disponiveis no Excel: {available}",
    ]
    return metadata_df


def xml_text(node: ET.Element, path: str, default: str = "") -> str:
    found = node.find(path, NFE_NAMESPACE)
    if found is None or found.text is None:
        return default
    return found.text.strip()


def xml_text_any_namespace(node: ET.Element, path: str, default: str = "") -> str:
    found = node.find(path)
    if found is None or found.text is None:
        return default
    text = found.text.strip()
    return text or default


def fallback_nf_from_filename(filename: str) -> str:
    digits = re.findall(r"\d+", filename)
    return digits[-1] if digits else ""


def extract_issue_date_from_xml(root: ET.Element) -> str:
    issue_date = find_xml_text_by_localname(root, ["dhEmi", "dEmi"])
    if not issue_date:
        issue_date = xml_text(root, ".//nfe:ide/nfe:dhEmi") or xml_text(root, ".//nfe:ide/nfe:dEmi")
    return format_single_date(issue_date)


def parse_xml_file(uploaded_xml) -> dict[str, object]:
    filename = getattr(uploaded_xml, "name", "arquivo.xml")

    try:
        root = ET.fromstring(uploaded_xml.getvalue())
    except Exception as exc:
        return {
            "NF": fallback_nf_from_filename(filename),
            "ChaveNFe": "",
            "Destinatario": "",
            "Municipio": "",
            "Status": f"Erro ao ler XML: {exc}",
            "PesoTotal": 0.0,
            "Items": [],
            "Arquivo": filename,
            "Erro": True,
            "TipoXML": "desconhecido",
            "nf_normalizada": normalize_nf(fallback_nf_from_filename(filename)),
        }

    xml_type = detect_xml_type(root)
    ch_nfe = normalize_chave_nfe(find_xml_text_by_localname(root, ["chNFe"]))
    nf = normalize_nf(find_xml_text_by_localname(root, ["nNF"])) or extract_nf_from_chave(ch_nfe)
    emitente = xml_text_any_namespace(root, ".//{*}emit/{*}xNome")
    destinatario = xml_text_any_namespace(root, ".//{*}dest/{*}xNome", "DESTINATARIO NAO INFORMADO")
    if emitente and normalize_label(destinatario) == normalize_label(emitente):
        destinatario = "ERRO: DESTINATARIO INCORRETO"
    municipio = xml_text_any_namespace(root, ".//{*}dest/{*}enderDest/{*}xMun") or find_xml_text_by_localname(root, ["xMun"])
    status = find_xml_text_by_localname(root, ["xMotivo"]) or "Status nao informado"
    data_emissao = extract_issue_date_from_xml(root)

    volume_total = 0.0
    peso_total = 0.0
    for volume_node in root.findall(".//nfe:transp/nfe:vol", NFE_NAMESPACE):
        volume_total += parse_float(xml_text(volume_node, "./nfe:qVol", "0"))
        peso_total += parse_float(xml_text(volume_node, "./nfe:pesoL", "0"))

    raw_items = []
    total_quantity = 0.0

    for det in root.findall(".//{*}det"):
        quantity = parse_float(find_xml_text_by_localname(det, ["qCom"]) or "0")
        raw_items.append(
            {
                "cProd": find_xml_text_by_localname(det, ["cProd"]),
                "Descricao": find_xml_text_by_localname(det, ["xProd"]),
                "Qtd": quantity,
                "Unidade": find_xml_text_by_localname(det, ["uCom"]),
            }
        )
        total_quantity += quantity

    peso_unitario = peso_total / total_quantity if total_quantity > 0 else 0.0
    if volume_total <= 0 and total_quantity > 0:
        volume_total = total_quantity
    items = []
    for item in raw_items:
        items.append({**item, "Peso": peso_unitario * item["Qtd"]})

    return {
        "NF": nf or fallback_nf_from_filename(filename),
        "nf_normalizada": nf or normalize_nf(fallback_nf_from_filename(filename)),
        "ChaveNFe": ch_nfe,
        "Data": data_emissao,
        "Destinatario": destinatario,
        "Municipio": municipio,
        "Status": status,
        "VolumeTotal": volume_total,
        "PesoTotal": peso_total,
        "Items": items,
        "Arquivo": filename,
        "Erro": False,
        "TipoXML": xml_type,
    }


def build_minuta_records(dataframe: pd.DataFrame) -> list[dict[str, object]]:
    if dataframe.empty:
        return []

    minuta_records: list[dict[str, object]] = []
    grouped_df = dataframe.groupby("NF", sort=False, dropna=False)

    for nf, group in grouped_df:
        first_row = group.iloc[0]
        produtos_df = group[
            group["Descricao"].astype(str).str.strip().ne("")
            | group["cProd"].astype(str).str.strip().ne("")
        ]

        produtos = [
            {
                "descricao": str(row["Descricao"] or "").strip(),
                "codigo": str(row["cProd"] or "").strip(),
                "qtd": parse_float(row["Qtd"]),
                "un": str(row["Unidade"] or "").strip(),
                "peso": parse_float(row["Peso"]),
            }
            for _, row in produtos_df.iterrows()
        ]

        data_emissao = str(first_row.get("Data", "") or "").strip()
        cliente = str(first_row.get("Destinatario", "") or "").strip()
        volume = parse_float(first_row.get("Volume", 0.0))
        if volume <= 0:
            volume = parse_float(group["Qtd"].sum())
        peso_total = parse_float(first_row.get("PesoTotalNF", group["Peso"].sum()))

        minuta_records.append(
            {
                "nf": str(nf),
                "data": data_emissao,
                "cliente": cliente,
                "rota": str(first_row.get("ROTA", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL).strip(),
                "volume": int(volume) if volume.is_integer() else volume,
                "peso_total": peso_total,
                "produtos": produtos,
            }
        )

    return minuta_records


def generate_minuta_pdf(
    dados_minuta: list[dict[str, object]],
    numero_carga: str,
    data_emissao: str,
    veiculo: str,
    motorista: str,
) -> bytes:
    regular_font, bold_font = register_pdf_fonts()
    mono_font = PDF_FONT_MONO
    mono_bold_font = PDF_FONT_MONO_BOLD
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    page_width, page_height = A4
    left_margin = 40
    right_margin = page_width - 40
    top_margin = page_height - 45
    bottom_margin = 55
    line_height = 12
    nf_row_padding = 8
    product_row_padding = 6
    route_block_height = 20

    table_columns = {
        "nota": {"x": left_margin, "width": 72},
        "emissao": {"x": left_margin + 74, "width": 78},
        "cliente": {"x": left_margin + 154, "width": 228},
        "vol": {"x": left_margin + 386, "width": 48},
        "peso": {"x": left_margin + 438, "width": 78},
    }

    product_columns = {
        "descricao": {"x": left_margin + 34, "width": 332},
        "qtd": {"x": left_margin + 374, "width": 58},
        "un": {"x": left_margin + 438, "width": 34},
        "peso": {"x": left_margin + 478, "width": 78},
    }

    def wrap_text(text: object, font_name: str, font_size: int, width: float) -> list[str]:
        lines = simpleSplit(str(text or "--"), font_name, font_size, width)
        return lines or ["--"]

    def draw_wrapped_text(x_pos: float, y_top: float, lines: list[str], font_name: str, font_size: int) -> None:
        pdf.setFont(font_name, font_size)
        text_y = y_top
        for line in lines:
            pdf.drawString(x_pos, text_y, line)
            text_y -= line_height

    def wrap_product_description(produto: dict[str, object]) -> list[str]:
        descricao = format_product_description(produto.get("descricao", ""), produto.get("codigo", ""))
        return wrap_text(descricao, mono_font, 10, product_columns["descricao"]["width"] - 14)

    def draw_product_description(x_pos: float, y_top: float, produto: dict[str, object]) -> None:
        descricao = format_product_description(produto.get("descricao", ""), produto.get("codigo", ""))
        descricao_lines = wrap_text(descricao, mono_font, 10, product_columns["descricao"]["width"] - 14)
        codigo = str(produto.get("codigo", "") or "").strip()

        for index, line in enumerate(descricao_lines):
            text_y = y_top - (index * line_height)
            prefix = "• " if index == 0 else "  "
            pdf.setFont(mono_font, 10)
            pdf.drawString(x_pos, text_y, prefix)

            current_x = x_pos + pdf.stringWidth(prefix, mono_font, 10)
            if codigo and codigo in line:
                before, _, after = line.rpartition(codigo)
                if before:
                    pdf.setFont(mono_font, 10)
                    pdf.drawString(current_x, text_y, before)
                    current_x += pdf.stringWidth(before, mono_font, 10)

                pdf.setFont(mono_bold_font, 10)
                pdf.drawString(current_x, text_y, codigo)
                current_x += pdf.stringWidth(codigo, mono_bold_font, 10)

                if after:
                    pdf.setFont(mono_font, 10)
                    pdf.drawString(current_x, text_y, after)
            else:
                pdf.setFont(mono_font, 10)
                pdf.drawString(current_x, text_y, line)

    def draw_right_aligned(
        x_pos: float,
        width: float,
        y_pos: float,
        text: object,
        font_name: str,
        font_size: int,
        padding_right: float = 8,
    ) -> None:
        pdf.setFont(font_name, font_size)
        pdf.drawRightString(x_pos + width - padding_right, y_pos, str(text or ""))

    def draw_centered(x_pos: float, width: float, y_pos: float, text: object, font_name: str, font_size: int) -> None:
        pdf.setFont(font_name, font_size)
        pdf.drawCentredString(x_pos + (width / 2), y_pos, str(text or ""))

    def draw_page_title(y_pos: float, continuation: bool = False) -> float:
        pdf.setFont(bold_font, 20 if not continuation else 14)
        title = "MINUTA DE CARREGAMENTO"
        if continuation:
            title = f"{title} - CONTINUACAO"
        pdf.drawCentredString(page_width / 2, y_pos, title)
        return y_pos - 24

    def draw_first_page_header() -> float:
        y_pos = top_margin
        y_pos = draw_page_title(y_pos)

        pdf.setFont(bold_font, 15)
        pdf.drawString(left_margin, y_pos, "BRIDA LUBRIFICANTES LTDA")
        y_pos -= 24

        pdf.setFont(regular_font, 11)
        pdf.drawString(left_margin, y_pos, f"Carregamento:   {numero_carga or '--'}")
        y_pos -= 18
        pdf.drawString(left_margin, y_pos, f"Emissao:   {data_emissao or '--'}")
        y_pos -= 16

        pdf.setStrokeColor(colors.HexColor("#b8b8b8"))
        pdf.line(left_margin, y_pos, right_margin, y_pos)
        y_pos -= 24

        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, y_pos, "TRANSPORTADORA:")
        pdf.setFont(regular_font, 11)
        pdf.drawString(left_margin + 118, y_pos, "BRIDA LUBRIFICANTES LTDA")
        y_pos -= 18

        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, y_pos, "VEICULO:")
        pdf.setFont(regular_font, 11)
        pdf.drawString(left_margin + 70, y_pos, veiculo or "--")
        y_pos -= 18

        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, y_pos, "MOTORISTA:")
        pdf.setFont(regular_font, 11)
        pdf.drawString(left_margin + 85, y_pos, motorista or "--")
        y_pos -= 18

        pdf.setDash(4, 3)
        pdf.line(left_margin, y_pos, right_margin, y_pos)
        pdf.setDash()
        return y_pos - 18

    def draw_continuation_header() -> float:
        y_pos = top_margin
        y_pos = draw_page_title(y_pos, continuation=True)
        pdf.setStrokeColor(colors.HexColor("#d0d0d0"))
        pdf.line(left_margin, y_pos, right_margin, y_pos)
        return y_pos - 18

    def draw_main_table_header(y_pos: float) -> float:
        header_height = 20
        pdf.setFillColor(colors.HexColor("#efefef"))
        pdf.roundRect(left_margin, y_pos - header_height + 4, right_margin - left_margin, header_height, 4, fill=1, stroke=0)
        pdf.setFillColor(colors.black)
        pdf.setFont(mono_bold_font, 10)
        pdf.drawString(table_columns["nota"]["x"] + 8, y_pos - 10, "Nota")
        pdf.drawString(table_columns["emissao"]["x"] + 8, y_pos - 10, "Emissao")
        pdf.drawString(table_columns["cliente"]["x"] + 8, y_pos - 10, "Cliente")
        pdf.drawRightString(table_columns["vol"]["x"] + table_columns["vol"]["width"] - 12, y_pos - 10, "Vol")
        pdf.drawRightString(table_columns["peso"]["x"] + table_columns["peso"]["width"] - 12, y_pos - 10, "Peso")
        return y_pos - 28

    def start_new_page() -> float:
        pdf.showPage()
        return draw_continuation_header()

    def ensure_space(current_y: float, required_height: float) -> float:
        if current_y - required_height < bottom_margin:
            return start_new_page()
        return current_y

    def compute_block_height(registro: dict[str, object]) -> float:
        cliente_lines = wrap_text(registro.get("cliente", ""), mono_font, 10, table_columns["cliente"]["width"] - 8)
        nf_row_height = max(18, len(cliente_lines) * line_height) + nf_row_padding
        block_height = route_block_height + nf_row_height + 18

        produtos = registro.get("produtos", []) or []
        if not produtos:
            return block_height + 10

        for produto in produtos:
            descricao_lines = wrap_product_description(produto)
            block_height += max(14, len(descricao_lines) * line_height) + product_row_padding

        return block_height + 8

    current_y = draw_first_page_header()
    current_y = draw_main_table_header(current_y)

    for index, registro in enumerate(dados_minuta):
        current_y = ensure_space(current_y, compute_block_height(registro) + 20)

        cliente_lines = wrap_text(registro.get("cliente", ""), mono_font, 10, table_columns["cliente"]["width"] - 8)
        nf_row_height = max(18, len(cliente_lines) * line_height) + nf_row_padding
        route_text = f"ROTA: {str(registro.get('rota', UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL).strip().upper()}"

        if index > 0:
            pdf.setStrokeColor(colors.HexColor("#444444"))
            pdf.setLineWidth(1.8)
            pdf.line(left_margin, current_y + 6, right_margin, current_y + 6)
            pdf.setLineWidth(1)

        pdf.setFillColor(colors.HexColor("#1F3A5F"))
        pdf.setFont(bold_font, 10)
        pdf.drawString(left_margin + 2, current_y - 8, route_text)

        row_top = current_y - route_block_height - 10
        pdf.setFillColor(colors.black)
        pdf.setFont(mono_font, 10)
        pdf.drawString(table_columns["nota"]["x"] + 2, row_top, str(registro.get("nf", "") or "--"))
        draw_centered(
            table_columns["emissao"]["x"],
            table_columns["emissao"]["width"],
            row_top,
            str(registro.get("data", "") or ""),
            mono_font,
            10,
        )
        draw_wrapped_text(table_columns["cliente"]["x"] + 2, row_top, cliente_lines, mono_font, 10)
        volume = registro.get("volume", 0)
        volume_value = parse_float(volume)
        volume_text = str(int(volume_value)) if volume_value.is_integer() else format_quantity_display(volume_value)
        draw_right_aligned(table_columns["vol"]["x"], table_columns["vol"]["width"], row_top, volume_text, mono_font, 10, padding_right=10)
        draw_right_aligned(
            table_columns["peso"]["x"],
            table_columns["peso"]["width"],
            row_top,
            format_decimal_br(registro.get("peso_total", 0.0)),
            mono_font,
            10,
            padding_right=12,
        )
        current_y -= route_block_height + nf_row_height

        pdf.setFont(bold_font, 10)
        pdf.drawString(left_margin + 18, current_y - 8, "• Produtos:")
        current_y -= 18

        produtos = registro.get("produtos", []) or []
        if not produtos:
            pdf.setFont(regular_font, 10)
            pdf.drawString(left_margin + 36, current_y - 8, "Sem produtos detalhados")
            current_y -= 18
        else:
            for produto in produtos:
                descricao_lines = wrap_product_description(produto)
                product_height = max(14, len(descricao_lines) * line_height) + product_row_padding
                current_y = ensure_space(current_y, product_height + 12)

                row_top = current_y - 8
                draw_product_description(product_columns["descricao"]["x"], row_top, produto)
                draw_right_aligned(
                    product_columns["qtd"]["x"],
                    product_columns["qtd"]["width"],
                    row_top,
                    format_quantity_display(produto.get("qtd", 0)),
                    mono_font,
                    10,
                    padding_right=8,
                )
                draw_centered(
                    product_columns["un"]["x"],
                    product_columns["un"]["width"],
                    row_top,
                    str(produto.get("un", "") or "--"),
                    mono_font,
                    10,
                )
                draw_right_aligned(
                    product_columns["peso"]["x"],
                    product_columns["peso"]["width"],
                    row_top,
                    format_decimal_br(produto.get("peso", 0.0)),
                    mono_font,
                    10,
                    padding_right=12,
                )
                current_y -= product_height

        current_y -= 6

    total_volume = sum(parse_float(registro.get("volume", 0)) for registro in dados_minuta)
    total_nf = len({str(registro.get("nf", "")).strip() for registro in dados_minuta if str(registro.get("nf", "")).strip()})
    total_peso = sum(parse_float(registro.get("peso_total", 0.0)) for registro in dados_minuta)
    total_block_height = 50
    signature_block_height = 90
    current_y = ensure_space(current_y, total_block_height + signature_block_height)

    pdf.setStrokeColor(colors.HexColor("#b8b8b8"))
    pdf.line(left_margin, current_y, right_margin, current_y)
    pdf.setFont(bold_font, 11)
    pdf.drawString(left_margin, current_y - 18, "TOTAL GERAL:")
    pdf.setFont(bold_font, 10)
    pdf.drawString(left_margin + 18, current_y - 36, f"Volumes: {format_quantity_display(total_volume)}")
    pdf.drawString(left_margin + 190, current_y - 36, f"NF: {total_nf}")
    pdf.drawString(left_margin + 320, current_y - 36, f"Peso: {format_decimal_br(total_peso)}")

    current_y -= total_block_height

    signature_y = max(bottom_margin + 32, current_y - 34)
    pdf.setStrokeColor(colors.HexColor("#6a6a6a"))
    pdf.line(page_width / 2 - 120, signature_y, page_width / 2 + 120, signature_y)
    pdf.setFont(regular_font, 12)
    pdf.drawCentredString(page_width / 2, signature_y - 18, "Ass. do conferente")

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


def build_xml_index(xml_files: list) -> tuple[dict[str, dict[str, object]], list[str]]:
    xml_index: dict[str, dict[str, object]] = {}
    issues: list[str] = []

    for xml_file in xml_files:
        xml_data = parse_xml_file(xml_file)
        nf = normalize_nf(xml_data.get("nf_normalizada", "") or xml_data.get("NF", ""))

        if not nf:
            issues.append(f"XML sem NF identificavel: {xml_data.get('Arquivo', 'arquivo.xml')}")
            continue

        if nf in xml_index:
            current_xml = xml_index[nf]
            if should_replace_xml(current_xml, xml_data):
                issues.append(f"NF {nf} encontrada em XML de evento e XML normal. Foi mantido o XML normal.")
                xml_index[nf] = xml_data
                continue
            if should_replace_xml(xml_data, current_xml):
                issues.append(f"NF {nf} encontrada em XML normal e XML de evento. Foi mantido o XML normal.")
                continue
            issues.append(f"NF {nf} duplicada nos XMLs. Foi mantido o ultimo arquivo enviado.")

        if xml_data.get("Erro"):
            issues.append(f"Erro no XML {xml_data.get('Arquivo', 'arquivo.xml')}: {xml_data.get('Status', '')}")
            continue

        xml_index[nf] = xml_data

    return xml_index, issues


def serialize_xml_record(xml_data: dict[str, object]) -> dict[str, object]:
    items = []
    for item in xml_data.get("Items", []) or []:
        items.append(
            {
                "cProd": str(item.get("cProd", "") or "").strip(),
                "Descricao": str(item.get("Descricao", "") or "").strip(),
                "Qtd": parse_float(item.get("Qtd", 0.0)),
                "Unidade": str(item.get("Unidade", "") or "").strip(),
                "Peso": parse_float(item.get("Peso", 0.0)),
            }
        )

    municipio = str(xml_data.get("Municipio", "") or "").strip()
    return {
        "NF": normalize_nf(xml_data.get("NF", "") or xml_data.get("nf_normalizada", "")),
        "nf_normalizada": normalize_nf(xml_data.get("nf_normalizada", "") or xml_data.get("NF", "")),
        "ChaveNFe": normalize_chave_nfe(xml_data.get("ChaveNFe", "")),
        "Data": str(xml_data.get("Data", "") or "").strip(),
        "Destinatario": str(xml_data.get("Destinatario", "") or "").strip(),
        "Municipio": municipio,
        "Status": str(xml_data.get("Status", "") or "").strip(),
        "VolumeTotal": parse_float(xml_data.get("VolumeTotal", 0.0)),
        "PesoTotal": parse_float(xml_data.get("PesoTotal", 0.0)),
        "Items": items,
        "Arquivo": str(xml_data.get("Arquivo", "") or "").strip(),
        "Erro": bool(xml_data.get("Erro", False)),
        "TipoXML": str(xml_data.get("TipoXML", "normal") or "normal").strip(),
        "ROTA": str(xml_data.get("ROTA", "") or get_route_for_municipio(municipio)).strip(),
    }


def build_xml_index_from_records(xml_records: list[dict[str, object]]) -> tuple[dict[str, dict[str, object]], list[str]]:
    xml_index: dict[str, dict[str, object]] = {}
    issues: list[str] = []

    for xml_record in xml_records or []:
        normalized_record = serialize_xml_record(xml_record)
        nf = normalize_nf(normalized_record.get("nf_normalizada", "") or normalized_record.get("NF", ""))

        if not nf:
            issues.append("Registro salvo sem NF identificavel foi ignorado.")
            continue

        xml_index[nf] = normalized_record

    return xml_index, issues


def resolve_xml_source(xml_source: object) -> tuple[dict[str, dict[str, object]], list[str]]:
    if not xml_source:
        return {}, []

    if isinstance(xml_source, list) and xml_source:
        first_item = xml_source[0]
        if isinstance(first_item, dict):
            return build_xml_index_from_records(xml_source)
        return build_xml_index(xml_source)

    return {}, []


def salvar_xmls_processados_json(xml_files: list) -> tuple[int, int, list[str]]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    xml_index, issues = build_xml_index(xml_files or [])
    serialized_records = [serialize_xml_record(xml_data) for xml_data in xml_index.values()]

    XMLS_PROCESSADOS_JSON_PATH.write_text(
        json.dumps(serialized_records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    carregar_xmls_processados_json.clear()
    return len(xml_index), max(0, len(xml_files or []) - len(xml_index)), issues


@st.cache_data(show_spinner=False)
def carregar_xmls_processados_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    path = Path(json_path)
    if not path.is_file():
        return [], ""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        return [], f"Os XMLs salvos no sistema nao puderam ser lidos ({exc}). Envie novos arquivos para atualizar a base."

    if not isinstance(payload, list):
        return [], "Os XMLs salvos no sistema estao em formato invalido. Envie novos arquivos para atualizar a base."

    try:
        records = [serialize_xml_record(item) for item in payload if isinstance(item, dict)]
    except Exception as exc:
        return [], f"Os XMLs salvos no sistema estao corrompidos ({exc}). Envie novos arquivos para atualizar a base."

    return records, ""


def get_xml_storage_status() -> tuple[bool, str]:
    if not XMLS_PROCESSADOS_JSON_PATH.is_file():
        return False, ""

    updated_at = datetime.fromtimestamp(XMLS_PROCESSADOS_JSON_PATH.stat().st_mtime)
    return True, format_datetime_display(updated_at)


def integrate_excel_with_xml(base_df: pd.DataFrame, xml_source: object) -> tuple[pd.DataFrame, dict[str, object], list[str], list[dict[str, str]]]:
    xml_index, issues = resolve_xml_source(xml_source)
    issues.extend(base_df.attrs.get("issues", []))
    rows: list[dict[str, object]] = []
    debug_rows: list[dict[str, str]] = []
    integration_mode = base_df.attrs.get("integration_mode", "excel_nf")

    if integration_mode == "xml_base":
        if not xml_index:
            raise ValueError("A planilha nao possui NF e nenhum XML valido foi enviado para montar a base.")

        metadata_records = base_df.to_dict(orient="records") or [{"Seq": "", "Seq_sort": None, "Data Saida": "", "Motorista": "", "Filial": "BRIDA"}]
        xml_records = list(xml_index.values())

        if len(metadata_records) == len(xml_records):
            paired_records = zip(metadata_records, xml_records)
            issues.append("A planilha nao possui NF. Os XMLs foram associados pela ordem de envio.")
        else:
            shared_record = metadata_records[0]
            paired_records = ((shared_record, xml_data) for xml_data in xml_records)
            if len(metadata_records) > 1:
                issues.append("A planilha nao possui NF para vinculo exato. Foi usada apenas a primeira linha do Excel como referencia geral.")

        for index, (metadata_row, xml_data) in enumerate(paired_records, start=1):
            seq_value = metadata_row.get("Seq", "")
            seq_sort = metadata_row.get("Seq_sort")
            if pd.isna(seq_sort):
                seq_sort = index
            if seq_value in (None, ""):
                seq_value = seq_sort

            if not xml_data["Items"]:
                rows.append(
                    {
                        "Seq": seq_value,
                        "Seq_sort": seq_sort,
                        "ChaveNFe": xml_data["ChaveNFe"],
                        "NF": xml_data["NF"],
                        "Data": xml_data["Data"],
                        "cProd": "",
                        "Descricao": "",
                        "Qtd": 0.0,
                        "Unidade": "",
                        "Volume": xml_data["VolumeTotal"],
                        "Peso": 0.0,
                        "PesoTotalNF": xml_data["PesoTotal"],
                        "Destinatario": xml_data["Destinatario"],
                        "Municipio": xml_data["Municipio"],
                        "Status": str(xml_data["Status"]),
                    }
                )
                continue

            for item in xml_data["Items"]:
                rows.append(
                    {
                        "Seq": seq_value,
                        "Seq_sort": seq_sort,
                        "ChaveNFe": xml_data["ChaveNFe"],
                        "NF": xml_data["NF"],
                        "Data": xml_data["Data"],
                        "cProd": item["cProd"],
                        "Descricao": item["Descricao"],
                        "Qtd": item["Qtd"],
                        "Unidade": item["Unidade"],
                        "Volume": xml_data["VolumeTotal"],
                        "Peso": item["Peso"],
                        "PesoTotalNF": xml_data["PesoTotal"],
                        "Destinatario": xml_data["Destinatario"],
                        "Municipio": xml_data["Municipio"],
                        "Status": str(xml_data["Status"]),
                    }
                )

        processed_df = pd.DataFrame(rows)
        if processed_df.empty:
            processed_df = create_empty_processed_df()
        else:
            processed_df = processed_df.sort_values(by=["Seq_sort", "NF"], ascending=[False, True], na_position="last")
        processed_df = apply_routes_to_dataframe(processed_df)

        display_df = processed_df[TABLE_COLUMNS].copy()
        error_mask = ~display_df["Status"].astype(str).str.contains("autorizado", case=False, na=False)
        item_mask = display_df["cProd"].astype(str).str.strip() != ""
        metadata = summarize_metadata(base_df)

        summary = {
            "filial": summarize_filial(base_df),
            "numero_carga": metadata["numero_carga"],
            "data_saida": metadata["data_saida"],
            "motorista": metadata["motorista"],
            "placa": metadata["placa"],
            "nf_count": int(display_df["NF"].nunique()),
            "item_count": int(item_mask.sum()),
            "peso_total": float(display_df["Peso"].sum()),
            "error_count": int(display_df.loc[error_mask, "NF"].nunique()),
        }

        return processed_df, summary, issues, debug_rows

    excel_nfs = set(base_df["nf_normalizada"].astype(str).tolist()) if "nf_normalizada" in base_df.columns else set()
    xml_nfs = set(xml_index.keys())

    unmatched_xml_nfs = sorted(xml_nfs - excel_nfs)
    missing_xml_nfs = sorted(excel_nfs - xml_nfs)

    if xml_source and not (excel_nfs & xml_nfs):
        issues.append("Nenhum XML enviado corresponde as NFs presentes no Excel.")

    if unmatched_xml_nfs:
        issues.append(f"XMLs ignorados por nao existirem no Excel: {', '.join(unmatched_xml_nfs)}")

    if missing_xml_nfs:
        issues.append(f"NFs do Excel sem XML correspondente: {', '.join(missing_xml_nfs)}")

    for row in base_df.to_dict(orient="records"):
        nf = row["NF"]
        nf_normalizada = normalize_nf(row.get("nf_normalizada", nf))
        xml_data = xml_index.get(nf_normalizada)

        debug_rows.append(
            {
                "NF Planilha": str(nf),
                "NF XML": str(xml_data.get("nf_normalizada", "")) if xml_data else "",
                "Tipo XML": str(xml_data.get("TipoXML", "")) if xml_data else "",
                "Arquivo XML": str(xml_data.get("Arquivo", "")) if xml_data else "",
                "Correspondencia": "OK" if xml_data else "XML nao encontrado",
            }
        )

        if not xml_data:
            rows.append(
                {
                    "Seq": row["Seq"],
                    "Seq_sort": row["Seq_sort"],
                    "ChaveNFe": "",
                    "NF": nf,
                    "Data": "",
                    "cProd": "",
                    "Descricao": "",
                    "Qtd": 0.0,
                    "Unidade": "",
                    "Volume": 0.0,
                    "Peso": 0.0,
                    "PesoTotalNF": 0.0,
                    "Destinatario": "",
                    "Municipio": "",
                    "Status": "XML nao encontrado",
                }
            )
            continue

        if not xml_data["Items"]:
            rows.append(
                {
                    "Seq": row["Seq"],
                    "Seq_sort": row["Seq_sort"],
                    "ChaveNFe": xml_data["ChaveNFe"],
                    "NF": nf,
                    "Data": xml_data["Data"],
                    "cProd": "",
                    "Descricao": "",
                    "Qtd": 0.0,
                    "Unidade": "",
                    "Volume": xml_data["VolumeTotal"],
                    "Peso": 0.0,
                    "PesoTotalNF": xml_data["PesoTotal"],
                    "Destinatario": xml_data["Destinatario"],
                    "Municipio": xml_data["Municipio"],
                    "Status": str(xml_data["Status"]),
                }
            )
            continue

        for item in xml_data["Items"]:
            rows.append(
                {
                    "Seq": row["Seq"],
                    "Seq_sort": row["Seq_sort"],
                    "ChaveNFe": xml_data["ChaveNFe"],
                    "NF": nf,
                    "Data": xml_data["Data"],
                    "cProd": item["cProd"],
                    "Descricao": item["Descricao"],
                    "Qtd": item["Qtd"],
                    "Unidade": item["Unidade"],
                    "Volume": xml_data["VolumeTotal"],
                    "Peso": item["Peso"],
                    "PesoTotalNF": xml_data["PesoTotal"],
                    "Destinatario": xml_data["Destinatario"],
                    "Municipio": xml_data["Municipio"],
                    "Status": str(xml_data["Status"]),
                }
            )

    processed_df = pd.DataFrame(rows)
    if processed_df.empty:
        processed_df = create_empty_processed_df()
    else:
        processed_df = processed_df.sort_values(by=["Seq_sort", "NF"], ascending=[False, True], na_position="last")
    processed_df = apply_routes_to_dataframe(processed_df)

    display_df = processed_df[TABLE_COLUMNS].copy()
    error_mask = ~display_df["Status"].astype(str).str.contains("autorizado", case=False, na=False)
    item_mask = display_df["cProd"].astype(str).str.strip() != ""
    metadata = summarize_metadata(base_df)

    summary = {
        "filial": summarize_filial(base_df),
        "numero_carga": metadata["numero_carga"],
        "data_saida": metadata["data_saida"],
        "motorista": metadata["motorista"],
        "placa": metadata["placa"],
        "nf_count": int(base_df["NF"].nunique()),
        "item_count": int(item_mask.sum()),
        "peso_total": float(display_df["Peso"].sum()),
        "error_count": int(display_df.loc[error_mask, "NF"].nunique()),
    }

    return processed_df, summary, issues, debug_rows


def create_empty_summary() -> dict[str, object]:
    return {
        "filial": "BRIDA",
        "numero_carga": "--",
        "data_saida": "--",
        "motorista": "--",
        "placa": "--",
        "nf_count": 0,
        "item_count": 0,
        "peso_total": 0.0,
        "error_count": 0,
    }


def create_empty_processed_df() -> pd.DataFrame:
    return pd.DataFrame(columns=["Seq_sort", "ChaveNFe", "Data", "Volume", "PesoTotalNF", "Municipio", *TABLE_COLUMNS])


def create_empty_nf_debug_df() -> pd.DataFrame:
    return pd.DataFrame(columns=NF_DEBUG_COLUMNS)


def build_table_column_config(dataframe: pd.DataFrame) -> dict[str, object]:
    return {
        "Seq": st.column_config.NumberColumn("Seq", format="%d", width="small"),
        "NF": st.column_config.TextColumn("NF", width="small"),
        "cProd": st.column_config.TextColumn("cProd", width="small"),
        "Descricao": st.column_config.TextColumn("Descricao", width="medium"),
        "Qtd": st.column_config.NumberColumn("Qtd", format="%.4f", width="small"),
        "Unidade": st.column_config.TextColumn("UN", width="small"),
        "Peso": st.column_config.NumberColumn("Peso", format="%.3f kg", width="small"),
        "Destinatario": st.column_config.TextColumn("Destinatario", width="medium"),
        "ROTA": st.column_config.TextColumn("ROTA", width="medium"),
        "Status": st.column_config.TextColumn("Status", width="medium"),
    }


def wrap_table_text(value: object, width: int) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return textwrap.fill(text, width=width, break_long_words=False, break_on_hyphens=False)


def build_display_table(dataframe: pd.DataFrame) -> pd.DataFrame:
    display_df = dataframe.copy()

    if "Descricao" in display_df.columns:
        if "cProd" in display_df.columns:
            display_df["Descricao"] = display_df.apply(
                lambda row: format_product_description(row.get("Descricao", ""), row.get("cProd", "")),
                axis=1,
            )
        display_df["Descricao"] = display_df["Descricao"].apply(lambda value: wrap_table_text(value, 32))

    if "Destinatario" in display_df.columns:
        display_df["Destinatario"] = display_df["Destinatario"].apply(lambda value: wrap_table_text(value, 28))

    if "ROTA" in display_df.columns:
        display_df["ROTA"] = display_df["ROTA"].apply(lambda value: wrap_table_text(value, 20))

    return display_df


def handle_login(username: str, password: str) -> bool:
    return username == LOGIN_USERNAME and password == LOGIN_PASSWORD


def initialize_login_state() -> None:
    auth_query_value = st.query_params.get(AUTH_QUERY_PARAM, "")
    if "logado" not in st.session_state:
        st.session_state["logado"] = auth_query_value == AUTH_QUERY_VALUE
        return

    if not st.session_state["logado"] and auth_query_value == AUTH_QUERY_VALUE:
        st.session_state["logado"] = True


def persist_login_state(is_logged_in: bool) -> None:
    if is_logged_in:
        st.query_params[AUTH_QUERY_PARAM] = AUTH_QUERY_VALUE
    else:
        st.query_params.clear()


def format_summary_value(value: object, default: str = "--") -> str:
    text = str(value or "").strip()
    return text or default


def is_authorized_status(value: object) -> bool:
    text = str(value or "").strip().lower()
    return "autoriz" in text


def style_status_cell(value: object) -> str:
    if is_authorized_status(value):
        return "; ".join(
            [
                "background-color: #EAF7EE",
                "color: #18794E",
                "font-weight: 700",
                "text-align: center",
                "border-radius: 8px",
            ]
        )

    return "; ".join(
        [
            "background-color: #FDECEC",
            "color: #B42318",
            "font-weight: 700",
            "text-align: center",
            "border-radius: 8px",
        ]
    )


def style_description_cell(value: object) -> str:
    if has_formatted_product_code(value):
        return "font-weight: 700"
    return ""


def style_route_cell(value: object) -> str:
    if str(value or "").strip().upper() != UNDEFINED_ROUTE_LABEL:
        return ""

    return "; ".join(
        [
            "background-color: #FEF3C7",
            "color: #9A3412",
            "font-weight: 700",
            "text-align: center",
            "border-radius: 8px",
        ]
    )


def build_status_styler(dataframe: pd.DataFrame):
    if dataframe.empty:
        return dataframe

    styler = dataframe.style

    if "Status" in dataframe.columns:
        styler = styler.map(style_status_cell, subset=["Status"]).set_properties(
            subset=["Status"], **{"text-align": "center"}
        )

    if "Descricao" in dataframe.columns:
        styler = styler.map(style_description_cell, subset=["Descricao"])

    if "ROTA" in dataframe.columns:
        styler = styler.map(style_route_cell, subset=["ROTA"])

    return styler


def render_info_card(title: str, value: object, icon_key: str, secondary: str = "") -> None:
    st.markdown(
        f"""
    <div class="erp-card erp-card-info">
        <div class="erp-card-header">
            {render_label_icon(ICON_MAP[icon_key])}
            <span class="erp-card-title">{title}</span>
        </div>
        <div class="erp-card-value">{format_summary_value(value)}</div>
        <div class="erp-card-secondary">{secondary or '&nbsp;'}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def render_metric_card(title: str, value: object, icon_key: str) -> None:
    st.markdown(
        f"""
    <div class="erp-card erp-card-kpi">
        <div class="erp-kpi-icon">{render_label_icon(ICON_MAP[icon_key])}</div>
        <div class="erp-kpi-value">{value}</div>
        <div class="erp-kpi-label">{title}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def render_section_heading(label: str, icon_key: str) -> None:
    st.markdown(
        f'''
    <div class="section-title-block with-icon">{render_label_icon(ICON_MAP[icon_key])}<span>{label}</span></div>
    ''',
        unsafe_allow_html=True,
    )


def render_login_screen() -> None:
    st.markdown(
        """
    <style>
    html, body {
        margin: 0;
    }
    .stApp {
        background: #F5F7FA;
    }
    .block-container {
        padding-top: 1.25rem !important;
        padding-bottom: 1.25rem !important;
        max-width: 1180px;
    }
    div[data-testid="stVerticalBlock"] {
        width: 100%;
    }
    div[data-testid="stVerticalBlock"] > div:empty {
        display: none;
    }
    .login-stage {
        width: 100%;
        max-width: 1040px;
        margin: 8vh auto 0;
        padding: 0 1.5rem;
        box-sizing: border-box;
    }
    .login-logo-wrap {
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
        padding: 1rem 0;
    }
    .login-shell {
        max-width: 400px;
        margin: 0 auto;
        display: flex;
        flex-direction: column;
        text-align: left;
    }
    .login-intro {
        max-width: 400px;
        margin: 0 0 20px;
        text-align: left;
    }
    .login-intro h2 {
        margin: 0 0 10px;
        color: #1F3A5F;
        font-size: 1.65rem;
        font-weight: 700;
    }
    .login-intro p {
        margin: 0;
        color: #607085;
        line-height: 1.5;
        font-size: 0.95rem;
    }
    div[data-testid="stForm"] {
        max-width: 420px;
        margin: 0;
        padding: 24px 24px 20px;
        border-radius: 10px;
        background: #FFFFFF;
        border: 1px solid rgba(31, 58, 95, 0.10);
        box-shadow: 0 6px 18px rgba(31, 58, 95, 0.04);
    }
    div[data-testid="stTextInputRootElement"] > div {
        border-radius: 8px;
    }
    div[data-testid="stTextInputRootElement"] input {
        border-radius: 8px;
    }
    div[data-testid="stCheckbox"] {
        max-width: 400px;
        margin: 0 0 12px;
        color: #607085;
        text-align: left;
    }
    .login-feedback-success {
        max-width: 420px;
        margin: 0 0 14px;
        margin-bottom: 14px;
        padding: 10px 12px;
        border-radius: 8px;
        background: rgba(46, 111, 149, 0.10);
        color: #1F3A5F;
        border: 1px solid rgba(46, 111, 149, 0.22);
        font-size: 0.92rem;
    }
    .login-feedback-error {
        max-width: 420px;
        margin: 0 0 14px;
        margin-bottom: 14px;
        padding: 10px 12px;
        border-radius: 8px;
        background: rgba(243, 112, 33, 0.10);
        color: #9A4310;
        border: 1px solid rgba(243, 112, 33, 0.26);
        font-size: 0.92rem;
    }
    div[data-testid="stFormSubmitButton"] button {
        background: #1F3A5F;
        color: #FFFFFF;
        border: 0;
        border-radius: 8px;
        min-height: 42px;
        font-weight: 700;
        box-shadow: none;
    }
    div[data-testid="stFormSubmitButton"] button:hover {
        background: #25486E;
        color: #FFFFFF;
    }
    @media (max-width: 960px) {
        .login-stage {
            margin-top: 5vh;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        .login-logo-wrap {
            margin-bottom: 1rem;
        }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="login-stage">', unsafe_allow_html=True)
    logo_col, login_col = st.columns([1, 1], gap="large", vertical_alignment="center")

    with logo_col:
        st.markdown('<div class="login-logo-wrap">', unsafe_allow_html=True)
        logo_path = get_logo_path()
        if logo_path is not None:
            st.image(str(logo_path), width=320)
        st.markdown('</div>', unsafe_allow_html=True)

    with login_col:
        st.markdown('<div class="login-shell">', unsafe_allow_html=True)
        st.markdown(
            """
        <div class="login-intro">
            <h2>Acessar sistema</h2>
            <p>Entre com suas credenciais.</p>
        </div>
        """,
            unsafe_allow_html=True,
        )

        show_password = st.checkbox("Mostrar senha")

        login_error = st.session_state.get("login_error", "")
        if login_error:
            st.markdown(f'<div class="login-feedback-error">{login_error}</div>', unsafe_allow_html=True)

        with st.form("login_form"):
            username = st.text_input("Usuario", placeholder="Digite seu usuario")
            password = st.text_input(
                "Senha",
                type="default" if show_password else "password",
                placeholder="Digite sua senha",
            )
            submitted = st.form_submit_button("Entrar", use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

        if submitted:
            if handle_login(username, password):
                st.session_state["logado"] = True
                persist_login_state(True)
                st.session_state["login_error"] = ""
                st.session_state["login_success"] = "Acesso validado com sucesso."
                st.rerun()
            else:
                st.session_state["login_error"] = "Usuario ou senha incorretos."
                st.session_state["login_success"] = ""

    st.markdown('</div>', unsafe_allow_html=True)


def logout() -> None:
    st.session_state["logado"] = False
    persist_login_state(False)
    st.session_state["login_error"] = ""
    st.session_state["login_success"] = ""
    st.rerun()


def toggle_menu() -> None:
    st.session_state["menu_aberto"] = not st.session_state.get("menu_aberto", True)


def apply_sidebar_visibility(menu_aberto: bool) -> None:
    if menu_aberto:
        st.markdown(
            """
        <style>
        [data-testid="stSidebar"] {
            display: block;
        }
        [data-testid="stSidebarCollapsedControl"] {
            display: none;
        }
        </style>
        """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        """
    <style>
    [data-testid="stSidebar"] {
        display: none;
    }
    [data-testid="stSidebarCollapsedControl"] {
        display: none;
    }
    [data-testid="stMainBlockContainer"] {
        max-width: none;
        width: 100%;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> tuple[list, object, bool]:
    with st.sidebar:
        st.markdown(
            f'''
        <div class="sidebar-heading with-icon">{render_label_icon(ICON_MAP["dados_gerais"])}<span>Navegacao</span></div>
        ''',
            unsafe_allow_html=True,
        )

        st.markdown("**Tela**")
        st.caption("Minuta de Carregamento")
        st.session_state["tela_atual"] = "minuta"

        st.markdown(
            f'''
        <div class="sidebar-heading with-icon">{render_label_icon(ICON_MAP["dados_gerais"])}<span>Arquivos</span></div>
        ''',
            unsafe_allow_html=True,
        )

        st.markdown(
            f'''
        <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["excel"])}<span>Atualizar Praças</span></div>
        ''',
            unsafe_allow_html=True,
        )

        pracas_file = st.file_uploader(
            "Atualizar Praças",
            type=["xlsx"],
            accept_multiple_files=False,
            key="pracas_upload_widget",
        )

        if pracas_file is None:
            st.session_state.pracas_upload_signature = ""
        else:
            file_bytes = pracas_file.getvalue()
            current_signature = hashlib.sha256(file_bytes).hexdigest()
            if current_signature != st.session_state.pracas_upload_signature:
                try:
                    st.session_state.pracas_upload_message = update_pracas_json(pracas_file)
                    st.session_state.pracas_upload_error = ""
                    st.session_state.pracas_upload_signature = current_signature
                except ValueError as exc:
                    st.session_state.pracas_upload_message = ""
                    st.session_state.pracas_upload_error = str(exc)
                    st.session_state.pracas_upload_signature = ""

        if st.session_state.pracas_upload_message:
            st.success(st.session_state.pracas_upload_message)
        if st.session_state.pracas_upload_error:
            st.error(st.session_state.pracas_upload_error)

        st.markdown(
            f'''
        <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["xml"])}<span>Upload XML</span></div>
        ''',
            unsafe_allow_html=True,
        )

        xml_files = st.file_uploader(
            "Selecionar XMLs",
            type=["xml"],
            accept_multiple_files=True,
            key="xml_upload_widget",
        )

        xml_records, xml_storage_error = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))

        if xml_files:
            upload_signature_parts: list[str] = []
            for uploaded_file in xml_files:
                file_bytes = uploaded_file.getvalue()
                upload_signature_parts.append(f"{uploaded_file.name}:{hashlib.sha256(file_bytes).hexdigest()}")

            current_signature = hashlib.sha256("|".join(upload_signature_parts).encode("utf-8")).hexdigest()
            if current_signature != st.session_state.xml_upload_signature or not XMLS_PROCESSADOS_JSON_PATH.is_file():
                valid_files, ignored_files, issues = salvar_xmls_processados_json(xml_files)
                xml_records, xml_storage_error = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
                st.session_state.xml_upload_signature = current_signature
                st.session_state.xml_upload_message = (
                    f"Base de XMLs atualizada: {valid_files} validos, {ignored_files} ignorados"
                )
                st.session_state.xml_upload_error = ""
                st.session_state.xml_upload_issues = issues

        has_xml_storage, xml_updated_at = get_xml_storage_status()
        if st.session_state.xml_upload_message:
            st.success(st.session_state.xml_upload_message)
        if xml_storage_error:
            st.warning(xml_storage_error)
        elif has_xml_storage:
            st.caption(f"Dados carregados do sistema • Ultima atualizacao: {xml_updated_at}")

        if st.session_state.get("xml_upload_issues"):
            with st.expander("Detalhes dos XMLs", expanded=False):
                for issue in st.session_state.xml_upload_issues:
                    st.warning(issue)

        st.markdown(
            f'''
        <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["excel"])}<span>Upload Excel</span></div>
        ''',
            unsafe_allow_html=True,
        )

        excel_file = st.file_uploader(
            "Selecionar Excel",
            type=["xlsx", "xls"],
        )

        process_clicked = st.button("Processar", use_container_width=True)

    return xml_records, excel_file, process_clicked


def render_processing_screen(process_clicked: bool, xml_records: list, excel_file) -> None:
    if process_clicked:
        if excel_file is None:
            st.error("Envie um arquivo Excel para iniciar o processamento.")
        else:
            try:
                excel_base = load_excel_base(excel_file)
                processed_df, summary, issues, nf_debug = integrate_excel_with_xml(excel_base, xml_records or [])
                st.session_state.processed_df = processed_df
                st.session_state.summary = summary
                st.session_state.issues = issues
                st.session_state.nf_debug = pd.DataFrame(nf_debug, columns=NF_DEBUG_COLUMNS)
                st.session_state.document_issue_at = format_datetime_display()

                if processed_df.empty:
                    st.warning("Nenhum dado foi processado. Verifique se o Excel possui NFs validas.")
                else:
                    st.success("Processamento concluido.")
            except ValueError as exc:
                st.session_state.processed_df = create_empty_processed_df()
                st.session_state.summary = create_empty_summary()
                st.session_state.issues = []
                st.session_state.nf_debug = create_empty_nf_debug_df()
                st.error(str(exc))
            except Exception as exc:
                st.session_state.processed_df = create_empty_processed_df()
                st.session_state.summary = create_empty_summary()
                st.session_state.issues = []
                st.session_state.nf_debug = create_empty_nf_debug_df()
                st.error(f"Erro inesperado ao processar os arquivos: {exc}")

    summary = st.session_state.summary
    processed_df = apply_routes_to_dataframe(st.session_state.processed_df.copy())
    st.session_state.processed_df = processed_df

    render_section_heading("Dados Gerais", "dados_gerais")
    dados_col_1, dados_col_2, dados_col_3 = st.columns(3, gap="medium")
    with dados_col_1:
        render_info_card("Filial", summary["filial"], "filial")
    with dados_col_2:
        render_info_card("Carregamento", summary["numero_carga"], "carregamento")
    with dados_col_3:
        render_info_card("Data Saida", summary["data_saida"], "data_saida")

    dados_col_4, dados_col_5 = st.columns(2, gap="medium")
    with dados_col_4:
        render_info_card("Motorista", summary["motorista"], "motorista")
    with dados_col_5:
        render_info_card("Placa", summary["placa"], "placa")

    render_section_heading("Resumo da Carga", "resumo_carga")
    resumo_col_1, resumo_col_2, resumo_col_3, resumo_col_4 = st.columns(4, gap="medium")
    with resumo_col_1:
        render_metric_card("NF", summary["nf_count"], "nf")
    with resumo_col_2:
        render_metric_card("Peso", f"{summary['peso_total'] / 1000:.3f} t", "peso")
    with resumo_col_3:
        render_metric_card("Itens", summary["item_count"], "itens")
    with resumo_col_4:
        render_metric_card("Erros", summary["error_count"], "erros")

    if DISPLAY_PROCESSING_WARNINGS and st.session_state.issues:
        for issue in st.session_state.issues:
            st.warning(issue)

    if not st.session_state.nf_debug.empty:
        with st.expander("Debug de correspondencia NF x XML", expanded=False):
            st.dataframe(st.session_state.nf_debug, use_container_width=True, hide_index=True)

    action_col_search, action_col_download = st.columns([2.0, 1.8], gap="medium")

    with action_col_search:
        st.markdown('<div class="section-title">Localizar registros</div>', unsafe_allow_html=True)
        search_term = st.text_input("Pesquisar (qualquer coluna)", placeholder="Buscar NF, produto, destinatario ou status")

    if search_term and not processed_df.empty:
        filtered_df = processed_df[
            processed_df.astype(str).apply(
                lambda column: column.str.contains(search_term, case=False, na=False)
            ).any(axis=1)
        ]
    else:
        filtered_df = processed_df

    display_df = build_display_table(filtered_df[TABLE_COLUMNS].copy())
    styled_display_df = build_status_styler(display_df)

    minuta_records = build_minuta_records(processed_df)
    pdf_bytes = b""
    if minuta_records:
        pdf_bytes = generate_minuta_pdf(
            dados_minuta=minuta_records,
            numero_carga=str(summary.get("numero_carga", "--") or "--"),
            data_emissao=str(st.session_state.document_issue_at or "--"),
            veiculo=str(summary.get("placa", "--") or "--"),
            motorista=str(summary.get("motorista", "--") or "--"),
        )

    with action_col_download:
        st.markdown('<div class="section-title export-title">Exportacao</div>', unsafe_allow_html=True)
        pdf_col_left, pdf_col_right = st.columns([1.0, 1.1], gap="small")
        with pdf_col_right:
            st.download_button(
                "Baixar PDF",
                data=pdf_bytes,
                file_name=f"minuta_carregamento_{sanitize_filename_part(summary.get('numero_carga'), 'brida')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                disabled=not bool(pdf_bytes),
            )

    st.markdown("### Painel de Notas e Itens")
    st.caption("Visualizacao consolidada da carga com detalhamento operacional por nota fiscal.")
    st.dataframe(
        styled_display_df,
        width="stretch",
        hide_index=True,
        column_config=build_table_column_config(display_df),
        row_height=56,
    )
def render_main_screen() -> None:
    if "menu_aberto" not in st.session_state:
        st.session_state["menu_aberto"] = True

    apply_sidebar_visibility(st.session_state["menu_aberto"])

    col_logo, col_header, col_menu_toggle, col_action = st.columns([1.1, 4.4, 1, 1], vertical_alignment="center")
    with col_logo:
        logo_path = get_logo_path()
        if logo_path is not None:
            st.image(str(logo_path), width=120)
    with col_header:
        st.markdown("## Minuta de Carregamento")
    with col_menu_toggle:
        st.button("Menu", use_container_width=True, on_click=toggle_menu)
    with col_action:
        st.button("Sair", use_container_width=True, on_click=logout)

    if "processed_df" not in st.session_state:
        st.session_state.processed_df = create_empty_processed_df()

    if "summary" not in st.session_state:
        st.session_state.summary = create_empty_summary()

    if "issues" not in st.session_state:
        st.session_state.issues = []

    if "nf_debug" not in st.session_state:
        st.session_state.nf_debug = create_empty_nf_debug_df()

    if "document_issue_at" not in st.session_state:
        st.session_state.document_issue_at = format_datetime_display()

    if "pracas_upload_signature" not in st.session_state:
        st.session_state.pracas_upload_signature = ""

    if "pracas_upload_message" not in st.session_state:
        st.session_state.pracas_upload_message = ""

    if "pracas_upload_error" not in st.session_state:
        st.session_state.pracas_upload_error = ""

    if "xml_upload_signature" not in st.session_state:
        st.session_state.xml_upload_signature = ""

    if "xml_upload_message" not in st.session_state:
        st.session_state.xml_upload_message = ""

    if "xml_upload_error" not in st.session_state:
        st.session_state.xml_upload_error = ""

    if "xml_upload_issues" not in st.session_state:
        st.session_state.xml_upload_issues = []

    if "tela_atual" not in st.session_state:
        st.session_state["tela_atual"] = "minuta"

    login_success = st.session_state.get("login_success", "")
    if login_success:
        st.success(login_success)
        st.session_state["login_success"] = ""

    st.markdown(
        """
    <style>
    .stApp {
        background: #F5F7FA;
    }
    .ui-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 16px;
        min-width: 16px;
        height: 16px;
        color: #7C8AA0;
    }
    .ui-icon svg {
        width: 16px;
        height: 16px;
        stroke: currentColor;
        stroke-width: 1.7;
        fill: none;
        stroke-linecap: round;
        stroke-linejoin: round;
    }
    .with-icon {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
    }
    .section-title-block {
        margin: 0 0 14px;
        color: #1F3A5F;
        font-size: 1rem;
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    .erp-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
        padding: 10px 12px;
        height: auto;
        min-height: 0;
        margin-bottom: 12px;
        color: #405468;
    }
    .erp-card-info {
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        gap: 0.3rem;
    }
    .erp-card-kpi {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 0.2rem;
    }
    .erp-card-header {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin-bottom: 0.25rem;
    }
    .erp-card-title {
        color: #6B7280;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        line-height: 1.2;
        text-transform: uppercase;
    }
    .erp-card-value {
        color: #1F2937;
        font-size: 0.96rem;
        font-weight: 600;
        line-height: 1.2;
        overflow-wrap: anywhere;
        word-break: break-word;
        text-wrap: pretty;
    }
    .erp-card-secondary {
        color: #9CA3AF;
        font-size: 0.76rem;
        line-height: 1.2;
        min-height: 0;
        margin-top: 0.15rem;
    }
    .erp-kpi-icon {
        display: flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 0.05rem;
    }
    .erp-kpi-value {
        color: #1F2937;
        font-size: 1.55rem;
        line-height: 1.1;
        font-weight: 700;
        letter-spacing: -0.03em;
    }
    .erp-kpi-label {
        color: #6B7280;
        font-size: 0.74rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        line-height: 1.2;
        text-transform: uppercase;
    }
    [data-testid="stDataFrame"] table td {
        vertical-align: middle;
    }
    .section-title {
        margin: 6px 0 10px;
        color: #1F3A5F;
        font-size: 0.95rem;
        font-weight: 700;
    }
    .export-title {
        text-align: right;
        margin-bottom: 0.75rem;
    }
    .print-action-button {
        width: 100%;
        min-height: 42px;
        padding: 0 1rem;
        border-radius: 10px;
        border: 0;
        background: #1F3A5F;
        color: #FFFFFF;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 0.45rem;
        font-size: 0.96rem;
        font-weight: 500;
        cursor: pointer;
        transition: background-color 0.15s ease;
    }
    .print-action-button:hover {
        background: #25486E;
    }
    .print-action-button .ui-icon {
        color: #FFFFFF;
    }
    .print-action-button.is-disabled {
        background: #9AA8B8;
        cursor: not-allowed;
    }
    .print-action-button.is-loading {
        background: #35597D;
        cursor: wait;
    }
    .table-shell {
        background: #FFFFFF;
        padding: 16px;
        border-radius: 8px;
        border: 1px solid rgba(31, 58, 95, 0.08);
        box-shadow: 0 4px 16px rgba(31, 58, 95, 0.04);
    }
    .table-shell h3 {
        margin: 0 0 6px;
        color: #1F3A5F;
        font-size: 0.98rem;
    }
    .table-shell p {
        margin: 0 0 16px;
        color: #617285;
        font-size: 0.9rem;
    }
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 8px;
        min-height: 40px;
        border: 1px solid rgba(31, 58, 95, 0.14);
        background: #FFFFFF;
        color: #1F3A5F;
        font-weight: 700;
        box-shadow: none;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: rgba(46, 111, 149, 0.30);
        color: #1F3A5F;
    }
    .stDownloadButton > button {
        background: #1F3A5F;
        color: #FFFFFF;
        border: 0;
        min-height: 42px;
        padding-left: 1rem;
        padding-right: 1rem;
        font-weight: 500;
    }
    .stDownloadButton > button:hover {
        color: #FFFFFF;
        background: #25486E;
    }
    .stTextInput > div > div input {
        border-radius: 8px;
    }
    [data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid rgba(31, 58, 95, 0.08);
    }
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #1F3A5F;
    }
    [data-testid="stSidebar"] label[data-testid="stWidgetLabel"] {
        color: #516478;
        font-weight: 500;
    }
    [data-testid="stSidebar"] .stFileUploader {
        padding: 10px;
        border-radius: 8px;
        background: #F8FAFC;
        border: 1px solid rgba(31, 58, 95, 0.08);
        box-shadow: none;
    }
    [data-testid="stSidebar"] .stButton > button {
        background: #1F3A5F;
        color: #FFFFFF;
        border: 0;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        color: #FFFFFF;
        background: #25486E;
    }
    div[data-testid="stAlert"] {
        border-radius: 8px;
        border: 1px solid rgba(31, 58, 95, 0.08);
        box-shadow: none;
    }
    div[data-testid="stAlert"] p {
        font-size: 0.92rem;
        line-height: 1.55;
    }
    @media (max-width: 900px) {
        .table-shell {
            padding: 16px;
        }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    xml_records = []
    excel_file = None
    process_clicked = False
    if st.session_state["menu_aberto"]:
        xml_records, excel_file, process_clicked = render_sidebar()

    st.session_state["tela_atual"] = "minuta"
    render_processing_screen(process_clicked, xml_records, excel_file)


def main() -> None:
    st.set_page_config(layout="wide")
    initialize_login_state()

    if "login_error" not in st.session_state:
        st.session_state["login_error"] = ""

    if "login_success" not in st.session_state:
        st.session_state["login_success"] = ""

    if not st.session_state["logado"]:
        render_login_screen()
        return

    render_main_screen()


if __name__ == "__main__":
    main()