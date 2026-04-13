from pathlib import Path
from datetime import datetime
from collections import Counter
import base64
import html
from io import BytesIO
import json
import re
import hashlib
import textwrap
import unicodedata
import xml.etree.ElementTree as ET

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle
import streamlit as st
import streamlit.components.v1 as components

BASE_DIR = Path(__file__).resolve().parent
FIXED_LOGO_PATH = BASE_DIR / "baixados.png"
WINDOWS_FONT_DIR = Path("C:/Windows/Fonts")
DATA_DIR = BASE_DIR / "data"
PRACAS_JSON_PATH = DATA_DIR / "pracas.json"
XMLS_PROCESSADOS_JSON_PATH = DATA_DIR / "xmls_processados.json"
CLASSIFICACAO_PRODUTOS_JSON_PATH = DATA_DIR / "classificacao_produtos.json"
SEPARACAO_JSON_PATH = DATA_DIR / "separacao.json"
LOTES_JSON_PATH = DATA_DIR / "lotes.json"
SEPARACAO_EXCLUIDOS_JSON_PATH = DATA_DIR / "separacao_excluidos.json"
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
SEPARATION_PENDING_STATUS = "Pendente"
SEPARATION_SEPARATED_STATUS = "Separado"
NF_STATUS_AUTHORIZED = "Autorizado o uso da NF-e"
NF_STATUS_CANCELED = "Cancelada"
LOT_STATUS_OPEN = "Aberto"
LOT_STATUS_CLOSED = "Fechado"
DATA_CLEANUP_TYPE_XML = "Apenas XMLs"
DATA_CLEANUP_TYPE_SEPARACAO = "Apenas Separação"
DATA_CLEANUP_TYPE_LOTES = "Apenas Lotes"
DATA_CLEANUP_TYPE_COMPLETE = "Limpeza completa"
DATA_CLEANUP_OPTIONS = [
    DATA_CLEANUP_TYPE_XML,
    DATA_CLEANUP_TYPE_SEPARACAO,
    DATA_CLEANUP_TYPE_LOTES,
    DATA_CLEANUP_TYPE_COMPLETE,
]
SEPARATION_VISIBLE_COLUMNS = ["NF", "Produto", "Qtd", "Tipo", "Cliente", "Setor", "Rota", "Lote", "Status NF"]
SEPARATION_SECTORS = ["Lubrificantes", "Paletas", "Filtros", "Arla", "Não Identificados"]
SECTOR_CLASSIFICATION_PRIORITY = {
    "Filtros": 0,
    "Arla": 1,
    "Lubrificantes": 2,
    "Paletas": 3,
    "Não Identificados": 99,
}
SECTOR_NAME_ALIASES = {
    "LUBRIFICANTE": "Lubrificantes",
    "LUBRIFICANTES": "Lubrificantes",
    "OLEO": "Lubrificantes",
    "OLEOS": "Lubrificantes",
    "PALETA": "Paletas",
    "PALETAS": "Paletas",
    "PALLET": "Paletas",
    "PALLETS": "Paletas",
    "FILTRO": "Filtros",
    "FILTROS": "Filtros",
    "ARLA": "Arla",
    "NAO IDENTIFICADO": "Não Identificados",
    "NAO IDENTIFICADOS": "Não Identificados",
    "NAO CLASSIFICADO": "Não Identificados",
    "SEM SETOR": "Não Identificados",
}
UNDEFINED_ROUTE_LABEL = "NÃO DEFINIDA"
DEFAULT_PRODUCT_CLASSIFICATION_RULES = [
    {"palavra_chave": "OLEO", "setor": "Lubrificantes"},
    {"palavra_chave": "MOBIL", "setor": "Lubrificantes"},
    {"palavra_chave": "LUBRIFICANTE", "setor": "Lubrificantes"},
    {"palavra_chave": "PALETA", "setor": "Paletas"},
    {"palavra_chave": "PALETAS", "setor": "Paletas"},
    {"palavra_chave": "PALLET", "setor": "Paletas"},
    {"palavra_chave": "PALLETS", "setor": "Paletas"},
    {"palavra_chave": "FILTRO", "setor": "Filtros"},
    {"palavra_chave": "WEGA", "setor": "Filtros"},
    {"palavra_chave": "ARLA", "setor": "Arla"},
]
SECTOR_COLOR_MAP = {
    "Lubrificantes": {"bg": "#E8F1FF", "fg": "#174EA6", "border": "#3B82F6"},
    "Filtros": {"bg": "#F4E8FF", "fg": "#6B21A8", "border": "#A855F7"},
    "Arla": {"bg": "#EAFBF0", "fg": "#166534", "border": "#22C55E"},
    "Paletas": {"bg": "#FFF1E6", "fg": "#C2410C", "border": "#F97316"},
    "Não Identificados": {"bg": "#FDECEC", "fg": "#B42318", "border": "#EF4444"},
    "Misto": {"bg": "#EEF2F7", "fg": "#334155", "border": "#94A3B8"},
}
PDF_FONT_REGULAR = "Helvetica"
PDF_FONT_BOLD = "Helvetica-Bold"
PDF_FONT_MONO = "Courier"
PDF_FONT_MONO_BOLD = "Courier-Bold"
LOGIN_USERNAME = "minuta"
LOGIN_PASSWORD = "minuta123"
AUTH_QUERY_PARAM = "auth"
AUTH_QUERY_VALUE = "1"
SCREEN_LOGIN = "login"
SCREEN_MENU = "menu"
SCREEN_MINUTA = "minuta"
SCREEN_SEPARACAO = "separacao"
SCREEN_LOTES = "lotes"
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
    "separacao": "box",
    "rota": "truck",
    "setor": "folder",
    "barcode": "receipt",
    "status_operacional": "chart",
    "lotes": "box",
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


def normalize_matching_text(value: object) -> str:
    text = normalize_praca_name(value)
    if not text:
        return ""
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokenize_matching_text(value: object) -> list[str]:
    text = normalize_matching_text(value)
    if not text:
        return []
    return text.split()


def keyword_matches_description(keyword: str, normalized_description: str, description_tokens: set[str]) -> bool:
    normalized_keyword = normalize_matching_text(keyword)
    if not normalized_keyword:
        return False

    keyword_tokens = normalized_keyword.split()
    if len(keyword_tokens) == 1:
        return keyword_tokens[0] in description_tokens

    return normalized_keyword in normalized_description


def normalize_sector_name(value: object) -> str:
    normalized = normalize_matching_text(value)
    sector_lookup = {normalize_matching_text(sector): sector for sector in SEPARATION_SECTORS}
    if normalized in sector_lookup:
        return sector_lookup[normalized]
    return SECTOR_NAME_ALIASES.get(normalized, "")


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


def get_sector_colors(setor: str) -> dict[str, str]:
    return SECTOR_COLOR_MAP.get(setor, SECTOR_COLOR_MAP["Não Identificados"])


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


def build_default_product_classification_records() -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for item in DEFAULT_PRODUCT_CLASSIFICATION_RULES:
        keyword = normalize_matching_text(item.get("palavra_chave", ""))
        sector = normalize_sector_name(item.get("setor", ""))
        if keyword and sector:
            records.append({"palavra_chave": keyword, "setor": sector})
    return sorted(records, key=lambda record: (-len(record["palavra_chave"]), record["palavra_chave"]))


def update_classificacao_produtos_json(uploaded_file) -> str:
    try:
        uploaded_file.seek(0)
        classificacao_df = pd.read_excel(uploaded_file)
        uploaded_file.seek(0)
    except Exception as exc:
        raise ValueError(f"Erro ao ler o arquivo de classificacao de produtos: {exc}") from exc

    keyword_column = find_column(
        list(classificacao_df.columns),
        ["Palavra Chave", "Palavra-chave", "Palavra", "Keyword", "Descricao", "Produto", "Chave"],
    )
    sector_column = find_column(
        list(classificacao_df.columns),
        ["Setor", "Classificacao", "Classificação", "Categoria", "Grupo"],
    )

    if keyword_column is None or sector_column is None:
        raise ValueError("A planilha precisa conter colunas de palavra-chave e setor para a classificacao.")

    records: dict[str, dict[str, str]] = {}
    for _, row in classificacao_df.iterrows():
        keyword = normalize_matching_text(row.get(keyword_column, ""))
        sector = normalize_sector_name(row.get(sector_column, ""))
        if keyword and sector:
            records[keyword] = {"palavra_chave": keyword, "setor": sector}

    if not records:
        raise ValueError("Nenhuma regra valida foi encontrada na planilha de classificacao.")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CLASSIFICACAO_PRODUTOS_JSON_PATH.write_text(
        json.dumps(sorted(records.values(), key=lambda record: (-len(record["palavra_chave"]), record["palavra_chave"])), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    carregar_classificacao_produtos_json.clear()
    return f"Classificacao de produtos atualizada com {len(records)} regra(s)."


@st.cache_data(show_spinner=False)
def carregar_classificacao_produtos_json(json_path: str, version_token: int = 0) -> tuple[list[dict[str, str]], str]:
    _ = version_token
    default_records = build_default_product_classification_records()
    path = Path(json_path)
    if not path.is_file():
        return default_records, ""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        return default_records, f"A base de classificacao nao pôde ser lida ({exc}). Foi usada a base padrao do sistema."

    if not isinstance(payload, list):
        return default_records, "A base de classificacao esta em formato invalido. Foi usada a base padrao do sistema."

    records: dict[str, dict[str, str]] = {}
    for item in payload:
        if not isinstance(item, dict):
            continue
        keyword = normalize_matching_text(item.get("palavra_chave", ""))
        sector = normalize_sector_name(item.get("setor", ""))
        if keyword and sector:
            records[keyword] = {"palavra_chave": keyword, "setor": sector}

    if not records:
        return default_records, "A base de classificacao estava vazia. Foi usada a base padrao do sistema."

    return sorted(records.values(), key=lambda record: (-len(record["palavra_chave"]), record["palavra_chave"])), ""


def get_classificacao_storage_status() -> tuple[bool, str]:
    if not CLASSIFICACAO_PRODUTOS_JSON_PATH.is_file():
        return False, "Usando base padrao do sistema"

    updated_at = datetime.fromtimestamp(CLASSIFICACAO_PRODUTOS_JSON_PATH.stat().st_mtime)
    return True, format_datetime_display(updated_at)


def classify_product_sector(description: object, classification_records: list[dict[str, str]]) -> str:
    normalized_description = normalize_matching_text(description)
    if not normalized_description:
        return "Não Identificados"

    description_tokens = set(tokenize_matching_text(description))
    sector_match_count: Counter[str] = Counter()
    sector_keyword_length: Counter[str] = Counter()

    for rule in classification_records or []:
        keyword = normalize_matching_text(rule.get("palavra_chave", ""))
        sector = normalize_sector_name(rule.get("setor", "")) or "Não Identificados"
        if keyword_matches_description(keyword, normalized_description, description_tokens):
            sector_match_count[sector] += 1
            sector_keyword_length[sector] += len(keyword)

    if sector_match_count:
        ranked_sector = max(
            sector_match_count,
            key=lambda sector: (
                -SECTOR_CLASSIFICATION_PRIORITY.get(sector, 999),
                sector_match_count[sector],
                sector_keyword_length[sector],
                sector,
            ),
        )
        return ranked_sector

    return "Não Identificados"


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


def parse_xml_datetime(value: object) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None

    iso_text = text.replace("Z", "+00:00")
    try:
        parsed_iso = datetime.fromisoformat(iso_text)
        if parsed_iso.tzinfo is not None:
            return parsed_iso.astimezone().replace(tzinfo=None)
        return parsed_iso
    except ValueError:
        pass

    parsed = pd.to_datetime(text, errors="coerce", utc=True, dayfirst=True)
    if pd.isna(parsed):
        return None

    return parsed.tz_convert(None).to_pydatetime()


def extract_xml_reference_datetime(root: ET.Element, xml_type: str) -> tuple[str, datetime | None]:
    if xml_type == "evento":
        raw_value = find_xml_text_by_localname(root, ["dhRegEvento", "dhEvento", "dhRecbto"])
    else:
        raw_value = find_xml_text_by_localname(root, ["dhEmi", "dEmi", "dhSaiEnt"])

    return raw_value, parse_xml_datetime(raw_value)


def normalize_nf_status(value: object) -> str:
    text = normalize_matching_text(value)
    if not text:
        return "Status nao informado"
    if "CANCEL" in text:
        return NF_STATUS_CANCELED
    if "AUTORIZ" in text:
        return NF_STATUS_AUTHORIZED
    return str(value or "").strip() or "Status nao informado"


def is_canceled_nf_status(value: object) -> bool:
    return normalize_nf_status(value) == NF_STATUS_CANCELED


def is_authorized_nf_status(value: object) -> bool:
    return normalize_nf_status(value) == NF_STATUS_AUTHORIZED


def get_nf_status_priority(value: object) -> int:
    status = normalize_nf_status(value)
    if status == NF_STATUS_CANCELED:
        return 2
    if status == NF_STATUS_AUTHORIZED:
        return 1
    return 0


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


def extract_xml_status(root: ET.Element, xml_type: str) -> str:
    if xml_type == "evento":
        event_code = find_xml_text_by_localname(root, ["tpEvento"])
        event_description = normalize_matching_text(find_xml_text_by_localname(root, ["descEvento", "xEvento", "xJust"]))
        if event_code == "110111" or "CANCEL" in event_description:
            return NF_STATUS_CANCELED

    status = find_xml_text_by_localname(root, ["xMotivo", "cStat", "descEvento"])
    normalized_status = normalize_nf_status(status)
    if xml_type == "normal" and normalized_status == "Status nao informado":
        return NF_STATUS_AUTHORIZED
    return normalized_status


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
    reference_raw, reference_datetime = extract_xml_reference_datetime(root, xml_type)
    status = extract_xml_status(root, xml_type)
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
        "DataReferencia": reference_raw,
        "DataReferenciaISO": reference_datetime.isoformat() if reference_datetime else "",
        "Destinatario": destinatario,
        "Municipio": municipio,
        "Status": status,
        "StatusNF": status,
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
        "DataReferencia": str(xml_data.get("DataReferencia", "") or "").strip(),
        "DataReferenciaISO": str(xml_data.get("DataReferenciaISO", "") or "").strip(),
        "Destinatario": str(xml_data.get("Destinatario", "") or "").strip(),
        "Municipio": municipio,
        "Status": normalize_nf_status(xml_data.get("StatusNF", xml_data.get("Status", ""))),
        "StatusNF": normalize_nf_status(xml_data.get("StatusNF", xml_data.get("Status", ""))),
        "VolumeTotal": parse_float(xml_data.get("VolumeTotal", 0.0)),
        "PesoTotal": parse_float(xml_data.get("PesoTotal", 0.0)),
        "Items": items,
        "Arquivo": str(xml_data.get("Arquivo", "") or "").strip(),
        "Erro": bool(xml_data.get("Erro", False)),
        "TipoXML": str(xml_data.get("TipoXML", "normal") or "normal").strip(),
        "ROTA": str(xml_data.get("ROTA", "") or get_route_for_municipio(municipio)).strip(),
    }


def get_xml_identity(xml_data: dict[str, object]) -> str:
    chave = normalize_chave_nfe(xml_data.get("ChaveNFe", ""))
    if chave:
        return chave
    return normalize_nf(xml_data.get("NF", "") or xml_data.get("nf_normalizada", ""))


def get_xml_reference_datetime(record: dict[str, object]) -> datetime | None:
    return parse_xml_datetime(record.get("DataReferenciaISO", "") or record.get("DataReferencia", "") or record.get("Data", ""))


def should_replace_xml_record(current_record: dict[str, object], new_record: dict[str, object]) -> bool:
    current_dt = get_xml_reference_datetime(current_record) or datetime.min
    new_dt = get_xml_reference_datetime(new_record) or datetime.min
    if new_dt != current_dt:
        return new_dt > current_dt

    current_priority = get_nf_status_priority(current_record.get("StatusNF", current_record.get("Status", "")))
    new_priority = get_nf_status_priority(new_record.get("StatusNF", new_record.get("Status", "")))
    if new_priority != current_priority:
        return new_priority > current_priority

    current_type = str(current_record.get("TipoXML", "normal") or "normal")
    new_type = str(new_record.get("TipoXML", "normal") or "normal")
    if current_type != new_type:
        return new_type == "evento"

    return False


def sort_xml_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    return sorted(
        records,
        key=lambda record: (
            normalize_nf(record.get("NF", "")),
            get_xml_reference_datetime(record) or datetime.min,
            str(record.get("Arquivo", "") or "").strip().upper(),
        ),
    )


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


def salvar_xmls_processados_json(xml_files: list) -> tuple[dict[str, int], list[str]]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    existing_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    existing_separacao_records, _ = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    locked_separacao_groups = group_separacao_records_by_identity(existing_separacao_records)
    locked_identities = {
        identity
        for identity, records in locked_separacao_groups.items()
        if is_separacao_group_locked(records)
    }
    storage_lookup: dict[str, dict[str, object]] = {}
    issues: list[str] = []
    summary = {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0, "ignoradas_duplicadas": 0}

    for existing_record in existing_records:
        normalized_record = serialize_xml_record(existing_record)
        identity = get_xml_identity(normalized_record)
        if identity:
            storage_lookup[identity] = normalized_record

    for xml_file in xml_files or []:
        xml_data = parse_xml_file(xml_file)
        if xml_data.get("Erro"):
            issues.append(f"Erro no XML {xml_data.get('Arquivo', 'arquivo.xml')}: {xml_data.get('Status', '')}")
            summary["ignoradas_duplicadas"] += 1
            continue

        serialized = serialize_xml_record(xml_data)
        identity = get_xml_identity(serialized)
        if not identity:
            issues.append(f"XML sem chave/NF identificavel: {serialized.get('Arquivo', 'arquivo.xml')}")
            summary["ignoradas_duplicadas"] += 1
            continue

        if identity in locked_identities:
            issues.append(f"NF {serialized.get('NF', '--')} ignorada no upload porque ja esta separada.")
            summary["ignoradas_separadas"] += 1
            continue

        current_record = storage_lookup.get(identity)
        if current_record is None:
            storage_lookup[identity] = serialized
            summary["novas"] += 1
            continue

        if should_replace_xml_record(current_record, serialized):
            storage_lookup[identity] = serialized
            summary["atualizadas"] += 1
            issues.append(f"NF {serialized.get('NF', '--')} atualizada pelo evento/XML mais recente.")
        else:
            summary["ignoradas_duplicadas"] += 1
            issues.append(f"XML duplicado ou desatualizado ignorado: {serialized.get('Arquivo', 'arquivo.xml')}")

    serialized_records = sort_xml_records(list(storage_lookup.values()))

    XMLS_PROCESSADOS_JSON_PATH.write_text(
        json.dumps(serialized_records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    carregar_xmls_processados_json.clear()
    return summary, issues


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


def serialize_separacao_record(record: dict[str, object]) -> dict[str, object]:
    descricao = str(record.get("Descricao", record.get("Produto", "")) or "").strip()
    codigo = str(record.get("cProd", "") or "").strip()
    produto = str(record.get("Produto", "") or "").strip()
    if not produto:
        produto = format_product_description(descricao, codigo)

    lote_id = str(record.get("lote_id", record.get("Lote", "")) or "").strip()
    data_hora_criacao = str(record.get("data_hora_criacao", record.get("Data Hora Criação", "")) or "").strip()
    status_lote = str(record.get("status_lote", record.get("Status Lote", "")) or "").strip()

    return {
        "NF": normalize_nf(record.get("NF", "")),
        "Chave": normalize_chave_nfe(record.get("Chave", "") or record.get("ChaveNFe", "")),
        "Produto": produto or "Sem produto detalhado",
        "Qtd": parse_float(record.get("Qtd", 0.0)),
        "Tipo": str(record.get("Tipo", record.get("Unidade", "")) or "").strip(),
        "Cliente": str(record.get("Cliente", record.get("Destinatario", "")) or "").strip(),
        "Setor": normalize_sector_name(record.get("Setor", "")) or "Não Identificados",
        "Rota": str(record.get("Rota", record.get("ROTA", UNDEFINED_ROUTE_LABEL)) or UNDEFINED_ROUTE_LABEL).strip() or UNDEFINED_ROUTE_LABEL,
        "Lote": lote_id,
        "Status NF": normalize_nf_status(record.get("Status NF", record.get("StatusNF", record.get("Status", "")))),
        "Status": str(record.get("Status", SEPARATION_PENDING_STATUS) or SEPARATION_PENDING_STATUS).strip(),
        "Municipio": str(record.get("Municipio", "") or "").strip(),
        "cProd": codigo,
        "Arquivo": str(record.get("Arquivo", "") or "").strip(),
        "Data Hora Criação": data_hora_criacao,
        "Status Lote": status_lote,
        "lote_id": lote_id,
        "data_hora_criacao": data_hora_criacao,
        "status_lote": status_lote,
    }


def create_empty_separacao_df() -> pd.DataFrame:
    return pd.DataFrame(columns=[*SEPARATION_VISIBLE_COLUMNS, "Status", "Municipio", "cProd", "Arquivo", "Chave", "Data Hora Criação", "Status Lote"])


def sort_separacao_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    status_order = {SEPARATION_PENDING_STATUS: 0, SEPARATION_SEPARATED_STATUS: 1}
    return sorted(
        records,
        key=lambda record: (
            is_canceled_nf_status(record.get("Status NF", "")) is False,
            status_order.get(str(record.get("Status", "")).strip(), 9),
            str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL).upper(),
            normalize_nf(record.get("NF", "")),
            str(record.get("Produto", "") or "").upper(),
        ),
    )


@st.cache_data(show_spinner=False)
def carregar_separacao_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    path = Path(json_path)
    if not path.is_file():
        return [], ""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        return [], f"A base de separacao nao pôde ser lida ({exc})."

    if not isinstance(payload, list):
        return [], "A base de separacao esta em formato invalido."

    return sort_separacao_records([serialize_separacao_record(item) for item in payload if isinstance(item, dict)]), ""


def salvar_separacao_json(records: list[dict[str, object]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    serialized_records = [serialize_separacao_record(record) for record in records]
    SEPARACAO_JSON_PATH.write_text(
        json.dumps(sort_separacao_records(serialized_records), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    carregar_separacao_json.clear()


def get_separacao_identity(record: dict[str, object]) -> str:
    normalized_record = serialize_separacao_record(record)
    chave = normalize_chave_nfe(normalized_record.get("Chave", ""))
    if chave:
        return chave
    return normalize_nf(normalized_record.get("NF", ""))


def build_lote_payload(lote_id: str, data_hora_criacao: str, status_lote: str) -> dict[str, str]:
    return {
        "lote_id": str(lote_id or "").strip(),
        "data_hora_criacao": str(data_hora_criacao or "").strip(),
        "status_lote": str(status_lote or "").strip(),
    }


def get_lote_info_from_record(record: dict[str, object]) -> dict[str, str]:
    normalized_record = serialize_separacao_record(record)
    return build_lote_payload(
        normalized_record.get("lote_id", normalized_record.get("Lote", "")),
        normalized_record.get("data_hora_criacao", normalized_record.get("Data Hora Criação", "")),
        normalized_record.get("status_lote", normalized_record.get("Status Lote", "")),
    )


def get_lote_sort_key(lote_info: dict[str, str]) -> tuple[datetime, str]:
    lote_datetime = parse_xml_datetime(lote_info.get("data_hora_criacao", "")) or datetime.min
    return lote_datetime, lote_info.get("lote_id", "")


def get_open_lotes(records: list[dict[str, object]]) -> list[dict[str, str]]:
    open_lotes: dict[str, dict[str, str]] = {}
    for record in records:
        lote_info = get_lote_info_from_record(record)
        lote_id = lote_info.get("lote_id", "")
        if not lote_id or lote_info.get("status_lote") != LOT_STATUS_OPEN:
            continue
        current = open_lotes.get(lote_id)
        if current is None or get_lote_sort_key(lote_info) > get_lote_sort_key(current):
            open_lotes[lote_id] = lote_info

    return sorted(open_lotes.values(), key=get_lote_sort_key, reverse=True)


def generate_lote_id(records: list[dict[str, object]]) -> str:
    today = datetime.now().strftime("%Y%m%d")
    sequence = 0
    for record in records:
        lote_id = str(record.get("lote_id", record.get("Lote", "")) or "").strip()
        match = re.fullmatch(r"LOTE-(\d{8})-(\d{3})", lote_id)
        if not match:
            continue
        lote_date, lote_sequence = match.groups()
        if lote_date == today:
            sequence = max(sequence, int(lote_sequence))

    return f"LOTE-{today}-{sequence + 1:03d}"


def create_new_lote(records: list[dict[str, object]]) -> dict[str, str]:
    return build_lote_payload(
        generate_lote_id(records),
        datetime.now().isoformat(timespec="seconds"),
        LOT_STATUS_OPEN,
    )


def ensure_lote_atual(records: list[dict[str, object]]) -> dict[str, str]:
    open_lotes = get_open_lotes(records)
    session_lote = st.session_state.get("lote_atual")
    if open_lotes:
        if isinstance(session_lote, dict) and session_lote.get("lote_id") in {lote.get("lote_id") for lote in open_lotes}:
            normalized_session_lote = build_lote_payload(
                session_lote.get("lote_id", ""),
                session_lote.get("data_hora_criacao", ""),
                session_lote.get("status_lote", LOT_STATUS_OPEN),
            )
            st.session_state["lote_atual"] = normalized_session_lote
            return normalized_session_lote

        st.session_state["lote_atual"] = open_lotes[0]
        return open_lotes[0]

    if isinstance(session_lote, dict) and session_lote.get("lote_id") and session_lote.get("status_lote") == LOT_STATUS_OPEN:
        normalized_session_lote = build_lote_payload(
            session_lote.get("lote_id", ""),
            session_lote.get("data_hora_criacao", ""),
            session_lote.get("status_lote", LOT_STATUS_OPEN),
        )
        st.session_state["lote_atual"] = normalized_session_lote
        return normalized_session_lote

    if not records:
        empty_lote = build_lote_payload("", "", "")
        st.session_state["lote_atual"] = empty_lote
        return empty_lote

    new_lote = create_new_lote(records)
    st.session_state["lote_atual"] = new_lote
    return new_lote


def get_lote_records(records: list[dict[str, object]], lote_id: str) -> list[dict[str, object]]:
    normalized_lote_id = str(lote_id or "").strip()
    if not normalized_lote_id:
        return []
    return group_lote_records(records).get(normalized_lote_id, [])


@st.cache_data(show_spinner=False)
def group_lote_records(records: list[dict[str, object]]) -> dict[str, list[dict[str, object]]]:
    grouped_records: dict[str, list[dict[str, object]]] = {}
    for record in records:
        normalized_record = serialize_separacao_record(record)
        lote_id = str(normalized_record.get("Lote", "") or "").strip()
        if not lote_id:
            continue
        grouped_records.setdefault(lote_id, []).append(normalized_record)
    return grouped_records


def assign_nf_to_lote(records: list[dict[str, object]], chave: str, lote_atual: dict[str, str]) -> list[dict[str, object]]:
    updated_records: list[dict[str, object]] = []
    for record in records:
        normalized = serialize_separacao_record(record)
        if normalized.get("Chave") == chave:
            normalized["Status"] = SEPARATION_SEPARATED_STATUS
            normalized["Lote"] = lote_atual.get("lote_id", "")
            normalized["lote_id"] = lote_atual.get("lote_id", "")
            normalized["Data Hora Criação"] = lote_atual.get("data_hora_criacao", "")
            normalized["data_hora_criacao"] = lote_atual.get("data_hora_criacao", "")
            normalized["Status Lote"] = lote_atual.get("status_lote", LOT_STATUS_OPEN)
            normalized["status_lote"] = lote_atual.get("status_lote", LOT_STATUS_OPEN)
        updated_records.append(normalized)
    return sort_separacao_records(updated_records)


def remove_nf_from_lote(records: list[dict[str, object]], nf: str, lote_id: str) -> list[dict[str, object]]:
    updated_records: list[dict[str, object]] = []
    normalized_nf = normalize_nf(nf)
    for record in records:
        normalized = serialize_separacao_record(record)
        if normalized.get("NF") == normalized_nf and normalized.get("Lote") == lote_id and normalized.get("Status Lote") == LOT_STATUS_OPEN:
            normalized["Status"] = SEPARATION_PENDING_STATUS
            normalized["Lote"] = ""
            normalized["lote_id"] = ""
            normalized["Data Hora Criação"] = ""
            normalized["data_hora_criacao"] = ""
            normalized["Status Lote"] = ""
            normalized["status_lote"] = ""
        updated_records.append(normalized)
    return sort_separacao_records(updated_records)


def close_lote(records: list[dict[str, object]], lote_id: str) -> list[dict[str, object]]:
    updated_records: list[dict[str, object]] = []
    for record in records:
        normalized = serialize_separacao_record(record)
        if normalized.get("Lote") == lote_id:
            normalized["Status Lote"] = LOT_STATUS_CLOSED
            normalized["status_lote"] = LOT_STATUS_CLOSED
        updated_records.append(normalized)
    return sort_separacao_records(updated_records)


def serialize_lote_record(record: dict[str, object]) -> dict[str, object]:
    lote_id = str(record.get("lote_id", "") or "").strip()
    status = str(record.get("status", record.get("status_lote", LOT_STATUS_OPEN)) or LOT_STATUS_OPEN).strip()
    data_abertura = str(record.get("data_abertura", record.get("data_hora_criacao", "")) or "").strip()
    data_fechamento = str(record.get("data_fechamento", "") or "").strip()
    raw_nfs = record.get("nfs", []) or []
    nfs = sorted({normalize_nf(nf) for nf in raw_nfs if normalize_nf(nf)})

    return {
        "lote_id": lote_id,
        "status": status if status in {LOT_STATUS_OPEN, LOT_STATUS_CLOSED} else LOT_STATUS_OPEN,
        "data_abertura": data_abertura,
        "data_fechamento": data_fechamento,
        "nfs": nfs,
    }


@st.cache_data(show_spinner=False)
def carregar_lotes_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    path = Path(json_path)
    if not path.is_file():
        return [], ""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        return [], f"A base de lotes nao pôde ser lida ({exc})."

    if not isinstance(payload, list):
        return [], "A base de lotes esta em formato invalido."

    return [serialize_lote_record(item) for item in payload if isinstance(item, dict)], ""


def salvar_lotes_json(records: list[dict[str, object]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    normalized_records: list[dict[str, object]] = []
    for record in records:
        normalized = serialize_lote_record(record)
        if normalized.get("lote_id"):
            normalized_records.append(normalized)

    normalized_records = enrich_lote_registry_dates(normalized_records)

    normalized_records = sorted(
        normalized_records,
        key=lambda record: (
            parse_xml_datetime(record.get("data_abertura", "")) or datetime.min,
            record.get("lote_id", ""),
        ),
        reverse=True,
    )
    LOTES_JSON_PATH.write_text(json.dumps(normalized_records, ensure_ascii=False, indent=2), encoding="utf-8")
    carregar_lotes_json.clear()


def enrich_lote_registry_dates(records: list[dict[str, object]]) -> list[dict[str, object]]:
    normalized_records = [serialize_lote_record(record) for record in records if serialize_lote_record(record).get("lote_id")]
    ordered_records = sorted(
        normalized_records,
        key=lambda record: (
            parse_xml_datetime(record.get("data_abertura", "")) or datetime.min,
            record.get("lote_id", ""),
        ),
    )

    for index, record in enumerate(ordered_records):
        if record.get("status") != LOT_STATUS_CLOSED:
            continue

        abertura_dt = parse_xml_datetime(record.get("data_abertura", ""))
        fechamento_dt = parse_xml_datetime(record.get("data_fechamento", ""))
        if abertura_dt is None:
            continue
        if fechamento_dt is not None and fechamento_dt >= abertura_dt:
            continue

        next_abertura_dt = None
        for next_record in ordered_records[index + 1 :]:
            candidate_dt = parse_xml_datetime(next_record.get("data_abertura", ""))
            if candidate_dt is not None and candidate_dt >= abertura_dt:
                next_abertura_dt = candidate_dt
                break

        inferred_fechamento = next_abertura_dt or abertura_dt
        record["data_fechamento"] = inferred_fechamento.isoformat(timespec="seconds")

    return ordered_records


def build_lote_registry_entry(
    lote_id: str,
    lote_records: list[dict[str, object]],
    existing_record: dict[str, object] | None = None,
    lote_info: dict[str, str] | None = None,
    status_override: str | None = None,
    fechamento_override: str | None = None,
) -> dict[str, object]:
    existing_record = serialize_lote_record(existing_record or {})
    lote_info = lote_info or {}
    normalized_records = [serialize_separacao_record(record) for record in lote_records]
    abertura_candidates = [
        str(record.get("Data Hora Criação", "") or "").strip()
        for record in normalized_records
        if str(record.get("Data Hora Criação", "") or "").strip()
    ]

    data_abertura = ""
    if abertura_candidates:
        parsed_candidates = [parse_xml_datetime(value) for value in abertura_candidates]
        parsed_candidates = [value for value in parsed_candidates if value is not None]
        if parsed_candidates:
            data_abertura = min(parsed_candidates).isoformat(timespec="seconds")
        else:
            data_abertura = abertura_candidates[0]

    if not data_abertura:
        data_abertura = str(lote_info.get("data_hora_criacao", existing_record.get("data_abertura", "")) or "").strip()

    inferred_closed = any(record.get("Status Lote") == LOT_STATUS_CLOSED for record in normalized_records)
    if status_override:
        status = status_override
    elif inferred_closed:
        status = LOT_STATUS_CLOSED
    else:
        status = str(lote_info.get("status_lote", existing_record.get("status", "")) or "").strip() or LOT_STATUS_OPEN

    data_fechamento = str(fechamento_override or existing_record.get("data_fechamento", "") or "").strip()
    if status == LOT_STATUS_CLOSED and not data_fechamento and abertura_candidates:
        parsed_candidates = [parse_xml_datetime(value) for value in abertura_candidates]
        parsed_candidates = [value for value in parsed_candidates if value is not None]
        if parsed_candidates:
            data_fechamento = max(parsed_candidates).isoformat(timespec="seconds")
        else:
            data_fechamento = abertura_candidates[-1]
    if status != LOT_STATUS_CLOSED:
        data_fechamento = ""

    nfs = sorted({record.get("NF", "") for record in normalized_records if record.get("NF", "")})
    return serialize_lote_record(
        {
            "lote_id": lote_id,
            "status": status,
            "data_abertura": data_abertura,
            "data_fechamento": data_fechamento,
            "nfs": nfs,
        }
    )


def sync_lote_registry_entry(
    lote_id: str,
    records: list[dict[str, object]],
    lote_info: dict[str, str] | None = None,
    status_override: str | None = None,
    fechamento_override: str | None = None,
) -> None:
    normalized_lote_id = str(lote_id or "").strip()
    if not normalized_lote_id:
        return

    lote_records = get_lote_records(records, normalized_lote_id)
    lote_registry, _ = carregar_lotes_json(str(LOTES_JSON_PATH))
    lookup = {record.get("lote_id", ""): serialize_lote_record(record) for record in lote_registry if record.get("lote_id", "")}
    lookup[normalized_lote_id] = build_lote_registry_entry(
        normalized_lote_id,
        lote_records,
        existing_record=lookup.get(normalized_lote_id),
        lote_info=lote_info,
        status_override=status_override,
        fechamento_override=fechamento_override,
    )
    salvar_lotes_json(list(lookup.values()))


def sync_lotes_registry(records: list[dict[str, object]], current_lote: dict[str, str] | None = None) -> None:
    lote_registry, _ = carregar_lotes_json(str(LOTES_JSON_PATH))
    lookup = {record.get("lote_id", ""): serialize_lote_record(record) for record in lote_registry if record.get("lote_id", "")}
    lote_ids = set(lookup)
    lote_ids.update(
        str(serialize_separacao_record(record).get("Lote", "") or "").strip()
        for record in records
        if str(serialize_separacao_record(record).get("Lote", "") or "").strip()
    )
    if isinstance(current_lote, dict) and current_lote.get("lote_id"):
        lote_ids.add(str(current_lote.get("lote_id", "") or "").strip())

    updated_lookup: dict[str, dict[str, object]] = {}
    for current_lote_id in lote_ids:
        lote_info = current_lote if isinstance(current_lote, dict) and current_lote.get("lote_id") == current_lote_id else None
        updated_lookup[current_lote_id] = build_lote_registry_entry(
            current_lote_id,
            get_lote_records(records, current_lote_id),
            existing_record=lookup.get(current_lote_id),
            lote_info=lote_info,
        )

    if list(updated_lookup.values()) != list(lookup.values()):
        salvar_lotes_json(list(updated_lookup.values()))


def excluir_lote(lote_id: str) -> list[dict[str, object]]:
    normalized_lote_id = str(lote_id or "").strip()
    if not normalized_lote_id:
        return []

    lotes_registry, _ = carregar_lotes_json(str(LOTES_JSON_PATH))
    updated_lotes = [record for record in lotes_registry if record.get("lote_id", "") != normalized_lote_id]
    salvar_lotes_json(updated_lotes)

    separacao_records, _ = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    updated_records: list[dict[str, object]] = []
    for record in separacao_records:
        normalized_record = serialize_separacao_record(record)
        if normalized_record.get("Lote") == normalized_lote_id or normalized_record.get("lote_id") == normalized_lote_id:
            normalized_record["Status"] = SEPARATION_PENDING_STATUS
            normalized_record["Lote"] = ""
            normalized_record["lote_id"] = ""
            normalized_record["Data Hora Criação"] = ""
            normalized_record["data_hora_criacao"] = ""
            normalized_record["Status Lote"] = ""
            normalized_record["status_lote"] = ""
        updated_records.append(normalized_record)

    salvar_separacao_json(updated_records)
    return sort_separacao_records(updated_records)


@st.cache_data(show_spinner=False)
def build_lote_catalog(records: list[dict[str, object]], lotes_metadata: list[dict[str, object]]) -> list[dict[str, object]]:
    lote_records_lookup = group_lote_records(records)
    catalog: list[dict[str, object]] = []
    for lote_record in lotes_metadata:
        lote_id = str(lote_record.get("lote_id", "") or "").strip()
        if not lote_id:
            continue
        lote_items = lote_records_lookup.get(lote_id, [])
        nfs = sorted({record.get("NF", "") for record in lote_items if record.get("NF", "")}) or lote_record.get("nfs", []) or []
        item_count = len(lote_items)
        catalog.append(
            {
                "Lote": lote_id,
                "Status": lote_record.get("status", LOT_STATUS_OPEN),
                "Abertura": lote_record.get("data_abertura", ""),
                "Fechamento": lote_record.get("data_fechamento", ""),
                "NFs": len(nfs),
                "Itens": item_count,
                "nfs": nfs,
            }
        )

    return sorted(
        catalog,
        key=lambda record: (parse_xml_datetime(record.get("Abertura", "")) or datetime.min, record.get("Lote", "")),
        reverse=True,
    )


def get_latest_closed_lote_summary(records: list[dict[str, object]]) -> dict[str, object] | None:
    lotes_metadata, _ = carregar_lotes_json(str(LOTES_JSON_PATH))
    catalog = build_lote_catalog(records, lotes_metadata)
    for record in catalog:
        if record.get("Status") == LOT_STATUS_CLOSED:
            return record
    return None


def format_lote_datetime_display(value: object) -> str:
    parsed = parse_xml_datetime(value)
    if parsed is None:
        return "--"
    return format_datetime_display(parsed)


def style_lote_status_badge(status: object) -> str:
    normalized_status = str(status or "").strip()
    if normalized_status == LOT_STATUS_CLOSED:
        bg_color = "#EAF7EE"
        fg_color = "#18794E"
    else:
        bg_color = "#EFF6FF"
        fg_color = "#1D4ED8"
    return (
        f"display:inline-block;padding:6px 12px;border-radius:999px;"
        f"background:{bg_color};color:{fg_color};font-weight:700;font-size:0.86rem;"
    )


@st.cache_data(show_spinner=False)
def build_lote_detail_dataframe(records: list[dict[str, object]], lote_id: str) -> pd.DataFrame:
    lote_records = group_lote_records(records).get(str(lote_id or "").strip(), [])
    if not lote_records:
        return pd.DataFrame(columns=["NF", "Código Produto", "Descrição", "Quantidade", "Tipo", "Cliente", "Setor", "Rota"])

    detail_df = pd.DataFrame(
        [
            {
                "NF": record.get("NF", ""),
                "Código Produto": record.get("cProd", ""),
                "Descrição": record.get("Produto", ""),
                "Quantidade": parse_float(record.get("Qtd", 0.0)),
                "Tipo": record.get("Tipo", ""),
                "Cliente": record.get("Cliente", ""),
                "Setor": record.get("Setor", "Não Identificados"),
                "Rota": record.get("Rota", UNDEFINED_ROUTE_LABEL),
            }
            for record in lote_records
        ]
    )
    return detail_df.sort_values(by=["Setor", "Rota", "NF"], ascending=[True, True, True], na_position="last")


@st.cache_data(show_spinner=False)
def build_lote_catalog_dataframe(catalog: list[dict[str, object]], lote_records_lookup: dict[str, list[dict[str, object]]]) -> pd.DataFrame:
    if not catalog:
        return pd.DataFrame(columns=["Lote", "Status", "Abertura", "Fechamento", "NFs", "Itens", "nfs", "AberturaData", "_lote_norm", "_search_blob"])

    catalog_df = pd.DataFrame(catalog)
    catalog_df["AberturaData"] = pd.to_datetime(catalog_df["Abertura"], errors="coerce").dt.date
    catalog_df["_lote_norm"] = catalog_df["Lote"].fillna("").astype(str).str.strip().str.upper()

    search_lookup: dict[str, str] = {}
    for lote_id, lote_records in lote_records_lookup.items():
        lote_df = pd.DataFrame(lote_records)
        search_lookup[lote_id] = " ".join(
            [
                str(lote_id or "").strip().lower(),
                " ".join(sorted({str(value or "").strip().lower() for value in lote_df.get("NF", pd.Series(dtype=str)).tolist() if str(value or "").strip()})),
                " ".join(sorted({str(value or "").strip().lower() for value in lote_df.get("Cliente", pd.Series(dtype=str)).tolist() if str(value or "").strip()})),
                " ".join(sorted({str(value or "").strip().lower() for value in lote_df.get("Produto", pd.Series(dtype=str)).tolist() if str(value or "").strip()})),
                " ".join(sorted({str(value or "").strip().lower() for value in lote_df.get("Setor", pd.Series(dtype=str)).tolist() if str(value or "").strip()})),
                " ".join(sorted({str(value or "").strip().lower() for value in lote_df.get("Rota", pd.Series(dtype=str)).tolist() if str(value or "").strip()})),
            ]
        ).strip()

    catalog_df["_search_blob"] = catalog_df["Lote"].map(lambda value: search_lookup.get(str(value or "").strip(), ""))
    return catalog_df




def build_lote_detail_styler(dataframe: pd.DataFrame):
    if dataframe.empty:
        return dataframe
    styler = dataframe.style
    if "Setor" in dataframe.columns:
        styler = styler.map(style_separacao_setor_cell, subset=["Setor"])
    return styler


def _generate_lote_pdf_document(
    lote_summary: dict[str, object],
    lote_records: list[dict[str, object]],
    report_type: str = "Completo",
    report_filter: str = "Todos",
    numero_carga: str = "--",
    data_emissao: str = "--",
) -> bytes:
    regular_font, bold_font = register_pdf_fonts()
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    page_width, page_height = A4
    left_margin = 40
    right_margin = page_width - 40
    top_margin = page_height - 45
    bottom_margin = 55
    line_height = 12
    section_gap = 16

    normalized_records = [serialize_separacao_record(record) for record in lote_records]
    report_type = str(report_type or "Completo").strip() or "Completo"
    report_filter = str(report_filter or "Todos").strip() or "Todos"
    normalized_report_filter = normalize_sector_name(report_filter)
    filters_group_by_nf = report_type == "Por Setor" and normalized_report_filter == "Filtros"

    xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    xml_lookup_by_identity: dict[str, dict[str, object]] = {}
    for xml_record in xml_records:
        serialized_xml = serialize_xml_record(xml_record)
        identity = get_xml_identity(serialized_xml)
        if identity:
            xml_lookup_by_identity[identity] = serialized_xml

    if filters_group_by_nf:
        grouped_records: dict[tuple[str, str, str], list[dict[str, object]]] = {}
        sorted_records = sorted(
            normalized_records,
            key=lambda item: (
                str(item.get("Rota", "")).upper(),
                normalize_nf(item.get("NF", "")),
                str(item.get("Produto", "")).upper(),
            ),
        )
        for record in sorted_records:
            group_key = (
                str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL),
                str(record.get("NF", "") or "--"),
                str(record.get("Cliente", "") or "--"),
            )
            grouped_records.setdefault(group_key, []).append(record)
    elif report_type == "Por Setor":
        grouped_records: dict[tuple[str, str, str], list[dict[str, object]]] = {}
        sorted_records = sorted(
            normalized_records,
            key=lambda item: (
                str(item.get("Rota", "")).upper(),
                str(item.get("Produto", "")).upper(),
                normalize_nf(item.get("NF", "")),
            ),
        )
        for record in sorted_records:
            group_key = (
                str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL),
                str(record.get("Produto", "Sem produto detalhado") or "Sem produto detalhado"),
                str(record.get("cProd", "") or "").strip(),
            )
            grouped_records.setdefault(group_key, []).append(record)
    elif report_type == "Por Rota":
        grouped_records = {}
        sorted_records = sorted(
            normalized_records,
            key=lambda item: (
                str(item.get("Setor", "")).upper(),
                str(item.get("Produto", "")).upper(),
                normalize_nf(item.get("NF", "")),
            ),
        )
        for record in sorted_records:
            group_key = (
                str(record.get("Setor", "Não Identificados") or "Não Identificados"),
                str(record.get("Produto", "Sem produto detalhado") or "Sem produto detalhado"),
                str(record.get("cProd", "") or "").strip(),
            )
            grouped_records.setdefault(group_key, []).append(record)
    else:
        grouped_records = {}
        sorted_records = sorted(
            normalized_records,
            key=lambda item: (
                str(item.get("Setor", "")).upper(),
                str(item.get("Rota", "")).upper(),
                str(item.get("Produto", "")).upper(),
                normalize_nf(item.get("NF", "")),
            ),
        )
        for record in sorted_records:
            group_key = (
                str(record.get("Setor", "Não Identificados") or "Não Identificados"),
                str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL),
                str(record.get("Produto", "Sem produto detalhado") or "Sem produto detalhado"),
                str(record.get("cProd", "") or "").strip(),
            )
            grouped_records.setdefault(group_key, []).append(record)

    lote_id = str(lote_summary.get("Lote", "--") or "--")
    report_type_label = report_type
    report_filter_label = report_filter

    unique_nf_identities = []
    seen_identities: set[str] = set()
    for record in normalized_records:
        identity = normalize_chave_nfe(record.get("Chave", "")) or normalize_nf(record.get("NF", ""))
        if identity and identity not in seen_identities:
            seen_identities.add(identity)
            unique_nf_identities.append(identity)

    total_volumes = 0.0
    total_peso = 0.0
    for identity in unique_nf_identities:
        xml_record = xml_lookup_by_identity.get(identity)
        if not xml_record:
            continue
        total_volumes += parse_float(xml_record.get("VolumeTotal", 0.0))
        total_peso += parse_float(xml_record.get("PesoTotal", 0.0))

    if total_volumes <= 0:
        total_volumes = float(len({record.get("NF", "") for record in normalized_records if record.get("NF", "")}))

    def format_qty(value: float) -> str:
        return format_quantity_display(parse_float(value))

    def wrap_text(text: object, font_name: str, font_size: int, width: float) -> list[str]:
        lines = simpleSplit(str(text or "--"), font_name, font_size, width)
        return lines or ["--"]

    def draw_header(y_pos: float, continuation: bool = False) -> float:
        pdf.setFont(bold_font, 19 if not continuation else 15)
        pdf.drawCentredString(page_width / 2, y_pos, "MINUTA DE SEPARAÇÃO")
        y_pos -= 24
        pdf.setFont(bold_font, 13)
        pdf.drawCentredString(page_width / 2, y_pos, "BRIDA LUBRIFICANTES LTDA")
        y_pos -= 24

        pdf.setFont(regular_font, 10)
        pdf.drawString(left_margin, y_pos, f"Carregamento: {numero_carga}")
        y_pos -= 15
        pdf.drawString(left_margin, y_pos, f"Data emissão: {data_emissao}")
        y_pos -= 15
        pdf.drawString(left_margin, y_pos, f"Lote: {lote_id}")
        y_pos -= 15
        pdf.drawString(left_margin, y_pos, f"Tipo: {report_type_label.upper()}")
        y_pos -= 15
        pdf.drawString(left_margin, y_pos, f"Filtro: {report_filter_label.upper()}")
        y_pos -= 16

        pdf.setStrokeColor(colors.HexColor("#C9D2DE"))
        pdf.line(left_margin, y_pos, right_margin, y_pos)
        return y_pos - 18

    def ensure_space(current_y: float, required_height: float) -> float:
        if current_y - required_height < bottom_margin:
            pdf.showPage()
            return draw_header(top_margin, continuation=True)
        return current_y

    def should_use_compact_layout(setor: str) -> bool:
        return normalize_sector_name(setor) == "Filtros"

    def draw_product_separator(current_y: float, compact_layout: bool) -> float:
        if compact_layout:
            return current_y - 6

        current_y -= 2
        current_y = ensure_space(current_y, 14)
        pdf.setStrokeColor(colors.HexColor("#2B2B2B"))
        pdf.setLineWidth(2)
        pdf.line(left_margin, current_y, right_margin, current_y)
        pdf.setLineWidth(1)
        return current_y - 16

    def get_zebra_background(item_row_index: int) -> str | None:
        return "#f2f2f2" if item_row_index % 2 == 1 else None

    table_text_style = ParagraphStyle(
        "MinutaTableText",
        fontName=regular_font,
        fontSize=10,
        leading=12,
        textColor=colors.black,
        alignment=TA_LEFT,
    )

    def build_table_paragraph(value: object) -> Paragraph:
        safe_text = html.escape(str(value or "--")).replace("\n", "<br/>")
        return Paragraph(safe_text, table_text_style)

    def build_items_table(
        headers: list[str],
        body_rows: list[list[object]],
        col_widths: list[float],
        paragraph_columns: set[int],
        right_align_columns: set[int],
        center_align_columns: set[int],
    ) -> Table:
        table_data: list[list[object]] = [headers]
        for row in body_rows:
            table_row: list[object] = []
            for column_index, value in enumerate(row):
                if column_index in paragraph_columns:
                    table_row.append(build_table_paragraph(value))
                else:
                    table_row.append(str(value or "--"))
            table_data.append(table_row)

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        style_commands: list[tuple[object, ...]] = [
            ("FONTNAME", (0, 0), (-1, 0), bold_font),
            ("FONTNAME", (0, 1), (-1, -1), regular_font),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("LEADING", (0, 0), (-1, -1), 12),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, 0), 0),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ("ALIGN", (0, 0), (-1, 0), "LEFT"),
            ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#E5E7EB")),
        ]

        for column_index in right_align_columns:
            style_commands.append(("ALIGN", (column_index, 0), (column_index, -1), "RIGHT"))
        for column_index in center_align_columns:
            style_commands.append(("ALIGN", (column_index, 0), (column_index, -1), "CENTER"))

        for row_index in range(len(body_rows)):
            background_color = get_zebra_background(row_index)
            if background_color:
                style_commands.append(
                    ("BACKGROUND", (0, row_index + 1), (-1, row_index + 1), colors.HexColor(background_color))
                )

        table.setStyle(TableStyle(style_commands))
        return table

    def draw_items_table(
        current_y: float,
        headers: list[str],
        body_rows: list[list[object]],
        col_widths: list[float],
        paragraph_columns: set[int],
        right_align_columns: set[int],
        center_align_columns: set[int],
    ) -> float:
        available_width = right_margin - left_margin
        pending_rows: list[list[object]] = list(body_rows)

        while pending_rows:
            chunk_rows: list[list[object]] = []
            last_fitting_table: Table | None = None
            last_fitting_height = 0.0
            available_height = current_y - bottom_margin

            for row in pending_rows:
                candidate_rows = [*chunk_rows, row]
                candidate_table = build_items_table(
                    headers,
                    candidate_rows,
                    col_widths,
                    paragraph_columns,
                    right_align_columns,
                    center_align_columns,
                )
                _, candidate_height = candidate_table.wrap(available_width, available_height)

                if chunk_rows and candidate_height > available_height:
                    break

                chunk_rows = candidate_rows
                last_fitting_table = candidate_table
                last_fitting_height = candidate_height

            if last_fitting_table is None:
                pdf.showPage()
                current_y = draw_header(top_margin, continuation=True)
                continue

            last_fitting_table.drawOn(pdf, left_margin, current_y - last_fitting_height)
            current_y -= last_fitting_height
            pending_rows = pending_rows[len(chunk_rows) :]

            if pending_rows:
                pdf.showPage()
                current_y = draw_header(top_margin, continuation=True)

        return current_y

    def draw_group_header(current_y: float, left_label: str, left_value: str, right_label: str = "", right_value: str = "") -> float:
        current_y = ensure_space(current_y, 24)
        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, current_y, f"{left_label}: {left_value.upper()}")
        if right_label and right_value:
            pdf.drawString(left_margin + 175, current_y, f"{right_label}: {right_value.upper()}")
        return current_y - 18

    def draw_product_header(current_y: float, produto: str, codigo: str, total_qtd: float) -> float:
        current_y = ensure_space(current_y, 22)
        produto_base = re.sub(r"\s*-\s*\([^()]+\)\s*$", "", str(produto or "").strip()) or str(produto or "--")
        produto_text = f"PRODUTO: {produto_base}"
        if codigo:
            produto_text = f"{produto_text} - ({codigo})"
        produto_text = f"{produto_text} - TOTAL {format_qty(total_qtd)}"
        lines = wrap_text(produto_text, bold_font, 10, right_margin - left_margin)
        pdf.setFont(bold_font, 10)
        for line in lines:
            pdf.drawString(left_margin, current_y, line)
            current_y -= line_height
        return current_y - 6

    def draw_nf_header(current_y: float, nf: str, cliente: str, total_qtd: float) -> float:
        current_y = ensure_space(current_y, 38)
        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, current_y, f"NF: {nf}")
        pdf.drawRightString(right_margin, current_y, f"TOTAL ITENS NF: {format_qty(total_qtd)}")
        current_y -= 16

        cliente_lines = wrap_text(f"CLIENTE: {cliente}", regular_font, 10, right_margin - left_margin)
        pdf.setFont(regular_font, 10)
        for line in cliente_lines:
            pdf.drawString(left_margin, current_y, line)
            current_y -= line_height
        return current_y - 4

    current_y = draw_header(top_margin)
    current_primary = None
    current_secondary = None
    previous_group_key = None

    for group_key, product_records in grouped_records.items():
        if filters_group_by_nf:
            rota, nf, cliente = group_key
            if rota != current_primary:
                if current_primary is not None:
                    current_y -= 8
                    current_y = ensure_space(current_y, 8)
                    pdf.setStrokeColor(colors.HexColor("#D8DEE8"))
                    pdf.line(left_margin, current_y, right_margin, current_y)
                    current_y -= section_gap
                current_y = draw_group_header(current_y, "ROTA", rota)
                current_primary = rota

            compact_layout = True

            total_qtd = sum(parse_float(record.get("Qtd", 0.0)) for record in product_records)
            current_y = draw_nf_header(current_y, nf, cliente, total_qtd)
            current_y = draw_items_table(
                current_y,
                ["PRODUTO", "QTDE", "UN"],
                [
                    [
                        record.get("Produto", "--"),
                        format_qty(record.get("Qtd", 0.0)),
                        str(record.get("Tipo", "") or "--"),
                    ]
                    for record in product_records
                ],
                [423, 52, 40],
                {0},
                {1},
                {2},
            )
            current_y = draw_product_separator(current_y, compact_layout)
            previous_group_key = rota
            continue
        elif report_type == "Por Setor":
            rota, produto, codigo = group_key
            if rota != current_primary:
                if current_primary is not None:
                    current_y -= 8
                    current_y = ensure_space(current_y, 8)
                    pdf.setStrokeColor(colors.HexColor("#D8DEE8"))
                    pdf.line(left_margin, current_y, right_margin, current_y)
                    current_y -= section_gap
                current_y = draw_group_header(current_y, "ROTA", rota)
                current_primary = rota
            compact_layout = should_use_compact_layout(normalized_report_filter)
            group_separator_key = rota
        elif report_type == "Por Rota":
            setor, produto, codigo = group_key
            if setor != current_primary:
                if current_primary is not None:
                    current_y -= 8
                    current_y = ensure_space(current_y, 8)
                    pdf.setStrokeColor(colors.HexColor("#D8DEE8"))
                    pdf.line(left_margin, current_y, right_margin, current_y)
                    current_y -= section_gap
                current_y = draw_group_header(current_y, "SETOR", setor)
                current_primary = setor
            compact_layout = should_use_compact_layout(setor)
            group_separator_key = setor
        else:
            setor, rota, produto, codigo = group_key
            if (setor, rota) != (current_primary, current_secondary):
                if current_primary is not None:
                    current_y -= 8
                    current_y = ensure_space(current_y, 8)
                    pdf.setStrokeColor(colors.HexColor("#D8DEE8"))
                    pdf.line(left_margin, current_y, right_margin, current_y)
                    current_y -= section_gap
                current_y = draw_group_header(current_y, "SETOR", setor, "ROTA", rota)
                current_primary, current_secondary = setor, rota
            compact_layout = should_use_compact_layout(setor)
            group_separator_key = (setor, rota)

        if previous_group_key is not None and previous_group_key != group_separator_key:
            current_y -= 2

        total_qtd = sum(parse_float(record.get("Qtd", 0.0)) for record in product_records)
        current_y = draw_product_header(current_y, produto, codigo, total_qtd)
        current_y = draw_items_table(
            current_y,
            ["NF", "QTDE", "UN", "CLIENTE"],
            [
                [
                    str(record.get("NF", "--") or "--"),
                    format_qty(record.get("Qtd", 0.0)),
                    str(record.get("Tipo", "") or "--"),
                    record.get("Cliente", "--"),
                ]
                for record in sorted(product_records, key=lambda item: normalize_nf(item.get("NF", "")))
            ],
            [95, 60, 45, 315],
            {3},
            {1},
            {2},
        )
        current_y = draw_product_separator(current_y, compact_layout)
        previous_group_key = group_separator_key

    current_y = ensure_space(current_y, 78)
    pdf.setStrokeColor(colors.HexColor("#C9D2DE"))
    pdf.line(left_margin, current_y, right_margin, current_y)
    current_y -= 20

    pdf.setFont(bold_font, 12)
    pdf.drawString(left_margin, current_y, "TOTAL GERAL")
    current_y -= 18
    pdf.setFont(regular_font, 10)
    pdf.drawString(left_margin + 4, current_y, f"Volumes: {format_qty(total_volumes)}")
    current_y -= 15
    pdf.drawString(left_margin + 4, current_y, f"Peso: {format_qty(total_peso)} kg")
    current_y -= 34

    pdf.setStrokeColor(colors.HexColor("#6B7280"))
    pdf.line(left_margin, current_y, left_margin + 180, current_y)
    current_y -= 14
    pdf.setFont(regular_font, 10)
    pdf.drawString(left_margin + 30, current_y, "Ass. do conferente")

    pdf.save()
    return buffer.getvalue()


@st.cache_data(show_spinner=False)
def generate_lote_pdf_cached(
    lote_summary: dict[str, object],
    lote_records: list[dict[str, object]],
    report_type: str,
    report_filter: str,
    numero_carga: str,
    data_emissao: str,
) -> bytes:
    return _generate_lote_pdf_document(lote_summary, lote_records, report_type, report_filter, numero_carga, data_emissao)


def generate_lote_pdf(
    lote_summary: dict[str, object],
    lote_records: list[dict[str, object]],
    report_type: str = "Completo",
    report_filter: str = "Todos",
) -> bytes:
    session_summary = st.session_state.get("summary", {}) if hasattr(st, "session_state") else {}
    numero_carga = "--"
    if isinstance(session_summary, dict):
        numero_carga = str(session_summary.get("numero_carga", "--") or "--")
    data_emissao = format_datetime_display()
    return generate_lote_pdf_cached(lote_summary, lote_records, report_type, report_filter, numero_carga, data_emissao)


def open_pdf_for_print(pdf_bytes: bytes, title: str) -> None:
    if not pdf_bytes:
        return
    encoded_pdf = base64.b64encode(pdf_bytes).decode("ascii")
    safe_title = html.escape(title)
    components.html(
        f"""
        <script>
        const pdfData = "data:application/pdf;base64,{encoded_pdf}";
        const printWindow = window.open("", "_blank");
        if (printWindow) {{
            printWindow.document.write(`
                <html>
                    <head><title>{safe_title}</title></head>
                    <body style=\"margin:0\">
                        <iframe src=\"${{pdfData}}\" style=\"border:0;width:100vw;height:100vh\"></iframe>
                    </body>
                </html>
            `);
            printWindow.document.close();
            setTimeout(() => printWindow.print(), 700);
        }}
        </script>
        """,
        height=0,
    )


def group_separacao_records_by_identity(records: list[dict[str, object]]) -> dict[str, list[dict[str, object]]]:
    grouped_records: dict[str, list[dict[str, object]]] = {}
    for record in records:
        normalized_record = serialize_separacao_record(record)
        identity = get_separacao_identity(normalized_record)
        if not identity:
            continue
        grouped_records.setdefault(identity, []).append(normalized_record)
    return grouped_records


def is_separacao_group_locked(records: list[dict[str, object]]) -> bool:
    return any(str(record.get("Status", "")).strip() == SEPARATION_SEPARATED_STATUS for record in records)


def build_separacao_records_from_xml_records(
    xml_records: list[dict[str, object]],
    classificacao_records: list[dict[str, str]],
    existing_records: list[dict[str, object]] | None = None,
    excluded_identities: set[str] | None = None,
) -> tuple[list[dict[str, object]], list[str], dict[str, int]]:
    existing_records = existing_records or []
    excluded_identities = excluded_identities or set()
    existing_groups = group_separacao_records_by_identity(existing_records)
    remaining_existing_groups = {identity: [serialize_separacao_record(record) for record in records] for identity, records in existing_groups.items()}

    separacao_records: list[dict[str, object]] = []
    issues: list[str] = []
    summary = {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0}
    for xml_record in xml_records or []:
        normalized_xml = serialize_xml_record(xml_record)
        chave = normalize_chave_nfe(normalized_xml.get("ChaveNFe", ""))
        nf = normalize_nf(normalized_xml.get("NF", "") or normalized_xml.get("nf_normalizada", ""))
        identity = chave or nf
        if not identity:
            issues.append(f"NF {nf or '--'} ignorada no mapa de separacao por nao possuir chave valida.")
            continue

        existing_group = remaining_existing_groups.pop(identity, [])
        if identity in excluded_identities:
            continue
        existing_group_locked = bool(existing_group and is_separacao_group_locked(existing_group))

        status_nf = normalize_nf_status(normalized_xml.get("StatusNF", normalized_xml.get("Status", "")))
        status_operacional = existing_group[0].get("Status", SEPARATION_PENDING_STATUS) if existing_group else SEPARATION_PENDING_STATUS
        if is_canceled_nf_status(status_nf) and not existing_group_locked:
            status_operacional = SEPARATION_PENDING_STATUS

        lote_payload = get_lote_info_from_record(existing_group[0]) if existing_group_locked else build_lote_payload("", "", "")
        if existing_group_locked:
            summary["ignoradas_separadas"] += 1

        route = str(normalized_xml.get("ROTA", "") or get_route_for_municipio(normalized_xml.get("Municipio", ""))).strip()
        route = route or UNDEFINED_ROUTE_LABEL
        items = normalized_xml.get("Items", []) or [{"cProd": "", "Descricao": "Sem produto detalhado", "Qtd": 0.0, "Unidade": ""}]

        for item in items:
            descricao = str(item.get("Descricao", "") or "").strip() or "Sem produto detalhado"
            separacao_records.append(
                serialize_separacao_record(
                    {
                        "NF": nf,
                        "Chave": chave,
                        "Descricao": descricao,
                        "Produto": format_product_description(descricao, item.get("cProd", "")),
                        "Qtd": item.get("Qtd", 0.0),
                        "Tipo": item.get("Unidade", ""),
                        "Cliente": normalized_xml.get("Destinatario", ""),
                        "Setor": classify_product_sector(descricao, classificacao_records),
                        "Rota": route,
                        "Status NF": status_nf,
                        "Status": status_operacional,
                        "Municipio": normalized_xml.get("Municipio", ""),
                        "cProd": item.get("cProd", ""),
                        "Arquivo": normalized_xml.get("Arquivo", ""),
                        "Lote": lote_payload.get("lote_id", ""),
                        "lote_id": lote_payload.get("lote_id", ""),
                        "Data Hora Criação": lote_payload.get("data_hora_criacao", ""),
                        "data_hora_criacao": lote_payload.get("data_hora_criacao", ""),
                        "Status Lote": lote_payload.get("status_lote", ""),
                        "status_lote": lote_payload.get("status_lote", ""),
                    }
                )
            )

        if existing_group:
            summary["atualizadas"] += 1
        else:
            summary["novas"] += 1

    for leftover_group in remaining_existing_groups.values():
        separacao_records.extend(leftover_group)

    return sort_separacao_records(separacao_records), issues, summary


def sincronizar_base_separacao(
    xml_records: list[dict[str, object]],
    classificacao_records: list[dict[str, str]],
) -> tuple[list[dict[str, object]], list[str], str, dict[str, int]]:
    existing_records, storage_error = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    excluded_identities = carregar_separacao_excluidos_json(str(SEPARACAO_EXCLUIDOS_JSON_PATH))
    if not xml_records:
        return existing_records, [], storage_error, {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0}

    rebuilt_records, issues, summary = build_separacao_records_from_xml_records(
        xml_records,
        classificacao_records,
        existing_records=existing_records,
        excluded_identities=excluded_identities,
    )

    current_records = sort_separacao_records([serialize_separacao_record(record) for record in existing_records])
    if storage_error or current_records != rebuilt_records:
        salvar_separacao_json(rebuilt_records)
        return rebuilt_records, issues, storage_error, summary

    return current_records, issues, storage_error, {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0}


def get_separacao_storage_status() -> tuple[bool, str]:
    if not SEPARACAO_JSON_PATH.is_file():
        return False, ""

    updated_at = datetime.fromtimestamp(SEPARACAO_JSON_PATH.stat().st_mtime)
    return True, format_datetime_display(updated_at)


@st.cache_data(show_spinner=False)
def carregar_separacao_excluidos_json(json_path: str) -> set[str]:
    path = Path(json_path)
    if not path.is_file():
        return set()

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return set()

    if not isinstance(payload, list):
        return set()

    return {
        str(item).strip()
        for item in payload
        if str(item or "").strip()
    }


def salvar_separacao_excluidos_json(identities: set[str]) -> None:
    normalized_identities = sorted({str(identity or "").strip() for identity in identities if str(identity or "").strip()})
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if not normalized_identities:
        if SEPARACAO_EXCLUIDOS_JSON_PATH.is_file():
            SEPARACAO_EXCLUIDOS_JSON_PATH.unlink()
        carregar_separacao_excluidos_json.clear()
        return

    SEPARACAO_EXCLUIDOS_JSON_PATH.write_text(
        json.dumps(normalized_identities, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    carregar_separacao_excluidos_json.clear()


def parse_flexible_datetime(value: object) -> datetime | None:
    parsed = parse_xml_datetime(value)
    if parsed is not None:
        return parsed

    text = str(value or "").strip()
    if not text:
        return None

    parsed_fallback = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(parsed_fallback):
        return None

    if isinstance(parsed_fallback, pd.Timestamp):
        if parsed_fallback.tzinfo is not None:
            parsed_fallback = parsed_fallback.tz_convert(None)
        return parsed_fallback.to_pydatetime()

    return None


def coerce_input_date(value: object):
    if isinstance(value, datetime):
        return value.date()

    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return value

    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        raise ValueError("Informe um período válido para a limpeza.")
    if isinstance(parsed, pd.Timestamp):
        return parsed.date()
    raise ValueError("Informe um período válido para a limpeza.")


def is_datetime_within_period(value: object, start_date, end_date) -> bool:
    parsed = parse_flexible_datetime(value)
    if parsed is None:
        return False
    return start_date <= parsed.date() <= end_date


def is_separacao_cleanup_status(value: object) -> bool:
    normalized = normalize_matching_text(value)
    return normalized in {"SEPARADO", "FINALIZADO"}


def get_xml_cleanup_reference(record: dict[str, object]) -> str:
    normalized_record = serialize_xml_record(record)
    return str(
        normalized_record.get("DataReferenciaISO", "")
        or normalized_record.get("DataReferencia", "")
        or normalized_record.get("Data", "")
    ).strip()


def build_lote_lookup(records: list[dict[str, object]]) -> dict[str, dict[str, object]]:
    return {
        str(serialize_lote_record(record).get("lote_id", "") or "").strip(): serialize_lote_record(record)
        for record in records
        if str(serialize_lote_record(record).get("lote_id", "") or "").strip()
    }


def rebuild_lote_registry_from_separacao(
    separacao_records: list[dict[str, object]],
    existing_lotes_records: list[dict[str, object]],
) -> list[dict[str, object]]:
    existing_lookup = build_lote_lookup(existing_lotes_records)
    remaining_lote_ids = sorted(
        {
            str(serialize_separacao_record(record).get("Lote", "") or "").strip()
            for record in separacao_records
            if str(serialize_separacao_record(record).get("Lote", "") or "").strip()
        }
    )

    rebuilt_records = [
        build_lote_registry_entry(
            lote_id,
            get_lote_records(separacao_records, lote_id),
            existing_record=existing_lookup.get(lote_id),
        )
        for lote_id in remaining_lote_ids
    ]
    return [serialize_lote_record(record) for record in rebuilt_records if str(record.get("lote_id", "") or "").strip()]


def get_separacao_cleanup_reference(record: dict[str, object], lote_lookup: dict[str, dict[str, object]]) -> str:
    normalized_record = serialize_separacao_record(record)
    data_hora_criacao = str(normalized_record.get("Data Hora Criação", "") or "").strip()
    if data_hora_criacao:
        return data_hora_criacao

    lote_id = str(normalized_record.get("Lote", normalized_record.get("lote_id", "")) or "").strip()
    lote_record = lote_lookup.get(lote_id, {}) if lote_id else {}
    return str(lote_record.get("data_fechamento", "") or lote_record.get("data_abertura", "") or "").strip()


def clear_lote_metadata_from_separacao_record(record: dict[str, object]) -> dict[str, object]:
    normalized_record = serialize_separacao_record(record)
    normalized_record["Lote"] = ""
    normalized_record["lote_id"] = ""
    normalized_record["Data Hora Criação"] = ""
    normalized_record["data_hora_criacao"] = ""
    normalized_record["Status Lote"] = ""
    normalized_record["status_lote"] = ""
    return normalized_record


def executar_limpeza_dados_sistema(data_inicial: object, data_final: object, tipo_limpeza: str) -> dict[str, object]:
    start_date = coerce_input_date(data_inicial)
    end_date = coerce_input_date(data_final)
    if start_date > end_date:
        raise ValueError("A data inicial não pode ser maior que a data final.")

    xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    separacao_records, _ = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    lotes_records, _ = carregar_lotes_json(str(LOTES_JSON_PATH))
    excluded_identities = carregar_separacao_excluidos_json(str(SEPARACAO_EXCLUIDOS_JSON_PATH))

    current_xml_records = [serialize_xml_record(record) for record in xml_records]
    current_separacao_records = [serialize_separacao_record(record) for record in separacao_records]
    current_lotes_records = [serialize_lote_record(record) for record in lotes_records]
    lote_lookup = build_lote_lookup(current_lotes_records)

    if tipo_limpeza == DATA_CLEANUP_TYPE_COMPLETE:
        removed_xmls = len(current_xml_records)
        removed_separacao = len(current_separacao_records)
        removed_lotes = len(current_lotes_records)

        salvar_separacao_json([])
        salvar_lotes_json([])
        XMLS_PROCESSADOS_JSON_PATH.write_text("[]", encoding="utf-8")
        carregar_xmls_processados_json.clear()
        salvar_separacao_excluidos_json(set())

        return {
            "tipo_limpeza": tipo_limpeza,
            "periodo": f"{start_date.strftime('%d/%m/%Y')} até {end_date.strftime('%d/%m/%Y')}",
            "xmls_removidos": removed_xmls,
            "separacao_removidos": removed_separacao,
            "lotes_removidos": removed_lotes,
            "xmls_protegidos": 0,
            "lotes_protegidos": 0,
            "total_removido": removed_xmls + removed_separacao + removed_lotes,
            "separacao_records": [],
        }

    removed_separacao_identities: set[str] = set()
    removed_xmls = 0
    removed_separacao = 0
    removed_lotes = 0
    protected_xmls = 0
    protected_lotes = 0
    separacao_changed = False
    xml_changed = False
    lotes_changed = False

    if tipo_limpeza == DATA_CLEANUP_TYPE_XML:
        removable_xml_identities: set[str] = set()
        remaining_xml_records: list[dict[str, object]] = []
        for record in current_xml_records:
            identity = get_xml_identity(record)
            within_period = is_datetime_within_period(get_xml_cleanup_reference(record), start_date, end_date)
            if within_period and identity:
                removable_xml_identities.add(identity)
                removed_xmls += 1
                continue
            remaining_xml_records.append(record)

        if removable_xml_identities:
            updated_separacao_records: list[dict[str, object]] = []
            for record in current_separacao_records:
                if get_separacao_identity(record) in removable_xml_identities:
                    removed_separacao += 1
                    continue
                updated_separacao_records.append(record)
            current_separacao_records = sort_separacao_records(updated_separacao_records)
            rebuilt_lotes_records = rebuild_lote_registry_from_separacao(current_separacao_records, current_lotes_records)
            removed_lotes = max(0, len(current_lotes_records) - len(rebuilt_lotes_records))
            current_lotes_records = rebuilt_lotes_records
            separacao_changed = True
            lotes_changed = True

        current_xml_records = sort_xml_records(remaining_xml_records)
        xml_changed = removed_xmls > 0

    if tipo_limpeza in {DATA_CLEANUP_TYPE_SEPARACAO, DATA_CLEANUP_TYPE_COMPLETE}:
        remaining_separacao_records: list[dict[str, object]] = []
        for record in current_separacao_records:
            is_open_lote = str(record.get("Status Lote", "") or "").strip() == LOT_STATUS_OPEN
            within_period = is_datetime_within_period(get_separacao_cleanup_reference(record, lote_lookup), start_date, end_date)
            if is_open_lote or not is_separacao_cleanup_status(record.get("Status", "")) or not within_period:
                remaining_separacao_records.append(record)
                continue

            removed_separacao += 1
            identity = get_separacao_identity(record)
            if identity:
                removed_separacao_identities.add(identity)

        current_separacao_records = sort_separacao_records(remaining_separacao_records)
        separacao_changed = removed_separacao > 0
        if separacao_changed and tipo_limpeza != DATA_CLEANUP_TYPE_XML:
            rebuilt_lotes_records = rebuild_lote_registry_from_separacao(current_separacao_records, current_lotes_records)
            removed_lotes += max(0, len(current_lotes_records) - len(rebuilt_lotes_records))
            current_lotes_records = rebuilt_lotes_records
            lotes_changed = True

    if tipo_limpeza in {DATA_CLEANUP_TYPE_LOTES, DATA_CLEANUP_TYPE_COMPLETE}:
        removable_lote_ids: set[str] = set()
        remaining_lotes_records: list[dict[str, object]] = []
        for record in current_lotes_records:
            lote_id = str(record.get("lote_id", "") or "").strip()
            status = str(record.get("status", "") or "").strip()
            reference_date = str(record.get("data_fechamento", "") or record.get("data_abertura", "") or "").strip()
            if status == LOT_STATUS_OPEN and is_datetime_within_period(reference_date, start_date, end_date):
                protected_lotes += 1
                remaining_lotes_records.append(record)
                continue

            if status == LOT_STATUS_CLOSED and is_datetime_within_period(reference_date, start_date, end_date):
                removable_lote_ids.add(lote_id)
                removed_lotes += 1
                continue

            remaining_lotes_records.append(record)

        if removable_lote_ids:
            updated_separacao_records: list[dict[str, object]] = []
            for record in current_separacao_records:
                lote_id = str(record.get("Lote", record.get("lote_id", "")) or "").strip()
                if lote_id in removable_lote_ids:
                    updated_separacao_records.append(clear_lote_metadata_from_separacao_record(record))
                else:
                    updated_separacao_records.append(record)
            current_separacao_records = sort_separacao_records(updated_separacao_records)
            separacao_changed = True

        current_lotes_records = remaining_lotes_records
        lotes_changed = removed_lotes > 0

    if tipo_limpeza == DATA_CLEANUP_TYPE_COMPLETE:
        referenced_identities = {
            get_separacao_identity(record)
            for record in current_separacao_records
            if get_separacao_identity(record)
        }
        remaining_xml_records: list[dict[str, object]] = []
        for record in current_xml_records:
            identity = get_xml_identity(record)
            within_period = is_datetime_within_period(get_xml_cleanup_reference(record), start_date, end_date)
            if not within_period:
                remaining_xml_records.append(record)
                continue

            if identity and identity in referenced_identities:
                protected_xmls += 1
                remaining_xml_records.append(record)
                continue

            removed_xmls += 1

        current_xml_records = sort_xml_records(remaining_xml_records)
        xml_changed = removed_xmls > 0

    remaining_xml_identities = {get_xml_identity(record) for record in current_xml_records if get_xml_identity(record)}
    updated_excluded_identities = {identity for identity in excluded_identities if identity in remaining_xml_identities}
    updated_excluded_identities.update(identity for identity in removed_separacao_identities if identity in remaining_xml_identities)

    if separacao_changed:
        salvar_separacao_json(current_separacao_records)
    if lotes_changed:
        salvar_lotes_json(current_lotes_records)
    if xml_changed:
        XMLS_PROCESSADOS_JSON_PATH.write_text(
            json.dumps(current_xml_records, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        carregar_xmls_processados_json.clear()
    salvar_separacao_excluidos_json(updated_excluded_identities)

    return {
        "tipo_limpeza": tipo_limpeza,
        "periodo": f"{start_date.strftime('%d/%m/%Y')} até {end_date.strftime('%d/%m/%Y')}",
        "xmls_removidos": removed_xmls,
        "separacao_removidos": removed_separacao,
        "lotes_removidos": removed_lotes,
        "xmls_protegidos": protected_xmls,
        "lotes_protegidos": protected_lotes,
        "total_removido": removed_xmls + removed_separacao + removed_lotes,
        "separacao_records": current_separacao_records,
    }


def format_file_size_mb(path: Path) -> str:
    try:
        size_bytes = path.stat().st_size
    except OSError:
        size_bytes = 0
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def invalidate_runtime_data() -> None:
    st.session_state["runtime_data_signature"] = None
    st.session_state["runtime_operational_signature"] = None
    st.session_state["runtime_xml_records"] = []
    st.session_state["runtime_classificacao_records"] = []
    st.session_state["runtime_refresh_required"] = True


def get_path_cache_token(path: Path) -> int:
    try:
        return path.stat().st_mtime_ns
    except OSError:
        return 0


def build_search_blob_series(dataframe: pd.DataFrame, columns: list[str]) -> pd.Series:
    if dataframe.empty:
        return pd.Series("", index=dataframe.index, dtype="object")

    parts: list[pd.Series] = []
    for column in columns:
        if column in dataframe.columns:
            parts.append(dataframe[column].fillna("").astype(str).str.lower())

    if not parts:
        return pd.Series("", index=dataframe.index, dtype="object")

    search_blob = parts[0]
    for part in parts[1:]:
        search_blob = search_blob.str.cat(part, sep=" ")

    return search_blob.str.replace(r"\s+", " ", regex=True).str.strip()


@st.cache_data(show_spinner=False)
def prepare_processed_search_dataframe(dataframe: pd.DataFrame, route_version: int) -> pd.DataFrame:
    _ = route_version
    prepared_df = apply_routes_to_dataframe(dataframe)
    prepared_df["_search_blob"] = build_search_blob_series(
        prepared_df,
        ["NF", "cProd", "Descricao", "Destinatario", "ROTA", "Status"],
    )
    return prepared_df


@st.cache_data(show_spinner=False)
def build_separacao_dataframe(records: list[dict[str, object]]) -> pd.DataFrame:
    if not records:
        return create_empty_separacao_df()
    return pd.DataFrame([serialize_separacao_record(record) for record in records])


def style_separacao_setor_cell(value: object) -> str:
    colors = get_sector_colors(str(value or "").strip() or "Não Identificados")
    return "; ".join(
        [
            f"background-color: {colors['bg']}",
            f"color: {colors['fg']}",
            f"border: 1px solid {colors['border']}",
            "font-weight: 700",
            "text-align: center",
            "border-radius: 8px",
        ]
    )


def style_lote_cell(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return "; ".join(
        [
            "background-color: #EEF4FF",
            "color: #1D4ED8",
            "font-weight: 700",
            "text-align: center",
            "border-radius: 8px",
        ]
    )


def summarize_separacao(records: list[dict[str, object]]) -> dict[str, int]:
    if not records:
        return {"nf_total": 0, "pendentes": 0, "separadas": 0, "canceladas": 0, "lotes_fechados": 0}

    df = build_separacao_dataframe(records)
    grouped = df.groupby("Chave", dropna=False).first().reset_index()
    return {
        "nf_total": int(grouped["Chave"].nunique()),
        "pendentes": int((grouped["Status"] == SEPARATION_PENDING_STATUS).sum()),
        "separadas": int((grouped["Status"] == SEPARATION_SEPARATED_STATUS).sum()),
        "canceladas": int(grouped["Status NF"].apply(is_canceled_nf_status).sum()),
        "lotes_fechados": int(grouped["Status Lote"].eq(LOT_STATUS_CLOSED).sum()),
    }


@st.cache_data(show_spinner=False)
def group_separacao_records_by_chave(records: list[dict[str, object]]) -> dict[str, list[dict[str, object]]]:
    grouped_records: dict[str, list[dict[str, object]]] = {}
    for record in records:
        normalized_record = serialize_separacao_record(record)
        chave = str(normalized_record.get("Chave", "") or "").strip()
        if not chave:
            continue
        grouped_records.setdefault(chave, []).append(normalized_record)
    return grouped_records


def render_scan_input_focus() -> None:
    components.html(
        """
        <script>
        const focusScanInput = () => {
            const input = window.parent.document.querySelector('input[aria-label="Bipar ou digitar chave da NF"]');
            if (input) {
                input.focus();
                input.select();
            }
        };
        window.parent.requestAnimationFrame(() => setTimeout(focusScanInput, 60));
        </script>
        """,
        height=0,
    )


def build_separacao_result(records: list[dict[str, object]], chave: str) -> dict[str, str] | None:
    matching_records = group_separacao_records_by_chave(records).get(chave, [])
    if not matching_records:
        return None

    setores = sorted({record.get("Setor", "Não Identificados") for record in matching_records})
    produtos = len(matching_records)
    return {
        "NF": matching_records[0].get("NF", "--"),
        "Cliente": matching_records[0].get("Cliente", "--") or "--",
        "Rota": matching_records[0].get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL,
        "Lote": matching_records[0].get("Lote", "") or "Sem lote",
        "Setor": setores[0] if len(setores) == 1 else "Misto",
        "Setores": ", ".join(setores),
        "Status NF": matching_records[0].get("Status NF", "Status nao informado"),
        "Status": matching_records[0].get("Status", SEPARATION_PENDING_STATUS),
        "Status Lote": matching_records[0].get("Status Lote", "") or "Sem lote",
        "Produtos": str(produtos),
    }


def apply_current_sector_classification(
    records: list[dict[str, object]],
    classification_records: list[dict[str, str]],
) -> list[dict[str, object]]:
    updated_records: list[dict[str, object]] = []
    for record in records:
        normalized_record = serialize_separacao_record(record)
        description_source = str(normalized_record.get("Produto", "") or "").strip()
        normalized_record["Setor"] = classify_product_sector(description_source, classification_records)
        updated_records.append(normalized_record)
    return updated_records


def atualizar_status_separacao_por_chave(records: list[dict[str, object]], chave: str) -> list[dict[str, object]]:
    updated_records: list[dict[str, object]] = []
    for record in records:
        normalized = serialize_separacao_record(record)
        if normalized.get("Chave") == chave and not is_canceled_nf_status(normalized.get("Status NF", "")):
            normalized["Status"] = SEPARATION_SEPARATED_STATUS
        updated_records.append(normalized)
    return sort_separacao_records(updated_records)


def render_highlight_card(title: str, value: object, accent_color: str, secondary: str = "") -> None:
    safe_title = html.escape(str(title or ""))
    safe_value = html.escape(format_summary_value(value))
    safe_secondary = html.escape(str(secondary or "")).replace("\n", "<br>") if secondary else "&nbsp;"
    st.markdown(
        f"""
    <div class="erp-card erp-card-info operation-card" style="border-top: 4px solid {accent_color};">
        <div class="erp-card-title">{safe_title}</div>
        <div class="erp-card-value">{safe_value}</div>
        <div class="erp-card-secondary">{safe_secondary}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )


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


def normalize_screen_name(value: object) -> str:
    screen = str(value or "").strip().lower()
    screen_aliases = {
        "gestao_lotes": SCREEN_LOTES,
        SCREEN_LOGIN: SCREEN_LOGIN,
        SCREEN_MENU: SCREEN_MENU,
        SCREEN_MINUTA: SCREEN_MINUTA,
        SCREEN_SEPARACAO: SCREEN_SEPARACAO,
        SCREEN_LOTES: SCREEN_LOTES,
    }
    return screen_aliases.get(screen, SCREEN_MENU)


def navegar(tela: str) -> None:
    target_screen = normalize_screen_name(tela)
    st.session_state["tela"] = target_screen
    if target_screen != SCREEN_LOGIN:
        st.session_state["menu_aberto"] = st.session_state.get("menu_aberto", True)
    st.rerun()


def initialize_navigation_state() -> None:
    legacy_screen = st.session_state.get("tela_atual", SCREEN_MINUTA)
    current_screen = normalize_screen_name(st.session_state.get("tela", legacy_screen))

    if not st.session_state.get("logado", False):
        st.session_state["tela"] = SCREEN_LOGIN
    else:
        st.session_state["tela"] = current_screen if current_screen != SCREEN_LOGIN else SCREEN_MENU


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


def render_box_open(extra_classes: str = "") -> None:
    classes = f"ui-section-box {extra_classes}".strip()
    st.markdown(f'<div class="{classes}">', unsafe_allow_html=True)


def render_box_close() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


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
                st.session_state["tela"] = SCREEN_MENU
                st.rerun()
            else:
                st.session_state["login_error"] = "Usuario ou senha incorretos."
                st.session_state["login_success"] = ""

    st.markdown('</div>', unsafe_allow_html=True)


def logout() -> None:
    st.session_state["logado"] = False
    st.session_state["tela"] = SCREEN_LOGIN
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
        render_box_open("is-sidebar is-soft")
        st.markdown(
            f'''
        <div class="sidebar-heading with-icon">{render_label_icon(ICON_MAP["dados_gerais"])}<span>Arquivos</span></div>
        ''',
            unsafe_allow_html=True,
        )
        render_box_close()

        render_box_open("is-sidebar")
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

        if PRACAS_JSON_PATH.is_file():
            route_updated_at = format_datetime_display(datetime.fromtimestamp(PRACAS_JSON_PATH.stat().st_mtime))
            st.caption(f"Base de rotas carregada • Ultima atualizacao: {route_updated_at}")
        render_box_close()

        render_box_open("is-sidebar")
        st.markdown(
            f'''
        <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["setor"])}<span>Classificação de Produtos</span></div>
        ''',
            unsafe_allow_html=True,
        )

        classificacao_file = st.file_uploader(
            "Atualizar classificação",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="classificacao_upload_widget",
        )

        if classificacao_file is None:
            st.session_state.classificacao_upload_signature = ""
        else:
            file_bytes = classificacao_file.getvalue()
            current_signature = hashlib.sha256(file_bytes).hexdigest()
            if current_signature != st.session_state.classificacao_upload_signature:
                try:
                    st.session_state.classificacao_upload_message = update_classificacao_produtos_json(classificacao_file)
                    st.session_state.classificacao_upload_error = ""
                    st.session_state.classificacao_upload_signature = current_signature
                    st.session_state["runtime_refresh_required"] = True
                except ValueError as exc:
                    st.session_state.classificacao_upload_message = ""
                    st.session_state.classificacao_upload_error = str(exc)
                    st.session_state.classificacao_upload_signature = ""

        if st.session_state.classificacao_upload_message:
            st.success(st.session_state.classificacao_upload_message)
        if st.session_state.classificacao_upload_error:
            st.error(st.session_state.classificacao_upload_error)

        has_classificacao_storage, classificacao_updated_at = get_classificacao_storage_status()
        if has_classificacao_storage:
            st.caption(f"Base de setores carregada • Ultima atualizacao: {classificacao_updated_at}")
        else:
            st.caption(classificacao_updated_at)
        render_box_close()

        render_box_open("is-sidebar")
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
                import_summary, issues = salvar_xmls_processados_json(xml_files)
                xml_records, xml_storage_error = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
                st.session_state.xml_upload_signature = current_signature
                st.session_state["runtime_refresh_required"] = True
                st.session_state.xml_upload_message = (
                    "Base atualizada: "
                    f"{import_summary.get('novas', 0)} novas NFs adicionadas • "
                    f"{import_summary.get('atualizadas', 0)} NFs atualizadas • "
                    f"{import_summary.get('ignoradas_separadas', 0)} NFs ignoradas (já separadas)"
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
        render_box_close()

        if normalize_screen_name(st.session_state.get("tela", SCREEN_MENU)) != SCREEN_MINUTA:
            return xml_records, None, False

        render_box_open("is-sidebar")
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
        render_box_close()

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
    route_version = get_path_cache_token(PRACAS_JSON_PATH)
    processed_df = prepare_processed_search_dataframe(st.session_state.processed_df, route_version)

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

    normalized_search_term = str(search_term or "").strip().lower()
    if normalized_search_term and not processed_df.empty:
        filtered_df = processed_df[processed_df["_search_blob"].str.contains(normalized_search_term, na=False)]
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


def render_separacao_screen(
    separacao_records: list[dict[str, object]],
    sync_issues: list[str],
    separacao_storage_error: str,
    import_summary: dict[str, int],
) -> None:
    current_records = st.session_state.get("separacao_records", separacao_records)
    separacao_lookup = group_separacao_records_by_chave(current_records)
    lote_atual = ensure_lote_atual(current_records)
    sync_lotes_registry(current_records, lote_atual)
    current_lote_records = get_lote_records(current_records, lote_atual.get("lote_id", ""))
    current_lote_nfs = sorted({record.get("NF", "") for record in current_lote_records if record.get("NF", "")})
    latest_closed_lote = get_latest_closed_lote_summary(current_records)
    latest_closed_lote_id = str((latest_closed_lote or {}).get("Lote", "") or "").strip()
    latest_closed_records = get_lote_records(current_records, latest_closed_lote_id) if latest_closed_lote_id else []
    latest_closed_pdf_bytes = b""
    if latest_closed_lote_id and latest_closed_records:
        latest_closed_lote["Abertura Formatada"] = format_datetime_display(parse_xml_datetime(latest_closed_lote.get("Abertura", ""))) or "--"
        latest_closed_lote["Fechamento Formatada"] = format_datetime_display(parse_xml_datetime(latest_closed_lote.get("Fechamento", ""))) or "--"
        latest_closed_pdf_bytes = generate_lote_pdf(latest_closed_lote, latest_closed_records)

    st.markdown(
        """
    <div class="page-hero">
        <h2>Mapa de Separação</h2>
        <p>Controle de picking por setor</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    has_storage, updated_at = get_separacao_storage_status()
    if separacao_storage_error:
        st.warning(separacao_storage_error)
    elif has_storage:
        st.caption(f"Base de separacao carregada automaticamente • Ultima atualizacao: {updated_at}")

    if any(import_summary.values()):
        st.success(
            "Base atualizada: "
            f"{import_summary.get('novas', 0)} novas NFs adicionadas • "
            f"{import_summary.get('atualizadas', 0)} NFs atualizadas • "
            f"{import_summary.get('ignoradas_separadas', 0)} NFs ignoradas (já separadas)"
        )

    if sync_issues:
        with st.expander("Avisos da sincronização", expanded=False):
            for issue in sync_issues:
                st.warning(issue)

    st.markdown(
        f"""
    <div class="lot-banner">
        <div class="lot-banner-label">LOTE ATUAL</div>
        <div class="lot-banner-value">{html.escape(lote_atual.get('lote_id', 'SEM LOTE'))}</div>
        <div class="lot-banner-meta">{html.escape(lote_atual.get('status_lote', LOT_STATUS_OPEN))} • {html.escape(format_single_date(lote_atual.get('data_hora_criacao', '')) or lote_atual.get('data_hora_criacao', '--'))}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    lote_action_col_1, lote_action_col_2 = st.columns([1.2, 1.2], gap="medium")
    with lote_action_col_1:
        if st.button("Iniciar novo lote", use_container_width=True):
            if current_lote_records:
                st.session_state["separacao_feedback"] = {
                    "type": "warning",
                    "message": f"Feche ou esvazie o lote {lote_atual.get('lote_id', '--')} antes de iniciar outro.",
                }
            else:
                novo_lote = create_new_lote(current_records)
                st.session_state["lote_atual"] = novo_lote
                sync_lotes_registry(current_records, novo_lote)
                st.session_state["separacao_feedback"] = {
                    "type": "success",
                    "message": f"Novo lote iniciado: {novo_lote.get('lote_id', '--')}",
                }
                st.rerun()
    with lote_action_col_2:
        if st.button("Fechar lote", use_container_width=True, disabled=not current_lote_records):
            updated_records = close_lote(current_records, lote_atual.get("lote_id", ""))
            salvar_separacao_json(updated_records)
            sync_lote_registry_entry(
                lote_atual.get("lote_id", ""),
                updated_records,
                lote_info=lote_atual,
                status_override=LOT_STATUS_CLOSED,
                fechamento_override=datetime.now().isoformat(timespec="seconds"),
            )
            st.session_state["separacao_records"] = updated_records
            novo_lote = create_new_lote(updated_records)
            st.session_state["lote_atual"] = novo_lote
            sync_lotes_registry(updated_records, novo_lote)
            st.session_state["separacao_feedback"] = {
                "type": "success",
                "message": f"Lote {lote_atual.get('lote_id', '--')} fechado. Novo lote disponível: {novo_lote.get('lote_id', '--')}",
            }
            st.rerun()

    reprint_col_1, reprint_col_2 = st.columns([1.2, 1.2], gap="medium")
    with reprint_col_1:
        if st.button(
            "Reimprimir último lote fechado",
            use_container_width=True,
            disabled=not bool(latest_closed_lote_id and latest_closed_pdf_bytes),
            key="reprint_last_closed_lote_button",
        ):
            open_pdf_for_print(latest_closed_pdf_bytes, f"Lote {latest_closed_lote_id}")
    with reprint_col_2:
        st.download_button(
            "Baixar PDF do último lote",
            data=latest_closed_pdf_bytes,
            file_name=f"lote_{sanitize_filename_part(latest_closed_lote_id, 'ultimo_lote')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            disabled=not bool(latest_closed_lote_id and latest_closed_pdf_bytes),
        )

    if latest_closed_lote_id:
        st.caption(f"Último lote fechado disponível para reimpressão: {latest_closed_lote_id}")

    if lote_atual.get("status_lote") == LOT_STATUS_OPEN:
        remove_col_1, remove_col_2 = st.columns([3.8, 1.2], gap="medium")
        with remove_col_1:
            nf_para_remover = st.selectbox(
                "Remover NF do lote atual",
                options=current_lote_nfs,
                key="nf_remover_lote",
                index=None,
                placeholder="Selecione uma NF do lote atual",
                disabled=not current_lote_nfs,
            )
        with remove_col_2:
            st.markdown('<div class="scan-button-spacer"></div>', unsafe_allow_html=True)
            if st.button("Remover NF", use_container_width=True, disabled=not bool(current_lote_nfs and nf_para_remover)):
                updated_records = remove_nf_from_lote(current_records, nf_para_remover or "", lote_atual.get("lote_id", ""))
                salvar_separacao_json(updated_records)
                sync_lotes_registry(updated_records, lote_atual)
                st.session_state["separacao_records"] = updated_records
                st.session_state["separacao_feedback"] = {
                    "type": "success",
                    "message": f"NF {nf_para_remover} removida do lote {lote_atual.get('lote_id', '--')}",
                }
                st.rerun()

    summary = summarize_separacao(current_records)
    render_box_open()
    render_section_heading("Visão Operacional", "separacao")
    metric_col_1, metric_col_2, metric_col_3, metric_col_4 = st.columns(4, gap="medium")
    with metric_col_1:
        render_metric_card("NF", summary["nf_total"], "nf")
    with metric_col_2:
        render_metric_card("Pendentes", summary["pendentes"], "processar")
    with metric_col_3:
        render_metric_card("Separadas", summary["separadas"], "status_operacional")
    with metric_col_4:
        render_metric_card("Lotes Fech.", summary["lotes_fechados"], "erros")
    render_box_close()

    render_box_open()
    render_section_heading("Entrada", "barcode")
    st.markdown('<div class="scan-shell">', unsafe_allow_html=True)
    with st.form("separacao_scan_form", clear_on_submit=True):
        scan_col, action_col = st.columns([4.4, 1.2], gap="medium")
        with scan_col:
            chave_digitada = st.text_input(
                "Bipar ou digitar chave da NF",
                key="input_chave",
                placeholder="Aguardando leitura...",
            )
        with action_col:
            st.markdown('<div class="scan-button-spacer"></div>', unsafe_allow_html=True)
            buscar = st.form_submit_button("Buscar", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    render_scan_input_focus()
    render_box_close()

    if buscar:
        chave_normalizada = normalize_chave_nfe(chave_digitada)
        current_records = st.session_state.get("separacao_records", separacao_records)
        lote_atual = ensure_lote_atual(current_records)
        if not chave_normalizada:
            st.session_state["separacao_feedback"] = {"type": "error", "message": "Informe uma chave valida com 44 digitos."}
            st.session_state["separacao_result"] = None
        else:
            matching_records = separacao_lookup.get(chave_normalizada, [])
            if not matching_records:
                st.session_state["separacao_feedback"] = {"type": "error", "message": "Chave nao encontrada na base de XMLs processados."}
                st.session_state["separacao_result"] = None
            elif is_canceled_nf_status(matching_records[0].get("Status NF", "")):
                st.session_state["separacao_feedback"] = {"type": "error", "message": "NF CANCELADA - NÃO REALIZAR SEPARAÇÃO"}
                st.session_state["separacao_result"] = build_separacao_result(current_records, chave_normalizada)
            elif matching_records[0].get("Lote"):
                st.session_state["separacao_feedback"] = {
                    "type": "warning",
                    "message": f"NF já vinculada ao lote {matching_records[0].get('Lote', '--')}",
                }
                st.session_state["separacao_result"] = build_separacao_result(current_records, chave_normalizada)
            elif matching_records[0].get("Status") == SEPARATION_SEPARATED_STATUS and matching_records[0].get("Status Lote") == LOT_STATUS_CLOSED:
                st.session_state["separacao_feedback"] = {
                    "type": "error",
                    "message": "NF já separada em lote fechado e não pode voltar ao fluxo.",
                }
                st.session_state["separacao_result"] = build_separacao_result(current_records, chave_normalizada)
            elif all(record.get("Status") == SEPARATION_SEPARATED_STATUS for record in matching_records):
                st.session_state["separacao_feedback"] = {"type": "warning", "message": "Esta NF ja foi separada e nao pode ser processada novamente."}
                st.session_state["separacao_result"] = build_separacao_result(current_records, chave_normalizada)
            else:
                updated_records = assign_nf_to_lote(current_records, chave_normalizada, lote_atual)
                salvar_separacao_json(updated_records)
                sync_lotes_registry(updated_records, lote_atual)
                st.session_state["separacao_records"] = updated_records
                st.session_state["separacao_feedback"] = {
                    "type": "success",
                    "message": f"NF vinculada ao lote {lote_atual.get('lote_id', '--')} com sucesso.",
                }
                st.session_state["separacao_result"] = build_separacao_result(updated_records, chave_normalizada)

    feedback = st.session_state.get("separacao_feedback", {})
    feedback_message = str(feedback.get("message", "") or "")
    if feedback_message:
        if feedback.get("type") == "success":
            st.success(feedback_message)
        elif feedback.get("type") == "warning":
            st.warning(feedback_message)
        else:
            st.error(feedback_message)

    separacao_result = st.session_state.get("separacao_result")
    if separacao_result:
        render_box_open()
        render_section_heading("Resultado da Separação", "status_operacional")
        result_col_1, result_col_2, result_col_3, result_col_4, result_col_5, result_col_6 = st.columns(6, gap="medium")
        with result_col_1:
            render_info_card("NF", separacao_result.get("NF", "--"), "nf")
        with result_col_2:
            render_info_card("Cliente", separacao_result.get("Cliente", "--"), "dados_gerais")
        with result_col_3:
            render_info_card("Rota", separacao_result.get("Rota", UNDEFINED_ROUTE_LABEL), "rota")
        with result_col_4:
            sector_colors = get_sector_colors(separacao_result.get("Setor", "Não Identificados"))
            render_highlight_card("Setor", separacao_result.get("Setor", "--"), sector_colors["border"], separacao_result.get("Setores", ""))
        with result_col_5:
            render_highlight_card("Lote", separacao_result.get("Lote", "Sem lote"), "#1D4ED8", separacao_result.get("Status Lote", "Sem lote"))
        with result_col_6:
            status_color = "#B42318" if is_canceled_nf_status(separacao_result.get("Status NF", "")) else "#22C55E"
            render_highlight_card("Status NF", separacao_result.get("Status NF", "--"), status_color, f"Produtos: {separacao_result.get('Produtos', '--')}")
        render_box_close()

    cleanup_feedback = st.session_state.get("data_cleanup_feedback")
    current_xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    current_lotes_registry, _ = carregar_lotes_json(str(LOTES_JSON_PATH))

    render_box_open()
    st.markdown(
        """
    <div class="page-hero" style="margin-top: 1.25rem;">
        <h2>Gestão de Dados do Sistema</h2>
        <p>Limpeza e controle de XMLs, separações e lotes</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    base_col_1, base_col_2, base_col_3 = st.columns(3, gap="medium")
    with base_col_1:
        render_metric_card("XMLs", len(current_xml_records), "xml")
    with base_col_2:
        render_metric_card("Separações", len(current_records), "separacao")
    with base_col_3:
        render_metric_card("Lotes", len(current_lotes_registry), "lotes")

    size_col_1, size_col_2, size_col_3 = st.columns(3, gap="medium")
    with size_col_1:
        render_info_card("Base XMLs", format_file_size_mb(XMLS_PROCESSADOS_JSON_PATH), "xml", "Arquivo json de XMLs processados")
    with size_col_2:
        render_info_card("Base Separação", format_file_size_mb(SEPARACAO_JSON_PATH), "separacao", "Arquivo json do mapa de separação")
    with size_col_3:
        render_info_card("Base Lotes", format_file_size_mb(LOTES_JSON_PATH), "lotes", "Arquivo json dos lotes registrados")

    st.warning("Essa ação não pode ser desfeita.")

    if isinstance(cleanup_feedback, dict) and cleanup_feedback:
        if cleanup_feedback.get("total_removido", 0) > 0:
            st.success("Limpeza realizada com sucesso")
        else:
            st.info("Nenhum registro foi removido com os critérios informados.")

        st.markdown(
            "\n".join(
                [
                    f"Período: {cleanup_feedback.get('periodo', '--')}",
                    f"XMLs removidos: {cleanup_feedback.get('xmls_removidos', 0)}",
                    f"Registros de separação removidos: {cleanup_feedback.get('separacao_removidos', 0)}",
                    f"Lotes removidos: {cleanup_feedback.get('lotes_removidos', 0)}",
                ]
            )
        )

        protected_xmls = cleanup_feedback.get("xmls_protegidos", 0)
        protected_lotes = cleanup_feedback.get("lotes_protegidos", 0)
        if protected_xmls:
            st.warning(f"{protected_xmls} XML(s) permaneceram na base por ainda estarem em uso na separação.")
        if protected_lotes:
            st.warning(f"{protected_lotes} lote(s) abertos permaneceram na base por proteção operacional.")

    cleanup_col_1, cleanup_col_2, cleanup_col_3 = st.columns([1.2, 1.2, 1.6], gap="medium")
    with cleanup_col_1:
        cleanup_start_date = st.date_input("Data inicial", key="data_cleanup_start_date")
    with cleanup_col_2:
        cleanup_end_date = st.date_input("Data final", key="data_cleanup_end_date")
    with cleanup_col_3:
        cleanup_type = st.selectbox(
            "Tipo de limpeza",
            options=DATA_CLEANUP_OPTIONS,
            key="data_cleanup_type",
        )

    cleanup_submitted = st.button("Limpar Dados", use_container_width=True, type="primary", key="execute_data_cleanup")

    if cleanup_submitted:
        try:
            with st.spinner("Limpando dados..."):
                cleanup_result = executar_limpeza_dados_sistema(cleanup_start_date, cleanup_end_date, cleanup_type)

            st.session_state["data_cleanup_feedback"] = cleanup_result
            st.session_state["separacao_records"] = cleanup_result.get("separacao_records", [])
            st.session_state["separacao_result"] = None
            st.session_state["separacao_feedback"] = {}
            st.session_state["lote_atual"] = build_lote_payload("", "", "")
            invalidate_runtime_data()
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))

    render_box_close()


def render_lotes_management_screen(separacao_records: list[dict[str, object]]) -> None:
    classificacao_records, _ = carregar_classificacao_produtos_json(
        str(CLASSIFICACAO_PRODUTOS_JSON_PATH),
        get_path_cache_token(CLASSIFICACAO_PRODUTOS_JSON_PATH),
    )
    separacao_records = apply_current_sector_classification(separacao_records, classificacao_records)
    current_lote = st.session_state.get("lote_atual") if isinstance(st.session_state.get("lote_atual"), dict) else None
    sync_lotes_registry(separacao_records, current_lote)
    lote_records_lookup = group_lote_records(separacao_records)
    lotes_metadata, lotes_storage_error = carregar_lotes_json(str(LOTES_JSON_PATH))
    catalog = build_lote_catalog(separacao_records, lotes_metadata)
    catalog_df = build_lote_catalog_dataframe(catalog, lote_records_lookup)
    feedback_message = str(st.session_state.get("gestao_lotes_feedback", "") or "")

    st.markdown(
        """
    <div class="page-hero">
        <h2>Gestão de Lotes de Separação</h2>
        <p>Controle e rastreabilidade dos lotes de picking</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    if lotes_storage_error:
        st.warning(lotes_storage_error)

    if feedback_message:
        st.success(feedback_message)
        st.session_state["gestao_lotes_feedback"] = ""

    if not catalog:
        st.info("Nenhum lote encontrado")
        return

    pending_delete_lote = str(st.session_state.get("gestao_lotes_pending_delete", "") or "").strip()
    if pending_delete_lote:
        confirm_col_1, confirm_col_2, confirm_col_3 = st.columns([3.2, 1.0, 1.0], gap="medium")
        with confirm_col_1:
            st.warning(f"Tem certeza que deseja excluir este lote? {pending_delete_lote}")
        with confirm_col_2:
            if st.button("Confirmar exclusão", use_container_width=True, key="confirmar_exclusao_lote"):
                updated_records = excluir_lote(pending_delete_lote)
                st.session_state["separacao_records"] = updated_records
                if isinstance(st.session_state.get("lote_atual"), dict) and st.session_state["lote_atual"].get("lote_id") == pending_delete_lote:
                    st.session_state["lote_atual"] = None
                st.session_state["gestao_lotes_pending_delete"] = ""
                st.session_state["lotes_filter_lote"] = "Todos"
                st.session_state["lotes_filter_search"] = ""
                st.session_state["gestao_lotes_feedback"] = f"Lote {pending_delete_lote} excluído com sucesso."
                st.rerun()
        with confirm_col_3:
            if st.button("Cancelar", use_container_width=True, key="cancelar_exclusao_lote"):
                st.session_state["gestao_lotes_pending_delete"] = ""
                st.rerun()

    render_box_open()
    report_col_1, report_col_2 = st.columns([2.0, 1.0], gap="medium")
    with report_col_1:
        report_type = st.radio(
            "Tipo de impressão",
            ["Completo", "Por Setor", "Por Rota"],
            horizontal=True,
            key="gestao_lotes_report_type",
        )
    with report_col_2:
        pesquisa_geral = st.text_input(
            "Pesquisar",
            key="lotes_filter_search",
            placeholder="🔍 Buscar...",
            label_visibility="collapsed",
        )

    search_query = str(pesquisa_geral or "").strip()
    search_term = search_query.lower()
    normalized_search_lote = search_query.upper()
    looks_like_lote_search = bool(re.fullmatch(r"\d{8}-\d{3,}", normalized_search_lote))

    filtered_catalog_df = catalog_df.copy()

    searched_lote_summary = None
    if normalized_search_lote:
        searched_match_df = filtered_catalog_df[filtered_catalog_df["_lote_norm"] == normalized_search_lote]
        if not searched_match_df.empty:
            searched_lote_summary = searched_match_df.iloc[0].to_dict()
            filtered_catalog_df = searched_match_df
        elif search_term:
            filtered_catalog_df = filtered_catalog_df[filtered_catalog_df["_search_blob"].str.contains(search_term, na=False)]
    elif search_term:
        filtered_catalog_df = filtered_catalog_df[filtered_catalog_df["_search_blob"].str.contains(search_term, na=False)]

    if filtered_catalog_df.empty:
        st.info("Nenhum lote encontrado")
        return

    filtered_catalog = filtered_catalog_df.to_dict("records")
    selected_summary = searched_lote_summary or filtered_catalog[0]
    selected_lote_id = str(selected_summary.get("Lote", "") or "").strip()
    selected_records = lote_records_lookup.get(selected_lote_id, [])
    selected_records = apply_current_sector_classification(selected_records, classificacao_records)
    detail_df = build_lote_detail_dataframe(selected_records, selected_lote_id)

    abertura_dt = parse_xml_datetime(selected_summary.get("Abertura", ""))
    fechamento_dt = parse_xml_datetime(selected_summary.get("Fechamento", ""))
    selected_summary["Abertura Formatada"] = format_datetime_display(abertura_dt) if abertura_dt else "--"
    selected_summary["Fechamento Formatada"] = format_datetime_display(fechamento_dt) if fechamento_dt else "--"

    search_term = str(pesquisa_geral or "").strip().lower()
    if search_term and not detail_df.empty:
        detail_df = detail_df.assign(
            _search_blob=build_search_blob_series(detail_df, ["NF", "Cliente", "Rota", "Descrição", "Setor", "Código Produto"])
        )
        detail_df = detail_df[detail_df["_search_blob"].str.contains(search_term, na=False)].drop(columns=["_search_blob"])

    report_filter_label = "Todos"
    report_records = list(selected_records)

    if report_type == "Por Setor" and selected_records:
        setor_counts: Counter[str] = Counter(
            str(record.get("Setor", "Não Identificados") or "Não Identificados")
            for record in selected_records
        )
        setor_options = [
            setor
            for setor, _ in sorted(
                setor_counts.items(),
                key=lambda item: (-item[1], item[0].upper()),
            )
        ]
        setor_widget_key = f"gestao_lotes_setor_{selected_lote_id or 'default'}"
        setor_context_key = f"{setor_widget_key}_context"
        preferred_setor = setor_options[0]
        current_selected_setor = str(st.session_state.get(setor_widget_key, "") or "")
        if st.session_state.get(setor_context_key) != selected_lote_id:
            st.session_state[setor_widget_key] = preferred_setor
            st.session_state[setor_context_key] = selected_lote_id
        elif current_selected_setor not in setor_options:
            st.session_state[setor_widget_key] = preferred_setor
        selected_setor = st.selectbox(
            "Setor",
            setor_options,
            key=setor_widget_key,
        )
        report_filter_label = selected_setor
        report_records = [
            record
            for record in selected_records
            if str(record.get("Setor", "Não Identificados") or "Não Identificados") == selected_setor
        ]
    elif report_type == "Por Setor":
        report_records = []
    elif report_type == "Por Rota" and selected_records:
        rota_options = sorted(
            {
                str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL)
                for record in selected_records
            },
            key=lambda item: item.upper(),
        )
        selected_rota = st.selectbox(
            "Rota",
            rota_options,
            key=f"gestao_lotes_rota_{selected_lote_id or 'default'}",
        )
        report_filter_label = selected_rota
        report_records = [
            record
            for record in selected_records
            if str(record.get("Rota", UNDEFINED_ROUTE_LABEL) or UNDEFINED_ROUTE_LABEL) == selected_rota
        ]
    elif report_type == "Por Rota":
        report_records = []

    lote_pdf_bytes = b""
    can_print = bool(selected_lote_id and selected_summary.get("Status") == LOT_STATUS_CLOSED and report_records)
    if can_print:
        lote_pdf_bytes = generate_lote_pdf(selected_summary, report_records, report_type, report_filter_label)

    file_suffix = sanitize_filename_part(f"{report_type}_{report_filter_label}", "completo")

    action_col_1, action_col_2 = st.columns([3.0, 1.0], gap="medium")
    with action_col_2:
        st.download_button(
            "📄 Exportar PDF",
            data=lote_pdf_bytes,
            file_name=f"lote_{sanitize_filename_part(selected_lote_id, 'lote')}_{file_suffix}.pdf",
            mime="application/pdf",
            use_container_width=True,
            disabled=not can_print,
        )

    if looks_like_lote_search and searched_lote_summary is None:
        st.warning("Lote não encontrado na busca. O campo continua funcionando como filtro geral.")
    elif selected_summary.get("Status") != LOT_STATUS_CLOSED:
        st.info("Lote precisa estar fechado para exportar ou imprimir em PDF.")
    elif not can_print:
        st.warning("Nenhum dado encontrado para o filtro selecionado.")
    render_box_close()

    render_box_open()
    render_section_heading("ListView de Lotes", "lotes")
    header_col_1, header_col_2, header_col_3, header_col_4, header_col_5, header_col_6, header_col_7 = st.columns(
        [1.4, 1.0, 1.4, 1.4, 1.0, 1.0, 1.2],
        gap="small",
    )
    header_col_1.markdown("**Lote**")
    header_col_2.markdown("**Status**")
    header_col_3.markdown("**Data abertura**")
    header_col_4.markdown("**Data fechamento**")
    header_col_5.markdown("**Qtd. NFs**")
    header_col_6.markdown("**Qtd. itens**")
    header_col_7.markdown("**Ações**")

    for lote_record in filtered_catalog:
        lote_id = str(lote_record.get("Lote", "") or "").strip()
        lote_records = lote_records_lookup.get(lote_id, [])
        row_pdf_bytes = b""
        can_generate_row_pdf = bool(lote_id and lote_record.get("Status") == LOT_STATUS_CLOSED and lote_records)
        if can_generate_row_pdf:
            lote_summary = dict(lote_record)
            lote_summary["Abertura Formatada"] = format_lote_datetime_display(lote_record.get("Abertura", ""))
            lote_summary["Fechamento Formatada"] = format_lote_datetime_display(lote_record.get("Fechamento", ""))
            row_pdf_bytes = generate_lote_pdf(lote_summary, lote_records)

        row_col_1, row_col_2, row_col_3, row_col_4, row_col_5, row_col_6, row_col_7 = st.columns(
            [1.4, 1.0, 1.4, 1.4, 1.0, 1.0, 1.2],
            gap="small",
        )
        with row_col_1:
            st.markdown(f"<div style='{style_lote_cell(lote_record.get('Lote', ''))}'>{html.escape(lote_record.get('Lote', '--'))}</div>", unsafe_allow_html=True)
        with row_col_2:
            st.markdown(f"<span style='{style_lote_status_badge(lote_record.get('Status', LOT_STATUS_OPEN))}'>{html.escape(lote_record.get('Status', LOT_STATUS_OPEN))}</span>", unsafe_allow_html=True)
        with row_col_3:
            st.write(format_lote_datetime_display(lote_record.get("Abertura", "")))
        with row_col_4:
            st.write(format_lote_datetime_display(lote_record.get("Fechamento", "")))
        with row_col_5:
            st.write(str(lote_record.get("NFs", 0)))
        with row_col_6:
            st.write(str(lote_record.get("Itens", 0)))
        with row_col_7:
            action_col_1, action_col_2 = st.columns([1, 1], gap="small")
            with action_col_1:
                if st.button("🗑️", key=f"excluir_lote_{lote_id}", use_container_width=True, help="Excluir lote"):
                    st.session_state["gestao_lotes_pending_delete"] = lote_id
                    st.rerun()
            with action_col_2:
                st.download_button(
                    "📄",
                    data=row_pdf_bytes,
                    file_name=f"lote_{sanitize_filename_part(lote_id, 'lote')}.pdf",
                    mime="application/pdf",
                    key=f"pdf_lote_{lote_id}",
                    use_container_width=True,
                    disabled=not can_generate_row_pdf,
                    help="Gerar PDF do lote" if can_generate_row_pdf else "Lote precisa estar fechado para gerar PDF",
                )
    render_box_close()

    render_box_open()
    render_section_heading("ListView do Lote", "itens")
    if detail_df.empty:
        st.info("Nenhum item encontrado para os filtros selecionados")
        render_box_close()
        return

    st.dataframe(
        build_lote_detail_styler(detail_df),
        width="stretch",
        hide_index=True,
        row_height=52,
        column_config={
            "NF": st.column_config.TextColumn("NF", width="small"),
            "Código Produto": st.column_config.TextColumn("Código Produto", width="small"),
            "Descrição": st.column_config.TextColumn("Descrição", width="large"),
            "Quantidade": st.column_config.NumberColumn("Quantidade", format="%.3f", width="small"),
            "Tipo": st.column_config.TextColumn("Tipo", width="small"),
            "Cliente": st.column_config.TextColumn("Cliente", width="medium"),
            "Setor": st.column_config.TextColumn("Setor", width="medium"),
            "Rota": st.column_config.TextColumn("Rota", width="medium"),
        },
    )
    render_box_close()


def initialize_app_state() -> None:
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

    if "classificacao_upload_signature" not in st.session_state:
        st.session_state.classificacao_upload_signature = ""

    if "classificacao_upload_message" not in st.session_state:
        st.session_state.classificacao_upload_message = ""

    if "classificacao_upload_error" not in st.session_state:
        st.session_state.classificacao_upload_error = ""

    if "runtime_refresh_required" not in st.session_state:
        st.session_state["runtime_refresh_required"] = False

    if "runtime_data_signature" not in st.session_state:
        st.session_state["runtime_data_signature"] = None

    if "runtime_operational_signature" not in st.session_state:
        st.session_state["runtime_operational_signature"] = None

    if "runtime_xml_records" not in st.session_state:
        st.session_state["runtime_xml_records"] = []

    if "runtime_classificacao_records" not in st.session_state:
        st.session_state["runtime_classificacao_records"] = []

    if "separacao_sync_issues" not in st.session_state:
        st.session_state["separacao_sync_issues"] = []

    if "separacao_storage_error" not in st.session_state:
        st.session_state["separacao_storage_error"] = ""

    if "separacao_import_summary_runtime" not in st.session_state:
        st.session_state["separacao_import_summary_runtime"] = {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0}

    if "separacao_records" not in st.session_state:
        st.session_state["separacao_records"] = []

    if "separacao_result" not in st.session_state:
        st.session_state["separacao_result"] = None

    if "separacao_feedback" not in st.session_state:
        st.session_state["separacao_feedback"] = {}

    if "data_cleanup_feedback" not in st.session_state:
        st.session_state["data_cleanup_feedback"] = {}

    if "lote_atual" not in st.session_state:
        st.session_state["lote_atual"] = None

    if "lotes_filter_lote" not in st.session_state:
        st.session_state["lotes_filter_lote"] = "Todos"


def render_global_app_styles() -> None:
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
    .ui-section-box {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 14px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
        padding: 18px;
        margin: 0 0 18px;
    }
    .ui-section-box.is-soft {
        background: #FAFBFC;
    }
    .ui-section-box.is-sidebar {
        padding: 14px;
        margin-bottom: 16px;
        box-shadow: 0 3px 10px rgba(15, 23, 42, 0.035);
    }
    .ui-section-box .section-title-block {
        margin-bottom: 12px;
    }
    .ui-section-box .stCaption {
        margin-bottom: 0;
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
    .page-hero {
        background: linear-gradient(135deg, #ffffff 0%, #eef4fb 100%);
        border: 1px solid rgba(31, 58, 95, 0.08);
        border-radius: 16px;
        box-shadow: 0 10px 24px rgba(31, 58, 95, 0.05);
        padding: 18px 20px;
        margin: 0 0 18px;
    }
    .page-hero h2 {
        margin: 0;
        color: #16324F;
        font-size: 1.55rem;
        font-weight: 800;
    }
    .page-hero p {
        margin: 0.35rem 0 0;
        color: #617285;
        font-size: 0.96rem;
    }
    .operation-card {
        min-height: 106px;
    }
    .scan-shell {
        background: #FAFBFC;
        border: 1px solid #E5E7EB;
        border-radius: 14px;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.6);
        padding: 14px 14px 2px;
        margin-bottom: 0;
    }
    .lot-banner {
        background: linear-gradient(135deg, #16324F 0%, #1F4B7A 100%);
        border-radius: 14px;
        box-shadow: 0 10px 24px rgba(22, 50, 79, 0.16);
        padding: 16px 18px;
        margin: 0 0 16px;
        color: #FFFFFF;
    }
    .lot-banner-label {
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.1em;
        opacity: 0.78;
        text-transform: uppercase;
    }
    .lot-banner-value {
        margin-top: 0.3rem;
        font-size: 1.35rem;
        font-weight: 800;
        letter-spacing: -0.02em;
    }
    .lot-banner-meta {
        margin-top: 0.35rem;
        font-size: 0.92rem;
        opacity: 0.9;
    }
    .scan-button-spacer {
        height: 1.85rem;
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
    .app-shell-header {
        margin-bottom: 1rem;
    }
    .dashboard-hero {
        background: linear-gradient(135deg, #16324F 0%, #25486E 100%);
        border-radius: 18px;
        padding: 28px;
        color: #FFFFFF;
        box-shadow: 0 16px 30px rgba(22, 50, 79, 0.16);
        margin-bottom: 1.25rem;
    }
    .dashboard-hero h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: -0.03em;
    }
    .dashboard-hero p {
        margin: 0.5rem 0 0;
        font-size: 1rem;
        opacity: 0.9;
    }
    .module-card {
        background: #FFFFFF;
        border: 1px solid rgba(31, 58, 95, 0.08);
        border-radius: 18px;
        padding: 20px;
        box-shadow: 0 10px 24px rgba(31, 58, 95, 0.06);
        min-height: 170px;
        margin-bottom: 12px;
    }
    .module-card h3 {
        margin: 0.7rem 0 0.35rem;
        color: #16324F;
        font-size: 1.1rem;
        font-weight: 800;
    }
    .module-card p {
        margin: 0 0 1rem;
        color: #617285;
        font-size: 0.94rem;
        line-height: 1.5;
    }
    .module-card-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 44px;
        height: 44px;
        border-radius: 14px;
        background: #EEF4FF;
        color: #1D4ED8;
    }
    .module-card-icon .ui-icon {
        width: 22px;
        min-width: 22px;
        height: 22px;
        color: inherit;
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
    .stButton > button[kind="primary"],
    .stFormSubmitButton > button[kind="primary"] {
        background: #B42318;
        color: #FFFFFF;
        border: 0;
    }
    .stButton > button[kind="primary"]:hover,
    .stFormSubmitButton > button[kind="primary"]:hover {
        background: #912018;
        color: #FFFFFF;
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
    .sidebar-heading {
        margin: 0 0 10px;
        color: #1F3A5F;
        font-size: 1rem;
        font-weight: 700;
    }
    .sidebar-field-label {
        margin: 0 0 10px;
        color: #334155;
        font-size: 0.95rem;
        font-weight: 700;
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
        .dashboard-hero {
            padding: 22px;
        }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )


def load_runtime_reference_data(force_refresh: bool = False) -> tuple[list[dict[str, object]], list[dict[str, str]]]:
    runtime_signature = (
        get_path_cache_token(XMLS_PROCESSADOS_JSON_PATH),
        get_path_cache_token(CLASSIFICACAO_PRODUTOS_JSON_PATH),
    )
    should_refresh = (
        force_refresh
        or st.session_state.get("runtime_data_signature") != runtime_signature
        or not st.session_state.get("runtime_xml_records")
        and XMLS_PROCESSADOS_JSON_PATH.is_file()
    )

    if should_refresh:
        with st.spinner("Atualizando bases de referência..."):
            xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
            classificacao_records, _ = carregar_classificacao_produtos_json(
                str(CLASSIFICACAO_PRODUTOS_JSON_PATH),
                get_path_cache_token(CLASSIFICACAO_PRODUTOS_JSON_PATH),
            )
        st.session_state["runtime_xml_records"] = xml_records
        st.session_state["runtime_classificacao_records"] = classificacao_records
        st.session_state["runtime_data_signature"] = runtime_signature

    return (
        st.session_state.get("runtime_xml_records", []),
        st.session_state.get("runtime_classificacao_records", []),
    )


def load_runtime_operational_data(force_refresh: bool = False) -> tuple[list[dict[str, object]], list[str], str, dict[str, int]]:
    xml_records, classificacao_records = load_runtime_reference_data(force_refresh=force_refresh)
    runtime_signature = (
        st.session_state.get("runtime_data_signature"),
        get_path_cache_token(SEPARACAO_JSON_PATH),
    )
    should_refresh = (
        force_refresh
        or st.session_state.get("runtime_operational_signature") != runtime_signature
        or "separacao_records" not in st.session_state
    )

    if should_refresh:
        with st.spinner("Atualizando base operacional..."):
            separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary = sincronizar_base_separacao(
                xml_records,
                classificacao_records,
            )
        st.session_state["separacao_records"] = separacao_records
        st.session_state["separacao_sync_issues"] = separacao_sync_issues
        st.session_state["separacao_storage_error"] = separacao_storage_error
        st.session_state["separacao_import_summary_runtime"] = separacao_import_summary
        st.session_state["runtime_operational_signature"] = runtime_signature
        st.session_state["runtime_refresh_required"] = False

    return (
        st.session_state.get("separacao_records", []),
        st.session_state.get("separacao_sync_issues", []),
        st.session_state.get("separacao_storage_error", ""),
        st.session_state.get("separacao_import_summary_runtime", {"novas": 0, "atualizadas": 0, "ignoradas_separadas": 0}),
    )


def render_active_screen(current_screen: str, process_clicked: bool, excel_file) -> None:
    force_refresh = bool(st.session_state.get("runtime_refresh_required", False))

    if current_screen == SCREEN_MINUTA:
        xml_records, _ = load_runtime_reference_data(force_refresh=force_refresh)
        tela_minuta(process_clicked, xml_records, excel_file)
        return

    separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary = load_runtime_operational_data(
        force_refresh=force_refresh
    )
    if current_screen == SCREEN_SEPARACAO:
        tela_separacao(separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary)
        return

    tela_lotes(separacao_records)


def render_screen_header(title: str, subtitle: str) -> None:
    col_logo, col_header, col_home, col_menu_toggle, col_action = st.columns([1.1, 4.1, 1.0, 1.0, 1.0], vertical_alignment="center")
    with col_logo:
        logo_path = get_logo_path()
        if logo_path is not None:
            st.image(str(logo_path), width=120)
    with col_header:
        st.markdown(f"## {title}")
        st.caption(subtitle)
    with col_home:
        if st.button("🏠 Menu", use_container_width=True, key=f"home_button_{title}"):
            navegar(SCREEN_MENU)
    with col_menu_toggle:
        st.button("Painel", use_container_width=True, on_click=toggle_menu, key=f"toggle_sidebar_{title}")
    with col_action:
        st.button("Sair", use_container_width=True, on_click=logout, key=f"logout_{title}")


def tela_menu() -> None:
    apply_sidebar_visibility(False)
    top_col_1, top_col_2, top_col_3 = st.columns([1.2, 4.8, 1.0], vertical_alignment="center")
    with top_col_1:
        logo_path = get_logo_path()
        if logo_path is not None:
            st.image(str(logo_path), width=130)
    with top_col_2:
        st.markdown(
            """
        <div class="dashboard-hero">
            <h1>Central Operacional</h1>
            <p>Selecione um módulo para continuar o fluxo de carregamento, separação e gestão de lotes.</p>
        </div>
        """,
            unsafe_allow_html=True,
        )
    with top_col_3:
        st.button("Sair", use_container_width=True, on_click=logout, key="logout_menu")

    card_col_1, card_col_2, card_col_3 = st.columns(3, gap="medium")
    menu_cards = [
        (card_col_1, SCREEN_MINUTA, "Minuta de Carregamento", "Processamento da carga com XML, Excel e geração da minuta operacional.", "excel", "📦 Minuta de Carregamento"),
        (card_col_2, SCREEN_SEPARACAO, "Mapa de Separação", "Conferência de picking, leitura de notas e controle operacional da separação.", "separacao", "📊 Mapa de Separação"),
        (card_col_3, SCREEN_LOTES, "Gestão de Lotes de Separação", "Consulta, exportação e controle dos lotes fechados e em andamento.", "lotes", "📑 Gestão de Lotes de Separação"),
    ]
    for column, target_screen, title, description, icon_key, button_label in menu_cards:
        with column:
            st.markdown(
                f"""
            <div class="module-card">
                <div class="module-card-icon">{render_label_icon(ICON_MAP[icon_key])}</div>
                <h3>{html.escape(title)}</h3>
                <p>{html.escape(description)}</p>
            </div>
            """,
                unsafe_allow_html=True,
            )
            if st.button(button_label, use_container_width=True, key=f"menu_nav_{target_screen}"):
                navegar(target_screen)


def tela_minuta(process_clicked: bool, xml_records: list[dict[str, object]], excel_file) -> None:
    render_screen_header("Minuta de Carregamento", "Processamento de carga com XMLs e rotas")
    render_processing_screen(process_clicked, xml_records, excel_file)


def tela_separacao(
    separacao_records: list[dict[str, object]],
    separacao_sync_issues: list[str],
    separacao_storage_error: str,
    separacao_import_summary: dict[str, int],
) -> None:
    render_screen_header("Mapa de Separação", "Controle de picking por setor")
    render_separacao_screen(separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary)


def tela_lotes(separacao_records: list[dict[str, object]]) -> None:
    render_screen_header("Gestão de Lotes de Separação", "Controle e rastreabilidade dos lotes de picking")
    render_lotes_management_screen(separacao_records)


def render_main_screen() -> None:
    initialize_app_state()
    render_global_app_styles()

    login_success = st.session_state.get("login_success", "")
    if login_success:
        st.success(login_success)
        st.session_state["login_success"] = ""
    current_screen = normalize_screen_name(st.session_state.get("tela", SCREEN_MENU))
    if current_screen == SCREEN_MENU:
        tela_menu()
        return

    if "menu_aberto" not in st.session_state:
        st.session_state["menu_aberto"] = True

    apply_sidebar_visibility(st.session_state["menu_aberto"])
    excel_file = None
    process_clicked = False
    if st.session_state["menu_aberto"]:
        _, excel_file, process_clicked = render_sidebar()

    render_active_screen(current_screen, process_clicked, excel_file)


def main() -> None:
    st.set_page_config(layout="wide")
    initialize_login_state()
    initialize_navigation_state()

    if "login_error" not in st.session_state:
        st.session_state["login_error"] = ""

    if "login_success" not in st.session_state:
        st.session_state["login_success"] = ""

    if not st.session_state["logado"]:
        st.session_state["tela"] = SCREEN_LOGIN
        render_login_screen()
        return

    render_main_screen()


if __name__ == "__main__":
    main()