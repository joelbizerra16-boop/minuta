import sys
from contextlib import contextmanager
from pathlib import Path

_APP_ROOT = Path(__file__).resolve().parent
if str(_APP_ROOT) not in sys.path:
    sys.path.insert(0, str(_APP_ROOT))
from datetime import datetime
from collections import Counter
from typing import Callable
import base64
import html
from io import BytesIO
import json
import re
import hashlib
import textwrap
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
import logging
import time

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

from utils.gerador_minuta import generate_minuta_entrega_pdf
from utils.document_download_package import build_documentos_download_package
from utils.minuta_carregamento import (
    MINUTA_CARREGAMENTO_CONFIG,
    MINUTA_MODULES,
    MinutaModuleConfig,
)
from core.bootstrap import configure_application_storage
from core.runtime_data_coherence import (
    get_classificacao_version_token,
    get_operational_data_signature,
    get_reference_data_signature,
    get_separacao_storage_status_from_db,
)
from core.startup_environment import run_startup_environment_checks
from core.startup_retention import run_startup_retention_once
from core.performance import (
    build_performance_report,
    bump_processed_data_version,
    invalidate_balcao_lookup_cache,
    invalidate_latest_closed_lote_pdf_cache,
    measure,
)
from auth.pages.login import render_login_page
from auth.pages.usuarios import render_usuarios_page
from auth.security import session as operador_session

clear_session_on_logout = operador_session.clear_session_on_logout
get_current_user = operador_session.get_current_user
get_logged_operator_display_name = operador_session.get_logged_operator_display_name
is_admin = operador_session.is_admin
is_logged_in = operador_session.is_logged_in
render_logged_user_badge = operador_session.render_logged_user_badge
require_admin = operador_session.require_admin
OPERADOR_NAO_IDENTIFICADO = operador_session.OPERADOR_NAO_IDENTIFICADO
from carregamentos.integration import (
    DECISAO_OPERACIONAL_LABELS,
    cancelar_operacao_pendente,
    clear_balcao_pending,
    clear_contexto_operacional,
    clear_reentrega_pending,
    clear_reimpressao_pending,
    OPERACIONAL_DECISAO_WIDGET_KEY,
    OPERACIONAL_CONTINUAR_HISTORICO_VALUE,
    confirmar_decisao_operacional_continuacao,
    executar_analise_operacional,
    executar_fechamento_balcao_para_pdf,
    executar_fechamento_veiculo_para_pdf,
    get_operacional_decisao,
    get_operacional_diagnostico,
    get_diagnostico_efetivo_fechamento,
    is_operacional_analise_confirmada,
    iniciar_entrega_balcao,
    on_baixar_pdf_click,
    on_operacional_decisao_widget_change,
    on_processing_panel_primary_click,
    on_processing_panel_secondary_click,
    persistir_pdfs_apos_fechamento,
    render_balcao_nf_preview,
    requer_confirmacao_explicita_historico,
    resolve_operational_panel_mode,
    snapshot_exportacao_documentos,
    sync_processing_context_for_excel,
)
from carregamentos.models.operacional import CenarioOperacional, DecisaoOperacional, DiagnosticoCarregamento
from carregamentos.ui.auditoria_nf_panel import render_auditoria_nf_expander, render_historico_nfs_contexto
from carregamentos.ui.xml_import_summary_panel import (
    build_xml_import_report,
    merge_xml_import_reports,
    render_xml_import_summary_panel,
)
from carregamentos.pages.consulta import render_consulta_carregamentos_page
from carregamentos.pages.gestao_dados import render_gestao_dados_page
from carregamentos.bootstrap import get_gestao_capacidade_service, get_gestao_dados_service
from carregamentos.services.nf_validation import localizar_nf_no_lote
from infrastructure.storage.config_storage import (
    CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS,
    CONFIG_CHAVE_LOTES,
    CONFIG_CHAVE_SEPARACAO,
    CONFIG_CHAVE_SEPARACAO_EXCLUIDOS,
    SqlJsonConfigStorage,
)
from infrastructure.storage.xml_storage import SqlXmlRecordRepository
from infrastructure.services.documento_xml_service import DocumentoXmlService, XmlDocumentalItem

_CONFIG_STORAGE = SqlJsonConfigStorage()
_LOGGER = logging.getLogger("minuta.documento_xml")
_DOCUMENTO_XML_SERVICE: DocumentoXmlService | None = None


def _get_documento_xml_service() -> DocumentoXmlService:
    global _DOCUMENTO_XML_SERVICE
    if _DOCUMENTO_XML_SERVICE is None:
        _DOCUMENTO_XML_SERVICE = DocumentoXmlService()
    return _DOCUMENTO_XML_SERVICE

BASE_DIR = Path(__file__).resolve().parent
FIXED_LOGO_PATH = BASE_DIR / "baixados.png"
WINDOWS_FONT_DIR = Path("C:/Windows/Fonts")
DATA_DIR = BASE_DIR / "data"
XMLS_PROCESSADOS_JSON_PATH = DATA_DIR / "xmls_processados.json"
CLASSIFICACAO_PRODUTOS_JSON_PATH = DATA_DIR / "classificacao_produtos.json"
SEPARACAO_JSON_PATH = DATA_DIR / "separacao.json"
LOTES_JSON_PATH = DATA_DIR / "lotes.json"
SEPARACAO_EXCLUIDOS_JSON_PATH = DATA_DIR / "separacao_excluidos.json"
NFE_NAMESPACE = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
UF_CODE_MAP = {
    "11": "RO",
    "12": "AC",
    "13": "AM",
    "14": "RR",
    "15": "PA",
    "16": "AP",
    "17": "TO",
    "21": "MA",
    "22": "PI",
    "23": "CE",
    "24": "RN",
    "25": "PB",
    "26": "PE",
    "27": "AL",
    "28": "SE",
    "29": "BA",
    "31": "MG",
    "32": "ES",
    "33": "RJ",
    "35": "SP",
    "41": "PR",
    "42": "SC",
    "43": "RS",
    "50": "MS",
    "51": "MT",
    "52": "GO",
    "53": "DF",
}
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
MAX_XML_UPLOAD_BATCH = 2000
ROUTE_FROM_INF_CPL_PATTERN = re.compile(r"(?im)(?:^|[\s\-])rota\s*:\s*([^\n\r]+)")
ROUTE_NEXT_FIELD_PATTERN = re.compile(
    r"\s+-\s+(?:Pedido|Cliente|Vendedor|Trib|Valor|ICMS|CNPJ)\s*:.*$",
    re.IGNORECASE,
)
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
SCREEN_LOGIN = "login"
SCREEN_MENU = "menu"
SCREEN_MINUTA = "minuta"
SCREEN_SEPARACAO = "separacao"
SCREEN_LOTES = "lotes"
SCREEN_USUARIOS = "usuarios"
SCREEN_CONSULTA_CARREGAMENTOS = "consulta_carregamentos"
SCREEN_GESTAO_DADOS = "gestao_dados"
SCREEN_GESTAO_RETENCAO = "gestao_retencao"
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
    "usuarios": "user_badge",
    "cadastro_usuarios": "sheet",
    "consulta_carregamentos": "truck",
    "gestao_dados": "folder",
    "gestao_retencao": "folder",
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


_PDF_FONTS_READY = False


def register_pdf_fonts() -> tuple[str, str]:
    global _PDF_FONTS_READY
    if _PDF_FONTS_READY:
        if "Arial" in pdfmetrics.getRegisteredFontNames():
            return "Arial", "Arial-Bold"
        return PDF_FONT_REGULAR, PDF_FONT_BOLD

    regular_font = WINDOWS_FONT_DIR / "arial.ttf"
    bold_font = WINDOWS_FONT_DIR / "arialbd.ttf"

    if regular_font.is_file() and bold_font.is_file():
        if "Arial" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("Arial", str(regular_font)))
        if "Arial-Bold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("Arial-Bold", str(bold_font)))
        _PDF_FONTS_READY = True
        return "Arial", "Arial-Bold"

    _PDF_FONTS_READY = True
    return PDF_FONT_REGULAR, PDF_FONT_BOLD


def get_logo_path() -> Path | None:
    return FIXED_LOGO_PATH if FIXED_LOGO_PATH.is_file() else None


_LOGO_DATA_URI_CACHE: str | None = None


def get_logo_data_uri() -> str:
    global _LOGO_DATA_URI_CACHE
    if _LOGO_DATA_URI_CACHE is not None:
        return _LOGO_DATA_URI_CACHE

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
    _LOGO_DATA_URI_CACHE = f"data:{mime_type};base64,{encoded_logo}"
    return _LOGO_DATA_URI_CACHE


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


def normalize_route_label(value: object) -> str:
    route = re.sub(r"\s+", " ", str(value or "").strip())
    return route or UNDEFINED_ROUTE_LABEL


def extract_route_from_inf_cpl(inf_cpl: object) -> str:
    text = str(inf_cpl or "").replace("\\n", "\n")
    if not text.strip():
        return ""

    match = ROUTE_FROM_INF_CPL_PATTERN.search(text)
    if not match:
        return ""

    route_value = match.group(1).strip()
    route_value = ROUTE_NEXT_FIELD_PATTERN.sub("", route_value).strip()
    return route_value


def ensure_route_column(dataframe: pd.DataFrame) -> pd.DataFrame:
    updated_df = dataframe.copy()
    if "ROTA" not in updated_df.columns:
        updated_df["ROTA"] = UNDEFINED_ROUTE_LABEL
        return updated_df

    updated_df["ROTA"] = updated_df["ROTA"].map(normalize_route_label)
    return updated_df


def apply_routes_from_xml_index(dataframe: pd.DataFrame, xml_index: dict[str, dict[str, object]]) -> pd.DataFrame:
    updated_df = dataframe.copy()
    if updated_df.empty:
        return ensure_route_column(updated_df)

    route_lookup: dict[str, str] = {}
    for key, xml_data in (xml_index or {}).items():
        nf_key = normalize_nf(xml_data.get("nf_normalizada", "") or xml_data.get("NF", "") or key)
        if nf_key:
            route_lookup[nf_key] = normalize_route_label(xml_data.get("ROTA", ""))

    if "NF" in updated_df.columns:
        updated_df["ROTA"] = updated_df["NF"].map(normalize_nf).map(route_lookup)
    else:
        updated_df["ROTA"] = UNDEFINED_ROUTE_LABEL

    return ensure_route_column(updated_df)


def get_sector_colors(setor: str) -> dict[str, str]:
    return SECTOR_COLOR_MAP.get(setor, SECTOR_COLOR_MAP["Não Identificados"])


def render_label_icon(icon_name: str) -> str:
    raw_svg = ICON_SVG.get(icon_name, ICON_SVG["folder"])
    return f'<span class="ui-icon" aria-hidden="true">{_normalize_icon_svg(raw_svg)}</span>'


def _normalize_icon_svg(svg_markup: str) -> str:
    """Atributos intrinsecos de tamanho e traco — icones corretos mesmo sem o bloco CSS no DOM."""
    if 'data-brida-icon="true"' in svg_markup:
        return svg_markup
    normalized = svg_markup.replace(
        '<svg viewBox="0 0 20 20"',
        (
            '<svg data-brida-icon="true" viewBox="0 0 20 20" '
            'width="16" height="16" '
            'fill="none" stroke="currentColor" stroke-width="1.7" '
            'stroke-linecap="round" stroke-linejoin="round"'
        ),
        1,
    )
    return normalized.replace("<path d=", '<path fill="none" stroke="currentColor" d=')


def resolve_menu_icon(icon_key: str) -> str:
    return ICON_MAP.get(str(icon_key or "").strip(), ICON_MAP["dados_gerais"])


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


def normalize_uf_value(value: object) -> str:
    text = str(value or "").strip().upper()
    text = re.sub(r"[^A-Z]", "", text)
    return text[:2] if len(text) >= 2 else text


def infer_uf_from_chave(chave_nfe: object) -> str:
    chave = normalize_chave_nfe(chave_nfe)
    if len(chave) != 44:
        return ""
    return UF_CODE_MAP.get(chave[:2], "")


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
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace("R$", "").replace(" ", "")
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") == 1 and text.count(".") == 0:
        text = text.replace(",", ".")
    elif text.count(".") > 1 and text.count(",") == 0:
        text = text.replace(".", "")

    text = re.sub(r"[^0-9.-]", "", text)
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


def first_non_empty(*values: object) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def extract_optional_excel_columns(base_df: pd.DataFrame) -> pd.DataFrame:
    optional_df = pd.DataFrame(index=base_df.index)
    optional_field_aliases = {
        "ClienteExcel": [
            "Cliente",
            "Destinatario",
            "Destinatário",
            "Razao Social",
            "Razão Social",
            "Nome Cliente",
        ],
        "CidadeExcel": ["Cidade", "Municipio", "Município", "Praca", "Praça"],
        "UFExcel": ["UF", "Estado"],
        "ValorExcel": ["Valor", "Valor NF", "Valor Nota", "Valor Total", "Vlr NF", "Total Nota", "Total"],
        "PesoExcel": ["Peso", "Peso Kg", "Peso NF", "Peso Total", "Peso Bruto"],
        "VolumeExcel": ["Volume", "Volumes", "Qtd Vol", "Quantidade Volume"],
        "TransportadoraExcel": ["Transportadora", "Transportador", "Transportes", "Nome Transportadora"],
        "MotoristaExcel": ["Motorista", "Nome Motorista", "Condutor"],
        "VeiculoExcel": ["Veiculo", "Veículo", "Placa", "Caminhao", "Caminhão"],
    }

    numeric_fields = {"ValorExcel", "PesoExcel", "VolumeExcel"}
    for field_name, aliases in optional_field_aliases.items():
        column_name = find_column(list(base_df.columns), aliases)
        if column_name is None:
            optional_df[field_name] = 0.0 if field_name in numeric_fields else ""
            continue

        if field_name in numeric_fields:
            optional_df[field_name] = base_df[column_name].apply(parse_float)
        else:
            optional_df[field_name] = base_df[column_name].fillna("").astype(str).str.strip()

    return optional_df


def build_default_product_classification_records() -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for item in DEFAULT_PRODUCT_CLASSIFICATION_RULES:
        keyword = normalize_matching_text(item.get("palavra_chave", ""))
        sector = normalize_sector_name(item.get("setor", ""))
        if keyword and sector:
            records.append({"palavra_chave": keyword, "setor": sector})
    return sorted(records, key=lambda record: (-len(record["palavra_chave"]), record["palavra_chave"]))


@st.cache_data(show_spinner=False)
def carregar_classificacao_produtos_records(
    json_path: str,
    version_token: tuple[int, str | None] | int = 0,
) -> tuple[list[dict[str, str]], str]:
    _ = (json_path, version_token)
    default_records = build_default_product_classification_records()
    try:
        payload = _CONFIG_STORAGE.load_list(CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS, default=[])
    except Exception as exc:
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


def carregar_classificacao_produtos_json(
    json_path: str,
    version_token: tuple[int, str | None] | int = 0,
) -> tuple[list[dict[str, str]], str]:
    return carregar_classificacao_produtos_records(json_path, version_token)


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
    transportadora_values = [value for value in base_df["Transportadora"].astype(str).tolist() if value] if "Transportadora" in base_df.columns else []
    carga_values = [value for value in base_df["Numero Carga"].astype(str).tolist() if value]

    unique_dates = sorted(set(data_saida_values))
    unique_motoristas = sorted(set(motorista_values))
    unique_placas = sorted(set(placa_values))
    unique_transportadoras = sorted(set(transportadora_values))
    unique_cargas = sorted(set(carga_values))

    return {
        "numero_carga": unique_cargas[0] if len(unique_cargas) == 1 else ("Multiplos" if unique_cargas else "--"),
        "data_saida": unique_dates[0] if len(unique_dates) == 1 else ("Multiplas" if unique_dates else "--"),
        "transportadora": unique_transportadoras[0] if len(unique_transportadoras) == 1 else ("Multiplas" if unique_transportadoras else "BRIDA LUBRIFICANTES LTDA"),
        "motorista": unique_motoristas[0] if len(unique_motoristas) == 1 else ("Multiplos" if unique_motoristas else "--"),
        "placa": unique_placas[0] if len(unique_placas) == 1 else ("Multiplas" if unique_placas else "--"),
    }


def detect_excel_structure(uploaded_excel) -> tuple[str, int | None, int | None, pd.DataFrame]:
    workbook = pd.ExcelFile(uploaded_excel)
    uploaded_excel.seek(0)

    overview_tokens = {"filial", "dtsaida", "data", "carga", "carregamento", "numerocarga", "motorista", "veiculo", "placa", "transportadora", "transportador"}
    detail_tokens = {"seqent", "numeronota", "nota", "nf", "carregamento", "numeropedido", "pesokg", "cliente", "cidade", "uf", "valor", "volume"}
    best_sheet = workbook.sheet_names[0]
    best_overview_row = None
    best_detail_row = None
    best_overview_score = -1
    best_detail_score = -1
    best_preview_df = pd.DataFrame()

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
                best_preview_df = preview_df

            if overview_score > best_overview_score:
                best_overview_score = overview_score
                best_overview_row = row_index

    uploaded_excel.seek(0)
    return best_sheet, best_detail_row, best_overview_row, best_preview_df


def extract_summary_metadata(preview_df: pd.DataFrame, overview_row: int | None) -> dict[str, str]:
    default_metadata = {"Filial": "BRIDA", "Numero Carga": "", "Data Saida": "", "Transportadora": "BRIDA LUBRIFICANTES LTDA", "Motorista": "", "Placa": ""}
    if overview_row is None or overview_row + 1 >= len(preview_df.index):
        return default_metadata

    header_values = [normalize_label(value) for value in preview_df.iloc[overview_row].tolist()]
    data_values = preview_df.iloc[overview_row + 1].tolist()
    mapping = dict(zip(header_values, data_values))

    filial = str(mapping.get("filial", "BRIDA") or "BRIDA").strip()
    numero_carga = str(mapping.get("carregamento", mapping.get("carga", mapping.get("numerocarga", ""))) or "").strip()
    data_saida = format_date_series(pd.Series([mapping.get("dtsaida", mapping.get("data", ""))])).iloc[0]
    transportadora = str(mapping.get("transportadora", mapping.get("transportador", mapping.get("transportes", "BRIDA LUBRIFICANTES LTDA"))) or "BRIDA LUBRIFICANTES LTDA").strip()
    motorista = str(mapping.get("motorista", "") or "").strip()
    placa = str(mapping.get("veiculo", mapping.get("placa", "")) or "").strip()

    return {
        "Filial": filial,
        "Numero Carga": numero_carga,
        "Data Saida": data_saida,
        "Transportadora": transportadora,
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
    metadata_df["Transportadora"] = summary_metadata.get("Transportadora", "BRIDA LUBRIFICANTES LTDA")
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
        sheet_name, detail_header_row, overview_row, preview_df = detect_excel_structure(uploaded_excel)

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
    metadata_df = metadata_df.join(extract_optional_excel_columns(base_df))
    if "TransportadoraExcel" in metadata_df.columns:
        transportadora_candidates = [value for value in metadata_df["TransportadoraExcel"].astype(str).tolist() if str(value).strip()]
        if transportadora_candidates and summary_metadata.get("Transportadora", "BRIDA LUBRIFICANTES LTDA") == "BRIDA LUBRIFICANTES LTDA":
            metadata_df["Transportadora"] = transportadora_candidates[0]
    if "MotoristaExcel" in metadata_df.columns:
        motorista_candidates = [value for value in metadata_df["MotoristaExcel"].astype(str).tolist() if str(value).strip()]
        if motorista_candidates and not summary_metadata.get("Motorista", ""):
            metadata_df["Motorista"] = motorista_candidates[0]
    if "VeiculoExcel" in metadata_df.columns:
        veiculo_candidates = [value for value in metadata_df["VeiculoExcel"].astype(str).tolist() if str(value).strip()]
        if veiculo_candidates and not summary_metadata.get("Placa", ""):
            metadata_df["Placa"] = veiculo_candidates[0]
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
    uf = xml_text_any_namespace(root, ".//{*}dest/{*}enderDest/{*}UF") or find_xml_text_by_localname(root, ["UF"])
    reference_raw, reference_datetime = extract_xml_reference_datetime(root, xml_type)
    status = extract_xml_status(root, xml_type)
    data_emissao = extract_issue_date_from_xml(root)
    valor_total = parse_float(
        xml_text_any_namespace(root, ".//{*}ICMSTot/{*}vNF")
        or find_xml_text_by_localname(root, ["vNF"])
        or "0"
    )
    inf_cpl = find_xml_text_by_localname(root, ["infCpl"]) or xml_text_any_namespace(root, ".//{*}infCpl")
    rota = extract_route_from_inf_cpl(inf_cpl)

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
        "UF": str(uf or "").strip().upper(),
        "Status": status,
        "StatusNF": status,
        "ValorNF": valor_total,
        "ROTA": rota,
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


def validate_delivery_record(record: dict[str, object]) -> list[str]:
    missing_fields: list[str] = []
    required_mapping = {
        "nota": "Nota",
        "data": "Data",
        "cliente": "Cliente",
        "cidade": "Cidade",
        "uf": "UF",
    }

    for field_name, label in required_mapping.items():
        if not str(record.get(field_name, "") or "").strip():
            missing_fields.append(label)

    if parse_float(record.get("valor", 0.0)) <= 0:
        missing_fields.append("Valor")
    if parse_float(record.get("peso", 0.0)) <= 0:
        missing_fields.append("Peso")
    return missing_fields


def build_minuta_entrega_records(dataframe: pd.DataFrame) -> tuple[list[dict[str, object]], list[str], dict[str, float | int]]:
    if dataframe.empty:
        return [], [], {"total_volumes": 0.0, "total_peso": 0.0, "total_valor": 0.0, "total_nfs": 0}

    valid_rows = dataframe[dataframe["Status"].map(is_authorized_nf_status)].copy()
    if valid_rows.empty:
        return [], ["Nenhuma NF autorizada esta disponivel para gerar o romaneio de entrega."], {"total_volumes": 0.0, "total_peso": 0.0, "total_valor": 0.0, "total_nfs": 0}

    issues: list[str] = []
    entrega_records: list[dict[str, object]] = []
    grouped_df = valid_rows.groupby("NF", sort=False, dropna=False)

    for nf, group in grouped_df:
        first_row = group.iloc[0]
        cliente_excel = first_non_empty(first_row.get("ClienteExcel"), first_row.get("Destinatario"))
        cidade_excel = first_non_empty(first_row.get("CidadeExcel"), first_row.get("Municipio"))
        uf_excel = first_non_empty(first_row.get("UFExcel"), first_row.get("UF"), infer_uf_from_chave(first_row.get("ChaveNFe", "")))
        valor_excel = parse_float(first_row.get("ValorExcel", 0.0))
        valor_xml = parse_float(first_row.get("ValorNF", 0.0))
        peso_excel = parse_float(first_row.get("PesoExcel", 0.0))
        peso_xml = parse_float(first_row.get("PesoTotalNF", group["Peso"].sum()))
        volume_excel = parse_float(first_row.get("VolumeExcel", 0.0))
        volume_xml = parse_float(first_row.get("Volume", 0.0))

        if first_non_empty(first_row.get("ClienteExcel")) and first_non_empty(first_row.get("Destinatario")):
            if normalize_matching_text(first_row.get("ClienteExcel")) != normalize_matching_text(first_row.get("Destinatario")):
                issues.append(f"NF {nf}: cliente diverge entre Excel e XML. Foi mantido o valor do Excel.")

        if first_non_empty(first_row.get("CidadeExcel")) and first_non_empty(first_row.get("Municipio")):
            if normalize_matching_text(first_row.get("CidadeExcel")) != normalize_matching_text(first_row.get("Municipio")):
                issues.append(f"NF {nf}: cidade diverge entre Excel e XML. Foi mantido o valor do Excel.")

        if first_non_empty(first_row.get("UFExcel")) and first_non_empty(first_row.get("UF")):
            if normalize_matching_text(first_row.get("UFExcel")) != normalize_matching_text(first_row.get("UF")):
                issues.append(f"NF {nf}: UF diverge entre Excel e XML. Foi mantido o valor do Excel.")

        if valor_excel > 0 and valor_xml > 0 and abs(valor_excel - valor_xml) > 0.01:
            issues.append(f"NF {nf}: valor diverge entre Excel e XML. Foi mantido o valor do Excel.")

        if peso_excel > 0 and peso_xml > 0 and abs(peso_excel - peso_xml) > 0.01:
            issues.append(f"NF {nf}: peso diverge entre Excel e XML. Foi mantido o valor do Excel.")

        item_volume = volume_xml if volume_xml > 0 else volume_excel

        record = {
            "item": item_volume,
            "nota": str(nf or "").strip(),
            "data": first_non_empty(first_row.get("Data")),
            "cliente": cliente_excel,
            "cidade": cidade_excel,
            "uf": normalize_uf_value(uf_excel),
            "valor": valor_excel if valor_excel > 0 else valor_xml,
            "peso": peso_excel if peso_excel > 0 else peso_xml,
            "volume": item_volume,
            "rota": first_non_empty(first_row.get("ROTA"), cidade_excel),
            "status": first_non_empty(first_row.get("Status")),
        }

        missing_fields = validate_delivery_record(record)
        if missing_fields:
            issues.append(f"NF {nf}: campos obrigatorios ausentes ou invalidos ({', '.join(missing_fields)}).")
            continue

        entrega_records.append(record)

    entrega_records = sorted(
        entrega_records,
        key=lambda record: (
            normalize_matching_text(record.get("rota", "") or record.get("cidade", "")),
            normalize_matching_text(record.get("cidade", "")),
            normalize_matching_text(record.get("cliente", "")),
            normalize_nf(record.get("nota", "")),
        ),
    )

    totals = {
        "total_volumes": float(sum(parse_float(record.get("volume", 0.0)) for record in entrega_records)),
        "total_peso": float(sum(parse_float(record.get("peso", 0.0)) for record in entrega_records)),
        "total_valor": float(sum(parse_float(record.get("valor", 0.0)) for record in entrega_records)),
        "total_nfs": int(len(entrega_records)),
    }
    return entrega_records, issues, totals


def process_minuta_inputs(process_clicked: bool, xml_records: list, excel_file) -> None:
    if not process_clicked:
        return

    if excel_file is None:
        st.error("Envie um arquivo Excel para iniciar o processamento.")
        return

    clear_contexto_operacional()

    try:
        with measure("process.load_excel"):
            excel_base = load_excel_base(excel_file)
        st.session_state["operacional_excel_nome"] = str(getattr(excel_file, "name", "") or "")
        with measure("process.integrate_excel_xml"):
            processed_df, summary, issues, nf_debug = integrate_excel_with_xml(excel_base, xml_records or [])
        st.session_state.processed_df = processed_df
        st.session_state.summary = summary
        st.session_state.issues = issues
        st.session_state.nf_debug = pd.DataFrame(nf_debug, columns=NF_DEBUG_COLUMNS)
        st.session_state.document_issue_at = format_datetime_display()
        bump_processed_data_version()

        if processed_df.empty:
            st.warning("Nenhum dado foi processado. Verifique se o Excel possui NFs validas.")
        else:
            with measure("process.analise_operacional"):
                executar_analise_operacional(processed_df)
            st.success("Processamento concluido.")
    except ValueError as exc:
        st.session_state.processed_df = create_empty_processed_df()
        st.session_state.summary = create_empty_summary()
        st.session_state.issues = []
        st.session_state.nf_debug = create_empty_nf_debug_df()
        bump_processed_data_version()
        st.error(str(exc))
    except Exception as exc:
        st.session_state.processed_df = create_empty_processed_df()
        st.session_state.summary = create_empty_summary()
        st.session_state.issues = []
        st.session_state.nf_debug = create_empty_nf_debug_df()
        bump_processed_data_version()
        st.error(f"Erro inesperado ao processar os arquivos: {exc}")


def generate_minuta_pdf(
    dados_minuta: list[dict[str, object]],
    numero_carga: str,
    data_emissao: str,
    veiculo: str,
    motorista: str,
    pdf_title: str = "MINUTA DE CARREGAMENTO",
    subject_label: str = "Carregamento",
    operador: str = "",
    impresso_em: str = "",
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
        title = pdf_title
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
        pdf.drawString(left_margin, y_pos, f"{subject_label}:   {numero_carga or '--'}")
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

    operador_label = str(operador or "").strip() or OPERADOR_NAO_IDENTIFICADO
    impresso_label = str(impresso_em or "").strip() or format_datetime_display()
    pdf.setFont(regular_font, 8)
    pdf.setFillColor(colors.HexColor("#6a6a6a"))
    pdf.drawString(left_margin, bottom_margin + 2, f"Operador: {operador_label}")
    pdf.drawString(left_margin, bottom_margin - 10, f"Impresso em: {impresso_label}")

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
    chave_nfe = normalize_chave_nfe(xml_data.get("ChaveNFe", ""))
    uf_value = normalize_uf_value(xml_data.get("UF", "")) or infer_uf_from_chave(chave_nfe)
    return {
        "NF": normalize_nf(xml_data.get("NF", "") or xml_data.get("nf_normalizada", "")),
        "nf_normalizada": normalize_nf(xml_data.get("nf_normalizada", "") or xml_data.get("NF", "")),
        "ChaveNFe": chave_nfe,
        "Data": str(xml_data.get("Data", "") or "").strip(),
        "DataReferencia": str(xml_data.get("DataReferencia", "") or "").strip(),
        "DataReferenciaISO": str(xml_data.get("DataReferenciaISO", "") or "").strip(),
        "Destinatario": str(xml_data.get("Destinatario", "") or "").strip(),
        "Municipio": municipio,
        "UF": uf_value,
        "Status": normalize_nf_status(xml_data.get("StatusNF", xml_data.get("Status", ""))),
        "StatusNF": normalize_nf_status(xml_data.get("StatusNF", xml_data.get("Status", ""))),
        "ValorNF": parse_float(xml_data.get("ValorNF", 0.0)),
        "VolumeTotal": parse_float(xml_data.get("VolumeTotal", 0.0)),
        "PesoTotal": parse_float(xml_data.get("PesoTotal", 0.0)),
        "Items": items,
        "Arquivo": str(xml_data.get("Arquivo", "") or "").strip(),
        "Erro": bool(xml_data.get("Erro", False)),
        "TipoXML": str(xml_data.get("TipoXML", "normal") or "normal").strip(),
        "ROTA": normalize_route_label(xml_data.get("ROTA", "")),
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


def build_xml_upload_signature(xml_files: list) -> str:
    upload_signature_parts: list[str] = []
    for uploaded_file in xml_files or []:
        file_bytes = uploaded_file.getvalue()
        upload_signature_parts.append(f"{uploaded_file.name}:{hashlib.sha256(file_bytes).hexdigest()}")
    return hashlib.sha256("|".join(upload_signature_parts).encode("utf-8")).hexdigest()


class StoredXmlUpload:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data

    def seek(self, _offset: int) -> None:
        return None


def build_xml_file_key(name: str, file_bytes: bytes) -> str:
    return f"{name}:{compute_xml_file_hash(file_bytes)}"


def compute_xml_file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def ensure_xml_upload_batch() -> dict[str, dict[str, object]]:
    if "xml_upload_batch" not in st.session_state:
        st.session_state.xml_upload_batch = {}
    return st.session_state.xml_upload_batch


def merge_uploaded_files_into_xml_batch(uploaded_files: list | None) -> tuple[int, int, list[str]]:
    if not uploaded_files:
        return 0, 0, []

    batch = ensure_xml_upload_batch()
    added = 0
    duplicates = 0
    messages: list[str] = []

    for uploaded_file in uploaded_files:
        if len(batch) >= MAX_XML_UPLOAD_BATCH:
            raise ValueError(f"Limite maximo de {MAX_XML_UPLOAD_BATCH} XMLs por lote.")

        file_bytes = uploaded_file.getvalue()
        file_hash = compute_xml_file_hash(file_bytes)
        file_key = f"{uploaded_file.name}:{file_hash}"
        if file_key in batch:
            duplicates += 1
            messages.append(f"XML duplicado ignorado: {uploaded_file.name}")
            continue

        batch[file_key] = {
            "name": uploaded_file.name,
            "data": file_bytes,
            "hash_sha256": file_hash,
            "imported": False,
        }
        added += 1

    return added, duplicates, messages


def remove_xml_from_upload_batch(file_key: str) -> None:
    ensure_xml_upload_batch().pop(file_key, None)


def get_pending_xml_batch_uploads() -> list[StoredXmlUpload]:
    batch = ensure_xml_upload_batch()
    pending: list[StoredXmlUpload] = []
    for item in batch.values():
        if not item.get("imported"):
            pending.append(StoredXmlUpload(str(item.get("name", "arquivo.xml")), bytes(item.get("data", b""))))
    return pending


def extract_rejected_xml_filenames_from_issues(
    issues: list[str],
    parsed_records: list[dict[str, object]],
) -> set[str]:
    rejected: set[str] = set()
    nf_to_arquivo = {
        str(record.get("NF", "")): str(record.get("Arquivo", ""))
        for record in parsed_records
        if record.get("NF") and record.get("Arquivo")
    }
    rejection_patterns = (
        r"^Erro no XML (.+?):",
        r"^XML sem chave/NF identificavel: (.+)$",
        r"^XML duplicado no lote ignorado: (.+)$",
        r"^XML duplicado ou desatualizado ignorado: (.+)$",
    )

    for issue in issues:
        matched = False
        for pattern in rejection_patterns:
            match = re.match(pattern, issue)
            if match:
                rejected.add(match.group(1).strip())
                matched = True
                break
        if matched:
            continue

        separada_match = re.match(
            r"^NF (.+?) ignorada no upload porque ja esta separada\.$",
            issue,
        )
        if separada_match:
            arquivo = nf_to_arquivo.get(separada_match.group(1).strip())
            if arquivo:
                rejected.add(arquivo)

    return rejected


def finalize_xml_upload_batch_after_import(
    imported_file_keys: list[str],
    parsed_records: list[dict[str, object]],
    issues: list[str],
) -> None:
    batch = ensure_xml_upload_batch()
    accepted_names = {
        str(record.get("Arquivo", ""))
        for record in parsed_records
        if record.get("Arquivo")
    }
    rejected_names = extract_rejected_xml_filenames_from_issues(issues, parsed_records)
    valid_names = {name for name in accepted_names if name not in rejected_names}

    for file_key in imported_file_keys:
        item = batch.get(file_key)
        if not item:
            continue

        filename = str(item.get("name", ""))
        if filename in valid_names:
            item["imported"] = True
            item["accepted"] = True
            continue

        batch.pop(file_key, None)


def get_accepted_xml_upload_batch_items() -> list[tuple[str, dict[str, object]]]:
    batch = ensure_xml_upload_batch()
    return [(file_key, item) for file_key, item in batch.items() if item.get("accepted")]


def has_accepted_xml_upload_batch() -> bool:
    return bool(get_accepted_xml_upload_batch_items())


def merge_import_summary(existing: dict[str, int], incoming: dict[str, int]) -> dict[str, int]:
    summary_keys = [
        "total_arquivos",
        "processados",
        "erros",
        "duplicados",
        "ignorados",
        "novas",
        "atualizadas",
        "ignoradas_separadas",
        "duplicados_lote",
        "duplicados_armazenamento",
    ]
    merged = {key: int((existing or {}).get(key, 0)) for key in summary_keys}
    for key in summary_keys:
        merged[key] += int((incoming or {}).get(key, 0))
    return merged


def import_pending_xml_batch(
    progress_callback: Callable[[int, int], None] | None = None,
) -> tuple[dict[str, int], list[str], list[str], list[dict[str, object]]]:
    batch = ensure_xml_upload_batch()
    pending_keys = [file_key for file_key, item in batch.items() if not item.get("imported")]
    pending_uploads = get_pending_xml_batch_uploads()
    if not pending_uploads:
        return (
            dict(st.session_state.get("xml_upload_summary", {})),
            list(st.session_state.get("xml_upload_issues", [])),
            [],
            [],
        )

    pending_filenames = [str(upload.name) for upload in pending_uploads]
    parsed_records, parse_summary, parse_issues = parse_xml_upload_batch(pending_uploads, progress_callback)
    with measure("import.xml_operacional_persist"):
        summary, issues = persist_xml_records(parsed_records, parse_summary, parse_issues)
    with measure("import.xml_documental_persist"):
        issues.extend(persist_documental_xml_batch_phase(parsed_records, issues, pending_uploads))
    finalize_xml_upload_batch_after_import(pending_keys, parsed_records, issues)
    return summary, issues, pending_filenames, parsed_records


def apply_xml_batch_import_result(
    summary: dict[str, int],
    issues: list[str],
    *,
    selected_count: int = 0,
    upload_duplicate_messages: list[str] | None = None,
    pending_filenames: list[str] | None = None,
    elapsed_seconds: float = 0.0,
    parsed_records: list[dict[str, object]] | None = None,
) -> None:
    merged_summary = merge_import_summary(st.session_state.get("xml_upload_summary", {}), summary)
    merged_issues = list(st.session_state.get("xml_upload_issues", [])) + list(issues)
    st.session_state["runtime_refresh_required"] = True
    invalidate_balcao_lookup_cache()
    clear_contexto_operacional()
    st.session_state.xml_upload_message = format_xml_import_summary_message(summary)
    st.session_state.xml_upload_summary = merged_summary
    st.session_state.xml_upload_error = ""
    st.session_state.xml_upload_issues = merged_issues

    nf_to_arquivo = {
        str(record.get("NF", "")): str(record.get("Arquivo", ""))
        for record in (parsed_records or [])
        if record.get("NF") and record.get("Arquivo")
    }
    incoming_report = build_xml_import_report(
        summary=summary,
        issues=issues,
        upload_duplicate_messages=upload_duplicate_messages,
        selected_count=selected_count,
        pending_filenames=pending_filenames,
        elapsed_seconds=elapsed_seconds,
        nf_to_arquivo=nf_to_arquivo,
    )
    st.session_state.xml_import_report = merge_xml_import_reports(
        st.session_state.get("xml_import_report"),
        incoming_report,
    )


def _record_upload_only_xml_import_report(
    *,
    selected_count: int,
    upload_duplicate_messages: list[str],
) -> None:
    if not upload_duplicate_messages and selected_count <= 0:
        return
    incoming_report = build_xml_import_report(
        summary={},
        issues=[],
        upload_duplicate_messages=upload_duplicate_messages,
        selected_count=selected_count,
    )
    st.session_state.xml_import_report = merge_xml_import_reports(
        st.session_state.get("xml_import_report"),
        incoming_report,
    )


def run_pending_xml_batch_import(
    *,
    selected_count: int = 0,
    upload_duplicate_messages: list[str] | None = None,
) -> None:
    pending_count = len(get_pending_xml_batch_uploads())
    if pending_count <= 0:
        return

    progress_bar = st.progress(0.0, text="Preparando importacao em lote...")
    progress_caption = st.empty()

    def update_import_progress(current: int, total: int) -> None:
        if total <= 0:
            return
        progress_value = min(float(current) / float(total), 1.0)
        progress_bar.progress(progress_value, text=f"Processando XMLs: {current}/{total}")
        progress_caption.caption(f"Lendo e validando arquivos: {current} de {total}")

    try:
        started_at = time.perf_counter()
        with measure("import.xml_batch"):
            summary, issues, pending_filenames, parsed_records = import_pending_xml_batch(update_import_progress)
        elapsed_seconds = time.perf_counter() - started_at
        apply_xml_batch_import_result(
            summary,
            issues,
            selected_count=selected_count,
            upload_duplicate_messages=upload_duplicate_messages,
            pending_filenames=pending_filenames,
            elapsed_seconds=elapsed_seconds,
            parsed_records=parsed_records,
        )
    except ValueError as exc:
        st.session_state.xml_upload_message = ""
        st.session_state.xml_upload_error = str(exc)
    finally:
        progress_bar.empty()
        progress_caption.empty()


def handle_xml_upload_selection(uploaded_files: list | None) -> None:
    if not uploaded_files:
        return

    try:
        selected_count = len(uploaded_files)
        added, _, duplicate_messages = merge_uploaded_files_into_xml_batch(uploaded_files)
        if added > 0:
            run_pending_xml_batch_import(
                selected_count=selected_count,
                upload_duplicate_messages=duplicate_messages,
            )
        elif duplicate_messages:
            _record_upload_only_xml_import_report(
                selected_count=selected_count,
                upload_duplicate_messages=duplicate_messages,
            )
    except ValueError as exc:
        st.session_state.xml_upload_message = ""
        st.session_state.xml_upload_error = str(exc)


def render_xml_upload_batch_list() -> None:
    visible_items = get_accepted_xml_upload_batch_items()
    if not visible_items:
        return

    st.caption(f"{len(visible_items)} arquivo(s) selecionado(s)")
    for file_key, item in visible_items:
        name_col, remove_col = st.columns([6, 1], gap="small")
        with name_col:
            st.markdown(f"**{html.escape(str(item.get('name', 'arquivo.xml')))}**")
        with remove_col:
            if st.button("❌", key=f"xml_remove_{hashlib.sha256(file_key.encode()).hexdigest()[:16]}", help="Remover"):
                remove_xml_from_upload_batch(file_key)
                st.rerun()

    st.markdown("---")


def format_xml_import_summary_message(summary: dict[str, int]) -> str:
    return (
        "Importacao concluida: "
        f"{summary.get('total_arquivos', 0)} arquivo(s) • "
        f"{summary.get('processados', 0)} processados • "
        f"{summary.get('erros', 0)} com erro • "
        f"{summary.get('duplicados', 0)} duplicados • "
        f"{summary.get('ignorados', 0)} ignorados"
    )


def _read_xml_upload_bytes(xml_file) -> bytes:
    if hasattr(xml_file, "getvalue"):
        return bytes(xml_file.getvalue())
    if isinstance(xml_file, (bytes, bytearray)):
        return bytes(xml_file)
    return bytes(getattr(xml_file, "data", b"") or b"")


def _build_documental_items_for_phase2(
    parsed_records: list[dict[str, object]],
    parse_issues: list[str],
    xml_files: list | None = None,
) -> list[XmlDocumentalItem]:
    rejected = extract_rejected_xml_filenames_from_issues(parse_issues, parsed_records)
    lookup: dict[str, tuple[bytes, str]] = {}

    for item in ensure_xml_upload_batch().values():
        name = str(item.get("name", "") or "").strip()
        if not name:
            continue
        data = bytes(item.get("data", b"") or b"")
        if not data:
            continue
        file_hash = str(item.get("hash_sha256", "") or "").strip() or compute_xml_file_hash(data)
        lookup[name] = (data, file_hash)

    for xml_file in xml_files or []:
        name = str(getattr(xml_file, "name", "arquivo.xml") or "arquivo.xml").strip()
        if name in lookup:
            continue
        data = _read_xml_upload_bytes(xml_file)
        if not data:
            continue
        lookup[name] = (data, compute_xml_file_hash(data))

    items: list[XmlDocumentalItem] = []
    seen_chaves: set[str] = set()
    for record in parsed_records:
        arquivo = str(record.get("Arquivo", "") or "").strip()
        if not arquivo or arquivo in rejected:
            continue
        chave = normalize_chave_nfe(record.get("ChaveNFe", ""))
        if not chave or chave in seen_chaves:
            continue
        seen_chaves.add(chave)
        data, file_hash = lookup.get(arquivo, (b"", ""))
        if not data:
            continue
        items.append(
            XmlDocumentalItem(
                file_bytes=data,
                hash_sha256=file_hash,
                original_filename=arquivo,
                chave_nfe=chave,
            )
        )
    return items


def persist_documental_xml_batch_phase(
    parsed_records: list[dict[str, object]],
    parse_issues: list[str],
    xml_files: list | None = None,
) -> list[str]:
    try:
        items = _build_documental_items_for_phase2(parsed_records, parse_issues, xml_files)
        if not items:
            return []
        user = get_current_user()
        usuario_id = int(user.id) if user and user.id else None
        result = _get_documento_xml_service().persist_raw_xml_batch(items, usuario_id=usuario_id)
        _LOGGER.info(
            "Fase documental XML concluida em %.1f ms (saved=%s reused=%s skipped=%s)",
            result.elapsed_ms,
            result.saved,
            result.reused,
            result.skipped,
        )
        return list(result.issues)
    except Exception as exc:
        _LOGGER.warning("Falha na fase documental XML: %s", exc, exc_info=True)
        return [f"Persistencia documental dos XMLs nao concluida: {exc}"]


def parse_xml_upload_batch(
    xml_files: list,
    progress_callback: Callable[[int, int], None] | None = None,
) -> tuple[list[dict[str, object]], dict[str, int], list[str]]:
    issues: list[str] = []
    summary = {
        "total_arquivos": len(xml_files or []),
        "erros": 0,
        "duplicados_lote": 0,
    }
    batch_lookup: dict[str, dict[str, object]] = {}
    total_files = len(xml_files or [])

    for index, xml_file in enumerate(xml_files or []):
        if progress_callback is not None:
            progress_callback(index, total_files)

        xml_data = parse_xml_file(xml_file)
        if xml_data.get("Erro"):
            summary["erros"] += 1
            issues.append(f"Erro no XML {xml_data.get('Arquivo', 'arquivo.xml')}: {xml_data.get('Status', '')}")
            continue

        serialized = serialize_xml_record(xml_data)
        identity = get_xml_identity(serialized)
        if not identity:
            summary["erros"] += 1
            issues.append(f"XML sem chave/NF identificavel: {serialized.get('Arquivo', 'arquivo.xml')}")
            continue

        current_record = batch_lookup.get(identity)
        if current_record is None:
            batch_lookup[identity] = serialized
            continue

        if should_replace_xml_record(current_record, serialized):
            batch_lookup[identity] = serialized
            issues.append(
                f"NF {serialized.get('NF', '--')} duplicada no lote. Foi mantido o arquivo mais recente: "
                f"{serialized.get('Arquivo', 'arquivo.xml')}"
            )
            continue

        summary["duplicados_lote"] += 1
        issues.append(f"XML duplicado no lote ignorado: {serialized.get('Arquivo', 'arquivo.xml')}")

    if progress_callback is not None:
        progress_callback(total_files, total_files)

    return list(batch_lookup.values()), summary, issues


def persist_xml_records(
    parsed_records: list[dict[str, object]],
    parse_summary: dict[str, int],
    parse_issues: list[str],
) -> tuple[dict[str, int], list[str]]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    existing_records, load_error = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    issues = list(parse_issues)
    summary = {
        "total_arquivos": int(parse_summary.get("total_arquivos", 0)),
        "erros": int(parse_summary.get("erros", 0)),
        "duplicados_lote": int(parse_summary.get("duplicados_lote", 0)),
        "novas": 0,
        "atualizadas": 0,
        "ignoradas_separadas": 0,
        "duplicados_armazenamento": 0,
        "duplicados": int(parse_summary.get("duplicados_lote", 0)),
        "processados": 0,
        "ignorados": 0,
    }
    if load_error:
        issues.append(f"Importacao de XMLs abortada: {load_error}")
        summary["ignorados"] = summary["total_arquivos"]
        return summary, issues

    existing_separacao_records, _ = carregar_separacao_json(str(SEPARACAO_JSON_PATH))
    locked_separacao_groups = group_separacao_records_by_identity(existing_separacao_records)
    locked_identities = {
        identity
        for identity, records in locked_separacao_groups.items()
        if is_separacao_group_locked(records)
    }
    storage_lookup: dict[str, dict[str, object]] = {}
    delta_records: list[dict[str, object]] = []

    for existing_record in existing_records:
        normalized_record = serialize_xml_record(existing_record)
        identity = get_xml_identity(normalized_record)
        if identity:
            storage_lookup[identity] = normalized_record

    for serialized in parsed_records:
        identity = get_xml_identity(serialized)
        if not identity:
            continue

        if identity in locked_identities:
            issues.append(f"NF {serialized.get('NF', '--')} ignorada no upload porque ja esta separada.")
            summary["ignoradas_separadas"] += 1
            continue

        current_record = storage_lookup.get(identity)
        if current_record is None:
            storage_lookup[identity] = serialized
            delta_records.append(serialized)
            summary["novas"] += 1
            continue

        if should_replace_xml_record(current_record, serialized):
            storage_lookup[identity] = serialized
            delta_records.append(serialized)
            summary["atualizadas"] += 1
            issues.append(f"NF {serialized.get('NF', '--')} atualizada pelo evento/XML mais recente.")
        else:
            summary["duplicados_armazenamento"] += 1
            summary["duplicados"] += 1
            issues.append(f"XML duplicado ou desatualizado ignorado: {serialized.get('Arquivo', 'arquivo.xml')}")

    if delta_records:
        SqlXmlRecordRepository().upsert_records(delta_records)
    mark_persistence_layer_stale(reference=True)

    summary["processados"] = summary["novas"] + summary["atualizadas"]
    summary["ignorados"] = summary["ignoradas_separadas"]
    return summary, issues


def import_xml_upload_batch(
    xml_files: list,
    progress_callback: Callable[[int, int], None] | None = None,
) -> tuple[dict[str, int], list[str]]:
    total_files = len(xml_files or [])
    if total_files > MAX_XML_UPLOAD_BATCH:
        raise ValueError(f"Limite maximo de {MAX_XML_UPLOAD_BATCH} XMLs por operacao.")

    parsed_records, parse_summary, parse_issues = parse_xml_upload_batch(xml_files, progress_callback)
    summary, issues = persist_xml_records(parsed_records, parse_summary, parse_issues)
    issues.extend(persist_documental_xml_batch_phase(parsed_records, issues, xml_files))
    return summary, issues


def salvar_xmls_processados_records(xml_files: list) -> tuple[dict[str, int], list[str]]:
    return import_xml_upload_batch(xml_files)


def salvar_xmls_processados_json(xml_files: list) -> tuple[dict[str, int], list[str]]:
    return salvar_xmls_processados_records(xml_files)


@st.cache_data(show_spinner=False)
def carregar_xmls_processados_records(json_path: str) -> tuple[list[dict[str, object]], str]:
    _ = json_path
    try:
        records = SqlXmlRecordRepository().list_all_records()
        return [serialize_xml_record(item) for item in records], ""
    except Exception as exc:
        return [], f"Os XMLs salvos no sistema nao puderam ser lidos ({exc}). Envie novos arquivos para atualizar a base."


def carregar_xmls_processados_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    return carregar_xmls_processados_records(json_path)


def get_xml_storage_status() -> tuple[bool, str]:
    repository = SqlXmlRecordRepository()
    if repository.count_records() <= 0:
        return False, ""
    updated_at = repository.get_last_updated_at()
    if updated_at is None:
        return True, ""
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
def carregar_separacao_records(json_path: str) -> tuple[list[dict[str, object]], str]:
    _ = json_path
    try:
        payload = _CONFIG_STORAGE.load_list(CONFIG_CHAVE_SEPARACAO, default=[])
    except Exception as exc:
        return [], f"A base de separacao nao pôde ser lida ({exc})."

    if not isinstance(payload, list):
        return [], "A base de separacao esta em formato invalido."

    return sort_separacao_records([serialize_separacao_record(item) for item in payload if isinstance(item, dict)]), ""


def carregar_separacao_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    return carregar_separacao_records(json_path)


def salvar_separacao_records(records: list[dict[str, object]]) -> None:
    serialized_records = [serialize_separacao_record(record) for record in records]
    _CONFIG_STORAGE.save_list(CONFIG_CHAVE_SEPARACAO, sort_separacao_records(serialized_records))
    mark_persistence_layer_stale(operational=True)
    invalidate_latest_closed_lote_pdf_cache()


def salvar_separacao_json(records: list[dict[str, object]]) -> None:
    salvar_separacao_records(records)


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
def carregar_lotes_records(json_path: str) -> tuple[list[dict[str, object]], str]:
    _ = json_path
    try:
        payload = _CONFIG_STORAGE.load_list(CONFIG_CHAVE_LOTES, default=[])
    except Exception as exc:
        return [], f"A base de lotes nao pôde ser lida ({exc})."

    if not isinstance(payload, list):
        return [], "A base de lotes esta em formato invalido."

    return [serialize_lote_record(item) for item in payload if isinstance(item, dict)], ""


def carregar_lotes_json(json_path: str) -> tuple[list[dict[str, object]], str]:
    return carregar_lotes_records(json_path)


def salvar_lotes_records(records: list[dict[str, object]]) -> None:
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
    _CONFIG_STORAGE.save_list(CONFIG_CHAVE_LOTES, normalized_records)
    mark_persistence_layer_stale(operational=True)


def salvar_lotes_json(records: list[dict[str, object]]) -> None:
    salvar_lotes_records(records)


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

        route = normalize_route_label(normalized_xml.get("ROTA", ""))
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
    has_records, revision = get_separacao_storage_status_from_db()
    if not has_records:
        return False, ""
    if revision:
        try:
            parsed = datetime.fromisoformat(str(revision).replace("Z", "+00:00"))
            return True, format_datetime_display(parsed)
        except ValueError:
            return True, str(revision)
    return True, ""


@st.cache_data(show_spinner=False)
def carregar_separacao_excluidos_records(json_path: str) -> set[str]:
    _ = json_path
    return _CONFIG_STORAGE.load_set(CONFIG_CHAVE_SEPARACAO_EXCLUIDOS)


def carregar_separacao_excluidos_json(json_path: str) -> set[str]:
    return carregar_separacao_excluidos_records(json_path)


def salvar_separacao_excluidos_records(identities: set[str]) -> None:
    normalized_identities = sorted({str(identity or "").strip() for identity in identities if str(identity or "").strip()})
    if not normalized_identities:
        _CONFIG_STORAGE.save_list(CONFIG_CHAVE_SEPARACAO_EXCLUIDOS, [])
    else:
        _CONFIG_STORAGE.save_list(CONFIG_CHAVE_SEPARACAO_EXCLUIDOS, normalized_identities)
    mark_persistence_layer_stale(operational=True)


def salvar_separacao_excluidos_json(identities: set[str]) -> None:
    salvar_separacao_excluidos_records(identities)


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
        SqlXmlRecordRepository().replace_all_records([])
        mark_persistence_layer_stale(reference=True, operational=True)
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
        SqlXmlRecordRepository().replace_all_records(current_xml_records)
        mark_persistence_layer_stale(reference=True)
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
    mark_persistence_layer_stale(reference=True, operational=True)
    st.session_state["runtime_xml_records"] = []
    st.session_state["runtime_classificacao_records"] = []
    invalidate_balcao_lookup_cache()
    invalidate_latest_closed_lote_pdf_cache()


def mark_persistence_layer_stale(*, reference: bool = False, operational: bool = False) -> None:
    """Invalida caches de leitura apos escrita confirmada no PostgreSQL."""
    if reference:
        st.session_state["runtime_data_signature"] = None
        carregar_xmls_processados_records.clear()
        carregar_classificacao_produtos_records.clear()
    if operational:
        st.session_state["runtime_operational_signature"] = None
        carregar_separacao_records.clear()
        carregar_lotes_records.clear()
        carregar_separacao_excluidos_records.clear()
    if reference or operational:
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
def prepare_processed_search_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    prepared_df = ensure_route_column(dataframe)
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
                        "ValorNF": xml_data.get("ValorNF", 0.0),
                        "Destinatario": xml_data["Destinatario"],
                        "Municipio": xml_data["Municipio"],
                        "UF": xml_data.get("UF", ""),
                        "ClienteExcel": metadata_row.get("ClienteExcel", ""),
                        "CidadeExcel": metadata_row.get("CidadeExcel", ""),
                        "UFExcel": metadata_row.get("UFExcel", ""),
                        "ValorExcel": metadata_row.get("ValorExcel", 0.0),
                        "PesoExcel": metadata_row.get("PesoExcel", 0.0),
                        "VolumeExcel": metadata_row.get("VolumeExcel", 0.0),
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
                        "ValorNF": xml_data.get("ValorNF", 0.0),
                        "Destinatario": xml_data["Destinatario"],
                        "Municipio": xml_data["Municipio"],
                        "UF": xml_data.get("UF", ""),
                        "ClienteExcel": metadata_row.get("ClienteExcel", ""),
                        "CidadeExcel": metadata_row.get("CidadeExcel", ""),
                        "UFExcel": metadata_row.get("UFExcel", ""),
                        "ValorExcel": metadata_row.get("ValorExcel", 0.0),
                        "PesoExcel": metadata_row.get("PesoExcel", 0.0),
                        "VolumeExcel": metadata_row.get("VolumeExcel", 0.0),
                        "Status": str(xml_data["Status"]),
                    }
                )

        processed_df = pd.DataFrame(rows)
        if processed_df.empty:
            processed_df = create_empty_processed_df()
        else:
            processed_df = processed_df.sort_values(by=["Seq_sort", "NF"], ascending=[False, True], na_position="last")
        processed_df = apply_routes_from_xml_index(processed_df, xml_index)

        display_df = processed_df[TABLE_COLUMNS].copy()
        error_mask = ~display_df["Status"].astype(str).str.contains("autorizado", case=False, na=False)
        item_mask = display_df["cProd"].astype(str).str.strip() != ""
        metadata = summarize_metadata(base_df)

        summary = {
            "filial": summarize_filial(base_df),
            "numero_carga": metadata["numero_carga"],
            "data_saida": metadata["data_saida"],
            "transportadora": metadata["transportadora"],
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
                    "ValorNF": 0.0,
                    "Destinatario": "",
                    "Municipio": "",
                    "UF": "",
                    "ClienteExcel": row.get("ClienteExcel", ""),
                    "CidadeExcel": row.get("CidadeExcel", ""),
                    "UFExcel": row.get("UFExcel", ""),
                    "ValorExcel": row.get("ValorExcel", 0.0),
                    "PesoExcel": row.get("PesoExcel", 0.0),
                    "VolumeExcel": row.get("VolumeExcel", 0.0),
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
                    "ValorNF": xml_data.get("ValorNF", 0.0),
                    "Destinatario": xml_data["Destinatario"],
                    "Municipio": xml_data["Municipio"],
                    "UF": xml_data.get("UF", ""),
                    "ClienteExcel": row.get("ClienteExcel", ""),
                    "CidadeExcel": row.get("CidadeExcel", ""),
                    "UFExcel": row.get("UFExcel", ""),
                    "ValorExcel": row.get("ValorExcel", 0.0),
                    "PesoExcel": row.get("PesoExcel", 0.0),
                    "VolumeExcel": row.get("VolumeExcel", 0.0),
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
                    "ValorNF": xml_data.get("ValorNF", 0.0),
                    "Destinatario": xml_data["Destinatario"],
                    "Municipio": xml_data["Municipio"],
                    "UF": xml_data.get("UF", ""),
                    "ClienteExcel": row.get("ClienteExcel", ""),
                    "CidadeExcel": row.get("CidadeExcel", ""),
                    "UFExcel": row.get("UFExcel", ""),
                    "ValorExcel": row.get("ValorExcel", 0.0),
                    "PesoExcel": row.get("PesoExcel", 0.0),
                    "VolumeExcel": row.get("VolumeExcel", 0.0),
                    "Status": str(xml_data["Status"]),
                }
            )

    processed_df = pd.DataFrame(rows)
    if processed_df.empty:
        processed_df = create_empty_processed_df()
    else:
        processed_df = processed_df.sort_values(by=["Seq_sort", "NF"], ascending=[False, True], na_position="last")
    processed_df = apply_routes_from_xml_index(processed_df, xml_index)

    display_df = processed_df[TABLE_COLUMNS].copy()
    error_mask = ~display_df["Status"].astype(str).str.contains("autorizado", case=False, na=False)
    item_mask = display_df["cProd"].astype(str).str.strip() != ""
    metadata = summarize_metadata(base_df)

    summary = {
        "filial": summarize_filial(base_df),
        "numero_carga": metadata["numero_carga"],
        "data_saida": metadata["data_saida"],
        "transportadora": metadata["transportadora"],
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
        "transportadora": "BRIDA LUBRIFICANTES LTDA",
        "motorista": "--",
        "placa": "--",
        "nf_count": 0,
        "item_count": 0,
        "peso_total": 0.0,
        "error_count": 0,
    }


def create_empty_processed_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "Seq_sort",
            "ChaveNFe",
            "Data",
            "Volume",
            "PesoTotalNF",
            "ValorNF",
            "Municipio",
            "UF",
            "ClienteExcel",
            "CidadeExcel",
            "UFExcel",
            "ValorExcel",
            "PesoExcel",
            "VolumeExcel",
            *TABLE_COLUMNS,
        ]
    )


def create_empty_nf_debug_df() -> pd.DataFrame:
    return pd.DataFrame(columns=NF_DEBUG_COLUMNS)


def build_balcao_lookup_dataframe(xml_records: list[dict[str, object]]) -> pd.DataFrame:
    xml_index, _ = build_xml_index_from_records(xml_records)
    if not xml_index:
        return create_empty_processed_df()

    rows: list[dict[str, object]] = []
    for index, xml_data in enumerate(xml_index.values(), start=1):
        if not xml_data.get("Items"):
            rows.append(
                {
                    "Seq": index,
                    "Seq_sort": index,
                    "ChaveNFe": xml_data.get("ChaveNFe", ""),
                    "NF": xml_data.get("NF", ""),
                    "Data": xml_data.get("Data", ""),
                    "cProd": "",
                    "Descricao": "",
                    "Qtd": 0.0,
                    "Unidade": "",
                    "Volume": xml_data.get("VolumeTotal", 0.0),
                    "Peso": 0.0,
                    "PesoTotalNF": xml_data.get("PesoTotal", 0.0),
                    "ValorNF": xml_data.get("ValorNF", 0.0),
                    "Destinatario": xml_data.get("Destinatario", ""),
                    "Municipio": xml_data.get("Municipio", ""),
                    "UF": xml_data.get("UF", ""),
                    "Status": str(xml_data.get("Status", "")),
                }
            )
            continue

        for item in xml_data.get("Items", []):
            rows.append(
                {
                    "Seq": index,
                    "Seq_sort": index,
                    "ChaveNFe": xml_data.get("ChaveNFe", ""),
                    "NF": xml_data.get("NF", ""),
                    "Data": xml_data.get("Data", ""),
                    "cProd": item.get("cProd", ""),
                    "Descricao": item.get("Descricao", ""),
                    "Qtd": item.get("Qtd", 0.0),
                    "Unidade": item.get("Unidade", ""),
                    "Volume": xml_data.get("VolumeTotal", 0.0),
                    "Peso": item.get("Peso", 0.0),
                    "PesoTotalNF": xml_data.get("PesoTotal", 0.0),
                    "ValorNF": xml_data.get("ValorNF", 0.0),
                    "Destinatario": xml_data.get("Destinatario", ""),
                    "Municipio": xml_data.get("Municipio", ""),
                    "UF": xml_data.get("UF", ""),
                    "Status": str(xml_data.get("Status", "")),
                }
            )

    processed_df = pd.DataFrame(rows)
    if processed_df.empty:
        return create_empty_processed_df()

    processed_df = processed_df.sort_values(by=["Seq_sort", "NF"], ascending=[False, True], na_position="last")
    return apply_routes_from_xml_index(processed_df, xml_index)


def build_balcao_summary() -> dict[str, object]:
    return {
        "filial": "BRIDA",
        "data_saida": datetime.now().strftime("%d/%m/%Y"),
        "nf_count": 0,
        "item_count": 0,
        "peso_total": 0.0,
    }


def get_balcao_lookup_dataframe(xml_records: list[dict[str, object]]) -> pd.DataFrame:
    runtime_signature = st.session_state.get("runtime_data_signature")
    cached_signature = st.session_state.get("_balcao_lookup_signature")
    cached_df = st.session_state.get("balcao_lookup_df")
    if cached_signature == runtime_signature and isinstance(cached_df, pd.DataFrame):
        return cached_df

    with measure("dataframe.balcao_lookup"):
        lookup_df = build_balcao_lookup_dataframe(xml_records)
    st.session_state["balcao_lookup_df"] = lookup_df
    st.session_state["_balcao_lookup_signature"] = runtime_signature
    return lookup_df


def get_prepared_processed_dataframe() -> pd.DataFrame:
    version = st.session_state.get("_processed_data_version", 0)
    if (
        st.session_state.get("_prepared_df_version") == version
        and isinstance(st.session_state.get("_prepared_processed_df"), pd.DataFrame)
    ):
        return st.session_state["_prepared_processed_df"]

    with measure("dataframe.prepare_processed"):
        prepared_df = prepare_processed_search_dataframe(st.session_state.processed_df)
    st.session_state["_prepared_processed_df"] = prepared_df
    st.session_state["_prepared_df_version"] = version
    return prepared_df


def get_display_table_dataframe(processed_df: pd.DataFrame) -> pd.DataFrame:
    version = st.session_state.get("_processed_data_version", 0)
    if (
        st.session_state.get("_display_table_version") == version
        and isinstance(st.session_state.get("_display_table_df"), pd.DataFrame)
    ):
        return st.session_state["_display_table_df"]

    with measure("dataframe.build_display_table"):
        display_df = build_display_table(processed_df[TABLE_COLUMNS].copy())
    st.session_state["_display_table_df"] = display_df
    st.session_state["_display_table_version"] = version
    return display_df


def get_latest_closed_lote_pdf_bytes(
    lote_summary: dict[str, object],
    lote_records: list[dict[str, object]],
    lote_id: str,
) -> bytes:
    if not lote_id or not lote_records:
        return b""

    signature = f"{lote_id}:{len(lote_records)}"
    cached_signature = st.session_state.get("_latest_closed_lote_pdf_sig")
    cached_pdf = st.session_state.get("_latest_closed_lote_pdf")
    if cached_signature == signature and isinstance(cached_pdf, (bytes, bytearray)) and cached_pdf:
        return bytes(cached_pdf)

    with measure("pdf.generate_lote"):
        pdf_bytes = generate_lote_pdf(lote_summary, lote_records)
    st.session_state["_latest_closed_lote_pdf"] = pdf_bytes
    st.session_state["_latest_closed_lote_pdf_sig"] = signature
    return pdf_bytes


from utils.streamlit_tables import build_table_column_config
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


def initialize_login_state() -> None:
    return


def normalize_screen_name(value: object) -> str:
    screen = str(value or "").strip().lower()
    screen_aliases = {
        "gestao_lotes": SCREEN_LOTES,
        SCREEN_LOGIN: SCREEN_LOGIN,
        SCREEN_MENU: SCREEN_MENU,
        SCREEN_MINUTA: SCREEN_MINUTA,
        SCREEN_SEPARACAO: SCREEN_SEPARACAO,
        SCREEN_LOTES: SCREEN_LOTES,
        SCREEN_USUARIOS: SCREEN_USUARIOS,
        SCREEN_CONSULTA_CARREGAMENTOS: SCREEN_CONSULTA_CARREGAMENTOS,
        SCREEN_GESTAO_DADOS: SCREEN_GESTAO_DADOS,
        SCREEN_GESTAO_RETENCAO: SCREEN_GESTAO_DADOS,
    }
    return screen_aliases.get(screen, SCREEN_MENU)


def get_minuta_module_config(screen_name: str) -> MinutaModuleConfig:
    return MINUTA_MODULES.get(screen_name, MINUTA_CARREGAMENTO_CONFIG)


def navegar(tela: str) -> None:
    target_screen = normalize_screen_name(tela)
    st.session_state["tela"] = target_screen
    if target_screen != SCREEN_LOGIN:
        st.session_state["menu_aberto"] = st.session_state.get("menu_aberto", True)
    st.rerun()


def initialize_navigation_state() -> None:
    legacy_screen = st.session_state.get("tela_atual", SCREEN_MINUTA)
    current_screen = normalize_screen_name(st.session_state.get("tela", legacy_screen))

    if not is_logged_in():
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


def render_metric_cards_row(cards: list[dict[str, object]]) -> None:
    normalized_cards = list(cards[:4])
    while len(normalized_cards) < 4:
        normalized_cards.append({"title": "--", "value": "--", "subtitle": "", "icon_key": "dados_gerais"})

    card_markup: list[str] = []
    for card in normalized_cards:
        title = html.escape(str(card.get("title", "") or ""))
        value = html.escape(str(card.get("value", "") or "--"))
        subtitle = html.escape(str(card.get("subtitle", "") or "")) or "&nbsp;"
        icon_key = str(card.get("icon_key", "dados_gerais") or "dados_gerais")
        icon_markup = render_label_icon(ICON_MAP.get(icon_key, ICON_MAP["dados_gerais"]))
        card_markup.append(
            "<div class=\"erp-card erp-card-kpi erp-card-kpi-fixed\">"
            f"<div class=\"erp-kpi-top\"><div class=\"erp-kpi-icon\">{icon_markup}</div><div class=\"erp-kpi-label\">{title}</div></div>"
            f"<div class=\"erp-kpi-value\">{value}</div>"
            f"<div class=\"erp-kpi-subtitle\">{subtitle}</div>"
            "</div>"
        )

    grid_markup = "<div class=\"erp-kpi-grid\">" + "".join(card_markup) + "</div>"
    st.markdown(grid_markup, unsafe_allow_html=True)


def render_section_heading(label: str, icon_key: str) -> None:
    st.markdown(
        f'''
    <div class="section-title-block with-icon">{render_label_icon(ICON_MAP[icon_key])}<span>{label}</span></div>
    ''',
        unsafe_allow_html=True,
    )


@contextmanager
def ui_section_box(extra_classes: str = ""):
    """Secao com borda nativa do Streamlit — evita divs HTML abertos/fechados em reruns."""
    _ = extra_classes
    with st.container(border=True):
        yield


_ui_section_stack: list = []


def render_box_open(extra_classes: str = "") -> None:
    """Abre secao com borda nativa (compatibilidade com blocos open/close legados)."""
    context = ui_section_box(extra_classes)
    _ui_section_stack.append(context)
    context.__enter__()


def render_box_close() -> None:
    """Fecha secao aberta por render_box_open."""
    if _ui_section_stack:
        context = _ui_section_stack.pop()
        context.__exit__(None, None, None)


def render_login_screen() -> None:
    render_login_page(
        logo_path=get_logo_path(),
        on_success_screen=SCREEN_MENU,
        navigate_callback=lambda screen: st.session_state.update({"tela": screen}),
    )


def logout() -> None:
    clear_session_on_logout()
    st.session_state["tela"] = SCREEN_LOGIN
    st.rerun()


def toggle_menu() -> None:
    st.session_state["menu_aberto"] = not st.session_state.get("menu_aberto", True)


def render_sidebar_dados_navigation() -> None:
    current_screen = normalize_screen_name(st.session_state.get("tela", SCREEN_MENU))
    button_type = "primary" if current_screen == SCREEN_GESTAO_DADOS else "secondary"
    if st.button(
        "🗃 Gestao de Dados",
        use_container_width=True,
        key="sidebar_nav_gestao_dados",
        type=button_type,
    ):
        navegar(SCREEN_GESTAO_DADOS)


def render_sidebar_retencao_navigation() -> None:
    render_sidebar_dados_navigation()


def maybe_prompt_capacidade_critica() -> None:
    if st.session_state.get("_capacidade_prompt_checked"):
        return
    st.session_state["_capacidade_prompt_checked"] = True
    if st.session_state.get("_capacidade_prompt_dismissed"):
        return
    try:
        capacidade = get_gestao_capacidade_service().avaliar_capacidade()
    except Exception:
        return
    if not capacidade.requer_dialogo_login:
        return
    st.session_state["_capacidade_prompt_snapshot"] = capacidade
    st.session_state["_capacidade_prompt_pending"] = True


def maybe_prompt_gestao_dados() -> None:
    if st.session_state.get("_gestao_dados_prompt_checked"):
        return
    st.session_state["_gestao_dados_prompt_checked"] = True
    if st.session_state.get("_gestao_dados_prompt_dismissed"):
        return
    try:
        if not get_gestao_dados_service().possui_carregamentos_elegiveis():
            return
    except Exception:
        return

    st.session_state["_gestao_dados_prompt_pending"] = True


def maybe_prompt_retencao_expirada() -> None:
    maybe_prompt_gestao_dados()


def _formatar_bytes_resumo(value: int) -> str:
    total = max(int(value), 0)
    if total >= 1024 * 1024:
        return f"{total / (1024 * 1024):.1f} MB"
    if total >= 1024:
        return f"{total / 1024:.1f} KB"
    return f"{total} B"


def render_capacidade_login_prompt() -> None:
    pending = bool(st.session_state.pop("_capacidade_prompt_pending", False))
    if not pending:
        return
    if st.session_state.get("_capacidade_prompt_dismissed"):
        return

    capacidade = st.session_state.pop("_capacidade_prompt_snapshot", None)
    if capacidade is None:
        return

    previa = None
    try:
        previa = get_gestao_capacidade_service().montar_previa_dia_mais_antigo()
    except Exception:
        pass

    usuario = get_current_user()
    if usuario is not None:
        try:
            get_gestao_capacidade_service().registrar_auditoria_capacidade(
                usuario_id=int(usuario.id),
                capacidade=capacidade,
                previa=previa,
            )
        except Exception:
            pass

    pct = capacidade.percentual
    pct_texto = f"{pct:.0f}%" if pct is not None else "90%"
    with st.container(border=True):
        st.markdown("#### Capacidade do Banco de Dados")
        st.warning(
            f"O banco de dados atingiu **{pct_texto}** da capacidade operacional.\n\n"
            "Para manter a estabilidade do sistema recomendamos executar a retencao do dia mais antigo.\n\n"
            "Deseja realizar agora?"
        )
        if previa is None:
            st.caption("Nao ha dia elegivel para retencao sugerida no momento.")
        col_executar, col_depois = st.columns(2)
        with col_executar:
            if st.button("Executar Retencao", use_container_width=True, key="capacidade_prompt_executar"):
                st.session_state["gestao_dados_capacidade_iniciar"] = True
                navegar(SCREEN_GESTAO_DADOS)
        with col_depois:
            if st.button("Depois", use_container_width=True, key="capacidade_prompt_depois"):
                st.session_state["_capacidade_prompt_dismissed"] = True
                st.rerun()


def render_gestao_dados_login_prompt() -> None:
    render_capacidade_login_prompt()
    pending = bool(
        st.session_state.pop("_gestao_dados_prompt_pending", False)
        or st.session_state.pop("_retencao_prompt_pending", False)
    )
    if not pending:
        return
    if st.session_state.get("_gestao_dados_prompt_dismissed") or st.session_state.get("_retencao_prompt_dismissed"):
        return

    try:
        painel = get_gestao_dados_service().obter_painel()
    except Exception:
        return

    pacote = painel.pacote
    with st.container(border=True):
        st.markdown("#### Gestao de Dados")
        st.warning(
            f"Foram encontrados **{pacote.carregamentos:,} carregamentos** "
            "fora da politica de retencao."
        )
        st.markdown(
            f"**Espaco estimado recuperavel:** {_formatar_bytes_resumo(pacote.espaco_recuperavel_bytes)}"
        )
        st.caption("Deseja visualizar a analise? Nenhum dado sera removido nesta etapa.")
        col_abrir, col_depois = st.columns(2)
        with col_abrir:
            if st.button("Abrir Gestao", use_container_width=True, key="gestao_dados_prompt_abrir"):
                navegar(SCREEN_GESTAO_DADOS)
        with col_depois:
            if st.button("Depois", use_container_width=True, key="gestao_dados_prompt_depois"):
                st.session_state["_gestao_dados_prompt_dismissed"] = True
                st.session_state["_retencao_prompt_dismissed"] = True
                st.rerun()


def render_retencao_login_prompt() -> None:
    render_gestao_dados_login_prompt()


def apply_sidebar_visibility(menu_aberto: bool) -> None:
    """Sem CSS no sidebar — expandir/recolher fica a cargo do Streamlit."""
    _ = menu_aberto


def _render_minuta_upload_content(xml_records: list) -> tuple[list, object, bool]:
    """Conteudo da coluna esquerda: XML, Excel, processar e estatisticas."""
    with ui_section_box("is-sidebar is-soft"):
        st.markdown(
            f'''
    <div class="sidebar-heading with-icon">{render_label_icon(ICON_MAP["dados_gerais"])}<span>Arquivos</span></div>
    ''',
            unsafe_allow_html=True,
        )

    if is_logged_in():
        with ui_section_box("is-sidebar is-soft"):
            render_logged_user_badge()

    with ui_section_box("is-sidebar"):
        st.markdown(
            f'''
    <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["xml"])}<span>Upload XML</span></div>
    ''',
            unsafe_allow_html=True,
        )

        if not has_accepted_xml_upload_batch():
            xml_files = st.file_uploader(
                "XMLs",
                type=["xml"],
                accept_multiple_files=True,
                key="xml_upload_widget",
            )
            if xml_files:
                handle_xml_upload_selection(xml_files)
        else:
            render_xml_upload_batch_list()

            if st.button("Adicionar XMLs", use_container_width=True, key="xml_add_files_button", type="secondary"):
                st.session_state["xml_add_uploader_open"] = True

            if st.session_state.get("xml_add_uploader_open", False):
                added_files = st.file_uploader(
                    "Adicionar XMLs",
                    type=["xml"],
                    accept_multiple_files=True,
                    key="xml_upload_add_widget",
                )
                if added_files:
                    handle_xml_upload_selection(added_files)

        render_xml_import_summary_panel(
            st.session_state.get("xml_import_report"),
            error_message=str(st.session_state.get("xml_upload_error", "") or ""),
        )

        has_xml_storage, xml_updated_at = get_xml_storage_status()
        if has_xml_storage:
            st.caption(f"Dados carregados do sistema • Ultima atualizacao: {xml_updated_at}")

    current_screen = normalize_screen_name(st.session_state.get("tela", SCREEN_MENU))
    if current_screen != SCREEN_MINUTA:
        return xml_records, None, False

    with ui_section_box("is-sidebar"):
        st.markdown(
            f'''
    <div class="sidebar-field-label with-icon">{render_label_icon(ICON_MAP["excel"])}<span>Upload Excel</span></div>
    ''',
            unsafe_allow_html=True,
        )

        excel_file = st.file_uploader(
            "Excel",
            type=["xlsx", "xls"],
        )

        process_clicked = st.button("Processar", use_container_width=True)
        if process_clicked and get_pending_xml_batch_uploads():
            run_pending_xml_batch_import()
            xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))

    return xml_records, excel_file, process_clicked


def render_sidebar() -> tuple[list, object, bool]:
    with measure("sql.carregar_xmls_runtime"):
        xml_records, _ = load_runtime_reference_data()
    with st.sidebar:
        render_logged_user_badge()
        render_sidebar_dados_navigation()
        st.divider()
        xml_storage_error = str(st.session_state.get("runtime_xml_storage_error", "") or "")
        if xml_storage_error:
            st.warning(xml_storage_error)
        excel_file = None
        process_clicked = False
        xml_records, excel_file, process_clicked = _render_minuta_upload_content(xml_records)

    current_screen = normalize_screen_name(st.session_state.get("tela", SCREEN_MENU))
    if current_screen != SCREEN_MINUTA:
        return xml_records, None, False
    return xml_records, excel_file, process_clicked


def build_processing_pdf_bytes(
    processed_df: pd.DataFrame,
    summary: dict[str, object],
    module_config: MinutaModuleConfig,
    *,
    gerar_minuta: bool,
    gerar_romaneio: bool,
) -> tuple[bytes | None, bytes | None]:
    carregamento_pdf_bytes = b""
    entrega_pdf_bytes = b""
    operador = get_logged_operator_display_name()
    impresso_em = format_datetime_display()
    if gerar_minuta:
        minuta_records = build_minuta_records(processed_df)
        if minuta_records:
            with measure("pdf.generate_minuta_carregamento"):
                carregamento_pdf_bytes = generate_minuta_pdf(
                    dados_minuta=minuta_records,
                    numero_carga=str(summary.get("numero_carga", "--") or "--"),
                    data_emissao=str(st.session_state.document_issue_at or "--"),
                    veiculo=str(summary.get("placa", "--") or "--"),
                    motorista=str(summary.get("motorista", "--") or "--"),
                    pdf_title=module_config.pdf_title,
                    subject_label=module_config.subject_label,
                    operador=operador,
                    impresso_em=impresso_em,
                )
    if gerar_romaneio:
        entrega_records, _, entrega_totals = build_minuta_entrega_records(processed_df)
        transportadora = str(summary.get("transportadora", "BRIDA LUBRIFICANTES LTDA") or "BRIDA LUBRIFICANTES LTDA")
        if entrega_records:
            with measure("pdf.generate_minuta_entrega"):
                entrega_pdf_bytes = generate_minuta_entrega_pdf(
                    records=entrega_records,
                    totals=entrega_totals,
                    numero_documento=str(summary.get("numero_carga", "--") or "--"),
                    data_emissao=str(st.session_state.document_issue_at or "--"),
                    transportadora=transportadora,
                    veiculo=str(summary.get("placa", "--") or "--"),
                    motorista=str(summary.get("motorista", "--") or "--"),
                    placa=str(summary.get("placa", "--") or "--"),
                    operador=operador,
                    impresso_em=impresso_em,
                )
    return (
        carregamento_pdf_bytes or None,
        entrega_pdf_bytes or None,
    )


def build_pdf_download_package(
    *,
    carregamento_pdf_bytes: bytes | None,
    entrega_pdf_bytes: bytes | None,
    carregamento_selected: bool,
    entrega_selected: bool,
    numero_carga: str,
    xml_selected: bool = False,
    xml_entries: list[tuple[str, bytes]] | None = None,
) -> tuple[bytes, str, str, str]:
    return build_documentos_download_package(
        carregamento_pdf_bytes=carregamento_pdf_bytes,
        entrega_pdf_bytes=entrega_pdf_bytes,
        carregamento_selected=carregamento_selected,
        entrega_selected=entrega_selected,
        xml_selected=xml_selected,
        xml_entries=xml_entries,
        numero_carga=numero_carga,
    )


def _run_baixar_pdf_pipeline(
    *,
    summary: dict[str, object],
    processed_df: pd.DataFrame,
    balcao_lookup_df: pd.DataFrame,
    balcao_summary: dict[str, object],
    module_config: MinutaModuleConfig,
    has_excel_loaded: bool,
    carregamento_selected: bool,
    entrega_selected: bool,
    xml_selected: bool = False,
    confirmar_reimpressao: bool = False,
    force_reentrega: bool = False,
    balcao_termo: str = "",
) -> None:
    can_close = has_excel_loaded and not processed_df.empty
    decisao = confirmar_decisao_operacional_continuacao()
    if decisao == DecisaoOperacional.REIMPRIMIR:
        confirmar_reimpressao = True
    somente_xml = xml_selected and not carregamento_selected and not entrega_selected
    gerar_minuta_fechamento = carregamento_selected or somente_xml
    gerar_romaneio_fechamento = entrega_selected
    if balcao_termo:
        finalize_status, fechamento_result = executar_fechamento_balcao_para_pdf(
            termo_busca=balcao_termo,
            summary=balcao_summary,
            lookup_df=balcao_lookup_df,
            gerar_minuta=gerar_minuta_fechamento,
            gerar_romaneio=gerar_romaneio_fechamento,
            force_reentrega=force_reentrega,
            confirmar_reimpressao=confirmar_reimpressao,
            standalone_balcao=True,
        )
    elif can_close:
        diagnostico = get_diagnostico_efetivo_fechamento()
        if diagnostico is not None:
            if diagnostico.bloqueia_fechamento:
                st.session_state["carregamento_finalize_error"] = (
                    "; ".join(diagnostico.mensagens) or "Operacao bloqueada pela analise operacional."
                )
                return
            if diagnostico.requer_decisao and decisao is None:
                st.session_state["carregamento_finalize_error"] = (
                    "Selecione como deseja continuar no painel operacional antes de gerar os documentos."
                )
                return
        finalize_status, fechamento_result = executar_fechamento_veiculo_para_pdf(
            summary=summary,
            processed_df=processed_df,
            gerar_minuta=gerar_minuta_fechamento,
            gerar_romaneio=gerar_romaneio_fechamento,
            force_reentrega=force_reentrega,
            confirmar_reimpressao=confirmar_reimpressao,
            diagnostico=diagnostico,
            decisao=decisao,
        )
    else:
        st.session_state["carregamento_finalize_error"] = (
            "Processe um Excel valido antes de baixar os documentos."
        )
        return

    if finalize_status in {"needs_reentrega", "balcao_needs_reentrega"}:
        return
    if finalize_status == "needs_reimpressao_confirm":
        st.session_state["carregamento_finalize_error"] = (
            "Confirme a reimpressao no painel operacional antes de gerar os documentos."
        )
        return

    if finalize_status in {"saved", "reimpressao", "complementacao", "balcao_saved"} and fechamento_result and fechamento_result.carregamento:
        pdf_df = processed_df
        pdf_summary = summary
        if balcao_termo:
            pdf_df = localizar_nf_no_lote(balcao_lookup_df, balcao_termo)
            pdf_summary = balcao_summary
        numero_documento = str(fechamento_result.carregamento.numero_carregamento or "--")
        pdf_summary = {**pdf_summary, "numero_carga": numero_documento}
        carregamento_pdf_bytes, entrega_pdf_bytes = build_processing_pdf_bytes(
            pdf_df,
            pdf_summary,
            module_config,
            gerar_minuta=carregamento_selected,
            gerar_romaneio=entrega_selected,
        )
        persistir_pdfs_apos_fechamento(
            fechamento_result.carregamento,
            minuta_pdf=carregamento_pdf_bytes,
            romaneio_pdf=entrega_pdf_bytes,
        )
        xml_entries: list[tuple[str, bytes]] = []
        if xml_selected:
            from carregamentos.bootstrap import get_xml_export_service

            export_result = get_xml_export_service().collect_xmls_for_carregamento(
                int(fechamento_result.carregamento.id)
            )
            xml_entries = [(entry.nome_arquivo, entry.conteudo) for entry in export_result.entries]
        download_payload, download_name, download_mime, validation_message = build_pdf_download_package(
            carregamento_pdf_bytes=carregamento_pdf_bytes,
            entrega_pdf_bytes=entrega_pdf_bytes,
            carregamento_selected=carregamento_selected,
            entrega_selected=entrega_selected,
            xml_selected=xml_selected,
            xml_entries=xml_entries,
            numero_carga=numero_documento,
        )
        if download_payload:
            st.session_state["pdf_download_payload"] = download_payload
            st.session_state["pdf_download_name"] = download_name
            st.session_state["pdf_download_mime"] = download_mime
            st.session_state["carregamento_finalize_message"] = (
                f"Carregamento {fechamento_result.carregamento.numero_carregamento} "
                "salvo no banco. Documentos prontos para download."
            )
            if validation_message:
                st.session_state["carregamento_finalize_warning"] = validation_message
            else:
                st.session_state.pop("carregamento_finalize_warning", None)
        else:
            st.session_state.pop("carregamento_finalize_warning", None)
            st.session_state["carregamento_finalize_error"] = validation_message or (
                "Nao foi possivel gerar os documentos apos o fechamento."
            )
        return

    if finalize_status not in {"invalid", "error"}:
        st.session_state["carregamento_finalize_error"] = (
            "Nao foi possivel concluir o fechamento do carregamento."
        )


def _process_processing_screen_actions(
    *,
    summary: dict[str, object],
    processed_df: pd.DataFrame,
    balcao_lookup_df: pd.DataFrame,
    balcao_summary: dict[str, object],
    module_config: MinutaModuleConfig,
    has_excel_loaded: bool,
    carregamento_selected: bool,
    entrega_selected: bool,
    xml_selected: bool,
) -> tuple[str | None, str | None, str | None]:
    finalize_message = str(st.session_state.pop("carregamento_finalize_message", "") or "")
    finalize_error = str(st.session_state.pop("carregamento_finalize_error", "") or "")
    finalize_warning = str(st.session_state.pop("carregamento_finalize_warning", "") or "")

    action = st.session_state.pop("_processing_action", None)
    if not isinstance(action, dict):
        return finalize_message or None, finalize_error or None, finalize_warning or None

    action_type = str(action.get("type", "") or "")

    if action_type == "finalizar_carregamento":
        finalize_message = (
            "O fechamento oficial do carregamento ocorre ao clicar em Baixar PDF, "
            "apos validacao e gravacao no banco de dados."
        )
    elif action_type == "operacional_cancel":
        cancelar_operacao_pendente()
        finalize_error = "Operacao cancelada pelo operador."
    elif action_type == "reimpressao_cancel":
        clear_reimpressao_pending()
        finalize_error = "Reimpressao cancelada pelo operador."
    elif action_type == "reentrega_cancel":
        clear_reentrega_pending()
        clear_balcao_pending()
        finalize_error = (
            "Operacao cancelada. A NF nao pode ser associada novamente sem autorizacao de reentrega."
        )
    elif action_type == "balcao_cancel":
        clear_balcao_pending()
        finalize_error = "Retirada no balcao cancelada."
    elif action_type == "balcao_iniciar":
        termo = str(action.get("termo", "") or st.session_state.get("entrega_balcao_termo", "") or "")
        preview_status = iniciar_entrega_balcao(
            termo_busca=termo,
            lookup_df=balcao_lookup_df,
            standalone_balcao=True,
        )
        if preview_status == "balcao_not_found":
            finalize_error = "NF nao encontrada nos XMLs carregados no sistema."
    elif action_type == "reentrega_confirm":
        reentrega_contexto = str(st.session_state.get("reentrega_contexto", "veiculo") or "veiculo")
        if reentrega_contexto == "balcao":
            balcao_termo = str(st.session_state.get("reentrega_balcao_termo", "") or "")
            st.session_state["balcao_pending_confirm"] = True
            st.session_state["balcao_confirm_termo"] = balcao_termo
            st.session_state["balcao_force_reentrega"] = True
            clear_reentrega_pending()
        elif has_excel_loaded and not processed_df.empty:
            clear_reentrega_pending()
            _run_baixar_pdf_pipeline(
                summary=summary,
                processed_df=processed_df,
                balcao_lookup_df=balcao_lookup_df,
                balcao_summary=balcao_summary,
                module_config=module_config,
                has_excel_loaded=has_excel_loaded,
                carregamento_selected=carregamento_selected,
                entrega_selected=entrega_selected,
                xml_selected=xml_selected,
                force_reentrega=True,
            )
            finalize_message = str(st.session_state.pop("carregamento_finalize_message", "") or "") or finalize_message
            finalize_error = str(st.session_state.pop("carregamento_finalize_error", "") or "") or finalize_error
            finalize_warning = str(st.session_state.pop("carregamento_finalize_warning", "") or "") or finalize_warning
        else:
            finalize_error = "Processe um Excel valido antes de confirmar a reentrega."
    elif action_type == "balcao_confirm":
        balcao_termo = str(st.session_state.get("balcao_confirm_termo", "") or "")
        force_reentrega = bool(st.session_state.pop("balcao_force_reentrega", False))
        _run_baixar_pdf_pipeline(
            summary=summary,
            processed_df=processed_df,
            balcao_lookup_df=balcao_lookup_df,
            balcao_summary=balcao_summary,
            module_config=module_config,
            has_excel_loaded=has_excel_loaded,
            carregamento_selected=carregamento_selected,
            entrega_selected=entrega_selected,
            xml_selected=xml_selected,
            force_reentrega=force_reentrega,
            balcao_termo=balcao_termo,
        )
        finalize_message = str(st.session_state.pop("carregamento_finalize_message", "") or "") or finalize_message
        finalize_error = str(st.session_state.pop("carregamento_finalize_error", "") or "") or finalize_error
        finalize_warning = str(st.session_state.pop("carregamento_finalize_warning", "") or "") or finalize_warning
    elif action_type == "baixar_pdf":
        if action.get("confirmar_reimpressao"):
            clear_reimpressao_pending()
        exportacao = snapshot_exportacao_documentos(module_config.screen_key)
        _run_baixar_pdf_pipeline(
            summary=summary,
            processed_df=processed_df,
            balcao_lookup_df=balcao_lookup_df,
            balcao_summary=balcao_summary,
            module_config=module_config,
            has_excel_loaded=has_excel_loaded,
            carregamento_selected=exportacao["carregamento_selected"],
            entrega_selected=exportacao["entrega_selected"],
            xml_selected=exportacao["xml_selected"],
            confirmar_reimpressao=bool(action.get("confirmar_reimpressao", False)),
            force_reentrega=bool(action.get("force_reentrega", False)),
            balcao_termo=str(action.get("balcao_termo", "") or ""),
        )
        finalize_message = str(st.session_state.pop("carregamento_finalize_message", "") or "") or finalize_message
        finalize_error = str(st.session_state.pop("carregamento_finalize_error", "") or "") or finalize_error
        finalize_warning = str(st.session_state.pop("carregamento_finalize_warning", "") or "") or finalize_warning

    return finalize_message or None, finalize_error or None, finalize_warning or None



def _humanizar_mensagem_operacional(mensagem: str) -> str:
    texto = str(mensagem or "").strip()
    if not texto:
        return texto
    substituicoes = {
        "A operacao foi bloqueada.": "Consulte o historico de cada NF para entender o contexto.",
        "Operacao bloqueada": "Ocorrencia operacional identificada",
    }
    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)
    if "carregamentos diferentes" in texto:
        return (
            "As NFs desta planilha possuem historico vinculado a carregamentos diferentes. "
            "Consulte o historico individual de cada NF abaixo."
        )
    if "pertencem ao carregamento" in texto:
        return texto.replace(
            "pertencem ao carregamento",
            "possuem historico no carregamento",
        )
    if "ja pertencem ao carregamento" in texto:
        return texto.replace(
            "ja pertencem ao carregamento",
            "ja possuem historico no carregamento",
        )
    return texto


def _build_operational_panel_copy(mode: str) -> tuple[str, str, str]:
    if mode == "reentrega":
        conflitos = st.session_state.get("reentrega_conflitos", [])
        warning = " ".join(str(item) for item in conflitos) if conflitos else ""
        return (
            "Confirmacao operacional",
            warning,
            "**Esta operacao e uma REENTREGA?**",
        )
    if mode == "reimpressao":
        info = st.session_state.get("reimpressao_info")
        detail = ""
        if info is not None:
            detail = (
                f"**Primeira impressao:** {info.primeira_impressao_data}\n\n"
                f"**Usuario:** {info.primeira_impressao_usuario}\n\n"
                f"**Quantidade de impressoes:** {info.quantidade_impressoes}\n\n"
                "**Deseja imprimir novamente?**"
            )
        return (
            "Confirmacao operacional",
            "Esta Minuta ja foi registrada anteriormente.",
            detail,
        )
    if mode == "balcao_confirm":
        return (
            "Entrega no Balcao",
            "",
            "**Confirmar retirada desta Nota Fiscal no Balcao?**",
        )
    if mode == "balcao":
        return (
            "Entrega no Balcao",
            "",
            "",
        )
    if mode == "carregamento_historico":
        return ("Ocorrencias anteriores encontradas", "", "")
    if mode == "carregamento_decisao":
        return ("Decisao operacional", "", "")
    if mode == "fechamento":
        return (
            "Fechamento operacional",
            "",
            "O fechamento e a impressao dos documentos ocorrem ao clicar em Baixar PDF.",
        )
    return ("Painel operacional", "", "")


def _operational_panel_button_labels(mode: str) -> tuple[str, str, bool, bool]:
    labels: dict[str, tuple[str, str, bool, bool]] = {
        "reentrega": ("SIM", "NAO", False, False),
        "reimpressao": ("Reimprimir", "Cancelar", False, False),
        "balcao_confirm": ("Confirmar", "Cancelar", False, False),
        "balcao": ("Registrar Entrega no Balcao", "—", False, True),
        "carregamento_historico": ("—", "Cancelar operacao", True, False),
        "carregamento_decisao": ("—", "Cancelar operacao", True, False),
        "fechamento": ("Finalizar Carregamento", "—", False, True),
        "idle": ("—", "—", True, True),
    }
    return labels.get(mode, ("—", "—", True, True))


_OPERATIONAL_ATTENTION_MODES = frozenset(
    {
        "reentrega",
        "reimpressao",
        "balcao_confirm",
        "carregamento_historico",
        "carregamento_decisao",
    }
)

_HISTORICO_PANEL_MODES = frozenset({"carregamento_historico", "carregamento_decisao"})


def _requires_operational_attention(mode: str) -> bool:
    return mode in _OPERATIONAL_ATTENTION_MODES


def _scroll_to_operational_panel() -> None:
    components.html(
        """
        <script>
        const scrollToOperationalPanel = () => {
            const doc = window.parent.document;
            const anchor = doc.getElementById("brida-painel-operacional");
            if (!anchor) {
                return;
            }
            anchor.scrollIntoView({ behavior: "smooth", block: "center" });
            const panel = anchor.nextElementSibling;
            if (panel) {
                panel.style.transition = "box-shadow 0.6s ease";
                panel.style.boxShadow = "0 0 0 3px rgba(37, 99, 235, 0.35)";
                setTimeout(() => {
                    panel.style.boxShadow = "";
                }, 1800);
            }
        };
        window.parent.requestAnimationFrame(() => setTimeout(scrollToOperationalPanel, 150));
        </script>
        """,
        height=0,
    )


def _scroll_to_decisao_operacional() -> None:
    components.html(
        """
        <script>
        const scrollToDecisao = () => {
            const doc = window.parent.document;
            const anchor = doc.getElementById("brida-decisao-operacional");
            if (!anchor) {
                return;
            }
            anchor.scrollIntoView({ behavior: "smooth", block: "center" });
        };
        window.parent.requestAnimationFrame(() => setTimeout(scrollToDecisao, 150));
        </script>
        """,
        height=0,
    )


def _render_conflito_nfs_detail(diagnostico: DiagnosticoCarregamento) -> None:
    grupos: dict[int, dict[str, object]] = {}
    for nf_resumo in diagnostico.nfs:
        vinculo = nf_resumo.vinculo
        if vinculo is None:
            continue
        grupo = grupos.setdefault(
            int(vinculo.carregamento_id),
            {
                "numero_carregamento": str(vinculo.numero_carregamento or "--"),
                "nfs": [],
            },
        )
        nfs = grupo["nfs"]
        assert isinstance(nfs, list)
        nfs.append(str(nf_resumo.nf or "--"))

    if len(grupos) <= 1:
        return

    st.warning("Conflito operacional: as NFs abaixo pertencem a carregamentos diferentes.")
    for carregamento_id, grupo in sorted(grupos.items()):
        nfs = grupo["nfs"]
        numero = grupo["numero_carregamento"]
        nfs_text = ", ".join(nfs) if isinstance(nfs, list) else str(nfs)
        st.markdown(
            f"- **Carregamento {numero}** (ID {carregamento_id}): {html.escape(nfs_text)}"
        )


def _render_operacional_decisao_radios(diagnostico: DiagnosticoCarregamento) -> None:
    if not diagnostico.opcoes_decisao:
        return

    for mensagem in diagnostico.mensagens:
        if diagnostico.bloqueia_fechamento:
            st.warning(_humanizar_mensagem_operacional(mensagem))
        else:
            st.info(_humanizar_mensagem_operacional(mensagem))

    if diagnostico.cenario == CenarioOperacional.CONFLITO_MULTIPLO:
        _render_conflito_nfs_detail(diagnostico)

    st.markdown('<div id="brida-decisao-operacional"></div>', unsafe_allow_html=True)

    confirmacao_explicita = requer_confirmacao_explicita_historico(diagnostico)
    if confirmacao_explicita:
        st.markdown("**Foram encontradas ocorrencias operacionais.**")
        st.markdown(
            "Algumas Notas Fiscais desta planilha ja possuem historico de utilizacao."
        )
        st.markdown("**Deseja realmente continuar o processamento desta carga?**")
        st.caption("Esta decisao sera registrada na auditoria.")
        opcoes = [OPERACIONAL_CONTINUAR_HISTORICO_VALUE, DecisaoOperacional.CANCELAR.value]

        def _formatar_confirmacao(value: object) -> str:
            texto = str(value)
            if texto == OPERACIONAL_CONTINUAR_HISTORICO_VALUE:
                return "Sim, desejo continuar."
            return "Nao, cancelar operacao."

        format_func = _formatar_confirmacao
    else:
        st.markdown("**Como deseja continuar?**")
        opcoes = [item.value for item in diagnostico.opcoes_decisao]
        format_func = lambda value: DECISAO_OPERACIONAL_LABELS.get(
            DecisaoOperacional(str(value)),
            str(value),
        )

    radio_kwargs: dict[str, object] = {
        "label": "Decisao operacional",
        "options": opcoes,
        "format_func": format_func,
        "key": OPERACIONAL_DECISAO_WIDGET_KEY,
        "label_visibility": "collapsed",
        "on_change": on_operacional_decisao_widget_change,
    }
    if OPERACIONAL_DECISAO_WIDGET_KEY not in st.session_state:
        radio_kwargs["index"] = None
    st.radio(**radio_kwargs)


def _render_operational_panel(mode: str, saved_id: object, processed_df=None) -> None:
    st.session_state["_operational_panel_mode"] = mode
    title, warning_text, detail_markdown = _build_operational_panel_copy(mode)
    primary_label, secondary_label, primary_disabled, secondary_disabled = _operational_panel_button_labels(mode)
    historico_phase_modes = {"carregamento_historico"}
    decisao_phase_modes = {"carregamento_decisao"}

    st.markdown('<div id="brida-painel-operacional"></div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown(f'<div class="section-title">{html.escape(title)}</div>', unsafe_allow_html=True)
        if mode not in historico_phase_modes and mode not in decisao_phase_modes and warning_text:
            st.warning(warning_text)

        if detail_markdown and mode not in historico_phase_modes and mode not in decisao_phase_modes:
            st.markdown(detail_markdown)
        elif mode == "fechamento":
            st.markdown(
                '<p class="fechamento-operacional-caption">'
                "O fechamento e a impressao dos documentos ocorrem ao clicar em Baixar PDF."
                "</p>",
                unsafe_allow_html=True,
            )

        if mode in historico_phase_modes and processed_df is not None:
            render_historico_nfs_contexto(processed_df, scope="painel", show_acoes=True)

        if mode in decisao_phase_modes:
            if st.session_state.pop("auditoria_nf_foco_decisao_painel", False):
                _scroll_to_decisao_operacional()
            diagnostico = get_operacional_diagnostico()
            if diagnostico is not None:
                if processed_df is not None:
                    render_historico_nfs_contexto(
                        processed_df,
                        scope="painel_decisao",
                        show_acoes=False,
                    )
                _render_operacional_decisao_radios(diagnostico)

        st.text_input(
            "Entrega no Balcao",
            key="entrega_balcao_termo",
            placeholder="NF ou Chave da NF para Entrega no Balcao",
            label_visibility="collapsed",
            disabled=mode != "balcao" or bool(st.session_state.get("balcao_pending_confirm")),
        )

        panel_col_primary, panel_col_secondary = st.columns(2, gap="medium")
        with panel_col_primary:
            st.button(
                primary_label,
                key="processing_panel_primary",
                use_container_width=True,
                disabled=primary_disabled,
                on_click=on_processing_panel_primary_click,
            )
        with panel_col_secondary:
            st.button(
                secondary_label,
                key="processing_panel_secondary",
                use_container_width=True,
                disabled=secondary_disabled,
                on_click=on_processing_panel_secondary_click,
            )

        saved_caption = (
            f"Carregamento registrado no banco (ID {saved_id})."
            if saved_id
            else " "
        )
        st.caption(saved_caption)


def _render_processing_export_panel(
    *,
    module_config: MinutaModuleConfig,
    carregamento_checkbox_key: str,
    entrega_checkbox_key: str,
    xml_checkbox_key: str,
    validation_message: str,
    can_close: bool,
    download_payload: bytes,
    download_name: str,
    download_mime: str,
) -> None:
    st.markdown(f'<div class="section-title export-title">{html.escape(module_config.export_label)}</div>', unsafe_allow_html=True)
    checkbox_col_1, checkbox_col_2, checkbox_col_3, button_col = st.columns([1.0, 0.9, 0.9, 1.1], gap="small")
    with checkbox_col_1:
        st.checkbox("Minuta", key=carregamento_checkbox_key)
    with checkbox_col_2:
        st.checkbox("Romaneio", key=entrega_checkbox_key)
    with checkbox_col_3:
        st.checkbox("XMLs", key=xml_checkbox_key)
    with button_col:
        st.button(
            "Gerar PDF",
            use_container_width=True,
            key="preparar_baixar_pdf_button",
            disabled=bool(validation_message),
            on_click=on_baixar_pdf_click,
        )
        st.download_button(
            "Baixar PDF",
            data=download_payload,
            file_name=download_name,
            mime=download_mime,
            use_container_width=True,
            key="baixar_pdf_download_button",
            disabled=not download_payload,
        )
    if validation_message and not download_payload:
        st.info(validation_message)


def render_processing_screen(
    process_clicked: bool,
    xml_records: list,
    excel_file,
    module_config: MinutaModuleConfig,
) -> None:
    process_minuta_inputs(process_clicked, xml_records, excel_file)

    summary = st.session_state.summary
    processed_df = get_prepared_processed_dataframe()
    has_excel_loaded = excel_file is not None
    balcao_lookup_df = get_balcao_lookup_dataframe(xml_records)
    balcao_summary = build_balcao_summary()
    carregamento_checkbox_key = f"{module_config.screen_key}_pdf_carregamento"
    entrega_checkbox_key = f"{module_config.screen_key}_pdf_entrega"
    xml_checkbox_key = f"{module_config.screen_key}_pdf_xmls"

    if carregamento_checkbox_key not in st.session_state:
        st.session_state[carregamento_checkbox_key] = True
    if entrega_checkbox_key not in st.session_state:
        st.session_state[entrega_checkbox_key] = False
    if xml_checkbox_key not in st.session_state:
        st.session_state[xml_checkbox_key] = False

    carregamento_selected = bool(st.session_state.get(carregamento_checkbox_key, True))
    entrega_selected = bool(st.session_state.get(entrega_checkbox_key, False))
    xml_selected = bool(st.session_state.get(xml_checkbox_key, False))

    sync_processing_context_for_excel(has_excel_loaded)
    finalize_message, finalize_error, finalize_warning = _process_processing_screen_actions(
        summary=summary,
        processed_df=processed_df,
        balcao_lookup_df=balcao_lookup_df,
        balcao_summary=balcao_summary,
        module_config=module_config,
        has_excel_loaded=has_excel_loaded,
        carregamento_selected=carregamento_selected,
        entrega_selected=entrega_selected,
        xml_selected=xml_selected,
    )
    operational_panel_mode = resolve_operational_panel_mode(
        has_excel_loaded=has_excel_loaded,
        processed_df_empty=processed_df.empty,
    )
    saved_id = st.session_state.get("carregamento_saved_id")

    filtered_df = processed_df
    display_df = get_display_table_dataframe(filtered_df)
    styled_display_df = build_status_styler(display_df)

    download_payload = st.session_state.get("pdf_download_payload") or b""
    download_name = str(st.session_state.get("pdf_download_name", "") or "documento.pdf")
    download_mime = str(st.session_state.get("pdf_download_mime", "application/pdf") or "application/pdf")
    can_close = has_excel_loaded and not processed_df.empty
    validation_message = ""
    diagnostico = get_operacional_diagnostico()
    decisao = get_operacional_decisao()
    if diagnostico and diagnostico.requer_decisao and decisao is None:
        if not is_operacional_analise_confirmada():
            validation_message = (
                "Revise o historico das NFs e clique em Continuar processamento para escolher a acao operacional."
            )
        else:
            validation_message = "Selecione como deseja continuar no painel operacional."
    elif decisao == DecisaoOperacional.CANCELAR:
        validation_message = "Selecione uma acao operacional para continuar ou use Cancelar operacao."
    elif (
        diagnostico
        and diagnostico.bloqueia_fechamento
        and decisao is None
        and diagnostico.cenario == CenarioOperacional.NF_CANCELADA
    ):
        validation_message = (
            "; ".join(_humanizar_mensagem_operacional(item) for item in diagnostico.mensagens)
            or "Existem NFs canceladas que impedem o processamento."
        )
    elif not carregamento_selected and not entrega_selected and not xml_selected:
        validation_message = "Selecione ao menos um tipo de documento para gerar o download"
    elif not can_close and not st.session_state.get("balcao_pending_confirm"):
        validation_message = "Processe um Excel valido para habilitar o fechamento via Baixar PDF."

    use_full_width_historico = (
        has_excel_loaded
        and not processed_df.empty
        and operational_panel_mode in _HISTORICO_PANEL_MODES
    )
    export_kwargs = {
        "module_config": module_config,
        "carregamento_checkbox_key": carregamento_checkbox_key,
        "entrega_checkbox_key": entrega_checkbox_key,
        "xml_checkbox_key": xml_checkbox_key,
        "validation_message": validation_message,
        "can_close": can_close,
        "download_payload": download_payload,
        "download_name": download_name,
        "download_mime": download_mime,
    }

    if use_full_width_historico:
        _render_operational_panel(
            operational_panel_mode,
            saved_id,
            processed_df,
        )
        if (
            process_clicked
            and _requires_operational_attention(operational_panel_mode)
        ):
            _scroll_to_operational_panel()
        if finalize_message:
            st.success(finalize_message)
        if finalize_warning:
            st.warning(finalize_warning)
        if finalize_error:
            st.error(finalize_error)

        export_spacer, export_col = st.columns([2.4, 2.6], gap="medium")
        with export_spacer:
            pass
        with export_col:
            _render_processing_export_panel(**export_kwargs)
    else:
        action_col_search, action_col_download = st.columns([1.7, 2.1], gap="medium")

        with action_col_search:
            _render_operational_panel(
                operational_panel_mode,
                saved_id,
                processed_df if has_excel_loaded and not processed_df.empty else None,
            )
            if (
                process_clicked
                and has_excel_loaded
                and not processed_df.empty
                and _requires_operational_attention(operational_panel_mode)
            ):
                _scroll_to_operational_panel()
            if finalize_message:
                st.success(finalize_message)
            if finalize_warning:
                st.warning(finalize_warning)
            if finalize_error:
                st.error(finalize_error)

        with action_col_download:
            _render_processing_export_panel(**export_kwargs)

    render_section_heading("Dados Gerais", "dados_gerais")
    dados_col_1, dados_col_2, dados_col_3 = st.columns(3, gap="medium")
    with dados_col_1:
        render_info_card("Filial", summary["filial"], "filial")
    with dados_col_2:
        render_info_card(module_config.subject_label, summary["numero_carga"], "carregamento")
    with dados_col_3:
        render_info_card("Data Saida", summary["data_saida"], "data_saida")

    dados_col_4, dados_col_5 = st.columns(2, gap="medium")
    with dados_col_4:
        render_info_card("Motorista", summary["motorista"], "motorista")
    with dados_col_5:
        render_info_card("Placa", summary["placa"], "placa")

    render_section_heading(module_config.summary_label, "resumo_carga")
    render_metric_cards_row(
        [
            {"title": "Notas", "value": summary["nf_count"], "icon_key": "nf"},
            {"title": "Peso Total", "value": f"{summary['peso_total'] / 1000:.3f} t", "icon_key": "peso"},
            {"title": "Itens", "value": summary["item_count"], "icon_key": "itens"},
            {"title": "Erros", "value": summary["error_count"], "icon_key": "erros"},
        ]
    )

    if DISPLAY_PROCESSING_WARNINGS and st.session_state.issues:
        for issue in st.session_state.issues:
            st.warning(issue)

    if not st.session_state.nf_debug.empty:
        with st.expander("Debug de correspondencia NF x XML", expanded=False):
            st.dataframe(st.session_state.nf_debug, use_container_width=True, hide_index=True)

    if has_excel_loaded and not processed_df.empty:
        render_auditoria_nf_expander(processed_df=processed_df)

    if not has_excel_loaded and st.session_state.get("balcao_pending_confirm"):
        preview_termo = str(
            st.session_state.get("balcao_confirm_termo", "")
            or st.session_state.get("entrega_balcao_termo", "")
            or ""
        )
        render_balcao_nf_preview(balcao_lookup_df, preview_termo)

    st.markdown(f"### {module_config.panel_title}")
    st.caption(module_config.panel_caption)
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
    st.session_state["separacao_records"] = separacao_records
    current_records = separacao_records
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
        latest_closed_pdf_bytes = get_latest_closed_lote_pdf_bytes(latest_closed_lote, latest_closed_records, latest_closed_lote_id)

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
    with ui_section_box():
        render_section_heading("Visão Operacional", "separacao")
        render_metric_cards_row(
            [
                {"title": "Notas", "value": summary["nf_total"], "icon_key": "nf"},
                {"title": "Pendentes", "value": summary["pendentes"], "icon_key": "processar"},
                {"title": "Separadas", "value": summary["separadas"], "icon_key": "status_operacional"},
                {"title": "Lotes Fech.", "value": summary["lotes_fechados"], "icon_key": "erros"},
            ]
        )

    with ui_section_box():
        render_section_heading("Entrada", "barcode")
        with st.container(border=True):
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
        render_scan_input_focus()

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
        with ui_section_box():
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

    cleanup_feedback = st.session_state.get("data_cleanup_feedback")
    current_xml_records, _ = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
    current_lotes_registry, _ = carregar_lotes_json(str(LOTES_JSON_PATH))

    with ui_section_box():
        st.markdown(
            """
    <div class="page-hero" style="margin-top: 1.25rem;">
        <h2>Gestão de Dados do Sistema</h2>
        <p>Limpeza e controle de XMLs, separações e lotes</p>
    </div>
    """,
            unsafe_allow_html=True,
        )

        render_metric_cards_row(
            [
                {"title": "XMLs", "value": len(current_xml_records), "icon_key": "xml"},
                {"title": "Separações", "value": len(current_records), "icon_key": "separacao"},
                {"title": "Lotes", "value": len(current_lotes_registry), "icon_key": "lotes"},
                {"title": "Total Bases", "value": len(current_xml_records) + len(current_records) + len(current_lotes_registry), "icon_key": "dados_gerais"},
            ]
        )

        size_col_1, size_col_2, size_col_3 = st.columns(3, gap="medium")
        with size_col_1:
            render_info_card("Base XMLs", f"{len(current_xml_records)} registro(s)", "xml", "PostgreSQL: nota_fiscal")
        with size_col_2:
            render_info_card("Base Separação", f"{len(current_records)} registro(s)", "separacao", "PostgreSQL: configuracao")
        with size_col_3:
            render_info_card("Base Lotes", f"{len(current_lotes_registry)} registro(s)", "lotes", "PostgreSQL: configuracao")

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


def render_lotes_management_screen(separacao_records: list[dict[str, object]]) -> None:
    classificacao_records, _ = carregar_classificacao_produtos_json(
        str(CLASSIFICACAO_PRODUTOS_JSON_PATH),
        get_classificacao_version_token(),
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

    if "xml_upload_batch" not in st.session_state:
        st.session_state.xml_upload_batch = {}

    if "xml_add_uploader_open" not in st.session_state:
        st.session_state.xml_add_uploader_open = False

    if "xml_upload_signature" not in st.session_state:
        st.session_state.xml_upload_signature = ""

    if "xml_upload_message" not in st.session_state:
        st.session_state.xml_upload_message = ""

    if "xml_upload_summary" not in st.session_state:
        st.session_state.xml_upload_summary = {}

    if "xml_upload_error" not in st.session_state:
        st.session_state.xml_upload_error = ""

    if "xml_upload_issues" not in st.session_state:
        st.session_state.xml_upload_issues = []

    if "xml_import_report" not in st.session_state:
        st.session_state.xml_import_report = None

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

    if "runtime_xml_storage_error" not in st.session_state:
        st.session_state["runtime_xml_storage_error"] = ""

    if "_processed_data_version" not in st.session_state:
        st.session_state["_processed_data_version"] = 0


def render_global_app_styles() -> None:
    # O Streamlit reconstrói o DOM a cada rerun; o CSS via st.markdown deve ser
    # reinjetado em toda execução. Cache apenas a string, nunca pular a injeção.
    st.markdown(
        """
    <style>
    :root {
        --brida-navy: #1F3A5F;
        --brida-navy-hover: #25486E;
        --brida-blue-soft: #E8F1FF;
        --brida-border: #1F3A5F;
        --brida-button-border: rgba(31, 58, 95, 0.18);
        --brida-shadow: 0 4px 12px rgba(31, 58, 95, 0.06);
        --brida-radius: 12px;
        --brida-gray-bg: #F5F7FA;
        --brida-text-muted: #617285;
        --brida-success: #166534;
        --brida-warning: #B45309;
        --brida-error: #B42318;
    }
    /* Borda institucional global — heranca automatica em todas as telas */
    section.main div[data-baseweb="input"] > div,
    section.main div[data-baseweb="textarea"],
    section.main div[data-baseweb="select"] > div,
    section.main div[data-baseweb="datepicker"] input,
    [data-testid="stSidebar"] div[data-baseweb="input"] > div,
    [data-testid="stSidebar"] div[data-baseweb="textarea"],
    [data-testid="stSidebar"] div[data-baseweb="select"] > div {
        border-color: var(--brida-border) !important;
    }
    section.main [data-testid="stTextInput"] > div > div,
    section.main [data-testid="stTextArea"] > div > div,
    section.main [data-testid="stNumberInput"] > div > div,
    section.main [data-testid="stDateInput"] > div > div,
    section.main [data-testid="stTimeInput"] > div > div,
    section.main [data-testid="stSelectbox"] > div > div,
    section.main [data-testid="stMultiSelect"] > div > div,
    [data-testid="stSidebar"] [data-testid="stTextInput"] > div > div,
    [data-testid="stSidebar"] [data-testid="stTextArea"] > div > div,
    [data-testid="stSidebar"] [data-testid="stNumberInput"] > div > div,
    [data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div,
    [data-testid="stSidebar"] [data-testid="stMultiSelect"] > div > div {
        border-color: var(--brida-border) !important;
    }
    section.main [data-testid="stFileUploader"] section[data-testid="stFileUploadDropzone"],
    [data-testid="stSidebar"] [data-testid="stFileUploader"] section[data-testid="stFileUploadDropzone"] {
        border: 1px solid var(--brida-border) !important;
    }
    section.main [data-testid="stExpander"] details,
    [data-testid="stSidebar"] [data-testid="stExpander"] details {
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        overflow: hidden;
    }
    section.main [data-testid="stTabs"],
    [data-testid="stSidebar"] [data-testid="stTabs"] {
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        padding: 0.25rem 0.5rem 0.65rem;
    }
    section.main [data-testid="stTabs"] [data-baseweb="tab-border"],
    section.main [data-testid="stTabs"] button[data-baseweb="tab"] {
        border-color: var(--brida-border) !important;
    }
    section.main [data-testid="stForm"],
    [data-testid="stSidebar"] [data-testid="stForm"] {
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
    }
    section.main [data-testid="stDataFrame"],
    section.main [data-testid="stDataFrameResizable"],
    [data-testid="stSidebar"] [data-testid="stDataFrame"],
    [data-testid="stSidebar"] [data-testid="stDataFrameResizable"] {
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        overflow: hidden;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"],
    div[data-testid="stRadio"] > div {
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        padding: 0.45rem 0.7rem;
    }
    div[data-testid="stAlert"],
    div[data-testid="stToast"] {
        border: 1px solid var(--brida-border) !important;
    }
    [data-testid="stDialog"],
    [data-testid="stModal"],
    div[role="dialog"][data-baseweb="modal"] {
        border: 1px solid var(--brida-border) !important;
    }
    .nf-historico-resumo-linha {
        border-bottom: 1px solid var(--brida-border);
    }
    .logged-user-sidebar-name {
        color: var(--brida-navy);
        font-size: 0.92rem;
        font-weight: 600;
        margin: 0;
        line-height: 1.35;
    }
    .stApp {
        background: var(--brida-gray-bg);
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
        flex-shrink: 0;
        fill: none;
        stroke: currentColor;
        stroke-width: 1.7;
        stroke-linecap: round;
        stroke-linejoin: round;
        vector-effect: non-scaling-stroke;
    }
    .ui-icon svg path {
        fill: none;
        stroke: inherit;
        vector-effect: non-scaling-stroke;
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
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        box-shadow: var(--brida-shadow);
        padding: 14px 16px;
        margin: 0 0 12px;
    }
    .ui-section-box.is-soft {
        background: #FAFBFC;
    }
    .ui-section-box.is-sidebar {
        padding: 11px 12px;
        margin-bottom: 10px;
        border-radius: 13px;
        box-shadow: 0 2px 8px rgba(31, 58, 95, 0.05);
    }
    section.main [data-testid="stVerticalBlockBorderWrapper"],
    [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF;
        border: 1px solid var(--brida-border) !important;
        border-radius: var(--brida-radius);
        box-shadow: var(--brida-shadow);
        padding: 14px 16px;
        margin: 0 0 12px;
    }
    [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FAFBFC;
        padding: 11px 12px;
        margin-bottom: 10px;
        border-radius: 13px;
        box-shadow: 0 2px 8px rgba(31, 58, 95, 0.05);
    }
    section.main [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FAFBFC;
        border-color: var(--brida-border) !important;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.6);
        padding: 12px 14px 4px;
        margin-bottom: 0;
    }
    section.main [data-testid="stVerticalBlockBorderWrapper"] h3 {
        margin: 0 0 6px;
        color: #1F3A5F;
        font-size: 0.98rem;
    }
    .ui-section-box .section-title-block {
        margin-bottom: 12px;
    }
    .ui-section-box .stCaption {
        margin-bottom: 0;
    }
    .erp-card {
        background: #FFFFFF;
        border: 1px solid var(--brida-border);
        border-radius: var(--brida-radius);
        box-shadow: var(--brida-shadow);
        padding: 10px 12px;
        height: auto;
        min-height: 0;
        margin-bottom: 10px;
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
        gap: 0.15rem;
    }
    .erp-kpi-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 12px;
        width: 100%;
        margin: 0 0 12px;
        align-items: stretch;
    }
    .erp-card-kpi-fixed {
        height: 90px;
        margin-bottom: 0;
        padding: 9px 12px;
        border: 1px solid var(--brida-border);
        border-radius: 13px;
        background: #FFFFFF;
        box-shadow: var(--brida-shadow);
        display: flex;
        flex-direction: column;
        align-items: stretch;
        justify-content: space-between;
        text-align: center;
        box-sizing: border-box;
        overflow: hidden;
    }
    .erp-kpi-top {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
        gap: 0.22rem;
        min-height: 32px;
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
        letter-spacing: 0.06em;
        line-height: 1.2;
        text-transform: uppercase;
    }
    .erp-card-value {
        color: var(--brida-navy);
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
        margin-bottom: 0;
    }
    .erp-card-kpi-fixed .erp-kpi-icon .ui-icon {
        width: 14px;
        min-width: 14px;
        height: 14px;
    }
    .erp-card-kpi-fixed .erp-kpi-icon .ui-icon svg {
        width: 14px;
        height: 14px;
    }
    .erp-kpi-value {
        color: var(--brida-navy);
        font-size: 1.8rem;
        line-height: 1.05;
        font-weight: 700;
        letter-spacing: -0.03em;
        display: flex;
        align-items: center;
        justify-content: center;
        flex: 1 1 auto;
        min-height: 42px;
        overflow: hidden;
        text-overflow: ellipsis;
        font-family: "Segoe UI", Calibri, Arial, sans-serif;
    }
    .erp-kpi-label {
        color: #475467;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        line-height: 1.15;
        text-transform: none;
        font-family: "Segoe UI", Calibri, Arial, sans-serif;
    }
    .erp-kpi-subtitle {
        color: #98A2B3;
        font-size: 0.67rem;
        line-height: 1.1;
        min-height: 11px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        font-family: "Segoe UI", Calibri, Arial, sans-serif;
    }
    [data-testid="stDataFrame"] table td {
        vertical-align: middle;
    }
    [data-testid="stDataFrame"] table thead th,
    [data-testid="stDataFrame"] table thead td {
        background: var(--brida-blue-soft) !important;
        color: var(--brida-navy) !important;
        font-weight: 600;
        border-bottom: 1px solid var(--brida-border) !important;
    }
    [data-testid="stDataFrame"] table tbody tr:nth-child(even) td {
        background: #F8FAFC;
    }
    [data-testid="stDataFrame"] table tbody tr:hover td {
        background: #EEF4FF;
    }
    [data-testid="stDataFrame"] table {
        width: 100%;
    }
    .section-title {
        margin: 4px 0 8px;
        color: var(--brida-navy);
        font-size: 0.95rem;
        font-weight: 700;
    }
    .fechamento-balcao-title {
        margin-bottom: 0.55rem;
        font-size: 0.95rem;
        font-weight: 700;
        color: var(--brida-navy);
    }
    .balcao-operacional-caption,
    .fechamento-operacional-caption {
        margin: 0 0 0.85rem;
        color: #617285;
        font-size: 0.82rem;
        line-height: 1.35;
    }
    .export-title {
        text-align: right;
        margin-bottom: 0.75rem;
    }
    .page-hero {
        background: linear-gradient(135deg, #ffffff 0%, #eef4fb 100%);
        border: 1px solid var(--brida-border);
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
        border: 1px solid var(--brida-border);
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
        padding: 14px 16px;
        border-radius: var(--brida-radius);
        border: 1px solid var(--brida-border);
        box-shadow: var(--brida-shadow);
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
        border-radius: 14px;
        padding: 22px 24px;
        color: #FFFFFF;
        box-shadow: 0 10px 22px rgba(22, 50, 79, 0.14);
        margin-bottom: 1rem;
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
        border: 1px solid var(--brida-border);
        border-radius: 14px;
        padding: 10px 16px 9px;
        box-shadow: var(--brida-shadow);
        height: 170px;
        margin-bottom: 5px;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        box-sizing: border-box;
        overflow: hidden;
    }
    .module-card h3 {
        margin: 0.35rem 0 0.15rem;
        color: var(--brida-navy);
        font-size: 1.05rem;
        font-weight: 800;
        line-height: 1.25;
        min-height: 38px;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
    }
    .module-card p {
        margin: 0;
        color: var(--brida-text-muted);
        font-size: 0.9rem;
        line-height: 1.35;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
        flex: 1 1 auto;
    }
    .module-card-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 34px;
        height: 34px;
        border-radius: 10px;
        background: var(--brida-blue-soft);
        color: var(--brida-navy);
        flex-shrink: 0;
    }
    .module-card-icon .ui-icon {
        width: 19px;
        min-width: 19px;
        height: 19px;
        color: inherit;
    }
    div[data-testid="stMarkdown"]:has(.module-card) {
        margin-bottom: 0;
    }
    div[data-testid="stMarkdown"]:has(.module-card) + div[data-testid="stVerticalBlock"] {
        margin-top: 0;
        padding-top: 0;
    }
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 8px;
        min-height: 38px;
        border: 1px solid var(--brida-button-border);
        background: #FFFFFF;
        color: var(--brida-navy);
        font-weight: 600;
        box-shadow: none;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: rgba(46, 111, 149, 0.30);
        color: #1F3A5F;
    }
    .stDownloadButton > button {
        background: var(--brida-navy);
        color: #FFFFFF;
        border: 0;
        min-height: 40px;
        padding-left: 1rem;
        padding-right: 1rem;
        font-weight: 500;
    }
    .stDownloadButton > button:hover {
        color: #FFFFFF;
        background: var(--brida-navy-hover);
    }
    .stButton > button[kind="primary"],
    .stFormSubmitButton > button[kind="primary"] {
        background: var(--brida-navy);
        color: #FFFFFF;
        border: 0;
    }
    .stButton > button[kind="primary"]:hover,
    .stFormSubmitButton > button[kind="primary"]:hover {
        background: var(--brida-navy-hover);
        color: #FFFFFF;
    }
    .stTextInput > div > div input,
    .stTextArea > div > div textarea,
    .stNumberInput > div > div input {
        border-radius: 8px;
        border-color: var(--brida-border) !important;
    }
    [data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid var(--brida-border);
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
        padding: 8px;
        border-radius: 8px;
        background: #F8FAFC;
        border: 1px solid var(--brida-border);
        box-shadow: none;
    }
    .sidebar-heading {
        margin: 0 0 6px;
        color: var(--brida-navy);
        font-size: 0.92rem;
        font-weight: 700;
    }
    .sidebar-field-label {
        margin: 0 0 8px;
        color: #334155;
        font-size: 0.88rem;
        font-weight: 700;
    }
    .sidebar-heading .ui-icon,
    .sidebar-field-label .ui-icon,
    .erp-card-header .ui-icon {
        color: var(--brida-navy);
    }
    [data-testid="stSidebar"] .stButton > button {
        background: var(--brida-navy);
        color: #FFFFFF;
        border: 0;
        min-height: 38px;
        font-weight: 600;
    }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"] {
        background: #FFFFFF;
        color: var(--brida-navy);
        border: 1px solid var(--brida-button-border);
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        color: #FFFFFF;
        background: var(--brida-navy-hover);
    }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
        color: var(--brida-navy);
        background: var(--brida-blue-soft);
    }
    section.main .block-container {
        padding-top: 0.85rem;
        padding-bottom: 1rem;
        max-width: 100%;
    }
    section.main [data-testid="stCaptionContainer"] {
        margin-bottom: 0.25rem;
    }
    div[data-testid="stAlert"][data-baseweb="notification"] {
        border-radius: 8px;
    }
    div[data-testid="stNotificationContentSuccess"] {
        background: #ECFDF3;
        border-color: rgba(22, 101, 52, 0.2);
    }
    div[data-testid="stNotificationContentWarning"] {
        background: #FFFBEB;
        border-color: rgba(180, 83, 9, 0.2);
    }
    div[data-testid="stNotificationContentError"] {
        background: #FEF3F2;
        border-color: rgba(180, 35, 24, 0.2);
    }
    div[data-testid="stAlert"] {
        border-radius: 8px;
        border: 1px solid var(--brida-border) !important;
        box-shadow: none;
    }
    div[data-testid="stAlert"] p {
        font-size: 0.92rem;
        line-height: 1.55;
    }
    .brida-users-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.9rem;
    }
    .brida-users-table thead th {
        background: var(--brida-blue-soft);
        color: var(--brida-navy);
        font-weight: 600;
        text-align: left;
        padding: 10px 12px;
        border-bottom: 1px solid var(--brida-border);
    }
    .brida-users-table tbody td {
        padding: 10px 12px;
        border-bottom: 1px solid rgba(31, 58, 95, 0.06);
        vertical-align: middle;
    }
    .brida-users-table tbody tr:nth-child(even) td {
        background: #F8FAFC;
    }
    .brida-users-table tbody tr:hover td {
        background: #EEF4FF;
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
    runtime_signature = get_reference_data_signature()
    should_refresh = (
        force_refresh
        or st.session_state.get("runtime_data_signature") != runtime_signature
        or not st.session_state.get("runtime_xml_records")
    )

    if should_refresh:
        with st.spinner("Atualizando bases de referência..."):
            with measure("sql.carregar_xmls"):
                xml_records, xml_error = carregar_xmls_processados_json(str(XMLS_PROCESSADOS_JSON_PATH))
            with measure("sql.carregar_classificacao"):
                classificacao_records, _ = carregar_classificacao_produtos_json(
                    str(CLASSIFICACAO_PRODUTOS_JSON_PATH),
                    get_classificacao_version_token(),
                )
        st.session_state["runtime_xml_records"] = xml_records
        st.session_state["runtime_classificacao_records"] = classificacao_records
        st.session_state["runtime_data_signature"] = runtime_signature
        st.session_state["runtime_xml_storage_error"] = xml_error
        invalidate_balcao_lookup_cache()

    return (
        st.session_state.get("runtime_xml_records", []),
        st.session_state.get("runtime_classificacao_records", []),
    )


def load_runtime_operational_data(force_refresh: bool = False) -> tuple[list[dict[str, object]], list[str], str, dict[str, int]]:
    xml_records, classificacao_records = load_runtime_reference_data(force_refresh=force_refresh)
    runtime_signature = (
        st.session_state.get("runtime_data_signature"),
        get_operational_data_signature(),
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
        tela_minuta(process_clicked, xml_records, excel_file, get_minuta_module_config(current_screen))
        return

    separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary = load_runtime_operational_data(
        force_refresh=force_refresh
    )
    if current_screen == SCREEN_SEPARACAO:
        tela_separacao(separacao_records, separacao_sync_issues, separacao_storage_error, separacao_import_summary)
        return

    if current_screen == SCREEN_USUARIOS:
        if not require_admin():
            st.session_state["auth_access_error"] = "Acesso restrito ao perfil Administrador."
            navegar(SCREEN_MENU)
            return
        tela_usuarios()
        return

    if current_screen == SCREEN_CONSULTA_CARREGAMENTOS:
        tela_consulta_carregamentos()
        return

    if current_screen == SCREEN_GESTAO_DADOS:
        tela_gestao_dados()
        return

    tela_lotes(separacao_records)



def build_menu_cards() -> list[dict[str, object]]:
    menu_cards: list[dict[str, object]] = [
        {
            "target_screen": MINUTA_CARREGAMENTO_CONFIG.screen_key,
            "nav_key": "minuta",
            "title": MINUTA_CARREGAMENTO_CONFIG.menu_title,
            "description": MINUTA_CARREGAMENTO_CONFIG.menu_description,
            "icon_key": MINUTA_CARREGAMENTO_CONFIG.menu_icon_key,
            "button_label": "📦 Minuta",
        },
        {
            "target_screen": SCREEN_CONSULTA_CARREGAMENTOS,
            "nav_key": "consulta_carregamentos",
            "title": "Consulta de NFs",
            "description": "Localize NFs e documentos do historico operacional.",
            "icon_key": "consulta_carregamentos",
            "button_label": "🔎 Consulta",
        },
    ]
    if is_admin():
        menu_cards.extend(
            [
                {
                    "target_screen": SCREEN_USUARIOS,
                    "nav_key": "list",
                    "title": "Usuarios",
                    "description": "Gerenciar operadores e acessos.",
                    "icon_key": "usuarios",
                    "button_label": "👤 Usuarios",
                },
                {
                    "target_screen": SCREEN_USUARIOS,
                    "nav_key": "create",
                    "title": "Cadastro",
                    "description": "Novo operador ou administrador.",
                    "icon_key": "cadastro_usuarios",
                    "button_label": "⚙ Cadastro",
                    "open_action": "create",
                },
            ]
        )
    return menu_cards


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
        if st.button("🏠 Painel", use_container_width=True, key=f"home_button_{title}"):
            navegar(SCREEN_MENU)
    with col_menu_toggle:
        st.button("Painel", use_container_width=True, on_click=toggle_menu, key=f"toggle_sidebar_{title}")
    with col_action:
        st.button("🚪 Sair", use_container_width=True, on_click=logout, key=f"logout_{title}", type="secondary")


def tela_menu() -> None:
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
            <p>Selecione o modulo para continuar o fluxo de carregamento.</p>
        </div>
        """,
            unsafe_allow_html=True,
        )
    with top_col_3:
        st.button("🚪 Sair", use_container_width=True, on_click=logout, key="logout_menu", type="secondary")

    auth_access_error = st.session_state.pop("auth_access_error", "")
    if auth_access_error:
        st.error(auth_access_error)

    menu_cards = build_menu_cards()
    row_columns = st.columns(len(menu_cards), gap="medium")
    for column, card in zip(row_columns, menu_cards):
        with column:
            icon_markup = render_label_icon(resolve_menu_icon(str(card["icon_key"])))
            st.markdown(
                f"""
            <div class="module-card">
                <div class="module-card-icon">{icon_markup}</div>
                <h3>{html.escape(str(card['title']))}</h3>
                <p>{html.escape(str(card['description']))}</p>
            </div>
            """,
                unsafe_allow_html=True,
            )
            nav_key = str(card.get("nav_key", card["target_screen"]))
            if st.button(
                str(card.get("button_label", "Acessar")),
                use_container_width=True,
                key=f"menu_nav_{nav_key}",
            ):
                if card.get("open_action") == "create":
                    st.session_state["usuarios_action"] = "create"
                    st.session_state.pop("usuarios_selected_id", None)
                else:
                    st.session_state.pop("usuarios_action", None)
                    st.session_state.pop("usuarios_selected_id", None)
                navegar(str(card["target_screen"]))


def tela_minuta(
    process_clicked: bool,
    xml_records: list[dict[str, object]],
    excel_file,
    module_config: MinutaModuleConfig,
) -> None:
    render_screen_header(module_config.header_title, module_config.header_subtitle)
    render_processing_screen(process_clicked, xml_records, excel_file, module_config)


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


def tela_usuarios() -> None:
    render_usuarios_page(
        render_header_callback=render_screen_header,
        navigate_callback=navegar,
        menu_screen=SCREEN_MENU,
    )


def tela_consulta_carregamentos() -> None:
    render_consulta_carregamentos_page(render_header_callback=render_screen_header)


def tela_gestao_dados() -> None:
    render_gestao_dados_page(render_header_callback=render_screen_header)


def tela_gestao_retencao() -> None:
    tela_gestao_dados()


def render_main_screen() -> None:
    with measure("ui.render_main_screen"):
        initialize_app_state()
        render_global_app_styles()

        login_success = st.session_state.get("login_success", "")
        if login_success:
            st.success(login_success)
            st.session_state["login_success"] = ""
        current_screen = normalize_screen_name(st.session_state.get("tela", SCREEN_MENU))
        if current_screen == SCREEN_MENU:
            with st.sidebar:
                render_logged_user_badge()
                render_sidebar_dados_navigation()
            tela_menu()
            render_gestao_dados_login_prompt()
            _render_performance_report_panel()
            return

        render_gestao_dados_login_prompt()

        if "menu_aberto" not in st.session_state:
            st.session_state["menu_aberto"] = True

        apply_sidebar_visibility(st.session_state["menu_aberto"])
        excel_file = None
        process_clicked = False
        if st.session_state["menu_aberto"]:
            _, excel_file, process_clicked = render_sidebar()

        with measure("ui.render_active_screen"):
            render_active_screen(current_screen, process_clicked, excel_file)
        _render_performance_report_panel()


def _render_performance_report_panel() -> None:
    if not st.session_state.get("_perf_panel_enabled", True):
        return

    report = build_performance_report()
    if "Nenhuma medição" in report:
        return

    with st.expander("Diagnóstico de performance (sessão atual)", expanded=False):
        st.markdown(report)


def main() -> None:
    _ui_section_stack.clear()
    st.set_page_config(layout="wide")
    with measure("startup.environment_checks"):
        run_startup_environment_checks()
    with measure("startup.configure_storage"):
        configure_application_storage()
    with measure("startup.retencao_automatica"):
        run_startup_retention_once()
    initialize_login_state()
    initialize_navigation_state()

    if "login_error" not in st.session_state:
        st.session_state["login_error"] = ""

    if "login_success" not in st.session_state:
        st.session_state["login_success"] = ""

    if not is_logged_in():
        st.session_state["tela"] = SCREEN_LOGIN
        with measure("ui.render_login"):
            render_login_screen()
        return

    maybe_prompt_capacidade_critica()
    maybe_prompt_gestao_dados()
    render_main_screen()


if __name__ == "__main__":
    main()
