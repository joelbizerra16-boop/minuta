"""Extração e normalização da rota individual a partir de <infCpl> da NF-e."""

from __future__ import annotations

import re

UNDEFINED_ROUTE_LABEL = "NÃO DEFINIDA"

# Aceita: Rota ABCD | Rota: ABCD | ROTA ABCD | ROTA: ABCD | Rota - ABCD | Rota=ABCD
ROUTE_FROM_INF_CPL_PATTERN = re.compile(
    r"(?im)(?:^|[\s\-])rota\s*(?:[:=\-]\s*|\s+)(.+?)(?="
    r"(?:\n|\r|$|"
    r"(?<!\w)(?:Trib|Valor\s+CBS|Valor\s+IBS|Vendedor|Pedido|Cliente)\b))"
)

ROUTE_TRAILING_JUNK_PATTERN = re.compile(r"[\s\-:;=]+$")


def normalize_route_label(value: object) -> str:
    if value is None:
        return UNDEFINED_ROUTE_LABEL
    try:
        # pandas/numpy NaN é truthy; não pode virar literal "nan"
        if value != value:  # noqa: PLR0124
            return UNDEFINED_ROUTE_LABEL
    except Exception:
        pass
    route = re.sub(r"\s+", " ", str(value).strip())
    if not route or route.casefold() == "nan":
        return UNDEFINED_ROUTE_LABEL
    return route


def is_undefined_route(value: object) -> bool:
    """True quando a rota está vazia, ausente ou é o rótulo padrão NÃO DEFINIDA."""
    route = normalize_route_label(value)
    return route.casefold() == UNDEFINED_ROUTE_LABEL.casefold()


def has_concrete_route(value: object) -> bool:
    return not is_undefined_route(value)


def should_enrich_xml_route(current_record: dict[str, object], new_record: dict[str, object]) -> bool:
    """Permite atualizar rota válida sobre registro existente sem rota concreta."""
    return is_undefined_route(current_record.get("ROTA")) and has_concrete_route(new_record.get("ROTA"))


def extract_route_from_inf_cpl(inf_cpl: object) -> str:
    """Extrai a rota de <infCpl>, sem incluir marcadores tributários ou campos seguintes."""
    text = str(inf_cpl or "").replace("\\n", "\n")
    if not text.strip():
        return ""

    match = ROUTE_FROM_INF_CPL_PATTERN.search(text)
    if not match:
        return ""

    route_value = re.sub(r"\s+", " ", match.group(1)).strip()
    route_value = ROUTE_TRAILING_JUNK_PATTERN.sub("", route_value).strip()
    return route_value
