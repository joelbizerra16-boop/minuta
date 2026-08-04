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
    route = re.sub(r"\s+", " ", str(value or "").strip())
    return route or UNDEFINED_ROUTE_LABEL


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
