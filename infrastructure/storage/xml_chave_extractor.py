from __future__ import annotations

import re


_CHNFE_PATTERN = re.compile(r"<(?:[\w:]*?)chNFe[^>]*>\s*(\d{44})\s*</(?:[\w:]*?)chNFe>", re.IGNORECASE)
_CHNFE_FALLBACK = re.compile(r"chNFe[^0-9]*(\d{44})")


def extract_chave_nfe_from_xml_bytes(data: bytes) -> str:
    if not data:
        return ""
    text = data.decode("utf-8", errors="ignore")
    match = _CHNFE_PATTERN.search(text)
    if match:
        return match.group(1)
    match = _CHNFE_FALLBACK.search(text)
    if match:
        return match.group(1)
    return ""


def extract_numero_nf_from_chave(chave_nfe: str) -> str:
    if len(chave_nfe) != 44 or not chave_nfe.isdigit():
        return ""
    raw = chave_nfe[25:34]
    return raw.lstrip("0") or raw
