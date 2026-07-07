from __future__ import annotations

import os

"""Politica centralizada de retencao e limites de armazenamento (unico ponto de configuracao)."""

RETENTION_DAYS = int(os.getenv("MINUTA_RETENTION_DAYS", "8"))

# Neon PostgreSQL Free — limite operacional oficial (500 MB)
DATABASE_STORAGE_LIMIT_BYTES = int(
    os.getenv("MINUTA_DATABASE_LIMIT_BYTES", str(500 * 1024 * 1024))
)

RETENTION_POLICY_STATUS = "Ativa"

# Faixas de capacidade operacional (%)
CAPACITY_GREEN_MAX_PERCENT = 79.0
CAPACITY_YELLOW_MIN_PERCENT = 80.0
CAPACITY_YELLOW_MAX_PERCENT = 89.0
CAPACITY_ORANGE_MIN_PERCENT = 90.0
CAPACITY_ORANGE_MAX_PERCENT = 94.0
CAPACITY_RED_MIN_PERCENT = 95.0


def retention_days_before_today() -> int:
    return max(RETENTION_DAYS - 1, 0)


def retention_policy_description() -> str:
    anteriores = retention_days_before_today()
    if anteriores <= 0:
        return "Somente o dia atual"
    if anteriores == 1:
        return "Hoje + ultimo dia"
    return f"Hoje + {anteriores} dias"
