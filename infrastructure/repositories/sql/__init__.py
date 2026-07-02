"""Implementacoes SQL dos contratos de repositorio (preparadas para fases M1+)."""

from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
from infrastructure.repositories.sql.documento_repository import SqlDocumentoRepository
from infrastructure.repositories.sql.historico_repository import SqlHistoricoRepository
from infrastructure.repositories.sql.nota_fiscal_repository import SqlNotaFiscalRepository

__all__ = [
    "SqlConfiguracaoRepository",
    "SqlDocumentoRepository",
    "SqlHistoricoRepository",
    "SqlNotaFiscalRepository",
]
