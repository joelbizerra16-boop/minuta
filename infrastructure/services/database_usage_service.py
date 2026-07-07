from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from sqlalchemy import text

from core.retention_policy import DATABASE_STORAGE_LIMIT_BYTES
from core.settings import get_settings
from infrastructure.unit_of_work import UnitOfWork


@dataclass(frozen=True)
class UsoBancoDados:
    motor: str
    bytes_ocupados: int | None
    bytes_limite: int | None
    bytes_disponiveis: int | None
    utilizacao_percentual: float | None
    medicao_direta: bool
    rotulo_ocupacao: str
    observacao: str | None = None


class DatabaseUsageService:
    """Mede ou estima o espaco ocupado pelo banco sem alterar dados."""

    def medir(self) -> UsoBancoDados:
        settings = get_settings()
        database_url = str(settings.database_url or "").strip().lower()
        limite = DATABASE_STORAGE_LIMIT_BYTES

        if database_url.startswith("postgresql"):
            return self._medir_postgresql(limite)
        if database_url.startswith("sqlite"):
            return self._medir_sqlite(database_url, limite)
        return UsoBancoDados(
            motor="Desconhecido",
            bytes_ocupados=None,
            bytes_limite=limite,
            bytes_disponiveis=None,
            utilizacao_percentual=None,
            medicao_direta=False,
            rotulo_ocupacao="Espaco ocupado",
            observacao="Nao foi possivel identificar o motor do banco para medir o espaco ocupado.",
        )

    def _medir_postgresql(self, limite: int) -> UsoBancoDados:
        try:
            with UnitOfWork() as uow:
                bytes_ocupados = int(
                    uow.session.scalar(text("SELECT pg_database_size(current_database())")) or 0
                )
        except Exception as exc:
            return UsoBancoDados(
                motor="PostgreSQL",
                bytes_ocupados=None,
                bytes_limite=limite,
                bytes_disponiveis=None,
                utilizacao_percentual=None,
                medicao_direta=False,
                rotulo_ocupacao="Espaco ocupado",
                observacao=f"Nao foi possivel medir o espaco do PostgreSQL ({exc}).",
            )

        return self._build_result(
            motor="PostgreSQL",
            bytes_ocupados=bytes_ocupados,
            limite=limite,
            medicao_direta=True,
            rotulo="Espaco utilizado",
        )

    def _medir_sqlite(self, database_url: str, limite: int) -> UsoBancoDados:
        sqlite_path = database_url.replace("sqlite:///", "", 1)
        db_file = Path(sqlite_path)
        if not db_file.is_file():
            return UsoBancoDados(
                motor="SQLite",
                bytes_ocupados=None,
                bytes_limite=limite,
                bytes_disponiveis=None,
                utilizacao_percentual=None,
                medicao_direta=False,
                rotulo_ocupacao="Espaco ocupado estimado",
                observacao="Arquivo do banco SQLite nao encontrado para medicao.",
            )

        bytes_ocupados = int(db_file.stat().st_size)
        return self._build_result(
            motor="SQLite",
            bytes_ocupados=bytes_ocupados,
            limite=limite,
            medicao_direta=False,
            rotulo="Espaco ocupado estimado",
            observacao="Medicao baseada no arquivo local SQLite (ambiente de desenvolvimento).",
        )

    @staticmethod
    def _build_result(
        *,
        motor: str,
        bytes_ocupados: int,
        limite: int,
        medicao_direta: bool,
        rotulo: str,
        observacao: str | None = None,
    ) -> UsoBancoDados:
        disponivel = max(limite - bytes_ocupados, 0) if limite > 0 else None
        utilizacao = round((bytes_ocupados / limite) * 100, 1) if limite > 0 else None
        return UsoBancoDados(
            motor=motor,
            bytes_ocupados=bytes_ocupados,
            bytes_limite=limite,
            bytes_disponiveis=disponivel,
            utilizacao_percentual=utilizacao,
            medicao_direta=medicao_direta,
            rotulo_ocupacao=rotulo,
            observacao=observacao,
        )
