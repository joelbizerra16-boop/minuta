from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from pathlib import Path

from sqlalchemy import text
from sqlalchemy.engine import Engine

from core.retention_policy import DATABASE_STORAGE_LIMIT_BYTES
from infrastructure.database import get_engine
from infrastructure.unit_of_work import UnitOfWork

_LOGGER = logging.getLogger("minuta.database_usage")


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
    """Mede o espaco ocupado pelo banco ativo via SQLAlchemy Engine (fonte unica de verdade)."""

    def medir(self) -> UsoBancoDados:
        limite = DATABASE_STORAGE_LIMIT_BYTES
        inicio = time.perf_counter()

        try:
            engine = get_engine()
        except RuntimeError as exc:
            _LOGGER.exception(
                "database_usage.engine_unavailable limite_bytes=%s erro=%s",
                limite,
                exc,
            )
            return UsoBancoDados(
                motor="Desconhecido",
                bytes_ocupados=None,
                bytes_limite=limite,
                bytes_disponiveis=None,
                utilizacao_percentual=None,
                medicao_direta=False,
                rotulo_ocupacao="Espaco ocupado",
                observacao="Banco de dados ainda nao configurado para medicao.",
            )

        driver = self._normalizar_driver(engine)
        database = str(engine.url.database or "").strip()
        engine_url = engine.url.render_as_string(hide_password=True)

        _LOGGER.info(
            "database_usage.inicio driver=%s engine_url=%s engine_database=%s limite_bytes=%s",
            driver,
            engine_url,
            database,
            limite,
        )

        try:
            if driver == "postgresql":
                resultado = self._medir_postgresql(engine, limite)
            elif driver == "sqlite":
                resultado = self._medir_sqlite(engine, limite)
            else:
                _LOGGER.error(
                    "database_usage.driver_nao_suportado driver=%s engine_url=%s engine_database=%s",
                    driver,
                    engine_url,
                    database,
                )
                resultado = UsoBancoDados(
                    motor=driver or "Desconhecido",
                    bytes_ocupados=None,
                    bytes_limite=limite,
                    bytes_disponiveis=None,
                    utilizacao_percentual=None,
                    medicao_direta=False,
                    rotulo_ocupacao="Espaco ocupado",
                    observacao=f"Motor de banco nao suportado para medicao: {driver or 'desconhecido'}.",
                )
        except Exception:
            _LOGGER.exception(
                "database_usage.falha driver=%s engine_url=%s engine_database=%s",
                driver,
                engine_url,
                database,
            )
            raise

        duracao_ms = (time.perf_counter() - inicio) * 1000.0
        _LOGGER.info(
            "database_usage.concluido driver=%s metodo=%s bytes_ocupados=%s bytes_livres=%s "
            "percentual=%s duracao_ms=%.2f observacao=%s",
            driver,
            resultado.rotulo_ocupacao,
            resultado.bytes_ocupados,
            resultado.bytes_disponiveis,
            resultado.utilizacao_percentual,
            duracao_ms,
            resultado.observacao,
        )
        return resultado

    @staticmethod
    def _normalizar_driver(engine: Engine) -> str:
        driver = str(engine.url.drivername or "").strip().lower()
        if "+" in driver:
            driver = driver.split("+", 1)[0]
        return driver

    def _medir_postgresql(self, engine: Engine, limite: int) -> UsoBancoDados:
        _LOGGER.info(
            "database_usage.postgresql metodo=pg_database_size engine_database=%s",
            engine.url.database,
        )
        with UnitOfWork() as uow:
            bytes_ocupados = int(
                uow.session.scalar(text("SELECT pg_database_size(current_database())")) or 0
            )

        return self._build_result(
            motor="PostgreSQL",
            bytes_ocupados=bytes_ocupados,
            limite=limite,
            medicao_direta=True,
            rotulo="Espaco utilizado",
            observacao="Medicao via pg_database_size(current_database()).",
        )

    def _medir_sqlite(self, engine: Engine, limite: int) -> UsoBancoDados:
        database = str(engine.url.database or "").strip()
        metodo = "sqlite_arquivo"

        if database and database != ":memory:":
            db_file = Path(database)
            if not db_file.is_absolute():
                db_file = db_file.resolve()
            if db_file.is_file():
                bytes_ocupados = int(db_file.stat().st_size)
                _LOGGER.info(
                    "database_usage.sqlite metodo=%s engine_database=%s arquivo=%s bytes=%s",
                    metodo,
                    database,
                    db_file,
                    bytes_ocupados,
                )
                return self._build_result(
                    motor="SQLite",
                    bytes_ocupados=bytes_ocupados,
                    limite=limite,
                    medicao_direta=False,
                    rotulo="Espaco ocupado estimado",
                    observacao="Medicao baseada no arquivo SQLite referenciado pelo Engine.",
                )

            _LOGGER.warning(
                "database_usage.sqlite arquivo_indisponivel engine_database=%s arquivo_resolvido=%s exists=%s",
                database,
                db_file,
                db_file.exists(),
            )

        metodo = "sqlite_pragma"
        bytes_ocupados = self._medir_sqlite_via_pragma(engine)
        _LOGGER.info(
            "database_usage.sqlite metodo=%s engine_database=%s bytes=%s",
            metodo,
            database or ":memory:",
            bytes_ocupados,
        )
        return self._build_result(
            motor="SQLite",
            bytes_ocupados=bytes_ocupados,
            limite=limite,
            medicao_direta=True,
            rotulo="Espaco ocupado estimado",
            observacao="Medicao via PRAGMA page_count x page_size no Engine ativo.",
        )

    @staticmethod
    def _medir_sqlite_via_pragma(engine: Engine) -> int:
        with engine.connect() as connection:
            page_count = int(connection.exec_driver_sql("PRAGMA page_count").scalar() or 0)
            page_size = int(connection.exec_driver_sql("PRAGMA page_size").scalar() or 0)
        return max(page_count * page_size, 0)

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
