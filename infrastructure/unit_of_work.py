from __future__ import annotations

import logging
import time
from types import TracebackType

from sqlalchemy.orm import Session

from infrastructure.database import get_session_factory

_LOGGER = logging.getLogger("minuta.unit_of_work")


class UnitOfWork:
    def __init__(self, session: Session | None = None) -> None:
        self._session = session
        self._owns_session = session is None
        self._started_at: float | None = None

    def __enter__(self) -> UnitOfWork:
        if self._session is None:
            self._session = get_session_factory()()
        self._started_at = time.perf_counter()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if self._session is None:
            return
        started_at = self._started_at or time.perf_counter()
        action = "rollback" if exc_type is not None else "commit"
        try:
            if exc_type is None:
                self._session.commit()
            else:
                self._session.rollback()
        finally:
            duration_ms = (time.perf_counter() - started_at) * 1000.0
            dialect = "unknown"
            bind = self._session.get_bind()
            if bind is not None:
                dialect = str(bind.dialect.name)
            _LOGGER.debug(
                "unit_of_work.%s dialect=%s owns_session=%s duracao_ms=%.2f",
                action,
                dialect,
                self._owns_session,
                duration_ms,
            )
            if self._owns_session:
                self._session.close()
            self._session = None
            self._started_at = None

    @property
    def session(self) -> Session:
        if self._session is None:
            raise RuntimeError("UnitOfWork session is not active.")
        return self._session
