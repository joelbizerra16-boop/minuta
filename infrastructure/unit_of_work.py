from __future__ import annotations

from types import TracebackType

from sqlalchemy.orm import Session

from infrastructure.database import get_session_factory


class UnitOfWork:
  def __init__(self, session: Session | None = None) -> None:
      self._session = session
      self._owns_session = session is None

  def __enter__(self) -> UnitOfWork:
      if self._session is None:
          self._session = get_session_factory()()
      return self

  def __exit__(
      self,
      exc_type: type[BaseException] | None,
      exc: BaseException | None,
      tb: TracebackType | None,
  ) -> None:
      if self._session is None:
          return
      try:
          if exc_type is None:
              self._session.commit()
          else:
              self._session.rollback()
      finally:
          if self._owns_session:
              self._session.close()
          self._session = None

  @property
  def session(self) -> Session:
      if self._session is None:
          raise RuntimeError("UnitOfWork session is not active.")
      return self._session
