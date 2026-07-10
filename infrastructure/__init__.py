"""Camada de infraestrutura: ORM, banco de dados, repositories SQL e Unit of Work."""

from __future__ import annotations

from typing import Any

__all__ = [
    "Base",
    "UnitOfWork",
    "configure_database",
    "get_engine",
    "get_session_factory",
]


def __getattr__(name: str) -> Any:
    if name == "Base":
        from infrastructure.models.base import Base

        return Base
    if name == "UnitOfWork":
        from infrastructure.unit_of_work import UnitOfWork

        return UnitOfWork
    if name == "configure_database":
        from infrastructure.database import configure_database

        return configure_database
    if name == "get_engine":
        from infrastructure.database import get_engine

        return get_engine
    if name == "get_session_factory":
        from infrastructure.database import get_session_factory

        return get_session_factory
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
