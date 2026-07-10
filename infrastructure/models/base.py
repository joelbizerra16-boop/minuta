from __future__ import annotations

from typing import Any

from sqlalchemy.orm import DeclarativeBase


def _merge_extend_existing(table_args: Any) -> Any:
    """Permite reimportar modulos ORM sem InvalidRequestError (ex.: rerun Streamlit)."""
    extend = {"extend_existing": True}
    if table_args is None:
        return extend
    if isinstance(table_args, dict):
        return {**table_args, **extend}
    if isinstance(table_args, tuple):
        if table_args and isinstance(table_args[-1], dict):
            return (*table_args[:-1], {**table_args[-1], **extend})
        return (*table_args, extend)
    return table_args


class Base(DeclarativeBase):
    def __init_subclass__(cls, **kwargs: object) -> None:
        if not cls.__dict__.get("__abstract__", False) and "__tablename__" in cls.__dict__:
            cls.__table_args__ = _merge_extend_existing(cls.__dict__.get("__table_args__"))
        super().__init_subclass__(**kwargs)
