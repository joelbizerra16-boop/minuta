from infrastructure.models.base import Base as ModelBase


def ensure_full_schema() -> None:
    from infrastructure.database import get_engine

    ModelBase.metadata.create_all(get_engine(), checkfirst=True)
