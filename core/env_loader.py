"""Carregamento do arquivo .env na raiz do projeto."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class DotenvLoadResult:
    found: bool
    loaded: bool
    path: Path
    message: str


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resolve_dotenv_path(explicit_path: Path | None = None) -> Path:
    if explicit_path is not None:
        return explicit_path.expanduser().resolve()
    return project_root() / ".env"


def load_project_dotenv(explicit_path: Path | None = None) -> DotenvLoadResult:
    """
    Carrega variaveis do .env para os.environ.

    Nao sobrescreve variaveis ja definidas no ambiente (override=False).
    Se o arquivo nao existir ou python-dotenv nao estiver instalado, retorna sem erro.
    """
    env_path = resolve_dotenv_path(explicit_path)

    try:
        from dotenv import load_dotenv
    except ImportError:
        return DotenvLoadResult(
            found=env_path.is_file(),
            loaded=False,
            path=env_path,
            message="Pacote python-dotenv nao instalado; usando apenas variaveis do sistema.",
        )

    if not env_path.is_file():
        return DotenvLoadResult(
            found=False,
            loaded=False,
            path=env_path,
            message="Arquivo .env nao localizado.",
        )

    load_dotenv(env_path, override=False)
    return DotenvLoadResult(
        found=True,
        loaded=True,
        path=env_path,
        message="Arquivo .env carregado com sucesso.",
    )
