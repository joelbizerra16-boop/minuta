from __future__ import annotations

from dataclasses import replace

from auth.models.usuario import UsuarioPublico
from carregamentos.models.rastreabilidade_nf import RastreabilidadeNfRelatorio
from carregamentos.repository.rastreabilidade_nf_repository import RastreabilidadeNfRepository
from carregamentos.repository.sql_rastreabilidade_nf_repository import SqlRastreabilidadeNfRepository
from utils.gerador_rastreabilidade_nf import generate_rastreabilidade_nf_pdf


class RastreabilidadeNfService:
    def __init__(self, repository: RastreabilidadeNfRepository | None = None) -> None:
        self._repository = repository or SqlRastreabilidadeNfRepository()

    def buscar_relatorio(self, termo_nf: str) -> RastreabilidadeNfRelatorio | None:
        return self._repository.buscar_por_termo(termo_nf)

    def gerar_relatorio_pdf(
        self,
        termo_nf: str,
        current_user: UsuarioPublico | None = None,
    ) -> bytes:
        relatorio = self.buscar_relatorio(termo_nf)
        if relatorio is None:
            raise ValueError("Nenhum historico operacional encontrado para a Nota Fiscal informada.")

        emitido_por = str(current_user.usuario if current_user else "sistema")
        relatorio = replace(relatorio, emitido_por=emitido_por)
        return generate_rastreabilidade_nf_pdf(relatorio)
