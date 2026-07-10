from __future__ import annotations

import json
import socket
from datetime import datetime

from auth.models.usuario import UsuarioPublico
from infrastructure.models.constants import AUDIT_CATEGORIA_CARREGAMENTO, AUDIT_EVENTO_DECISAO_OPERACIONAL
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository


class DecisaoOperacionalAuditoriaService:
    """Registra decisões confirmadas pelo operador com rastreabilidade completa."""

    @staticmethod
    def resolver_estacao(ip_origem: str | None = None) -> str:
        host = socket.gethostname() or "desconhecida"
        ip = ip_origem or "--"
        return f"{host}|{ip}"

    @staticmethod
    def build_payload(
        *,
        usuario_nome: str,
        motivo: str,
        decisao: str,
        situacao_anterior: dict,
        situacao_posterior: dict,
        nfs_envolvidas: list[str],
        impactos: list[str],
        riscos: list[str],
        recomendacao: str,
        estacao: str,
        extras: dict | None = None,
    ) -> str:
        payload = {
            "usuario_nome": usuario_nome,
            "data": datetime.now().strftime("%d/%m/%Y"),
            "hora": datetime.now().strftime("%H:%M:%S"),
            "estacao": estacao,
            "motivo": motivo,
            "decisao": decisao,
            "situacao_anterior": situacao_anterior,
            "situacao_posterior": situacao_posterior,
            "nfs_envolvidas": nfs_envolvidas,
            "impactos_apresentados": impactos,
            "riscos_apresentados": riscos,
            "recomendacao_apresentada": recomendacao,
        }
        if extras:
            payload.update(extras)
        return json.dumps(payload, ensure_ascii=False, sort_keys=True)

    def registrar(
        self,
        audit_repo: SqlEventoAuditoriaRepository,
        *,
        usuario_id: int | None,
        usuario_nome: str,
        carregamento_id: int | None,
        motivo: str,
        decisao: str,
        situacao_anterior: dict,
        situacao_posterior: dict,
        nfs_envolvidas: list[str],
        impactos: list[str],
        riscos: list[str],
        recomendacao: str,
        ip_origem: str | None = None,
        extras: dict | None = None,
    ) -> EventoAuditoriaRecord:
        estacao = self.resolver_estacao(ip_origem)
        return audit_repo.append(
            EventoAuditoriaRecord(
                id=0,
                categoria=AUDIT_CATEGORIA_CARREGAMENTO,
                evento=AUDIT_EVENTO_DECISAO_OPERACIONAL,
                usuario_id=usuario_id,
                entidade_tipo="carregamento" if carregamento_id else None,
                entidade_id=carregamento_id,
                descricao=f"Decisao operacional: {decisao}",
                metadados_json=self.build_payload(
                    usuario_nome=usuario_nome,
                    motivo=motivo,
                    decisao=decisao,
                    situacao_anterior=situacao_anterior,
                    situacao_posterior=situacao_posterior,
                    nfs_envolvidas=nfs_envolvidas,
                    impactos=impactos,
                    riscos=riscos,
                    recomendacao=recomendacao,
                    estacao=estacao,
                    extras=extras,
                ),
                ip_origem=ip_origem,
            )
        )

    @staticmethod
    def snapshot_carregamento(carregamento) -> dict:
        return {
            "id": int(carregamento.id),
            "numero_carregamento": str(carregamento.numero_carregamento or ""),
            "status": str(carregamento.status or ""),
            "quantidade_itens": int(carregamento.quantidade_itens or 0),
            "quantidade_nf": int(carregamento.quantidade_nf or 0),
            "reentrega": bool(carregamento.reentrega),
        }

    @staticmethod
    def usuario_nome(current_user: UsuarioPublico | None) -> str:
        if current_user and current_user.nome:
            return str(current_user.nome)
        if current_user and current_user.usuario:
            return str(current_user.usuario)
        return "sistema"
