from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime, time, timezone
from sqlalchemy import text
from sqlalchemy.orm import Session

from carregamentos.models.carregamento import MODALIDADE_BALCAO, MODALIDADE_VEICULO, normalize_chave_nfe, normalize_nf_number
from carregamentos.models.rastreabilidade_nf import (
    RastreabilidadeDocumentoLinha,
    RastreabilidadeEstatisticas,
    RastreabilidadeHistoricoLinha,
    RastreabilidadeModalidadeLinha,
    RastreabilidadeNfRelatorio,
    RastreabilidadeNfResumo,
    RastreabilidadeReentregaLinha,
    RastreabilidadeTimelineEvento,
    RastreabilidadeUsuarioLinha,
    RastreabilidadeVeiculoLinha,
)
from carregamentos.repository.rastreabilidade_nf_repository import RastreabilidadeNfRepository
from infrastructure.models.constants import DOC_TIPO_MINUTA, DOC_TIPO_ROMANEIO, HISTORICO_EVENTO_ENTREGA_BALCAO, HISTORICO_EVENTO_REENTREGA
from infrastructure.unit_of_work import UnitOfWork

RASTREABILIDADE_NF_SQL = """
WITH nf_ref AS (
    SELECT
        nf.id AS nf_id,
        nf.numero_nf,
        nf.chave_nfe,
        nf.destinatario,
        nf.peso_total AS nf_peso_total,
        nf.status_nf,
        nf.criado_em AS nf_criado_em,
        (
            SELECT COUNT(*)
            FROM item_nota_fiscal inf
            WHERE inf.nota_fiscal_id = nf.id
        ) AS qtd_itens_nf
    FROM nota_fiscal nf
    WHERE (:chave_nfe IS NOT NULL AND nf.chave_nfe = :chave_nfe)
       OR (:numero_nf IS NOT NULL AND (
            nf.numero_nf = :numero_nf
            OR nf.numero_nf = :numero_nf_raw
            OR TRIM(CAST(nf.numero_nf AS TEXT), '0') = TRIM(CAST(:numero_nf AS TEXT), '0')
       ))
    ORDER BY nf.id DESC
    LIMIT 1
),
ic_anchor AS (
    SELECT
        ic.numero_nf,
        ic.chave_nfe,
        MAX(ic.destinatario) AS destinatario,
        MAX(ic.status_nf) AS status_nf,
        SUM(COALESCE(ic.peso, 0)) AS peso_sum,
        COUNT(*) AS qtd_itens_ic
    FROM item_carregamento ic
    WHERE (:chave_nfe IS NOT NULL AND ic.chave_nfe = :chave_nfe)
       OR (:numero_nf IS NOT NULL AND (
            ic.numero_nf = :numero_nf
            OR ic.numero_nf = :numero_nf_raw
            OR TRIM(CAST(ic.numero_nf AS TEXT), '0') = TRIM(CAST(:numero_nf AS TEXT), '0')
       ))
    GROUP BY ic.numero_nf, ic.chave_nfe
    ORDER BY MAX(ic.id) DESC
    LIMIT 1
),
nf_resumo AS (
    SELECT
        COALESCE((SELECT nf_id FROM nf_ref), 0) AS nf_id,
        COALESCE((SELECT numero_nf FROM nf_ref), (SELECT numero_nf FROM ic_anchor)) AS numero_nf,
        COALESCE((SELECT chave_nfe FROM nf_ref), (SELECT chave_nfe FROM ic_anchor), '') AS chave_nfe,
        COALESCE((SELECT destinatario FROM nf_ref), (SELECT destinatario FROM ic_anchor), '--') AS destinatario,
        COALESCE((SELECT nf_peso_total FROM nf_ref), (SELECT peso_sum FROM ic_anchor), 0) AS peso_total,
        COALESCE((SELECT status_nf FROM nf_ref), (SELECT status_nf FROM ic_anchor), '--') AS status_nf,
        (SELECT nf_criado_em FROM nf_ref) AS nf_criado_em,
        COALESCE((SELECT qtd_itens_nf FROM nf_ref), (SELECT qtd_itens_ic FROM ic_anchor), 0) AS quantidade_itens
),
carregamentos_nf AS (
    SELECT DISTINCT c.id AS carregamento_id
    FROM carregamento c
    INNER JOIN item_carregamento ic ON ic.carregamento_id = c.id
    CROSS JOIN nf_resumo nr
    WHERE nr.numero_nf IS NOT NULL
      AND (
            (nr.nf_id > 0 AND ic.nota_fiscal_id = nr.nf_id)
            OR (:chave_nfe IS NOT NULL AND ic.chave_nfe = :chave_nfe)
            OR ic.numero_nf = nr.numero_nf
            OR TRIM(CAST(ic.numero_nf AS TEXT), '0') = TRIM(CAST(nr.numero_nf AS TEXT), '0')
      )
)
SELECT
    nr.numero_nf,
    nr.chave_nfe,
    nr.destinatario,
    nr.peso_total,
    nr.status_nf,
    nr.nf_criado_em,
    nr.quantidade_itens,
    c.id AS carregamento_id,
    c.numero_carregamento,
    c.data AS carregamento_data,
    c.hora AS carregamento_hora,
    c.criado_em AS carregamento_criado_em,
    c.modalidade,
    c.status AS carregamento_status,
    c.reentrega,
    c.motorista,
    c.placa,
    c.quantidade_itens AS carregamento_qtd_itens,
    c.peso_total AS carregamento_peso,
    c.quantidade_impressoes,
    c.ultima_impressao_em,
    u.usuario AS usuario_login,
    ui.usuario AS ultima_impressao_usuario,
    v.placa AS veiculo_placa,
    (
        SELECT ic.rota
        FROM item_carregamento ic
        CROSS JOIN nf_resumo nr2
        WHERE ic.carregamento_id = c.id
          AND (
                (nr2.nf_id > 0 AND ic.nota_fiscal_id = nr2.nf_id)
                OR ic.numero_nf = nr2.numero_nf
                OR TRIM(CAST(ic.numero_nf AS TEXT), '0') = TRIM(CAST(nr2.numero_nf AS TEXT), '0')
          )
        ORDER BY ic.sequencia, ic.id
        LIMIT 1
    ) AS rota_nf,
    (
        SELECT json_group_array(
            json_object(
                'evento', ho.evento,
                'criado_em', ho.criado_em,
                'descricao', COALESCE(ho.descricao, ''),
                'usuario', hu.usuario
            )
        )
        FROM historico_operacional ho
        LEFT JOIN usuario hu ON hu.id = ho.usuario_id
        WHERE ho.carregamento_id = c.id
        ORDER BY ho.criado_em, ho.id
    ) AS historicos_json,
    (
        SELECT json_group_array(
            json_object(
                'tipo', d.tipo,
                'criado_em', d.criado_em,
                'usuario', du.usuario
            )
        )
        FROM documento d
        LEFT JOIN usuario du ON du.id = d.usuario_id
        WHERE d.carregamento_id = c.id
        ORDER BY d.criado_em, d.id
    ) AS documentos_json
FROM nf_resumo nr
INNER JOIN carregamentos_nf cn ON 1 = 1
INNER JOIN carregamento c ON c.id = cn.carregamento_id
INNER JOIN usuario u ON u.id = c.usuario_id
LEFT JOIN usuario ui ON ui.id = c.ultima_impressao_usuario_id
LEFT JOIN veiculo v ON v.id = c.veiculo_id
ORDER BY c.data ASC, c.hora ASC, c.id ASC
"""

EMPRESA_PADRAO = "BRIDA LUBRIFICANTES LTDA"


class SqlRastreabilidadeNfRepository(RastreabilidadeNfRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    def buscar_por_termo(self, termo: str) -> RastreabilidadeNfRelatorio | None:
        chave_nfe, numero_nf, numero_nf_raw = self._resolve_termo(termo)
        if chave_nfe is None and numero_nf is None:
            return None

        with UnitOfWork(self._session) as uow:
            rows = uow.session.execute(
                text(RASTREABILIDADE_NF_SQL),
                {
                    "chave_nfe": chave_nfe,
                    "numero_nf": numero_nf,
                    "numero_nf_raw": numero_nf_raw,
                },
            ).mappings().all()

        if not rows:
            return None

        return self._build_relatorio(rows)

    @staticmethod
    def _resolve_termo(termo: str) -> tuple[str | None, str | None, str | None]:
        normalized = str(termo or "").strip()
        if not normalized:
            return None, None, None

        chave = normalize_chave_nfe(normalized)
        if chave:
            return chave, None, None

        numero = normalize_nf_number(normalized)
        if numero:
            return None, numero, normalized

        digits = re.sub(r"\D", "", normalized)
        if digits:
            return None, digits.lstrip("0") or digits, normalized

        return None, normalized.lower(), normalized

    def _build_relatorio(self, rows: list) -> RastreabilidadeNfRelatorio:
        first = rows[0]
        carregamentos = [self._parse_carregamento_row(row) for row in rows]

        datas_saida = [item["data_hora"] for item in carregamentos if item["data_hora"]]
        reentregas_rows = [item for item in carregamentos if item["reentrega"]]

        resumo = RastreabilidadeNfResumo(
            numero_nf=str(first["numero_nf"] or "--"),
            chave_nfe=str(first["chave_nfe"] or "--"),
            destinatario=str(first["destinatario"] or "--"),
            quantidade_itens=int(first["quantidade_itens"] or 0),
            peso_total=float(first["peso_total"] or 0),
            quantidade_carregamentos=len(carregamentos),
            quantidade_reentregas=len(reentregas_rows),
            primeira_saida=self._format_datetime_br(min(datas_saida)) if datas_saida else "--",
            ultima_saida=self._format_datetime_br(max(datas_saida)) if datas_saida else "--",
            status_atual=str(first["status_nf"] or "--"),
            nf_criado_em=self._parse_datetime(first["nf_criado_em"]),
        )

        historico = [self._build_historico_linha(item) for item in carregamentos]
        reentregas = [self._build_reentrega_linha(item) for item in reentregas_rows]
        veiculos = self._build_veiculos(carregamentos)
        usuarios = self._build_usuarios(carregamentos)
        modalidades = self._build_modalidades(carregamentos)
        documentos = [self._build_documento_linha(item) for item in carregamentos]
        estatisticas = self._build_estatisticas(carregamentos, resumo)
        timeline = self._build_timeline(resumo, carregamentos)

        return RastreabilidadeNfRelatorio(
            empresa=EMPRESA_PADRAO,
            emitido_em=datetime.now(timezone.utc),
            emitido_por="",
            resumo=resumo,
            historico=historico,
            reentregas=reentregas,
            veiculos=veiculos,
            usuarios=usuarios,
            modalidades=modalidades,
            documentos=documentos,
            estatisticas=estatisticas,
            timeline=timeline,
        )

    def _parse_carregamento_row(self, row) -> dict:
        historicos = self._parse_json_array(row.get("historicos_json"))
        documentos = self._parse_json_array(row.get("documentos_json"))
        data_hora = self._resolve_data_hora(row, historicos)

        placa = str(row.get("placa") or row.get("veiculo_placa") or "--")
        motorista = str(row.get("motorista") or "--")
        veiculo = placa if placa != "--" else "--"

        minuta = next((doc for doc in documentos if doc.get("tipo") == DOC_TIPO_MINUTA), None)
        romaneio = next((doc for doc in documentos if doc.get("tipo") == DOC_TIPO_ROMANEIO), None)
        modalidade = str(row.get("modalidade") or MODALIDADE_VEICULO)

        if modalidade == MODALIDADE_BALCAO:
            pdf_gerado = "BALCAO"
            documento_label = "Retirada"
        elif minuta is not None:
            pdf_gerado = "Minuta"
            documento_label = "Romaneio" if romaneio is not None else "--"
        else:
            pdf_gerado = "--"
            documento_label = "Romaneio" if romaneio is not None else "--"

        historico_reentrega = next(
            (item for item in historicos if item.get("evento") == HISTORICO_EVENTO_REENTREGA),
            None,
        )

        return {
            "carregamento_id": int(row["carregamento_id"]),
            "numero_carregamento": str(row["numero_carregamento"] or "--"),
            "data_hora": data_hora,
            "modalidade": modalidade,
            "status": str(row.get("carregamento_status") or "--"),
            "usuario": str(row.get("usuario_login") or "--"),
            "motorista": motorista,
            "veiculo": veiculo,
            "placa": placa,
            "rota": str(row.get("rota_nf") or "--"),
            "pdf_gerado": pdf_gerado,
            "documento": documento_label,
            "reentrega": bool(row.get("reentrega")),
            "balcao": modalidade == MODALIDADE_BALCAO,
            "historicos": historicos,
            "documentos": documentos,
            "quantidade_impressoes": int(row.get("quantidade_impressoes") or 0),
            "ultima_impressao_em": self._parse_datetime(row.get("ultima_impressao_em")),
            "ultima_impressao_usuario": str(row.get("ultima_impressao_usuario") or "--"),
            "carregamento_qtd_itens": int(row.get("carregamento_qtd_itens") or 0),
            "carregamento_peso": float(row.get("carregamento_peso") or 0),
            "motivo_reentrega": str((historico_reentrega or {}).get("descricao") or "Reentrega"),
            "carregamento_data": row.get("carregamento_data"),
            "carregamento_hora": row.get("carregamento_hora"),
        }

    @staticmethod
    def _parse_json_array(raw_value: object) -> list[dict]:
        if raw_value in (None, "", "[]"):
            return []
        try:
            payload = json.loads(str(raw_value))
        except json.JSONDecodeError:
            return []
        if not isinstance(payload, list):
            return []
        return [item for item in payload if isinstance(item, dict)]

    @staticmethod
    def _parse_datetime(value: object) -> datetime | None:
        if value is None:
            return None
        if isinstance(value, datetime):
            if value.tzinfo is None:
                return value.replace(tzinfo=timezone.utc)
            return value
        text_value = str(value).strip()
        if not text_value:
            return None
        normalized = text_value.replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError:
            return None
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed

    def _resolve_data_hora(self, row, historicos: list[dict]) -> datetime | None:
        if historicos:
            parsed = [self._parse_datetime(item.get("criado_em")) for item in historicos]
            valid = [item for item in parsed if item is not None]
            if valid:
                return min(valid)

        criado_em = self._parse_datetime(row.get("carregamento_criado_em"))
        if criado_em is not None:
            return criado_em

        data_value = row.get("carregamento_data")
        hora_value = row.get("carregamento_hora")
        if data_value is None:
            return None

        hora = hora_value if isinstance(hora_value, time) else time(0, 0, 0)
        if isinstance(data_value, datetime):
            base_date = data_value.date()
        else:
            base_date = data_value
        combined = datetime.combine(base_date, hora)
        return combined.replace(tzinfo=timezone.utc)

    @staticmethod
    def _build_historico_linha(item: dict) -> RastreabilidadeHistoricoLinha:
        data_hora = item["data_hora"] or datetime.now(timezone.utc)
        return RastreabilidadeHistoricoLinha(
            data_hora=data_hora,
            numero_carregamento=item["numero_carregamento"],
            modalidade=item["modalidade"],
            status=item["status"],
            usuario=item["usuario"],
            motorista=item["motorista"] if item["motorista"] != "--" else "--",
            veiculo=item["veiculo"],
            placa=item["placa"] if item["placa"] != "--" else "--",
            rota=item["rota"],
            pdf_gerado=item["pdf_gerado"],
            documento=item["documento"],
            reentrega=item["reentrega"],
            balcao=item["balcao"],
        )

    @staticmethod
    def _build_reentrega_linha(item: dict) -> RastreabilidadeReentregaLinha:
        return RastreabilidadeReentregaLinha(
            data=SqlRastreabilidadeNfRepository._format_datetime_br(item["data_hora"]),
            carregamento=item["numero_carregamento"],
            usuario=item["usuario"],
            motivo=item["motivo_reentrega"],
            status=item["status"],
        )

    @staticmethod
    def _build_veiculos(carregamentos: list[dict]) -> list[RastreabilidadeVeiculoLinha]:
        agrupado: dict[tuple[str, str], dict] = {}
        for item in carregamentos:
            if item["modalidade"] != MODALIDADE_VEICULO:
                continue
            placa = item["placa"] if item["placa"] != "--" else "--"
            chave = (item["veiculo"], placa)
            bucket = agrupado.setdefault(
                chave,
                {"veiculo": item["veiculo"], "placa": placa, "viagens": 0, "motoristas": set()},
            )
            bucket["viagens"] += 1
            if item["motorista"] != "--":
                bucket["motoristas"].add(item["motorista"])

        linhas = []
        for bucket in agrupado.values():
            motoristas = sorted(bucket["motoristas"])
            linhas.append(
                RastreabilidadeVeiculoLinha(
                    veiculo=bucket["veiculo"],
                    placa=bucket["placa"],
                    quantidade_viagens=bucket["viagens"],
                    motorista=", ".join(motoristas) if motoristas else "--",
                )
            )
        return sorted(linhas, key=lambda item: (-item.quantidade_viagens, item.placa))

    @staticmethod
    def _build_usuarios(carregamentos: list[dict]) -> list[RastreabilidadeUsuarioLinha]:
        agrupado: dict[str, list[datetime]] = {}
        for item in carregamentos:
            usuario = item["usuario"]
            data_hora = item["data_hora"]
            if data_hora is None:
                continue
            agrupado.setdefault(usuario, []).append(data_hora)

        linhas = []
        for usuario, datas in agrupado.items():
            datas_sorted = sorted(datas)
            linhas.append(
                RastreabilidadeUsuarioLinha(
                    usuario=usuario,
                    quantidade_operacoes=len(datas_sorted),
                    primeira_operacao=SqlRastreabilidadeNfRepository._format_datetime_br(datas_sorted[0]),
                    ultima_operacao=SqlRastreabilidadeNfRepository._format_datetime_br(datas_sorted[-1]),
                )
            )
        return sorted(linhas, key=lambda item: (-item.quantidade_operacoes, item.usuario))

    @staticmethod
    def _build_modalidades(carregamentos: list[dict]) -> list[RastreabilidadeModalidadeLinha]:
        counter: Counter[str] = Counter()
        for item in carregamentos:
            if item["reentrega"]:
                counter["REENTREGA"] += 1
            counter[item["modalidade"]] += 1

        ordem = [MODALIDADE_VEICULO, MODALIDADE_BALCAO, "REENTREGA"]
        linhas = []
        for modalidade in ordem:
            quantidade = counter.get(modalidade, 0)
            if quantidade:
                linhas.append(RastreabilidadeModalidadeLinha(modalidade=modalidade, quantidade=quantidade))
        for modalidade, quantidade in sorted(counter.items()):
            if modalidade not in ordem:
                linhas.append(RastreabilidadeModalidadeLinha(modalidade=modalidade, quantidade=quantidade))
        return linhas

    @staticmethod
    def _build_documento_linha(item: dict) -> RastreabilidadeDocumentoLinha:
        minuta = "Sim" if item["pdf_gerado"] == "Minuta" else ("Balcao" if item["balcao"] else "Nao")
        romaneio = "Sim" if item["documento"] == "Romaneio" else ("Retirada" if item["balcao"] else "Nao")
        ultima = item["ultima_impressao_em"]
        return RastreabilidadeDocumentoLinha(
            minuta=minuta,
            romaneio=romaneio,
            data=SqlRastreabilidadeNfRepository._format_datetime_br(item["data_hora"]),
            usuario=item["usuario"],
            quantidade_impressoes=item["quantidade_impressoes"],
            ultima_impressao=SqlRastreabilidadeNfRepository._format_datetime_br(ultima) if ultima else "--",
        )

    @staticmethod
    def _build_estatisticas(carregamentos: list[dict], resumo: RastreabilidadeNfResumo) -> RastreabilidadeEstatisticas:
        motoristas = {
            item["motorista"]
            for item in carregamentos
            if item["motorista"] not in ("", "--") and item["modalidade"] == MODALIDADE_VEICULO
        }
        placas = {
            item["placa"]
            for item in carregamentos
            if item["placa"] not in ("", "--") and item["modalidade"] == MODALIDADE_VEICULO
        }
        usuarios = {item["usuario"] for item in carregamentos if item["usuario"]}
        peso_expedido = sum(item["carregamento_peso"] for item in carregamentos)
        itens_expedidos = sum(item["carregamento_qtd_itens"] for item in carregamentos)

        return RastreabilidadeEstatisticas(
            total_carregamentos=len(carregamentos),
            total_itens_expedidos=itens_expedidos,
            peso_expedido=peso_expedido,
            total_reentregas=resumo.quantidade_reentregas,
            total_balcao=sum(1 for item in carregamentos if item["balcao"]),
            veiculos_diferentes=len(placas),
            motoristas_diferentes=len(motoristas),
            usuarios_envolvidos=len(usuarios),
        )

    @staticmethod
    def _build_timeline(
        resumo: RastreabilidadeNfResumo,
        carregamentos: list[dict],
    ) -> list[RastreabilidadeTimelineEvento]:
        eventos: list[RastreabilidadeTimelineEvento] = []

        if resumo.nf_criado_em is not None:
            eventos.append(
                RastreabilidadeTimelineEvento(rotulo="Recebimento XML", data_hora=resumo.nf_criado_em)
            )

        for item in carregamentos:
            data_hora = item["data_hora"]
            if data_hora is None:
                continue

            eventos.append(
                RastreabilidadeTimelineEvento(
                    rotulo=f"Associada ao carregamento {item['numero_carregamento']}",
                    data_hora=data_hora,
                )
            )

            for historico in item["historicos"]:
                parsed = SqlRastreabilidadeNfRepository._parse_datetime(historico.get("criado_em"))
                if parsed is None:
                    continue
                evento = str(historico.get("evento") or "")
                if evento == HISTORICO_EVENTO_REENTREGA:
                    rotulo = "Reentrega"
                elif evento == HISTORICO_EVENTO_ENTREGA_BALCAO:
                    rotulo = "Entrega Balcao"
                else:
                    rotulo = "Carga Finalizada"
                eventos.append(RastreabilidadeTimelineEvento(rotulo=rotulo, data_hora=parsed))

            for documento in item["documentos"]:
                parsed = SqlRastreabilidadeNfRepository._parse_datetime(documento.get("criado_em"))
                if parsed is None:
                    continue
                tipo = str(documento.get("tipo") or "")
                if tipo == DOC_TIPO_MINUTA:
                    rotulo = "Minuta emitida"
                elif tipo == DOC_TIPO_ROMANEIO:
                    rotulo = "Romaneio emitido"
                else:
                    rotulo = f"Documento {tipo}"
                eventos.append(RastreabilidadeTimelineEvento(rotulo=rotulo, data_hora=parsed))

        eventos.sort(key=lambda item: (item.data_hora, item.rotulo))
        return eventos

    @staticmethod
    def _format_datetime_br(value: datetime | None) -> str:
        if value is None:
            return "--"
        local = value.astimezone(timezone.utc)
        return local.strftime("%d/%m/%Y %H:%M")
