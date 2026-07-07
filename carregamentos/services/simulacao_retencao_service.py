from __future__ import annotations

import hashlib
import time
from datetime import date
from pathlib import Path

from carregamentos.models.simulacao_retencao import (
    DocumentoPdfValidacao,
    DocumentoXmlValidacao,
    PacoteRetencaoUnitario,
    ProblemaIntegridade,
    RelatorioSimulacaoRetencao,
    ResumoSaudeSimulacao,
    SaudePacote,
)
from carregamentos.repository.simulacao_retencao_repository import (
    ArvoreCarregamentoRaw,
    SimulacaoRetencaoRepository,
)
from carregamentos.repository.sql_simulacao_retencao_repository import SqlSimulacaoRetencaoRepository
from carregamentos.services.gestao_dados_service import GestaoDadosService

_BYTES_ESTIMADOS_POR_REGISTRO = {
    "carregamento": 520,
    "item_carregamento": 220,
    "documento": 320,
    "historico": 280,
    "evento": 420,
    "item_nota_fiscal": 160,
}


class SimulacaoRetencaoService:
    """Simulacao executavel da retencao — valida a arvore sem alterar dados."""

    def __init__(
        self,
        repository: SimulacaoRetencaoRepository | None = None,
        gestao_dados_service: GestaoDadosService | None = None,
        pdf_storage_dir: Path | None = None,
        xml_storage_dir: Path | None = None,
    ) -> None:
        self._repository = repository or SqlSimulacaoRetencaoRepository()
        self._gestao_dados = gestao_dados_service or GestaoDadosService(pdf_storage_dir=pdf_storage_dir)
        self._pdf_storage_dir = pdf_storage_dir
        self._xml_storage_dir = xml_storage_dir

    def executar_simulacao(self, referencia: date | None = None) -> RelatorioSimulacaoRetencao:
        inicio = time.perf_counter()
        data_corte = self._gestao_dados.calcular_data_corte(referencia)
        arvores = self._repository.carregar_arvores_elegiveis(data_corte)
        carregamento_ids = [arvore.carregamento_id for arvore in arvores]

        todas_chaves = sorted({chave for arvore in arvores for chave in arvore.chaves_nfe})
        documentos_xml = self._repository.carregar_documentos_xml_por_chaves(todas_chaves)
        orfaos = tuple(self._repository.detectar_orfaos(data_corte, carregamento_ids))

        pacotes = tuple(
            self._montar_pacote_unitario(arvore, documentos_xml) for arvore in arvores
        )
        resumo = self._montar_resumo(pacotes, orfaos)
        registros = sum(p.total_registros for p in pacotes)
        duracao_ms = (time.perf_counter() - inicio) * 1000.0

        return RelatorioSimulacaoRetencao(
            data_corte=data_corte,
            pacotes=pacotes,
            orfaos=orfaos,
            resumo=resumo,
            registros_analisados=registros,
            arquivos_pdf=sum(p.documentos_pdf for p in pacotes),
            arquivos_xml=sum(p.documentos_xml for p in pacotes),
            espaco_elegivel_bytes=sum(p.espaco_recuperavel_bytes for p in pacotes),
            duracao_ms=duracao_ms,
            simulacao=True,
        )

    def obter_pacote_por_carregamento(
        self,
        carregamento_id: int,
        referencia: date | None = None,
    ) -> PacoteRetencaoUnitario | None:
        data_corte = self._gestao_dados.calcular_data_corte(referencia)
        for arvore in self._repository.carregar_arvores_elegiveis(data_corte):
            if arvore.carregamento_id == carregamento_id:
                documentos_xml = self._repository.carregar_documentos_xml_por_chaves(list(arvore.chaves_nfe))
                return self._montar_pacote_unitario(arvore, documentos_xml)
        return None

    def _montar_pacote_unitario(
        self,
        arvore: ArvoreCarregamentoRaw,
        documentos_xml_map: dict,
    ) -> PacoteRetencaoUnitario:
        problemas: list[str] = []
        severidade_max = SaudePacote.SAUDAVEL

        pdfs = self._validar_pdfs(arvore, problemas)
        xmls = self._validar_xmls(arvore, documentos_xml_map, problemas)

        if not arvore.item_ids:
            problemas.append("Carregamento elegivel sem itens vinculados.")
            severidade_max = self._elevar_severidade(severidade_max, SaudePacote.ATENCAO)

        if arvore.chaves_nfe and not xmls:
            problemas.append("Itens possuem chave NFe, mas nenhum Documento XML foi localizado.")
            severidade_max = self._elevar_severidade(severidade_max, SaudePacote.ATENCAO)

        for chave in arvore.chaves_nfe:
            if chave not in documentos_xml_map:
                problemas.append(f"Documento XML nao encontrado para chave {chave}.")
                severidade_max = self._elevar_severidade(severidade_max, SaudePacote.ATENCAO)

        if arvore.documento_ids and not arvore.item_ids:
            problemas.append("PDF registrado sem itens de carregamento.")
            severidade_max = self._elevar_severidade(severidade_max, SaudePacote.CRITICO)

        if not arvore.historico_ids:
            problemas.append("Carregamento sem historico operacional registrado.")
            severidade_max = self._elevar_severidade(severidade_max, SaudePacote.ATENCAO)

        arquivos_encontrados = sum(1 for pdf in pdfs if pdf.existe_arquivo) + sum(1 for xml in xmls if xml.existe_arquivo)
        arquivos_ausentes = sum(1 for pdf in pdfs if not pdf.existe_arquivo) + sum(1 for xml in xmls if not xml.existe_arquivo)
        arquivos_esperados = len(pdfs) + len(xmls)
        integridade = 100.0 if arquivos_esperados == 0 else round((arquivos_encontrados / arquivos_esperados) * 100, 1)

        if arquivos_ausentes > 0:
            severidade_max = self._elevar_severidade(severidade_max, SaudePacote.ATENCAO)

        notas_distintas = len({chave for chave in arvore.chaves_nfe if chave} | {nf for nf in arvore.numeros_nf if nf})
        espaco_pdfs = sum(pdf.tamanho_bytes for pdf in pdfs if pdf.existe_arquivo)
        espaco_xmls = sum(xml.tamanho_bytes for xml in xmls if xml.existe_arquivo)

        espaco_sql = (
            _BYTES_ESTIMADOS_POR_REGISTRO["carregamento"]
            + len(arvore.item_ids) * _BYTES_ESTIMADOS_POR_REGISTRO["item_carregamento"]
            + len(arvore.documento_ids) * _BYTES_ESTIMADOS_POR_REGISTRO["documento"]
            + len(arvore.historico_ids) * _BYTES_ESTIMADOS_POR_REGISTRO["historico"]
            + len(arvore.evento_ids) * _BYTES_ESTIMADOS_POR_REGISTRO["evento"]
            + len(arvore.item_nota_fiscal_ids) * _BYTES_ESTIMADOS_POR_REGISTRO["item_nota_fiscal"]
        )

        apto = severidade_max != SaudePacote.CRITICO

        return PacoteRetencaoUnitario(
            carregamento_id=arvore.carregamento_id,
            numero_carregamento=arvore.numero_carregamento,
            data_carregamento=arvore.data,
            itens_carregamento=len(arvore.item_ids),
            notas_fiscais=notas_distintas,
            itens_nota_fiscal=len(arvore.item_nota_fiscal_ids),
            documentos_xml=len(xmls),
            documentos_pdf=len(pdfs),
            historicos=len(arvore.historico_ids),
            eventos=len(arvore.evento_ids),
            arquivos_encontrados=arquivos_encontrados,
            arquivos_ausentes=arquivos_ausentes,
            integridade_percentual=integridade,
            saude=severidade_max,
            apto_retencao=apto,
            problemas=tuple(problemas),
            pdfs=pdfs,
            xmls=xmls,
            espaco_pdfs_bytes=espaco_pdfs,
            espaco_xmls_bytes=espaco_xmls,
            espaco_metadados_sql_bytes=espaco_sql,
        )

    def _validar_pdfs(
        self,
        arvore: ArvoreCarregamentoRaw,
        problemas: list[str],
    ) -> tuple[DocumentoPdfValidacao, ...]:
        pdf_dir = self._resolver_pdf_dir()
        validacoes: list[DocumentoPdfValidacao] = []
        for index, caminho in enumerate(arvore.documento_caminhos):
            doc_id = arvore.documento_ids[index] if index < len(arvore.documento_ids) else 0
            tipo = arvore.documento_tipos[index] if index < len(arvore.documento_tipos) else "PDF"
            nome = arvore.documento_nomes[index] if index < len(arvore.documento_nomes) else Path(caminho).name
            digest = arvore.documento_hashes[index] if index < len(arvore.documento_hashes) else ""
            absolute = self._resolver_arquivo(pdf_dir, caminho)
            existe = absolute.is_file()
            tamanho = int(absolute.stat().st_size) if existe else 0
            if not existe:
                problemas.append(f"Arquivo PDF ausente: {caminho}")
            elif digest and digest != "0" * 64:
                try:
                    conteudo = absolute.read_bytes()
                    calculado = hashlib.sha256(conteudo).hexdigest()
                    if calculado != digest:
                        problemas.append(f"Hash divergente no PDF {nome}.")
                except OSError:
                    problemas.append(f"Nao foi possivel ler o PDF {nome}.")
            validacoes.append(
                DocumentoPdfValidacao(
                    documento_id=int(doc_id),
                    tipo=str(tipo),
                    caminho_arquivo=str(caminho),
                    nome_arquivo=str(nome),
                    hash_sha256=str(digest),
                    existe_arquivo=existe,
                    tamanho_bytes=tamanho,
                )
            )
        return tuple(validacoes)

    def _validar_xmls(
        self,
        arvore: ArvoreCarregamentoRaw,
        documentos_xml_map: dict,
        problemas: list[str],
    ) -> tuple[DocumentoXmlValidacao, ...]:
        xml_dir = self._resolver_xml_dir()
        validacoes: list[DocumentoXmlValidacao] = []
        for chave in arvore.chaves_nfe:
            registro = documentos_xml_map.get(chave)
            if registro is None:
                validacoes.append(
                    DocumentoXmlValidacao(
                        chave_nfe=chave,
                        numero_nf="",
                        documento_xml_id=None,
                        caminho_arquivo=None,
                        hash_sha256=None,
                        registro_ativo=False,
                        existe_arquivo=False,
                        tamanho_bytes=0,
                    )
                )
                continue
            caminho = str(registro.caminho_arquivo or "")
            absolute = self._resolver_arquivo(xml_dir, caminho)
            existe = absolute.is_file()
            tamanho = int(absolute.stat().st_size) if existe else int(registro.tamanho)
            if not registro.ativo:
                problemas.append(f"Documento XML inativo para chave {chave}.")
            if not existe:
                problemas.append(f"Arquivo XML ausente: {caminho or chave}")
            elif registro.hash_sha256:
                try:
                    conteudo = absolute.read_bytes()
                    calculado = hashlib.sha256(conteudo).hexdigest()
                    if calculado != registro.hash_sha256:
                        problemas.append(f"Hash divergente no XML da chave {chave}.")
                except OSError:
                    problemas.append(f"Nao foi possivel ler o XML da chave {chave}.")
            validacoes.append(
                DocumentoXmlValidacao(
                    chave_nfe=chave,
                    numero_nf=str(registro.numero_nf),
                    documento_xml_id=int(registro.id),
                    caminho_arquivo=caminho or None,
                    hash_sha256=str(registro.hash_sha256),
                    registro_ativo=bool(registro.ativo),
                    existe_arquivo=existe,
                    tamanho_bytes=tamanho,
                )
            )
        return tuple(validacoes)

    @staticmethod
    def _montar_resumo(
        pacotes: tuple[PacoteRetencaoUnitario, ...],
        orfaos: tuple[ProblemaIntegridade, ...],
    ) -> ResumoSaudeSimulacao:
        _ = orfaos
        total = len(pacotes)
        saudaveis = sum(1 for p in pacotes if p.saude == SaudePacote.SAUDAVEL)
        atencao = sum(1 for p in pacotes if p.saude == SaudePacote.ATENCAO)
        criticos = sum(1 for p in pacotes if p.saude == SaudePacote.CRITICO)
        integros = sum(1 for p in pacotes if p.apto_retencao and p.arquivos_ausentes == 0)

        arquivos_esperados = sum(len(p.pdfs) + len(p.xmls) for p in pacotes)
        arquivos_ok = sum(p.arquivos_encontrados for p in pacotes)
        integridade_geral = 100.0 if arquivos_esperados == 0 else round((arquivos_ok / arquivos_esperados) * 100, 1)

        todos_pdfs = all(pdf.existe_arquivo for p in pacotes for pdf in p.pdfs) if pacotes else True
        todos_xmls = all(xml.existe_arquivo for p in pacotes for xml in p.xmls) if pacotes else True

        return ResumoSaudeSimulacao(
            pacotes_elegiveis=total,
            pacotes_integros=integros,
            pacotes_com_inconsistencia=max(total - integros, 0),
            integridade_geral_percentual=integridade_geral,
            todos_pdfs_encontrados=todos_pdfs,
            todos_xmls_encontrados=todos_xmls,
            pacotes_saudaveis=saudaveis,
            pacotes_atencao=atencao,
            pacotes_criticos=criticos,
        )

    def _resolver_pdf_dir(self) -> Path:
        if self._pdf_storage_dir is not None:
            return self._pdf_storage_dir
        from infrastructure.database import get_pdf_storage_dir

        return get_pdf_storage_dir()

    def _resolver_xml_dir(self) -> Path:
        if self._xml_storage_dir is not None:
            return self._xml_storage_dir
        from infrastructure.database import get_xml_storage_dir

        return get_xml_storage_dir()

    @staticmethod
    def _resolver_arquivo(base_dir: Path, relative_path: str) -> Path:
        candidate = Path(relative_path)
        if candidate.is_absolute():
            return candidate
        return base_dir / relative_path

    @staticmethod
    def _elevar_severidade(atual: SaudePacote, nova: SaudePacote) -> SaudePacote:
        ordem = {SaudePacote.SAUDAVEL: 0, SaudePacote.ATENCAO: 1, SaudePacote.CRITICO: 2}
        return nova if ordem[nova] > ordem[atual] else atual
