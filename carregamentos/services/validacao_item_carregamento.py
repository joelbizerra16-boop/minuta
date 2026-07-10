from __future__ import annotations

from dataclasses import dataclass

from carregamentos.models.carregamento import Carregamento, CarregamentoItem, normalize_chave_nfe, normalize_nf_number


@dataclass(frozen=True)
class ChaveUnicaItemCarregamento:
    """Alinhada à UNIQUE (carregamento_id, numero_nf, codigo_produto, sequencia)."""

    numero_nf: str
    codigo_produto: str
    sequencia: int


def _nf_key(item: CarregamentoItem) -> str:
    return normalize_nf_number(item.nf) or normalize_chave_nfe(item.chave_nfe)


def extrair_chaves_itens_existentes(carregamento: Carregamento) -> set[ChaveUnicaItemCarregamento]:
    chaves: set[ChaveUnicaItemCarregamento] = set()
    for index, item in enumerate(carregamento.itens, start=1):
        nf_norm = _nf_key(item)
        cprod = str(item.cprod or "").strip()
        if not nf_norm or not cprod:
            continue
        chaves.add(
            ChaveUnicaItemCarregamento(
                numero_nf=nf_norm,
                codigo_produto=cprod,
                sequencia=index,
            )
        )
    return chaves


def extrair_chaves_logicas_existentes(carregamento: Carregamento) -> set[tuple[str, str, str]]:
    """Chave lógica (nf, chave_nfe, cprod) para detecção de duplicidade operacional."""
    return {
        (
            normalize_nf_number(item.nf),
            normalize_chave_nfe(item.chave_nfe),
            str(item.cprod or "").strip(),
        )
        for item in carregamento.itens
        if _nf_key(item) and str(item.cprod or "").strip()
    }


def item_ja_existe_no_carregamento(carregamento: Carregamento, item: CarregamentoItem) -> bool:
    chave = (
        normalize_nf_number(item.nf),
        normalize_chave_nfe(item.chave_nfe),
        str(item.cprod or "").strip(),
    )
    return chave in extrair_chaves_logicas_existentes(carregamento)


def validar_lista_final_itens(
    carregamento: Carregamento,
    itens_finais: list[CarregamentoItem],
) -> tuple[bool, list[str]]:
    """
    Simula a atribuição de sequencia 1..N (como no repository) e verifica violação da UNIQUE.
    Retorna (valido, erros).
    """
    erros: list[str] = []
    chaves_vistas: set[ChaveUnicaItemCarregamento] = set()

    for index, item in enumerate(itens_finais, start=1):
        nf_norm = _nf_key(item)
        cprod = str(item.cprod or "").strip()
        if not nf_norm:
            erros.append(f"Item na posicao {index} sem NF identificavel.")
            continue
        if not cprod:
            erros.append(f"Item NF {item.nf} na posicao {index} sem codigo de produto.")
            continue
        chave = ChaveUnicaItemCarregamento(
            numero_nf=nf_norm,
            codigo_produto=cprod,
            sequencia=index,
        )
        if chave in chaves_vistas:
            erros.append(
                f"Duplicidade detectada: NF {item.nf}, produto {cprod}, sequencia {index}."
            )
        chaves_vistas.add(chave)

    return len(erros) == 0, erros


def filtrar_itens_novos_para_insercao(
    carregamento: Carregamento,
    candidatos: list[CarregamentoItem],
) -> tuple[list[CarregamentoItem], list[CarregamentoItem]]:
    """
    Separa itens que podem ser inseridos dos que já existem (duplicidade operacional).
    """
    chaves_existentes = extrair_chaves_logicas_existentes(carregamento)
    novos: list[CarregamentoItem] = []
    duplicados: list[CarregamentoItem] = []
    chaves_batch: set[tuple[str, str, str]] = set()

    for item in candidatos:
        chave = (
            normalize_nf_number(item.nf),
            normalize_chave_nfe(item.chave_nfe),
            str(item.cprod or "").strip(),
        )
        if not chave[0] or not chave[2]:
            duplicados.append(item)
            continue
        if chave in chaves_existentes or chave in chaves_batch:
            duplicados.append(item)
            continue
        chaves_batch.add(chave)
        novos.append(item)

    return novos, duplicados


def montar_lista_itens_pos_complementacao(
    carregamento: Carregamento,
    itens_novos: list[CarregamentoItem],
) -> tuple[list[CarregamentoItem], list[str]]:
    """Monta lista final e valida antes da persistência."""
    itens_filtrados, duplicados = filtrar_itens_novos_para_insercao(carregamento, itens_novos)
    lista_final = list(carregamento.itens) + itens_filtrados
    valido, erros = validar_lista_final_itens(carregamento, lista_final)
    if duplicados:
        for item in duplicados:
            erros.append(
                f"Item NF {item.nf} produto {item.cprod} ja pertence ao carregamento — reutilizacao."
            )
    if not valido:
        return lista_final, erros
    return lista_final, erros
