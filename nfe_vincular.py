import os
import datetime
import xml.etree.ElementTree as ET
from pathlib import Path

import streamlit as st


def validar_nfe(caminho_xml: str) -> bool:
    """
    Valida um arquivo XML de NF-e.

    Regras:
    - Se for um arquivo <procEventoNFe> contendo <descEvento>Cancelamento</descEvento>, retorna False
    - Se não houver nenhum elemento <cStat> com texto '100', retorna False
    - Caso contrário, retorna True

    Recebe o caminho do XML e usa xml.etree.ElementTree para parse.
    """
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
    except Exception:
        return False

    # Normalizar nomes de tags (pegar só a parte local sem namespace)
    def local_name(elem):
        if '}' in elem.tag:
            return elem.tag.split('}', 1)[1]
        return elem.tag

    # Detectar evento de cancelamento: root é procEventoNFe e existe descEvento == 'Cancelamento'
    if local_name(root).lower() == 'proceventonfe':
        for de in root.findall('.//'):
            if local_name(de).lower() == 'descevento' and (de.text or '').strip().lower() == 'cancelamento':
                return False

    # Procurar por qualquer elemento cStat e aceitar somente se existir pelo menos um com '100'
    cstats = [el.text.strip() for el in root.findall('.//') if local_name(el).lower() == 'cstat' and el.text]
    if cstats:
        # Se existir algum cStat==100, considerar válido
        if any(cs == '100' for cs in cstats):
            return True
        return False

    # Se não encontrou cStat, considerar inválido
    return False


def find_xml_by_chave(chave: str, search_dir: str = 'data') -> str | None:
    """
    Procura recursivamente por arquivos .xml em `search_dir` cujo conteúdo contenha a chave da NF-e.

    Estratégia heurística:
    - Verifica elementos `<chNFe>` cujo texto seja a chave
    - Verifica atributos `Id` como 'NFe{chave}'
    - Como fallback, procura qualquer ocorrência textual exata da chave no XML
    """
    chave = chave.strip()
    if not chave or len(chave) != 44 or not chave.isdigit():
        return None

    base = Path(search_dir)
    if not base.exists():
        return None

    for p in base.rglob('*.xml'):
        try:
            tree = ET.parse(p)
            root = tree.getroot()
        except Exception:
            # arquivo inválido -> pular
            continue

        # helper to local name
        def local_name(elem):
            if '}' in elem.tag:
                return elem.tag.split('}', 1)[1]
            return elem.tag

        found = False

        # 1) chNFe element exact match
        for el in root.findall('.//'):
            if local_name(el).lower() == 'chnfe' and el.text and el.text.strip() == chave:
                found = True
                break

        if found:
            return str(p)

        # 2) attribute Id on infNFe (like Id="NFe{chave}")
        for el in root.iter():
            if local_name(el).lower().endswith('infnfe'):
                id_attr = el.get('Id') or el.get('id')
                if id_attr and id_attr.endswith(chave):
                    return str(p)

        # 3) fallback: any text node equals chave
        text_found = any((el.text or '').strip() == chave for el in root.findall('.//'))
        if text_found:
            return str(p)

    return None


def generate_lote_id() -> str:
    today = datetime.date.today().strftime('%Y%m%d')
    # Simple fixed suffix; in a real app you might increment
    return f'LOTE-{today}-001'


def main():
    st.title('Vinculação de NF-e por Chave')

    if 'lote_id' not in st.session_state:
        st.session_state.lote_id = generate_lote_id()
    if 'lote_nfes' not in st.session_state:
        st.session_state.lote_nfes = []

    st.markdown(f"**Lote atual:** {st.session_state.lote_id}")

    chave = st.text_input('Bipar ou digitar chave da NF (44 dígitos)', key='chave_input')
    buscar = st.button('Buscar')

    if buscar:
        chave_val = (chave or '').strip()
        if not chave_val or len(chave_val) != 44 or not chave_val.isdigit():
            st.error('Chave inválida. Deve conter 44 dígitos numéricos.')
        else:
            xml_path = find_xml_by_chave(chave_val)
            if not xml_path:
                st.error('XML não encontrado localmente para esta chave.')
            else:
                is_valid = validar_nfe(xml_path)
                if not is_valid:
                    st.error('Nota cancelada ou inválida — não pode ser adicionada ao lote')
                else:
                    # verifica duplicidade
                    if chave_val in st.session_state.lote_nfes:
                        st.warning(f'NF já vinculada ao lote {st.session_state.lote_id}')
                    else:
                        st.session_state.lote_nfes.append(chave_val)
                        st.success(f'NF vinculada com sucesso ao lote {st.session_state.lote_id}')

    if st.session_state.lote_nfes:
        st.subheader('NFs vinculadas no lote atual')
        for n in st.session_state.lote_nfes:
            st.write('-', n)


if __name__ == '__main__':
    main()
