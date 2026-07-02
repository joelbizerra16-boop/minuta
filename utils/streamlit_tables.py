from __future__ import annotations

import pandas as pd
import streamlit as st


def build_table_column_config(dataframe: pd.DataFrame) -> dict[str, object]:
    """Configuracao responsiva de colunas para st.dataframe (sem largura fixa em px)."""
    descricao_width = "large" if "Produto" in dataframe.columns else "medium"
    configs: dict[str, object] = {
        "Seq": st.column_config.NumberColumn("Seq", format="%d", width="small"),
        "NF": st.column_config.TextColumn("NF", width="small"),
        "cProd": st.column_config.TextColumn("cProd", width="small"),
        "Produto": st.column_config.TextColumn("Produto", width="small"),
        "Descricao": st.column_config.TextColumn("Descricao", width=descricao_width),
        "Qtd": st.column_config.NumberColumn("Qtd", format="%.4f", width="small"),
        "Quantidade": st.column_config.NumberColumn("Quantidade", format="%.4f", width="small"),
        "Unidade": st.column_config.TextColumn("UN", width="small"),
        "Peso": st.column_config.NumberColumn("Peso", format="%.3f", width="small"),
        "Destinatario": st.column_config.TextColumn("Destinatario", width="medium"),
        "ROTA": st.column_config.TextColumn("ROTA", width="medium"),
        "Rota": st.column_config.TextColumn("Rota", width="medium"),
        "Status": st.column_config.TextColumn("Status", width="small"),
        "Data": st.column_config.TextColumn("Data", width="small"),
        "Carregamento": st.column_config.TextColumn("Carregamento", width="small"),
        "Motorista": st.column_config.TextColumn("Motorista", width="medium"),
        "Placa": st.column_config.TextColumn("Placa", width="small"),
        "Usuario": st.column_config.TextColumn("Usuario", width="small"),
        "Modalidade": st.column_config.TextColumn("Modalidade", width="small"),
    }
    return {column: configs[column] for column in dataframe.columns if column in configs}


def build_consulta_listagem_column_config(dataframe: pd.DataFrame) -> dict[str, object]:
    """Larguras da ListView da Consulta de NFs Carregadas."""
    configs: dict[str, object] = {
        "Data": st.column_config.TextColumn("Data", width=102),
        "Carregamento": st.column_config.TextColumn("Carregamento", width=128),
        "NF": st.column_config.TextColumn("NF", width=92),
        "Produto": st.column_config.TextColumn("Produto", width=78),
        "Descricao": st.column_config.TextColumn("Descricao", width=520),
        "Quantidade": st.column_config.NumberColumn("Quantidade", format="%.4f", width=98),
        "Peso": st.column_config.NumberColumn("Peso", format="%.3f", width=88),
        "Destinatario": st.column_config.TextColumn("Destinatario", width=320),
        "Rota": st.column_config.TextColumn("Rota", width=132),
        "Motorista": st.column_config.TextColumn("Motorista", width=168),
        "Placa": st.column_config.TextColumn("Placa", width=84),
        "Usuario": st.column_config.TextColumn("Usuario", width=92),
        "Modalidade": st.column_config.TextColumn("Modalidade", width=108),
        "Status": st.column_config.TextColumn("Status", width=118),
    }
    return {column: configs[column] for column in dataframe.columns if column in configs}


def build_auditoria_nf_expansion_column_config(dataframe: pd.DataFrame) -> dict[str, object]:
    """Larguras do extrato operacional exibido ao expandir uma NF."""
    configs: dict[str, object] = {
        "Etapa": st.column_config.TextColumn("Etapa", width=130),
        "Veiculo": st.column_config.TextColumn("Veiculo", width=110),
        "Placa": st.column_config.TextColumn("Placa", width=90),
        "Motorista": st.column_config.TextColumn("Motorista", width=180),
        "Data": st.column_config.TextColumn("Data", width=100),
        "Hora": st.column_config.TextColumn("Hora", width=80),
        "Usuario": st.column_config.TextColumn("Usuario", width=100),
        "Carregamento": st.column_config.TextColumn("Carregamento", width=110),
        "IdCarga": st.column_config.TextColumn("Id Carga", width=90),
        "Rota": st.column_config.TextColumn("Rota", width=120),
        "Observacao": st.column_config.TextColumn("Observacao", width=280),
    }
    return {column: configs[column] for column in dataframe.columns if column in configs}


def build_auditoria_nf_history_column_config(dataframe: pd.DataFrame) -> dict[str, object]:
    """Alias legado para compatibilidade."""
    return build_auditoria_nf_expansion_column_config(dataframe)
