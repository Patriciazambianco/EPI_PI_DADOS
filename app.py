import pandas as pd
import streamlit as st

@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Encontra as colunas relevantes (case insensitive)
    col_tec = [col for col in df.columns if 'TECNICO' in col.upper()]
    col_prod = [col for col in df.columns if 'PRODUTO' in col.upper()]
    col_data = [col for col in df.columns if 'INSPECAO' in col.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Verifique se o arquivo contém colunas de TÉCNICO, PRODUTO e INSPEÇÃO.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    # Padroniza nomes para facilitar depois
    df.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO',
        data_col: 'DATA_INSPECAO',
        'GERENTE': 'GERENTE_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    # Converte a data da inspeção pra datetime
    df['DATA_INSPECAO'] = pd.to_datetime(df['DATA_INSPECAO'], errors='coerce')

    # Cria chave única TECNICO + PRODUTO
    df['CHAVE'] = df['TECNICO'].astype(str).str.strip() + "|" + df['PRODUTO'].astype(str).str.strip()

    # Separa os com data e sem data (pendentes)
    df_com_data = df.dropna(subset=['DATA_INSPECAO']).copy()
    df_sem_data = df[df['DATA_INSPECAO'].isna()].copy()

    # Ordena do mais recente para o mais antigo e remove duplicados pela chave, mantendo a última inspeção
    df_com_data.sort_values('DATA_INSPECAO', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    # Remove do pendentes quem já tem inspeção registrada (pra não duplicar)
    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    # Junta os dados: últimos inspecionados + pendentes sem inspeção
    df_final = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    # Ajusta a coluna status para maiúsculas (se existir)
    if 'STATUS CHECK LIST' in df_final.columns:
        df_final['STATUS CHECK LIST'] = df_final['STATUS CHECK LIST'].astype(str).str.upper()

    # Calcula dias sem inspeção e se está vencido (mais de 180 dias)
    hoje = pd.Timestamp.now().normalize()
    df_final['DIAS_SEM_INSPECAO'] = (hoje - df_final['DATA_INSPECAO']).dt.days
    df_final['DIAS_SEM_INSPECAO'] = df_final['DIAS_SEM_INSPECAO'].fillna(-1)  # -1 pra pendentes sem data
    df_final['VENCIDO'] = df_final['DIAS_SEM_INSPECAO'] > 180

    # Remove coluna auxiliar CHAVE
    df_final.drop(columns=['CHAVE'], inplace=True)

    return df_final
