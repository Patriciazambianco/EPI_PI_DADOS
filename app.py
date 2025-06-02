import streamlit as st
import pandas as pd

@st.cache_data
def carregar_dados():
    # (Aqui vai a função completa que te passei antes)
    # Vou colocar só um exemplo aqui por simplicidade
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    col_tec = [col for col in df.columns if 'TECNICO' in col.upper()]
    col_prod = [col for col in df.columns if 'PRODUTO' in col.upper()]
    col_data = [col for col in df.columns if 'INSPECAO' in col.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Arquivo inválido. Falta coluna TECNICO, PRODUTO ou INSPECAO.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO',
        data_col: 'DATA_INSPECAO'
    }, inplace=True)

    df['DATA_INSPECAO'] = pd.to_datetime(df['DATA_INSPECAO'], errors='coerce')
    df['CHAVE'] = df['TECNICO'].astype(str).str.strip() + "|" + df['PRODUTO'].astype(str).str.strip()

    df_com_data = df.dropna(subset=['DATA_INSPECAO']).copy()
    df_sem_data = df[df['DATA_INSPECAO'].isna()].copy()

    df_com_data.sort_values('DATA_INSPECAO', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    df_final = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    st.write("### Data final carregada (Últimas inspeções + pendentes)")
    return df_final.drop(columns=['CHAVE'])


def main():
    df = carregar_dados()
    if df.empty:
        st.warning("⚠️ DataFrame está vazio, verifique o arquivo.")
    else:
        st.dataframe(df)  # Exibe a tabela na tela

if __name__ == "__main__":
    main()
