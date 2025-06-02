import streamlit as st
import pandas as pd

# Testa se tem o decorator cache_data, senão usa cache (fallback)
try:
    cache_decorator = st.cache_data
except AttributeError:
    cache_decorator = st.cache

@cache_decorator
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    col_tec = [col for col in df.columns if 'TECNICO' in col.upper()]
    col_prod = [col for col in df.columns if 'PRODUTO' in col.upper()]
    col_data = [col for col in df.columns if 'INSPECAO' in col.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Arquivo precisa ter colunas TECNICO, PRODUTO e INSPEÇÃO.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df.rename(columns={
        'GERENTE': 'GERENTE_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    df['Data_Inspecao'] = pd.to_datetime(df[data_col], errors='coerce')
    df['CHAVE'] = df[tecnico_col].astype(str).str.strip() + "|" + df[produto_col].astype(str).str.strip()

    df_com_data = df.dropna(subset=['Data_Inspecao']).copy()
    df_sem_data = df[df['Data_Inspecao'].isna()].copy()

    df_com_data.sort_values('Data_Inspecao', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    df_resultado = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    df_resultado.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO'
    }, inplace=True)

    if 'STATUS CHECK LIST' in df_resultado.columns:
        df_resultado['STATUS CHECK LIST'] = df_resultado['STATUS CHECK LIST'].str.upper()

    hoje = pd.Timestamp.now().normalize()
    df_resultado['Dias_Sem_Inspecao'] = (hoje - df_resultado['Data_Inspecao']).dt.days
    df_resultado['Vencido'] = df_resultado['Dias_Sem_Inspecao'] > 180

    return df_resultado.drop(columns=['CHAVE'])

def main():
    st.title("Painel EPI")
    df = carregar_dados()
    if not df.empty:
        st.dataframe(df)

if __name__ == "__main__":
    main()
