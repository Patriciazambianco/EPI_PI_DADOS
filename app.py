iimport streamlit as st
import pandas as pd

# Função para carregar e preparar os dados
@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Ajustar nomes de colunas conforme sua planilha
    df.rename(columns={
        'GERENTE': 'GERENTE_IMEDIATO',
        'COORDENADOR': 'COORDENADOR_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    # Padronizar STATUS para maiúsculas
    df['STATUS CHECK LIST'] = df['STATUS CHECK LIST'].str.upper()

    return df

# Função para calcular percentuais por grupo (gerente ou coordenador)
def calcula_percentuais(df, grupo_col):
    resumo = df.groupby(grupo_col).agg(
        total = ('STATUS CHECK LIST', 'count'),
        ok = ('STATUS CHECK LIST', lambda x: (x == 'OK').sum()),
        pendente = ('STATUS CHECK LIST', lambda x: (x != 'OK').sum())
    )
    resumo['% OK'] = (resumo['ok'] / resumo['total'] * 100).round(2)
    resumo['% Pendente'] = (resumo['pendente'] / resumo['total'] * 100).round(2)
    resumo = resumo.reset_index()
    return resumo

# Função para exibir cards no Streamlit
def exibe_cards(df):
    st.subheader("Indicadores por Gerente")
    resumo_gerente = calcula_percentuais(df, 'GERENTE_IMEDIATO')

    for _, row in resumo_gerente.iterrows():
        st.markdown(f"### {row['GERENTE_IMEDIATO']}")
        col1, col2 = st.columns(2)
        col1.metric(label="% EPI OK", value=f"{row['% OK']}%")
        col2.metric(label="% EPI Pendente", value=f"{row['% Pendente']}%")

    st.markdown("---")
    st.subheader("Indicadores por Coordenador")
    resumo_coord = calcula_percentuais(df, 'COORDENADOR_IMEDIATO')

    for _, row in resumo_coord.iterrows():
        st.markdown(f"### {row['COORDENADOR_IMEDIATO']}")
        col1, col2 = st.columns(2)
        col1.metric(label="% EPI OK", value=f"{row['% OK']}%")
        col2.metric(label="% EPI Pendente", value=f"{row['% Pendente']}%")

def main():
    st.title("Dashboard EPI - Status por Gerente e Coordenador")

    df = carregar_dados()

    if df.empty:
        st.error("❌ Dados não encontrados ou arquivo vazio.")
        return

    exibe_cards(df)

if __name__ == "__main__":
    main()
