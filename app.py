import streamlit as st
import pandas as pd
import plotly.express as px

# Sua função de carregar dados (ajuste conforme seu arquivo)
@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Padronizando nomes das colunas importantes
    df.rename(columns={
        'GERENTE': 'GERENTE_IMEDIATO',
        'COORDENADOR': 'COORDENADOR_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    # Ajusta status para caixa alta
    if 'STATUS CHECK LIST' in df.columns:
        df['STATUS CHECK LIST'] = df['STATUS CHECK LIST'].str.upper()

    return df

def calcula_percentuais(df, grupo_col):
    resumo = df.groupby(grupo_col).agg(
        total = ('STATUS CHECK LIST', 'count'),
        ok = (lambda x: (x == 'OK').sum()),
        pendente = (lambda x: (x != 'OK').sum())
    )
    resumo['% OK'] = (resumo['ok'] / resumo['total'] * 100).round(2)
    resumo['% Pendente'] = (resumo['pendente'] / resumo['total'] * 100).round(2)
    resumo = resumo.reset_index()
    return resumo

def exibe_cards(df):
    st.header("Indicadores de EPI por Gerente e Coordenador")

    # Filtro por gerente
    gerentes = df['GERENTE_IMEDIATO'].dropna().unique()
    gerente_selecionado = st.selectbox("Selecione o Gerente", options=sorted(gerentes))

    df_gerente = df[df['GERENTE_IMEDIATO'] == gerente_selecionado]

    # Cards por Gerente
    perc_gerente = calcula_percentuais(df_gerente, 'GERENTE_IMEDIATO')
    for _, row in perc_gerente.iterrows():
        col1, col2 = st.columns(2)
        col1.metric(label=f"{row['GERENTE_IMEDIATO']} - % EPI OK", value=f"{row['% OK']}%")
        col2.metric(label=f"{row['GERENTE_IMEDIATO']} - % EPI Pendente", value=f"{row['% Pendente']}%")

    # Cards e gráfico por Coordenador
    if 'COORDENADOR_IMEDIATO' in df_gerente.columns:
        perc_coord = calcula_percentuais(df_gerente, 'COORDENADOR_IMEDIATO')
        st.subheader(f"Por Coordenador do Gerente {gerente_selecionado}")
        for _, row in perc_coord.iterrows():
            col1, col2 = st.columns(2)
            col1.metric(label=f"{row['COORDENADOR_IMEDIATO']} - % EPI OK", value=f"{row['% OK']}%")
            col2.metric(label=f"{row['COORDENADOR_IMEDIATO']} - % EPI Pendente", value=f"{row['% Pendente']}%")

        # Gráfico barras lado a lado para coordenadores
        fig = px.bar(
            perc_coord,
            x='COORDENADOR_IMEDIATO',
            y=['% OK', '% Pendente'],
            barmode='group',
            labels={'value': 'Percentual (%)', 'COORDENADOR_IMEDIATO': 'Coordenador'},
            title=f'Percentual de EPI OK e Pendente por Coordenador do Gerente {gerente_selecionado}'
        )
        st.plotly_chart(fig, use_container_width=True)

def main():
    st.title("Dashboard de EPI - Status por Gerente e Coordenador")
    df = carregar_dados()
    if df.empty:
        st.error("Erro ao carregar dados ou dados vazios.")
        return

    exibe_cards(df)

if __name__ == "__main__":
    main()
