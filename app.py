import streamlit as st
import pandas as pd
import io
import plotly.express as px

@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    col_tec = [col for col in df.columns if 'TECNICO' in col.upper()]
    col_prod = [col for col in df.columns if 'PRODUTO' in col.upper()]
    col_data = [col for col in df.columns if 'INSPECAO' in col.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Verifique se o arquivo contém colunas de TÉCNICO, PRODUTO e INSPEÇÃO.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df.rename(columns={
        'GERENTE': 'GERENTE_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    df['Data_Inspecao'] = pd.to_datetime(df[data_col], errors='coerce')

    # Cria uma chave única TECNICO + PRODUTO
    df['CHAVE'] = df[tecnico_col].astype(str).str.strip() + "|" + df[produto_col].astype(str).str.strip()

    # Separar com e sem data
    df_com_data = df.dropna(subset=['Data_Inspecao']).copy()
    df_sem_data = df[df['Data_Inspecao'].isna()].copy()

    # Pegar a última inspeção por chave (técnico + produto)
    df_com_data.sort_values('Data_Inspecao', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    # Filtrar pendentes SEM inspeção (sem estar entre os últimos)
    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    # Junta os dois (última inspeção + pendentes nunca inspecionados)
    df_resultado = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    # Renomeia para padronizar
    df_resultado.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO'
    }, inplace=True)

    # Ajusta status e vencimento
    if 'STATUS CHECK LIST' in df_resultado.columns:
        df_resultado['STATUS CHECK LIST'] = df_resultado['STATUS CHECK LIST'].str.upper()

    hoje = pd.Timestamp.now().normalize()
    df_resultado['Dias_Sem_Inspecao'] = (hoje - df_resultado['Data_Inspecao']).dt.days
    df_resultado['Vencido'] = df_resultado['Dias_Sem_Inspecao'] > 180

    return df_resultado.drop(columns=['CHAVE'])


def calcula_percentuais(df, grupo_col):
    resumo = df.groupby(grupo_col).agg({
        'STATUS CHECK LIST': ['count', lambda x: (x == 'OK').sum(), lambda x: (x != 'OK').sum()]
    })
    resumo.columns = ['total', 'ok', 'pendente']
    resumo = resumo.reset_index()
    resumo['% OK'] = 100 * resumo['ok'] / resumo['total']
    resumo['% Pendente'] = 100 * resumo['pendente'] / resumo['total']
    return resumo


def exibe_cards(df):
    st.header("Indicadores de Inspeção de EPI")

    perc_gerente = calcula_percentuais(df, 'GERENTE_IMEDIATO')
    perc_coordenador = calcula_percentuais(df, 'COORDENADOR_IMEDIATO') if 'COORDENADOR_IMEDIATO' in df.columns else None

    st.subheader("Por Gerente")
    cols = st.columns(len(perc_gerente))
    for i, row in perc_gerente.iterrows():
        with cols[i]:
            st.markdown(f"### {row['GERENTE_IMEDIATO']}")
            st.metric(label="% EPI OK", value=f"{row['% OK']:.1f}%")
            st.metric(label="% EPI Pendente", value=f"{row['% Pendente']:.1f}%")


    if perc_coordenador is not None:
        st.subheader("Por Coordenador")
        cols = st.columns(len(perc_coordenador))
        for i, row in perc_coordenador.iterrows():
            with cols[i]:
                st.markdown(f"### {row['COORDENADOR_IMEDIATO']}")
                st.metric(label="% EPI OK", value=f"{row['% OK']:.1f}%")
                st.metric(label="% EPI Pendente", value=f"{row['% Pendente']:.1f}%")


def gera_grafico_barra(df, grupo_col, titulo):
    resumo = calcula_percentuais(df, grupo_col)
    fig = px.bar(resumo, x=grupo_col, y=['% OK', '% Pendente'],
                 title=f"Percentual de EPI OK/Pendente por {titulo}",
                 labels={grupo_col: titulo, "value": "Percentual (%)"},
                 barmode='group',
                 height=400)
    st.plotly_chart(fig, use_container_width=True)


def botao_exportar(df_pendentes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_pendentes.to_excel(writer, index=False, sheet_name='Pendentes')
        writer.save()
    processed_data = output.getvalue()
    st.download_button(
        label="📥 Exportar Pendentes para Excel",
        data=processed_data,
        file_name="pendentes_epi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def main():
    st.title("Dashboard de Inspeção de EPIs")

    df = carregar_dados()
    if df.empty:
        return

    exibe_cards(df)

    st.markdown("---")

    gera_grafico_barra(df, 'GERENTE_IMEDIATO', 'Gerente')

    if 'COORDENADOR_IMEDIATO' in df.columns:
        gera_grafico_barra(df, 'COORDENADOR_IMEDIATO', 'Coordenador')

    st.markdown("---")
    st.header("Tabela Completa")
    st.dataframe(df)

    df_pendentes = df[df['STATUS CHECK LIST'] != 'OK']
    botao_exportar(df_pendentes)


if __name__ == "__main__":
    main()
