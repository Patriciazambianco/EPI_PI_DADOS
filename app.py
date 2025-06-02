import streamlit as st
import pandas as pd
import io
import plotly.express as px

@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Encontra as colunas
    col_tec = [c for c in df.columns if 'TECNICO' in c.upper()]
    col_prod = [c for c in df.columns if 'PRODUTO' in c.upper()]
    col_data = [c for c in df.columns if 'INSPECAO' in c.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Arquivo deve conter colunas de Técnico, Produto e Inspeção.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df.rename(columns={'GERENTE': 'GERENTE_IMEDIATO',
                       'COORDENADOR': 'COORDENADOR_IMEDIATO',
                       'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'}, inplace=True)

    df['Data_Inspecao'] = pd.to_datetime(df[data_col], errors='coerce')

    # Criar chave única Técnico+Produto
    df['CHAVE'] = df[tecnico_col].astype(str).str.strip() + '|' + df[produto_col].astype(str).str.strip()

    # Separar linhas com e sem data
    df_com_data = df.dropna(subset=['Data_Inspecao']).copy()
    df_sem_data = df[df['Data_Inspecao'].isna()].copy()

    # Pega última inspeção por chave
    df_com_data.sort_values('Data_Inspecao', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    # Pendentes são os sem inspeção que não estão nos últimos
    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    # Junta tudo
    df_resultado = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    # Padroniza nomes
    df_resultado.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO'
    }, inplace=True)

    # Status check list padronizado
    if 'STATUS CHECK LIST' in df_resultado.columns:
        df_resultado['STATUS CHECK LIST'] = df_resultado['STATUS CHECK LIST'].str.upper()

    hoje = pd.Timestamp.now().normalize()
    df_resultado['Dias_Sem_Inspecao'] = (hoje - df_resultado['Data_Inspecao']).dt.days.fillna(-1)
    df_resultado['Vencido'] = df_resultado['Dias_Sem_Inspecao'] > 180

    return df_resultado.drop(columns=['CHAVE'])

def calcula_percentuais(df, grupo_col):
    resumo = df.groupby(grupo_col).agg(
        total=('STATUS CHECK LIST', 'count'),
        ok=('STATUS CHECK LIST', lambda x: (x == 'OK').sum()),
        pendente=('STATUS CHECK LIST', lambda x: (x != 'OK').sum())
    ).reset_index()

    resumo['% OK'] = (resumo['ok'] / resumo['total'] * 100).round(1)
    resumo['% Pendente'] = (resumo['pendente'] / resumo['total'] * 100).round(1)
    return resumo

def gera_grafico_barra(df, grupo_col, titulo):
    resumo = calcula_percentuais(df, grupo_col)
    fig = px.bar(
        resumo,
        x=grupo_col,
        y=['% OK', '% Pendente'],
        barmode='group',
        title=f'% EPI OK e Pendentes por {titulo}',
        labels={grupo_col: titulo, 'value': '%'},
        text_auto=True,
        height=400
    )
    return fig

def exibe_cards(df):
    resumo_gerente = calcula_percentuais(df, 'GERENTE_IMEDIATO')
    resumo_coordenador = calcula_percentuais(df, 'COORDENADOR_IMEDIATO')

    st.markdown("### Indicadores por Gerente")
    for _, row in resumo_gerente.iterrows():
        st.metric(label=f"{row['GERENTE_IMEDIATO']}",
                  value=f"{row['% OK']}% OK",
                  delta=f"{row['% Pendente']}% Pendentes")

    st.markdown("### Indicadores por Coordenador")
    for _, row in resumo_coordenador.iterrows():
        st.metric(label=f"{row['COORDENADOR_IMEDIATO']}",
                  value=f"{row['% OK']}% OK",
                  delta=f"{row['% Pendente']}% Pendentes")

def botao_exportar(df_pendentes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_pendentes.to_excel(writer, index=False, sheet_name='Pendentes')
    processed_data = output.getvalue()
    st.download_button(
        label="📥 Exportar Pendentes para Excel",
        data=processed_data,
        file_name="epi_pendentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def main():
    st.title("Dashboard EPI - Última Inspeção + Pendentes")

    df = carregar_dados()
    if df.empty:
        return

    # Mostrar cards resumidos
    exibe_cards(df)

    # Mostrar gráficos
    st.plotly_chart(gera_grafico_barra(df, 'GERENTE_IMEDIATO', 'Gerente'))
    st.plotly_chart(gera_grafico_barra(df, 'COORDENADOR_IMEDIATO', 'Coordenador'))

    # Tabela completa
    st.markdown("### Tabela Completa de Inspeções")
    st.dataframe(df)

    # Exportar pendentes
    df_pendentes = df[df['STATUS CHECK LIST'] != 'OK']
    botao_exportar(df_pendentes)

if __name__ == "__main__":
    main()
