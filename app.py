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
        st.error("❌ Arquivo inválido. Falta coluna TECNICO, PRODUTO ou INSPECAO.")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO',
        data_col: 'DATA_INSPECAO',
        'GERENTE': 'GERENTE_IMEDIATO',
        'COORDENADOR': 'COORDENADOR_IMEDIATO',
        'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'
    }, inplace=True)

    df['DATA_INSPECAO'] = pd.to_datetime(df['DATA_INSPECAO'], errors='coerce')

    # Chave única TECNICO+PRODUTO
    df['CHAVE'] = df['TECNICO'].astype(str).str.strip() + "|" + df['PRODUTO'].astype(str).str.strip()

    # Separar com e sem data
    df_com_data = df.dropna(subset=['DATA_INSPECAO']).copy()
    df_sem_data = df[df['DATA_INSPECAO'].isna()].copy()

    # Última inspeção por chave (técnico + produto)
    df_com_data.sort_values('DATA_INSPECAO', ascending=False, inplace=True)
    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    chaves_com_data = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_com_data)]

    df_final = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    # Ajusta status e cálculo pendência
    if 'STATUS CHECK LIST' in df_final.columns:
        df_final['STATUS CHECK LIST'] = df_final['STATUS CHECK LIST'].astype(str).str.upper()
    else:
        st.warning("⚠️ Coluna 'STATUS CHECK LIST' não encontrada, usando padrão para pendente/ok.")
        df_final['STATUS CHECK LIST'] = 'PENDENTE'

    hoje = pd.Timestamp.now().normalize()
    df_final['Dias_Sem_Inspecao'] = (hoje - df_final['DATA_INSPECAO']).dt.days.fillna(-1)
    df_final['Vencido'] = df_final['Dias_Sem_Inspecao'] > 180

    return df_final.drop(columns=['CHAVE'])


def calcula_percentuais(df, grupo_col):
    resumo = df.groupby(grupo_col).agg(
        total=('STATUS CHECK LIST', 'count'),
        ok=lambda x: (x == 'OK').sum(),
        pendente=lambda x: (x != 'OK').sum()
    ).reset_index()
    resumo['% OK'] = 100 * resumo['ok'] / resumo['total']
    resumo['% Pendente'] = 100 * resumo['pendente'] / resumo['total']
    return resumo


def gera_grafico_pizza(df, grupo_col, titulo):
    resumo = calcula_percentuais(df, grupo_col)
    fig = px.pie(
        resumo,
        names=grupo_col,
        values='% OK',
        title=f'% EPI OK por {titulo}',
        hole=0.4
    )
    return fig


def gera_grafico_barra(df, grupo_col, titulo):
    resumo = calcula_percentuais(df, grupo_col)
    fig = px.bar(
        resumo,
        x=grupo_col,
        y=['% OK', '% Pendente'],
        barmode='group',
        title=f'% EPI OK e Pendente por {titulo}'
    )
    return fig


def botao_exportar(df):
    df_pendentes = df[df['STATUS CHECK LIST'] != 'OK']
    if df_pendentes.empty:
        st.info("Nenhum EPI pendente para exportar.")
        return

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_pendentes.to_excel(writer, index=False, sheet_name='Pendentes')
        writer.save()
    buffer.seek(0)

    st.download_button(
        label="📥 Exportar Pendentes para Excel",
        data=buffer,
        file_name='epi_pendentes.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def main():
    st.title("Painel de Verificação de EPI")
    df = carregar_dados()
    if df.empty:
        st.warning("⚠️ Nenhum dado carregado.")
        return

    st.header("Gráficos de Status por Gerente")
    fig_gerente = gera_grafico_barra(df, 'GERENTE_IMEDIATO', 'Gerente')
    st.plotly_chart(fig_gerente, use_container_width=True)

    st.header("Gráficos de Status por Coordenador")
    fig_coordenador = gera_grafico_barra(df, 'COORDENADOR_IMEDIATO', 'Coordenador')
    st.plotly_chart(fig_coordenador, use_container_width=True)

    st.header("Tabela Completa dos Dados")
    st.dataframe(df)

    st.header("Exportar Pendentes")
    botao_exportar(df)


if __name__ == "__main__":
    main()
