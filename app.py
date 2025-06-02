import streamlit as st
import pandas as pd
import io

# Função para carregar dados e eliminar duplicados
def carregar_dados_sem_duplicados(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, engine='openpyxl')
    df.columns = df.columns.str.strip()

    col_tec = [col for col in df.columns if 'TECNICO' in col.upper()]
    col_prod = [col for col in df.columns if 'PRODUTO' in col.upper()]
    col_data = [col for col in df.columns if 'INSPECAO' in col.upper()]

    if not col_tec or not col_prod or not col_data:
        st.error("❌ Verifique colunas TECNICO, PRODUTO e INSPECAO no arquivo")
        return pd.DataFrame()

    tecnico_col = col_tec[0]
    produto_col = col_prod[0]
    data_col = col_data[0]

    df['Data_Inspecao'] = pd.to_datetime(df[data_col], errors='coerce')

    df['CHAVE'] = df[tecnico_col].astype(str).str.strip() + '|' + df[produto_col].astype(str).str.strip()

    df_com_data = df.dropna(subset=['Data_Inspecao']).copy()
    df_sem_data = df[df['Data_Inspecao'].isna()].copy()

    df_com_data.sort_values(['CHAVE', 'Data_Inspecao'], ascending=[True, False], inplace=True)

    df_ultimos = df_com_data.drop_duplicates(subset='CHAVE', keep='first')

    chaves_inspecionadas = set(df_ultimos['CHAVE'])

    df_pendentes = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_inspecionadas)].copy()

    df_final = pd.concat([df_ultimos, df_pendentes], ignore_index=True)

    df_final.drop(columns=['CHAVE'], inplace=True)

    # Ajustar colunas para facilitar filtro
    if 'GERENTE' in df_final.columns:
        df_final.rename(columns={'GERENTE': 'GERENTE_IMEDIATO'}, inplace=True)
    
    # Criar coluna status pendente / ok
    df_final['STATUS'] = df_final['Data_Inspecao'].apply(lambda x: 'PENDENTE' if pd.isna(x) else 'OK')

    return df_final, df_pendentes

# Função para criar cards de percentual
def exibe_cards(df):
    if df.empty:
        st.warning("Sem dados para exibir cards.")
        return
    
    # Agrupa por gerente e coordenador (se tiver coordenador, senão só gerente)
    for grupo in ['GERENTE_IMEDIATO', 'COORDENADOR_IMEDIATO'] if 'COORDENADOR_IMEDIATO' in df.columns else ['GERENTE_IMEDIATO']:
        st.markdown(f"### Percentual por {grupo}")

        resumo = df.groupby(grupo)['STATUS'].value_counts(normalize=True).unstack(fill_value=0) * 100

        for nome, linha in resumo.iterrows():
            ok = linha.get('OK', 0)
            pendente = linha.get('PENDENTE', 0)
            st.markdown(f"**{nome}**")
            col1, col2 = st.columns(2)
            col1.metric(label="EPI OK (%)", value=f"{ok:.1f}%")
            col2.metric(label="EPI Pendentes (%)", value=f"{pendente:.1f}%")
            st.markdown("---")

# Função para exportar pendentes para excel
def botao_exportar(df_pendentes):
    if df_pendentes.empty:
        st.info("Não há pendentes para exportar.")
        return

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_pendentes.to_excel(writer, index=False, sheet_name='Pendentes')
        writer.save()
    buffer.seek(0)

    st.download_button(
        label="Exportar Pendentes para Excel",
        data=buffer,
        file_name="epi_pendentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Main app
def main():
    st.title("Dashboard de Inspeção EPI")

    uploaded_file = st.file_uploader("Escolha o arquivo Excel com a lista de verificação EPI", type=["xlsx"])
    if uploaded_file is not None:
        df, df_pendentes = carregar_dados_sem_duplicados(uploaded_file)

        if not df.empty:
            exibe_cards(df)

            st.markdown("### Tabela completa dos dados")
            st.dataframe(df)

            botao_exportar(df_pendentes)

if __name__ == "__main__":
    main()
