@st.cache_data
def carregar_dados():
    df = pd.read_excel("LISTA DE VERIFICAÇÃO EPI.xlsx", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Normaliza nomes das colunas
    tecnico_col = [c for c in df.columns if 'TECNICO' in c.upper()][0]
    produto_col = [c for c in df.columns if 'PRODUTO' in c.upper()][0]
    data_col = [c for c in df.columns if 'INSPECAO' in c.upper()][0]

    df.rename(columns={'GERENTE': 'GERENTE_IMEDIATO',
                       'COORDENADOR': 'COORDENADOR_IMEDIATO',
                       'SITUAÇÃO CHECK LIST': 'STATUS CHECK LIST'}, inplace=True)

    # Normaliza e cria chave única
    df['CHAVE'] = (df[tecnico_col].astype(str).str.strip().str.lower() + '|' +
                   df[produto_col].astype(str).str.strip().str.lower())

    df['Data_Inspecao'] = pd.to_datetime(df[data_col], errors='coerce')

    # Separe os registros com e sem data
    df_com_data = df.dropna(subset=['Data_Inspecao']).copy()
    df_sem_data = df[df['Data_Inspecao'].isna()].copy()

    # Ordena para garantir que o drop_duplicates mantenha a última data
    df_com_data.sort_values('Data_Inspecao', ascending=False, inplace=True)

    # Pega só a última inspeção para cada técnico+produto
    df_ultimos = df_com_data.drop_duplicates(subset=['CHAVE'], keep='first')

    # Agora exclui do sem data as chaves que já aparecem na última inspeção
    chaves_ultimos = set(df_ultimos['CHAVE'])
    df_sem_data = df_sem_data[~df_sem_data['CHAVE'].isin(chaves_ultimos)]

    # Junta tudo (últimos + pendentes)
    df_final = pd.concat([df_ultimos, df_sem_data], ignore_index=True)

    # Renomeia colunas para facilitar
    df_final.rename(columns={
        tecnico_col: 'TECNICO',
        produto_col: 'PRODUTO'
    }, inplace=True)

    # Padroniza status em maiúsculo
    if 'STATUS CHECK LIST' in df_final.columns:
        df_final['STATUS CHECK LIST'] = df_final['STATUS CHECK LIST'].str.upper()

    # Dias desde a última inspeção
    hoje = pd.Timestamp.now().normalize()
    df_final['Dias_Sem_Inspecao'] = (hoje - df_final['Data_Inspecao']).dt.days.fillna(-1)
    df_final['Vencido'] = df_final['Dias_Sem_Inspecao'] > 180

    return df_final.drop(columns=['CHAVE'])
