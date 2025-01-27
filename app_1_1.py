import os
import subprocess

# Garantir instalação de dependências ausentes
REQUIRED_LIBRARIES = ["seaborn", "matplotlib", "pillow", "openpyxl", "xlsxwriter"]

for lib in REQUIRED_LIBRARIES:
    try:
        __import__(lib)
    except ModuleNotFoundError:
        subprocess.check_call([os.sys.executable, "-m", "pip", "install", lib])

# Imports necessários
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO

# Configurações globais para visualização de gráficos
sns.set_theme(style="ticks", rc={"axes.spines.right": False, "axes.spines.top": False})

# Função para carregar os dados (com cache para melhor performance)
@st.cache_data(show_spinner=True)
def load_data(file_data):
    """Carrega dados em formato CSV ou Excel."""
    try:
        return pd.read_csv(file_data, sep=';')
    except Exception:
        return pd.read_excel(file_data, engine='openpyxl')

# Função para converter dataframe em formato Excel
@st.cache_data
def to_excel(df):
    """Converte dataframe para formato Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Função para aplicar filtros em múltiplas seleções
def multiselect_filter(df, column, selected_options):
    """Aplica filtro baseado em seleções múltiplas."""
    if 'all' in selected_options:
        return df
    return df[df[column].isin(selected_options)].reset_index(drop=True)

# Função para filtrar dados por faixa etária
def filter_by_age(df, age_range):
    """Filtra dados por uma faixa etária."""
    return df[(df['age'] >= age_range[0]) & (df['age'] <= age_range[1])]

# Função para criar botão de download
def create_download_button(data, filename):
    """Cria botão para baixar os dados filtrados em formato Excel."""
    df_xlsx = to_excel(data)
    st.download_button(
        label="📥 Baixar Dados Filtrados (Excel)",
        data=df_xlsx,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Função principal
def main():
    # Configuração inicial da página no Streamlit
    st.set_page_config(
        page_title='Telemarketing Analysis Dashboard',
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Título da página
    st.title("📊 Telemarketing Analysis Dashboard")

    # Sidebar para upload do arquivo
    st.sidebar.header("📁 Upload de Dados")
    data_file = st.sidebar.file_uploader(
        "Selecione um arquivo CSV ou Excel:", type=['csv', 'xlsx']
    )

    # Verifica se o arquivo foi enviado
    if data_file:
        # Carrega os dados
        bank_data = load_data(data_file)

        # Exibe os dados carregados
        st.subheader("📋 Pré-visualização dos Dados")
        st.dataframe(bank_data.head())

        # Verifica se a coluna 'y' existe
        if 'y' not in bank_data.columns:
            st.error("Erro: A coluna 'y' não está presente nos dados fornecidos.")
            return

        # Configuração dos filtros
        st.sidebar.header("🔍 Filtros")
        age_min, age_max = int(bank_data['age'].min()), int(bank_data['age'].max())
        age_range = st.sidebar.slider("Selecione a faixa etária:", age_min, age_max, (age_min, age_max))

        job_options = ['all'] + bank_data['job'].dropna().unique().tolist()
        selected_jobs = st.sidebar.multiselect("Selecione profissões:", job_options, default='all')

        marital_options = ['all'] + bank_data['marital'].dropna().unique().tolist()
        selected_marital = st.sidebar.multiselect("Selecione estado civil:", marital_options, default='all')

        # Aplica os filtros
        filtered_data = filter_by_age(bank_data, age_range)
        filtered_data = multiselect_filter(filtered_data, 'job', selected_jobs)
        filtered_data = multiselect_filter(filtered_data, 'marital', selected_marital)

        # Exibe os dados filtrados
        st.subheader("📊 Dados Filtrados")
        st.dataframe(filtered_data)

        # Botão para baixar os dados filtrados
        create_download_button(filtered_data, "dados_filtrados.xlsx")

        # Análise gráfica
        st.subheader("📈 Análise Gráfica - Distribuição da Coluna 'y'")
        fig, axes = plt.subplots(1, 2, figsize=(14, 6))

        y_distribution = filtered_data['y'].value_counts(normalize=True) * 100
        sns.barplot(x=y_distribution.index, y=y_distribution.values, ax=axes[0])
        axes[0].set_title("Distribuição (Barras)")
        axes[0].set_ylabel("Porcentagem")
        axes[0].set_xlabel("Valores de 'y'")
        axes[0].bar_label(axes[0].containers[0])

        y_distribution.plot.pie(autopct='%1.1f%%', ax=axes[1], startangle=90, labels=y_distribution.index)
        axes[1].set_ylabel("")
        axes[1].set_title("Distribuição (Pizza)")

        st.pyplot(fig)
    else:
        st.warning("Por favor, faça o upload de um arquivo para começar.")

# Executa o app
if __name__ == "__main__":
    main()
