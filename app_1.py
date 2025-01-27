# Imports
import os
import subprocess
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import streamlit as st
from PIL import Image
from io import BytesIO

# Garantir instalação de dependências ausentes
REQUIRED_LIBRARIES = ["seaborn", "matplotlib", "pillow", "openpyxl", "xlsxwriter"]

for lib in REQUIRED_LIBRARIES:
    try:
        __import__(lib)
    except ModuleNotFoundError:
        subprocess.check_call([os.sys.executable, "-m", "pip", "install", lib])

# Configurações para visualização
custom_params = {"axes.spines.right": False, "axes.spines.top": False}
sns.set_theme(style="ticks", rc=custom_params)

# Funções utilitárias
@st.cache_data(show_spinner=True)
def load_data(file_data):
    """Carrega dados em formato CSV ou Excel."""
    try:
        return pd.read_csv(file_data, sep=';')
    except:
        return pd.read_excel(file_data, engine='openpyxl')

@st.cache_data
def to_excel(df):
    """Converte dataframe para formato Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

@st.cache_data
def multiselect_filter(relatorio, col, selecionados):
    """Aplica filtro multisseleção em colunas específicas."""
    if 'all' in selecionados:
        return relatorio
    else:
        return relatorio[relatorio[col].isin(selecionados)].reset_index(drop=True)

def filter_by_age(df, idades):
    """Filtra os dados por faixa etária."""
    return df[(df['age'] >= idades[0]) & (df['age'] <= idades[1])]

def create_download_button(data, filename):
    """Cria o botão de download para os dados filtrados."""
    df_xlsx = to_excel(data)
    st.download_button(
        label="📥 Download dos Dados Filtrados (Excel)",
        data=df_xlsx,
        file_name=filename,
        mime="application/vnd.ms-excel"
    )

# Função principal
def main():
    # Configuração inicial da página
    st.set_page_config(
        page_title='Telemarketing Analysis',
        layout="wide",
        initial_sidebar_state='expanded'
    )

    # Título e imagem na barra lateral
    st.title('Telemarketing Analysis 📊')
    st.sidebar.image('https://via.placeholder.com/300x150.png?text=Telemarketing+Analysis')

    # Upload do arquivo
    st.sidebar.write("## Suba o arquivo de dados")
    data_file = st.sidebar.file_uploader("Selecione um arquivo CSV ou Excel", type=['csv', 'xlsx'])

    if data_file:
        # Carregamento e exibição inicial dos dados
        bank_raw = load_data(data_file)
        bank = bank_raw.copy()

        st.write("### Dados carregados:")
        st.write(bank.head())

        # Verifica se a coluna 'y' está presente
        if 'y' not in bank.columns:
            st.error("A coluna 'y' não está presente nos dados. Verifique o arquivo carregado.")
            return

        # Filtros interativos
        with st.sidebar.form(key='filter_form'):
            st.write("### Filtros")

            # Filtro de idade
            min_age, max_age = int(bank.age.min()), int(bank.age.max())
            idades = st.slider("Selecione a faixa etária", min_value=min_age, max_value=max_age, value=(min_age, max_age))

            # Filtro de profissões
            jobs_list = bank.job.unique().tolist() + ['all']
            jobs_selected = st.multiselect("Selecione as profissões", jobs_list, ['all'])

            # Filtro de estado civil
            marital_list = bank.marital.unique().tolist() + ['all']
            marital_selected = st.multiselect("Selecione o estado civil", marital_list, ['all'])

            # Botão para aplicar os filtros
            submit_button = st.form_submit_button(label='Aplicar Filtros')

        if submit_button:
            # Aplicação dos filtros
            bank = filter_by_age(bank, idades)
            bank = (bank.pipe(multiselect_filter, 'job', jobs_selected)
                        .pipe(multiselect_filter, 'marital', marital_selected))

        # Exibição e download dos dados filtrados
        st.write("### Dados Após Filtros:")
        st.write(bank.head())
        create_download_button(bank, "dados_filtrados.xlsx")

        # Análise de proporção de aceite (coluna 'y')
        st.write("### Proporção de Aceite:")
        fig, ax = plt.subplots(1, 2, figsize=(12, 6))

        target_perc = bank['y'].value_counts(normalize=True).to_frame() * 100
        target_perc.columns = ['Percentage']
        
        # Gráfico de barras
        sns.barplot(x=target_perc.index, y='Percentage', data=target_perc, ax=ax[0])
        ax[0].set_title("Distribuição de Y (Barras)")
        for container in ax[0].containers:
            ax[0].bar_label(container)

        # Gráfico de pizza
        target_perc.plot.pie(y='Percentage', autopct='%1.1f%%', ax=ax[1])
        ax[1].set_title("Distribuição de Y (Pizza)")

        st.pyplot(fig)

    else:
        st.warning("Por favor, faça o upload de um arquivo.")

if __name__ == '__main__':
    main()
