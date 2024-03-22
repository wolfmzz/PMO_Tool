# ========================================================== PRAGMATIS - TELECOM ===========================================================
# ------------------------------------------------------ SIZING | GANTT CHART BUILDER ------------------------------------------------------



## Contexto
# XXX

## Propósito deste código
# XXX

## Sumário
# 01. Introdução
# 02. Cockpit
# 03. Funções
# 04. Manipulação de Dados
# 05. Visualização de Dados

## Contatos
# Julia Wolf Mazzuia | Intern



# ------------------------------------------------------------- 01. INTRODUÇÃO -------------------------------------------------------------



# Bibliotecas de Gerenciamento do Sistema
import os  # Gerenciamento de Arquivos
import sys  # Gerenciamento de Sistema
import warnings  # Gerenciamento de Avisos

# Bibliotecas de Manipulação de Dados
import numpy as np  # Funções Matemáticas
import pandas as pd  # Manipulação de DataFrames
from io import BytesIO  # Leitura de Bytes

# Bibliotecas relacionadas ao Tempo
import time  # Funções relacionadas ao tempo
from datetime import datetime  # Funções relacionadas ao tempo
from datetime import timedelta  # Função de duração do tempo

# Bibliotecas de Visualização
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Adicione o caminho para a pasta principal ao caminho do sistema
sys.path.append(os.path.join(os.path.dirname(sys.path[0])))

# Limpar memória
import gc
gc.collect()

# Ponto de início do horário do script
begin_time = time.monotonic()

# Verifique onde este arquivo está salvo
wdir = os.getcwd()  # Determine o caminho do diretório de trabalho
wdir = wdir.replace("\\", "/")  # Substitua o caractere do caminho para corresponder à sintaxe Python
os.chdir(wdir)  # Defina o diretório de trabalho

# Caminho adicional para armazenar inputs e outputs
input_path = wdir + "/Input"
output_path = wdir + "/Output"

# Adiciona na lista de tempos de execução
time_chapter = []

# Encerra este capítulo
time_chapter.append(0)
print("01. Introdução | OK")



# -------------------------------------------------------------- 02. COCKPIT ---------------------------------------------------------------



# Grava horário no formato para salvar o arquivo
version_save = datetime.now().strftime("%Y-%m-%d--%Hh%M")

# Caminho para o arquivo de configuração
result_path = output_path + "/Mock-up Project Tool_Preenchido.xlsx"

# Definir o número máximo de linhas e colunas para exibir
pd.set_option("display.max_rows", 100)  # Definir o número máximo de linhas
pd.set_option("display.max_columns", 20)  # Definir o número máximo de colunas

# Encerra este capítulo
time_chapter.append(0)
print("02. Cockpit | OK")



# ------------------------------------------------------------- 03. FUNCTIONS --------------------------------------------------------------



# Função que retorna valor determinado pelo usuário no arquivo de configuração
def get_config_value(
    label_raw: pd.DataFrame,
    config_label: str,
    ):
    """
    Função que complementa check de colunas com quantidade esperada de linhas para colunas atreladas a certas mecanicas

    Args:
        df_nan (DataFrame): DataFrame com quantidade de linhas com missing value por coluna.

    Returns:
        df_check (DataFrame): DataFrame com quantidade de linhas com missing value por coluna e quantidade esperada de linhas por mecanica
    """
    # Puxa valor da configuração desejada do arquivo de configuração
    value = label_raw.query(f"Label == '{config_label}'").Value.values[0]

    # Retorna valor
    return value


# Função que constroi o gráfico de Gantt
def build_gantt_chart(
    df: pd.DataFrame,
    config: pd.DataFrame
    ):
    """
    Função que substitui os valores da coluna pelos calculados pelo script

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): Planilha do Excel
        df (pd.DataFrame): DataFrame com os valores a serem substituidos
        column (str): Nome da coluna do DataFrame
        column_letter (str): Letra da coluna no Excel

    Returns:
        True: Retorna True se a função foi executada com sucesso
    """    
    # Transforma conteúdo das colunas em listas
    atividades = df["Atividade"].tolist()
    data_inicio = df["Data_Inicio"].tolist()
    data_fim = df["Data_Planejado"].tolist()
    duration = df["Duration"].tolist()

    # Leitura das configurações de visualização
    title = get_config_value(config, "Title")
    Xlabel = get_config_value(config, "Xlabel")
    Ylabel = get_config_value(config, "Ylabel")
    num_weeks_to_display = get_config_value(config, "Week_Display")
    height_bars = get_config_value(config, "Height_Bars")

    # Calculo de configurações de visualização 
    num_atividades = len(atividades)

    # Inicializa gráfico
    fig, ax = plt.subplots()

    # Preenche eixo y com atividades
    ax.set_yticks(np.arange(num_atividades))
    ax.set_yticklabels(atividades)

    # Plota cada atividade como uma barra horizontal
    for ativ in range(num_atividades):
        start_date = pd.to_datetime(data_inicio[ativ])
        end_date = start_date + pd.DateOffset(days = duration[ativ])
        ax.barh(ativ, end_date - start_date, left = start_date, height = height_bars, align = 'center')

    # Determina os limites do eixo de datas
    min_date = pd.to_datetime(min(data_inicio))
    max_date = pd.to_datetime(max(data_inicio)) + pd.DateOffset(days = max(duration))
    ax.set_xlim(min_date, max_date)

    # Customize the x-axis ticks and labels
    weeks_interval = max(int((max_date - min_date).days / 7 / num_weeks_to_display), 1)
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval = weeks_interval))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))

    # Customiza o eixo de datas
    ax.set_xlabel(Xlabel)
    ax.set_ylabel(Ylabel)
    ax.set_title(title)

    # Step 8: Display the chart
    plt.grid(True)
    plt.show()

    return fig


# Finaliza o capitulo
time_chapter.append(0)
print("03. Funções | OK")



# ---------------------------------------------------------- 04. DATA MANIPULATION ---------------------------------------------------------



# Marcar o ponto de partida do capítulo
start_time = time.monotonic()

# Le arquivo com resultados que foram feasible e relatório de feasibility
atividades = pd.read_excel(result_path, sheet_name = "Painel")
config = pd.read_excel(result_path, sheet_name = "Config")

# Mostar as iniciativas possíveis
list_iniciativas = list(atividades.Iniciativa.unique())

# Ends this chapter
end_time = time.monotonic()
duration = timedelta(seconds = end_time - start_time)
time_chapter.append(duration.total_seconds())
print("03. Data Manipulation | OK")
print(f"    Duration: {duration}")
print(" ")



# ------------------------------------------------------------- 05. APP BUILDER ------------------------------------------------------------



# Titulo do principal
st.title("Gráfico de Gantt das Atividades")

# Titulo do SideBar
st.sidebar.header("Filtros")  

# Permite usuário dar upload de arquivo com dados do Quinto Andar
Upload_Data = st.file_uploader("Upload excel com atividades", type = "xlsx")

# Caso usuário tenha feito upload de arquivo
if Upload_Data is not None:

    # Pega valor de bytes do arquivo
    bytes_data = Upload_Data.getvalue()

    # Le arquivo excel
    atividades = pd.read_excel(BytesIO(bytes_data))  

# Conteúdo que pode ser selecionado
iniciativa_selected = st.sidebar.multiselect("Escolha Iniciativa", list_iniciativas + ["Todas"])

# Caso a iniciativa selecionada seja Todas
if iniciativa_selected == "Todas" or iniciativa_selected == []:

    # Não filtra a iniciativa
    atividades_filtered = atividades

# Caso seja selecionada alguma iniciativa
else:

    # Filtra a iniciativa selecionada
    atividades_filtered = atividades.query(f"Iniciativa in {iniciativa_selected}")

# Cria gráfico de Gantt com a iniciativa selecionada
fig = build_gantt_chart(atividades_filtered, config)

# Plota gráfico na tela
st.pyplot(fig)

# Ends this chapter
end_time = time.monotonic()
duration = timedelta(seconds = end_time - start_time)
time_chapter.append(duration.total_seconds())
print("04. App Builder | OK")
print(f"    Duration: {duration}")
print(" ")



# ------------------------------------------------------------- 99. SCRAPYARD --------------------------------------------------------------



# # Caso o conteúdo selecionado seja Tabela
# if content_selected == "Tabela":

#     # Tipo de tabela que será selecionada
#     vision_selected = st.sidebar.selectbox("Visão", list_vision)

#     # Caixas de seleção que montam as tabelas
#     type_selected = st.sidebar.selectbox("Tipo de Escolha", list_type_selected)
#     constraint_selected = st.sidebar.selectbox("Constraint", list_run_constraint)

#     # Caso seja selecionado a visão Financeiro_Geral
#     if vision_selected == "Financeiro_Geral":

#         # Determina parte do nome do dicionário de tabelas a ser selecionado
#         report_selected = f"tabela_metricas_dict"
#         table_selection = f"{constraint_selected}_{type_selected}"
#         data_selection = eval(report_selected)[table_selection]

#         # Insere a tabela selecionada
#         st.write(data_selection)

#         # Cria gráfico para métricas financeiras
#         chart_fig = chart_metrics(data_selection)
#         st.pyplot(chart_fig)

#     # Caso seja selecionada a visão Produto
#     else:

#         # Caso seja selecionado a visão Produto, cria barra de seleção de taxonomia e métrica
#         taxonomia = st.sidebar.selectbox("Taxonomia", list_taxonomia)
#         metric = st.sidebar.selectbox("Métrica", list_metric)

#         # Determina parte do nome do dicionário de tabelas a ser selecionado
#         report_selected = f"tabela_{taxonomia}_{metric}_dict"
#         table_selection = f"{constraint_selected}_{type_selected}"
#         data_selection = eval(report_selected)[table_selection]

#         # Insere a tabela selecionada
#         st.write(data_selection)

#         # Mostra lista de opções para seleção da taxonomia
#         list_taxonomy_options = list(data_selection[taxonomia].unique())

#         # Cria seleções para gráfico
#         cycle_selected = st.selectbox("Escolha de ciclo para gráfico", list_cycles)
#         taxonomy_selected = st.selectbox(f"Escolha de {taxonomia} para gráfico", list_taxonomy_options)
        
#         # Cria gráfico para distribuição de sku por taxonomia em cada em cada intervalo de métrica
#         chart_fig = chart_distribution(data_selection, taxonomy_selected, cycle_selected, taxonomia, metric)
#         st.pyplot(chart_fig)

# # Caso o conteúdo selecionado seja Resumo
# else:

#     # Insere o relatório de feasibility
#     st.write(report_feasiblity)
