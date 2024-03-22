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
import openpyxl  # Manipulação de Excel

# Bibliotecas relacionadas ao Tempo
import time  # Funções relacionadas ao tempo
from datetime import datetime  # Funções relacionadas ao tempo
from datetime import timedelta  # Função de duração do tempo

# Bibliotecas de Visualização
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Bibliotecas de Manipulação de Strings
import ast

# Limpar memória
import gc
gc.collect()

# Ponto de início do horário do script
begin_time = time.monotonic()

# # Adiciona Path alternativo
# sys.path.append(os.path.join(os.path.dirname(sys.path[0])))

# # Caminho adicional para armazenar inputs
# input_path = os.path.abspath(
#     os.path.join(wdir, "Input")
# )

# # Caminho adicional para armazenar output
# output_path = os.path.abspath(
#     os.path.join(wdir, "../Output")
# )

# Verifique onde este arquivo está salvo
wdir = os.getcwd()  # Determine o caminho do diretório de trabalho
wdir = wdir.replace("\\", "/")  # Substitua o caractere do caminho para corresponder à sintaxe Python
os.chdir(wdir)  # Defina o diretório de trabalho

# Caminhos adicionais para armazenar arquivos
input_path = wdir + "/Input"
output_path = wdir + "/Output"
# output_path = wdir

# Adiciona na lista de tempos de execução
time_chapter = []

# Encerra este capítulo
time_chapter.append(0)
print("01. Introdução | OK")



# -------------------------------------------------------------- 02. COCKPIT ---------------------------------------------------------------



# Grava horário no formato para salvar o arquivo
version_save = datetime.now().strftime("%Y-%m-%d--%Hh%M")

# Ignorar Avisos
warnings.filterwarnings(
    "ignore",
    message = "Cannot parse header or footer so it will be ignored",
    category = UserWarning,
)

# Caminho do arquivo de input
database_path = input_path + "/Mock-up_Raw.xlsx"
result_path = output_path + "/Mock-up Project Tool_Preenchido.xlsx"

# Encerra este capítulo
time_chapter.append(0)
print("02. Cockpit | OK")



# -------------------------------------------------------------- 03. FUNÇÕES ---------------------------------------------------------------



# Adiciona "[]" nas colunas de conjuntos
def add_brackets(
    df: pd.DataFrame,
    column_eval: str,
    column_new: str
    ):  
    """
    Adiciona "[]" nas colunas de conjuntos

    Args:
        df (pd.DataFrame): Dataframe com colunas de conjuntos
        column_eval (str): Nome coluna a ser avaliada
        column_new (str): Nome coluna a ser criada

    Returns:
        df: Dataframe com colunas de conjuntos com "[]" adicionados.
    """
    # Função para adicionar "[]" nas colunas de conjuntos
    def process_row(s):  

        s = s.replace(' ', '')
        s = s.replace(',', ', ')

        if not s.startswith("["):  
            s = "[" + s  
        if not s.endswith("]"):  
            s = s + "]" 

        return s

    # Efetivamente adiciona "[]" nas colunas de conjuntos
    df[column_eval] = df[column_eval].astype(str)
    df[column_new] = df[column_eval].apply(process_row) 

    return df  


# Ajusta colunas necessárias para permitir o uso da tabela na função que preenche datas
def clean_df(
    df_raw: pd.DataFrame,
    ):
    """
    Função que complementa check de colunas com quantidade esperada de linhas para colunas atreladas a certas mecanicas

    Args:
        df_raw (DataFrame): DataFrame a ser limpo

    Returns:
        df (DataFrame): Retorna DataFrame pronto para preenchimento de datas
    """
    # Separa tabela com Data Inicio preenchida e não preenchida
    df_raw = df_raw.replace(np.nan, "", regex = True)

    # Ajusta coluna Atividade_Dependente
    df_raw = add_brackets(df_raw, "Atividade_Dependente", "Atividade_Dependente_New")

    # Converte coluna de Atividade Dependente para lista
    df_raw = (
        df_raw
        .copy()
        .assign(len_atv_dep = lambda _: _.Atividade_Dependente.apply(lambda x: len(x.split())))
        .assign(Atividade_Dependente = lambda _: np.where(_.len_atv_dep == 1, _.Atividade_Dependente, _.Atividade_Dependente_New))
        .assign(Atividade_Dependente = lambda _: _.Atividade_Dependente.astype(str))
        .assign(Atividade_Dependente = lambda _: _.Atividade_Dependente.apply(lambda x: ast.literal_eval(x) if x.startswith('[') else x))
        )
    
    # df_raw['Data_Inicio'] = pd.to_datetime(df_raw['Data_Inicio'])
    # df_raw['Data_Planejado'] = pd.to_datetime(df_raw['Data_Planejado'])

    # Transforma data em string
    df_filled = (
        df_raw
        .copy()
        .query("Data_Inicio == Data_Inicio")
        .assign(Data_Inicio = lambda _: _.Data_Inicio.dt.strftime('%d-%m-%y'))
        .assign(Data_Planejado = lambda _: _.Data_Planejado.dt.strftime('%d-%m-%y'))
        .assign(Data_Inicio = lambda _: _.Data_Inicio.astype(str))
        .assign(Data_Planejado = lambda _: _.Data_Planejado.astype(str))
        .assign(Data_Fim = lambda _: _.Data_Planejado)
        .assign(Data_Fim = lambda _: np.where(_.Fim_Efetivo == "", _.Data_Fim, _.Fim_Efetivo))
    )

    # Troca NaT por ""
    df_to_fill = (
        df_raw
        .copy()
        .query("Data_Inicio != Data_Inicio")
        .assign(Data_Inicio = "")
        .assign(Data_Planejado = "")
        )

    # Junta de novo as tabelas
    df = pd.concat([df_filled, df_to_fill], ignore_index = True)

    # Retorna DataFrame pronto para preenchimento de datas
    return df


# Preenche as datas de inicio e fim das atividades
def fill_dates(
    df_raw: pd.DataFrame,
    ):
    """
    Função que complementa check de colunas com quantidade esperada de linhas para colunas atreladas a certas mecanicas

    Args:
        df (DataFrame): DataFrame com a coluna que será preenchida

    Returns:
        df_check (DataFrame): DataFrame com quantidade de linhas com missing value por coluna e quantidade esperada de linhas por mecanica
    """
    # Separa tabela com Data Inicio preenchida e não preenchida
    df = clean_df(df_raw)

    # Preenche datas
    for index, row in df.iterrows():
        if row["Data_Inicio"] == "":
            if row["Atividade_Dependente"] == "":
                start_date = datetime.strptime(df.iloc[0]["Data_Inicio"], "%d-%m-%y")
            elif isinstance(row["Atividade_Dependente"], list):
                max_end_date = datetime.min
                for dep_id in row["Atividade_Dependente"]:
                    dep_row = df[df["ID"] == dep_id]
                    if not dep_row.empty:
                        dep_row = dep_row.iloc[0]
                        if dep_row["Data_Fim"] == "":
                            fill_dates(df)
                        dep_end_date = datetime.strptime(dep_row["Data_Fim"], "%d-%m-%y")
                        max_end_date = max(max_end_date, dep_end_date)
                start_date = max_end_date + timedelta(days=1)
            else:
                dep_row = df[df["ID"] == row["Atividade_Dependente"]]
                if not dep_row.empty:
                    dep_row = dep_row.iloc[0]
                    if dep_row["Data_Fim"] == "":
                        fill_dates(df)
                    start_date = datetime.strptime(dep_row["Data_Fim"], "%d-%m-%y") + timedelta(days=1)
                else:
                    start_date = datetime.strptime(df.iloc[0]["Data_Inicio"], "%d-%m-%y")
            df.at[index, "Data_Inicio"] = start_date.strftime("%d-%m-%y")
            end_date = start_date + timedelta(days=row["Duration"])
            df.at[index, "Data_Fim"] = end_date.strftime("%d-%m-%y")

    # Ajustes finais na formatacao da tabela
    df = (
        df
        .copy()
        .assign(Data_Planejado = lambda _: _.Data_Fim)
        .drop(["Atividade_Dependente_New", "len_atv_dep", "Data_Fim"], axis = 1)
        )

    return df


# Função adiciona uma coluna com qual das atividades dependentes é a com data mais longe
def add_latest_dependent(
    df: pd.DataFrame
    ):
    """
    Função adiciona uma coluna com qual das atividades dependentes é a com data mais longe

    Args:
        df (DataFrame): DataFrame com a coluna que será preenchida

    Returns:
        df_check (DataFrame): DataFrame com quantidade de linhas com missing value por coluna e quantidade esperada de linhas por mecanica
    """
    # Add a new column to store the latest dependent activity
    df["Atividade_Dependente_Gargalo"] = None

    for index, row in df.iterrows():
        latest_dep_id = None
        latest_dep_date = None

        if isinstance(row["Atividade_Dependente"], list):
            max_end_date = datetime.min
            for dep_id in row["Atividade_Dependente"]:
                dep_row = df[df["ID"] == dep_id]
                if not dep_row.empty:
                    dep_row = dep_row.iloc[0]
                    dep_end_date = datetime.strptime(dep_row["Data_Planejado"], "%d-%m-%y")
                    if dep_end_date > max_end_date:
                        max_end_date = dep_end_date
                        latest_dep_id = dep_id
        elif row["Atividade_Dependente"] != "":
            latest_dep_id = row["Atividade_Dependente"]

        df.at[index, "Atividade_Dependente_Gargalo"] = latest_dep_id

    df["Atividade_Dependente_Gargalo"] = df["Atividade_Dependente_Gargalo"].replace("-", 0)
    df["Atividade_Dependente_Gargalo"] = df["Atividade_Dependente_Gargalo"].astype(int)

    return df


# Função que substitui os valores da coluna pelos calculados pelo script
def replace_column(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    df: pd.DataFrame, 
    column: str):
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
    # Encontra em que posição da planilha está a coluna
    index_position = list(df.columns).index(column)

    # Converte a posição para letra (+1 pelo python começar a contar no 0 e +64 por ser a pasdrão ASCII de caracteres)
    column_letter = chr(index_position + 1 + 64)

    # Substitui os valores da coluna
    for row, value in enumerate(df[column], start = 2):
        sheet[f"{column_letter}{row}"] = value

    return True


# Função que ajusta formato da coluna de datas
def clean_date_format(
    date_string: str
):
    """
    Função que ajusta formato da coluna de datas
    
    Args:
        date_string (str): string de data

    Returns:
        formatted_date: string de data formatada
    """

    # date_object = datetime.strptime(date_string, "%m-%d-%y")
    # formatted_date = date_object.strftime("%m/%d/%Y")

    day, month, year = date_string.split('-')
    year = '20' + year
    formatted_date = f"{day}-{month}-{year}"

    return formatted_date


# Finaliza o capitulo
time_chapter.append(0)
print("03. Funções | OK")



# -------------------------------------------------------- 04. MANIPULAÇÃO DE DADOS --------------------------------------------------------



# Marcar o ponto de partida do capítulo
start_time = time.monotonic()

# Leitura do arquivo de input
activities_raw = pd.read_excel(database_path, sheet_name = "Painel", engine = "openpyxl")
config = pd.read_excel(database_path, sheet_name = "Config", engine = "openpyxl")

# Preenche as datas de inicio e fim das atividades
atividades = fill_dates(activities_raw)

# Adiciona coluna com gargalo das listas de atividades dependentes de cada atividade
atividades = add_latest_dependent(atividades)  

# Ajusta formato da coluna de data de inicio
atividades["Data_Inicio"] = atividades["Data_Inicio"].apply(clean_date_format)

# Escreve excel com datas preenchidas
# atividades.to_excel("Mock-up Project Tool_Preenchido.xlsx", index=False)

# Carrega arquivo original pelo openpyxl
workbook = openpyxl.load_workbook(database_path)
sheet = workbook["Painel"]

# Substitui os valores da coluna calculado pelo python
replace_column(sheet, atividades, "Data_Inicio")
replace_column(sheet, atividades, "Atividade_Dependente_Gargalo")

# Save the updated Excel file
workbook.save(result_path)



# ------------------------------------------------------- 05. VISUALIZAÇÃO DE DADOS --------------------------------------------------------



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

    return True


# Mostar as iniciativas possíveis
print(atividades.Iniciativa.unique())

# Filtra uma iniciativa em especifico
atividades_filtered = atividades.query("Iniciativa == 'Regionalização'")

# Cria gráfico de Gantt
build_gantt_chart(atividades_filtered, config)



# ------------------------------------------------------------- 99. SCRAPYARD --------------------------------------------------------------



# # Preenche as datas de inicio e fim das atividades
# def fill_dates(
#     df: pd.DataFrame,
#     ):
#     """
#     Função que complementa check de colunas com quantidade esperada de linhas para colunas atreladas a certas mecanicas

#     Args:
#         df (DataFrame): DataFrame com a coluna que será preenchida

#     Returns:
#         df_check (DataFrame): DataFrame com quantidade de linhas com missing value por coluna e quantidade esperada de linhas por mecanica
#     """

#     # Preenche as datas de inicio e fim das atividades
#     for index, row in df.iterrows():
#         if row["Data_Inicio"] == "":
#             if row["Atividade_Dependente"] == "":
#                 start_date = datetime.strptime(df.iloc[0]["Data_Inicio"], "%d-%m-%y")
#                 # start_date = datetime.strptime(df.iloc[0]["Data_Inicio"], "%d-%m-%y")
#             elif isinstance(row["Atividade_Dependente"], list):
#                 max_end_date = datetime.min
#                 for dep_id in row["Atividade_Dependente"]:
#                     dep_row = df[df["ID_Atividade"] == dep_id].iloc[0]
#                     if dep_row["Data_Planejado"] == "":
#                         fill_dates(df)
#                     dep_end_date = datetime.strptime(dep_row["Data_Planejado"], "%d-%m-%y")
#                     max_end_date = max(max_end_date, dep_end_date)
#                 start_date = max_end_date + timedelta(days=1)
#             else:
#                 dep_row = df[df["ID_Atividade"] == row["Atividade_Dependente"]].iloc[0]
#                 if dep_row["Data_Planejado"] == "":
#                     fill_dates(df)
#                 start_date = datetime.strptime(dep_row["Data_Planejado"], "%d-%m-%y") + timedelta(days=1)
#                 start_date = datetime.strptime(dep_row["Data_Planejado"], "%d-%m-%y") + timedelta(days=1)
#             df.at[index, "Data_Inicio"] = start_date.strftime("%d-%m-%y")
#             end_date = start_date + timedelta(days=row["Duration"])
#             df.at[index, "Data_Planejado"] = end_date.strftime("%d-%m-%y")

#     # Retorna dataframe com a coluna preenchida
#     return df

#     data = {
#     'ID_Atividade': [1, 2, 3, 4, 5, 6],
#     'Atividade': ['aaa', 'bbb', 'ccc', 'ddd', 'eee', 'fff'],
#     'Atividade_Dependente': ['', 1, [1, 2], 3, '', [3, 4]],
#     'Data_Inicio': ['20-01-23', '', '', '', '', ''],
#     'Duration': [3, 4, 2, 4, 1, 5],
#     'Data_Planejado': ['23-01-23', '', '', '', '', '']
#     }

#     df = pd.DataFrame(data)


#     atividades = fill_dates(df)







# # Adiciona "[]" nas colunas de conjuntos
# def add_brackets(
#     df: pd.DataFrame,
#     column_eval: str,
#     column_new: str
#     ):  
#     """
#     Adiciona "[]" nas colunas de conjuntos

#     Args:
#         df (pd.DataFrame): Dataframe com colunas de conjuntos
#         column_eval (str): Nome coluna a ser avaliada
#         column_new (str): Nome coluna a ser criada

#     Returns:
#         df: Dataframe com colunas de conjuntos com "[]" adicionados.
#     """
#     # Função para adicionar "[]" nas colunas de conjuntos
#     def process_row(s):  

#         # if s.startswith("("):    
#         #     s = "[" + s[1:]    
#         # elif not s.startswith("["):    
#         #     s = "[" + s  
    
#         # if s.endswith(")"):    
#         #     s = s[:-1] + "]"    
#         # elif not s.endswith("]"):    
#         #     s = s + "]"

#         s = s.replace(' ', '')
#         s = s.replace(',', ', ')

#         if not s.startswith("["):  
#             s = "[" + s  
#         if not s.endswith("]"):  
#             s = s + "]"  


#         # s = s.replace('[', '').replace(']', '')  
#         # if ',' in s:  
#         #     return [int(i.strip()) for i in s.split(',')]  
#         # elif s.isdigit():  
#         #     return int(s)  
#         # else:  
#         #     return s

#         # s = s.replace('[', '').replace(']', '').replace(' ', '')  
#         # if ',' in s:  
#         #     return [int(i) for i in s.split(',')]  
#         # elif s.isdigit():  
#         #     return [int(s)]  
#         # else:  
#         #     return s

#     # Efetivamente adiciona "[]" nas colunas de conjuntos
#     df[column_eval] = df[column_eval].astype(str)
#     df[column_new] = df[column_eval].apply(process_row) 
#     # df[column_new] = df[column_eval].apply(lambda x: f'[{x}]' if ',' in x else x)  
#     # df[column_new] = df[column_eval].apply(lambda x: f'[{", ".join(x.split(","))}]' if ',' in x else x)  

#     return df  