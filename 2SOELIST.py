#%%
# Importa bibliotecas
import pandas as pd
import glob
from collections import Counter
from datetime import datetime


'''
Algoritmo que importa a tagname completa e o acrônimo para Open e Close (0 e 1) do BD_SCADA
'''


# Seta a pasta BD
arquivos = glob.glob('BD\*.xlsx') 


# Subnam, PntNam, Acronimo
sostat = pd.read_excel(arquivos[0], sheet_name='RANGER_SOSTAT',usecols=[2,3,18])

# SAT1NO, STEXT0, STEXT1
sosat = pd.read_excel(arquivos[0], sheet_name='RANGER_SOSAT1',usecols=[0,11,12])

# Agrupa o acronimo 
sostat_sosat = pd.merge(sostat, sosat, left_on=['ACRONM'], right_on=['SAT1NO'],how='outer')
sostat_sosat['Tagname'] = sostat_sosat['SUBNAM']+'.'+sostat_sosat['PNTNAM']

# Cria a variável com os dados do BD do SCADA desejados
sostat_sosat = sostat_sosat.filter(items=['Tagname', 'STEXT0', 'STEXT1'])


# Mais fácil entender este nome
sostat = sostat_sosat


# Deleta o que ñ vou usar daqui pra FRENTE
del sostat_sosat, arquivos, sosat




# %%
'''
Cria um df contendo todas as correspondências do scratch no BD_SCADA
Este df também possui uma coluna Event Flag
    Que consiste em destacar quantas correspondências um alarme 
    do scratch possui no BD_SCADA
'''


# importa o scratchpad
scratch = pd.read_csv('scratch 17.txt', index_col=False)

# Extrai a coluna Event do scratchpad importado
alarme = pd.Series(index=scratch.index, dtype=str)


# Deixa o alarme no formato padrão da SOELIST, porém ainda truncado
for index, row in scratch.iterrows():
    
    # Caso o ponto seja do tipo STATUS PT
    if 'STATUS PT.' in row['Event']:
        # 12 primeiros caracteres
        alarme[index] = row['Event'][12:]

        # Separa em subNam e pointNam, por 8 caracteres
        subNam   = alarme[index][:8]
        pointNam = alarme[index][8:]

        # Strip para retirar espaços em branco
        subNam = str.strip(subNam)
        pointNam = str.strip(pointNam)
        
        # Formato padrão da SOELIST subnam.pointnam (ainda truncado!)
        alarme[index] = subNam + '.' + pointNam
    else:
        # Caso seja comentário, ANALOG PT> ou outra coisa
        # Ñ há a necessidade de colocar no formato da SOELIST
        # Pois este ponto não existe no df soelist
        alarme[index] = row['Event']


# Deleta variáveis auxiliares
del index, row


# Inicializa dois dicionários

# Corresp representa todas as correspondências do alarme do DTS no BD SCADA
corresp = dict()

# Status é o STEXE0 e STEXT1 para cada correspondência
status = dict()

# Inicializa o dataframe final com o cabeçalho padrão da SOELIST
resultado = pd.DataFrame(columns=["Event Flag", "Event Time", "Previous Event", "Status", "Tagname"])


# Percorre a Serie alarme
for index, i in enumerate(alarme):
    # Valores True para as correspondências do Tagname do scratchpad
    # Caso não haja correspondência, ou o valor ñ seja uma string, preenche com False
    aux = sostat['Tagname'].map(lambda x: i in x if isinstance(x,str) else False)

    # Caso tenha encontrado algum evento
    if aux.any():
        # Colocar a SOSTAT encontrada no formato da SOELIST
        # Divide o Event em subnam.pointnam        
        subNam   = sostat['Tagname'][aux].map(lambda x: x.split(".", maxsplit=1)[0] if isinstance(x, str) else "")
        pointNam = sostat['Tagname'][aux].map(lambda x: x.split(".", maxsplit=1)[1] if isinstance(x, str) else "")

        # strip em possíveis caracteres espúrios
        pointNam = pointNam.str.strip() 
        subNam   = subNam.str.strip() 

        # Sendo subnam necessariamente com 8 caracteres, preenchidos com espaços em branco
        subNam   = subNam + ' '*(8 - len(subNam.values[0]))

        # Preenche o dicionário corresp com todas as correspondências do ponto no formato da SOELIST
        corresp = subNam + '.' + pointNam

        # CLOSE(1) ou OPEN(0)
        # Preenche o status com o texto correspondente
        if "CLOSED" in scratch.iloc[index]['Description']:
            status = list(sostat['STEXT1'][aux])

        else:
            status = list(sostat['STEXT0'][aux])

        # Start time vindo do DTS
        startTime = pd.Series(scratch.iloc[index]['Start Time'])

        # Formato do SOELIST
        startTime = datetime.strptime(startTime[0], "%d/%b/%Y %H:%M:%S")
        startTime = startTime.strftime("%d/%m/%y %H:%M:%S") + '.000'

        # Índice(s) do scratch da(s) correspondência(s) na SOSTAT 
        flag = [index]*len(corresp)

        # Salva no DF final todas as correspondências com seus respectivos status e start time
        aux_df = pd.DataFrame({"Event Flag": flag, "Event Time": startTime, "Status": status, "Tagname": corresp})
        resultado = pd.concat([resultado, aux_df], ignore_index=True)

    # Caso encontre nenhuma correspondência
    else:
        # Start time vindo do DTS
        startTime = pd.Series(scratch.iloc[index]['Start Time'])

        # Formato do SOELIST
        startTime = datetime.strptime(startTime[0], "%d/%b/%Y %H:%M:%S")
        startTime = startTime.strftime("%d/%m/%y %H:%M:%S") + '.000'

        # Detecta se isso se trata de um comentário (flag = -1)
        # Ou de um ponto que deverá ser inserido pelo usuário (flag = -2)
        if ' '*48 in scratch['Description'].iloc[index]:
            # Preenche o status com vazio
            status = pd.Series('')
            flag = -1
        else:
            # Preenche o status com set ou setoverride xx.xx
            status = scratch['Description'].iloc[index]
            flag = -2

        # O dicionário corresp é preenchido com o comentário literal do Event
        corresp   = pd.Series(scratch.iloc[index]['Event'])

        # Preenche o dataframe com o comentário encontrado
        aux_df = pd.DataFrame({"Event Flag": flag, "Event Time": startTime, "Status": status, "Tagname": corresp})
        resultado = pd.concat([resultado, aux_df], ignore_index=True)


# Deleta variáveis auxiliares
del corresp, status, aux_df, startTime, i, index, flag, aux
del alarme, scratch



#%%
'''
Algoritmo que plota um df em um arquivo xlsx
Event Flag
    +n para as correspondências dos scratchpads na soelist
    -1 para comentários
    -2 para eventos não encontrados no BD_SCADA (ANALOG PT. e.g.)

Este xlsx é separado por cores, para facilitar o usuário a preencher a soelist corretamente
    AZUL para alarmes com correspondência única
    VERDE claro/escuro para alarmes com mais de uma correspondência
    LARANJA para comentários
    VERMELHO para alarmes sem correspondência no sostat

'''

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Criar o arquivo XLSX de saída
output_filename = 'LISTA.xlsx'
resultado.to_excel(output_filename, index=False, sheet_name='Sheet1')


# Carregar o arquivo XLSX com openpyxl
workbook = load_workbook(output_filename)
worksheet = workbook['Sheet1']


# Definir as cores de preenchimento
orang_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='lightUp')
green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
dark_green_fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
blue_fill = PatternFill(start_color='91BCE3', end_color='91BCE3', fill_type='solid')
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')


# Pintar as linhas de acordo com as condições
event_flag_counter = Counter(resultado['Event Flag'])

for row_idx, event_flag in enumerate(resultado['Event Flag'], start=2):

    if event_flag == -1:
        for cell in worksheet[row_idx]:
            cell.fill = orang_fill

    elif event_flag == -2:
        for cell in worksheet[row_idx]:
            cell.fill = red_fill

    elif event_flag_counter[event_flag] > 1:
        if event_flag % 2 == 0:
            for cell in worksheet[row_idx]:
                cell.fill = green_fill

        else:
            for cell in worksheet[row_idx]:
                cell.fill = dark_green_fill

    else:
        for cell in worksheet[row_idx]:
            cell.fill = blue_fill


# Salva o xlsx colorido com o nome de LISTA.xlsx
workbook.save(output_filename)

# Deleta variáveis auxiliares
del blue_fill, cell, dark_green_fill, event_flag, event_flag_counter, green_fill, orang_fill
del output_filename, PatternFill, row_idx, workbook, worksheet