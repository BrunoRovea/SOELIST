#%%
'''
#

'''
# Importa bibliotecas
import pandas as pd
import glob


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
#

'''
# importa o scratchpad
scratch = pd.read_csv('scratch 17.txt', index_col=False)
# scratch = scratch.drop('Index', axis=1)


alarme = pd.Series(index=scratch.index, dtype=str)



for index, row in scratch.iterrows():
    if 'STATUS PT.' in row['Event']:
        alarme[index] = row['Event'][12:]

        subNam   = alarme[index][:8]
        pointNam = alarme[index][8:]

        subNam = str.strip(subNam)
        pointNam = str.strip(pointNam)
        
        alarme[index] = subNam + '.' + pointNam
    else:
        alarme[index] = row['Event']


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
        # Divide o Event em subnam.pointnam        
        subNam   = sostat['Tagname'][aux].map(lambda x: x.split(".", maxsplit=1)[0] if isinstance(x, str) else "")

        pointNam = sostat['Tagname'][aux].map(lambda x: x.split(".", maxsplit=1)[1] if isinstance(x, str) else "")

        pointNam = pointNam.str.strip() 
        subNam   = subNam.str.strip() 

        # Sendo subnam necessariamente com 8 caracteres, preenchidos com espaços em branco
        subNam   = subNam + ' '*(8 - len(subNam.values[0]))

        # Preenche o dicionário corresp com todas as correspondências do ponto no formato da SOELIST
        corresp = subNam + '.' + pointNam

        # CLOSE ou OPEN (0) ou (1), sostat
        # Preenche o status com o texto correspondente

        if "CLOSED" in scratch.iloc[index]['Description']:
            status = list(sostat['STEXT1'][aux])

        else:
            status = list(sostat['STEXT0'][aux])

        # Preenche o start time de todas as correspondências do sostat
        startTime = [scratch.iloc[index]['Start Time'] + '.000',] * len(corresp)

        flag = [index]*len(corresp)

        # Salva no DF final todas as correspondências com seus respectivos status e start time
        aux_df = pd.DataFrame({"Event Flag": flag, "Event Time": startTime, "Status": status, "Tagname": corresp})
        resultado = pd.concat([resultado, aux_df], ignore_index=True)


    else:
        # Start time no formato padrão da SOELIST
        startTime = pd.Series(scratch.iloc[index]['Start Time']) + '.000'

        flag = -1

        # Status é preenchido com uma string vazia
        status    = pd.Series('')

        # O dicionário corresp é preenchido com o comentário literal do Event
        corresp   = pd.Series(scratch.iloc[index]['Event'])

        # Preenche o dataframe com o comentário encontrado
        aux_df = pd.DataFrame({"Event Flag": flag, "Event Time": startTime, "Status": status, "Tagname": corresp})
        resultado = pd.concat([resultado, aux_df], ignore_index=True)

del corresp, status, aux_df, startTime, i, index, flag, aux, pointNam, subNam
del alarme, scratch

# Escreve o arquivo csv com a SOELIST completa
# resultado.to_csv('SOELIST.csv')
resultado.to_excel('SOELIST.xlsx')

#%%






from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Ler o arquivo XLSX
filename = 'SOELIST.xlsx'
df = pd.read_excel(filename)

# Criar o arquivo XLSX de saída
output_filename = 'SOELIST_pintado.xlsx'

# Carregar o DataFrame para o arquivo XLSX
df.to_excel(output_filename, index=False, sheet_name='Sheet1')

# Carregar o arquivo XLSX com openpyxl
workbook = load_workbook(output_filename)
worksheet = workbook['Sheet1']

# Definir a cor vermelha
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

# Pintar as linhas de vermelho com base na condição na coluna "Event Flag"
for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
    if row[0] == -1:
        for cell in worksheet.iter_rows(min_row=row[0], max_row=row[0]):
            for c in cell:
                c.fill = red_fill

# Salvar o arquivo XLSX de saída
workbook.save(output_filename)

print(f"As linhas do arquivo '{filename}' foram pintadas de vermelho e salvas em '{output_filename}'.")
