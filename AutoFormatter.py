import os, sys
import openpyxl
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

#file = input('Nome do arquivo: ')

print("--------------------------------")
print("          CARREGANDO . . .      ")
print("--------------------------------")

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)


os.chdir(application_path)
for root, dirs, files in os.walk("."):
    for filename in files:
        if filename.endswith('.xlsx'):
            f = filename

#Encontrar o arquivo .xlsx
config_path = os.path.join(application_path, f)
print(config_path)

#carregar arquivo 
wb = load_workbook(config_path)
ws = wb.active

#deletar colunas
ws.delete_cols(1,4)
ws.delete_cols(2,1)
ws.delete_cols(4,4)
ws.delete_cols(5,18)
wb.save(filename = 'Planilha_model.xlsx')

df = pd.read_excel('Planilha_model.xlsx')
print(df)

df['Abertura'] = pd.to_datetime(df['Abertura'], format='%d/%m/%Y')
df['Abertura'] = df['Abertura'].dt.strftime('%d/%m/%Y')

df.sort_values(by='Abertura',  ascending = True)

print(df)

df.to_excel('Planilha_model.xlsx', sheet_name='Planilha1')







