import pandas as pd
import dotenv
import os
import sys
from datetime import datetime
import openpyxl

#-------------------------------------------------------------------------------
#-------CALCULO FEITO COM VOLUME ATUALIZADO - ADD LOGICA SALVAR ANTIGAS---------
#-------------------------------------------------------------------------------

if len(sys.argv) < 2:
    print("Uso: python calculo_final.py <CNPJ> <Ano_edit>")
    sys.exit(1)
cnpj = sys.argv[1]
ano = sys.argv[2]

#cnpj = "13.338.734/0001-27"
#ano = 2024
print(ano)
CNPJ_FILTRADO = cnpj

Ano_edit = int(ano)  # Usa o valor exatamente como veio
print(Ano_edit)
filepath_pl = r'Rating\db.csv'
filepath_env = r'Quantitativo\Alavancagem\horas.env'
filepath_pld = r'PLD.xlsm'
filepath_volume = r'Quantitativo\Alavancagem\soma_12m_por_cnpj.csv'

dotenv.load_dotenv(filepath_env)

wb = openpyxl.load_workbook(filepath_pld, data_only=True)
ws = wb['Planilha1']
pld_12m = float(ws['C14'].value)
horas_ano = int(os.environ.get('horas_ano', 0))

db_pl = pd.read_csv(filepath_pl, sep=',')

# Apenas garante que as colunas sejam string
db_pl['CNPJ'] = db_pl['CNPJ']
db_pl['Ano_edit'] = db_pl['Ano_edit']
db_volume = pd.read_csv(filepath_volume, sep=',')
db_volume['CNPJ'] = db_volume['CNPJ']


# Filtra os dados
db_pl = db_pl[(db_pl['CNPJ'] == CNPJ_FILTRADO) & (db_pl['Ano_edit'] == Ano_edit)]
db_volume = db_volume[db_volume['CNPJ'] == CNPJ_FILTRADO]

# Merge dos dados
df = pd.merge(db_pl, db_volume, on='CNPJ', how='left')

# Cálculo do grau de alavancagem
pl = pd.to_numeric(df['Patrimônio Líquido'], errors='coerce')
volume = pd.to_numeric(df['TOTAL_VENDA_ULT12'], errors='coerce') /12
df['grau_alavancagem'] = (pld_12m * volume * horas_ano) / pl
df.loc[pl <= 0, 'grau_alavancagem'] = 51

df['periodo'] = db_volume['RANGE_MESES']

# Adiciona colunas
output_path = r'Quantitativo\Alavancagem\alavancagem.csv'
now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

# Carrega CSV existente se houver
if os.path.exists(output_path):
    df_out = pd.read_csv(output_path)
    mask = (df_out['CNPJ'] == CNPJ_FILTRADO) & (df_out['Ano_edit'] == Ano_edit)
    df_out.loc[mask, 'status'] = 'editada'
else:
    df_out = pd.DataFrame()

# Prepara nova entrada
df['status'] = 'ativa'
df['DataHora'] = now_str
df['Ano_edit'] = Ano_edit

df_final = df[['CNPJ', 'grau_alavancagem', 'periodo', 'Ano_edit', 'status', 'DataHora']]

# Ensure 'Ano_edit' has exactly 4 digits
df_final = df_final[df_final['Ano_edit'].astype(str).str.len() == 4]

# Concatena e salva
if not df_out.empty:
    df_out = pd.concat([df_out, df_final], ignore_index=True)
else:
    df_out = df_final

df_out.to_csv(output_path, index=False)