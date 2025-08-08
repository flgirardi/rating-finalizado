import pandas as pd
import sys
import os
from datetime import datetime

# Garante que um CNPJ foi fornecido
if len(sys.argv) < 2:
    print("Erro: forneça um CNPJ como argumento.")
    sys.exit(1)
cnpj_arg = sys.argv[1].strip()
ano = sys.argv[2].strip()

filepath = r'Rating\db.csv'
db = pd.read_csv(filepath, sep=',')

# Filtra apenas o CNPJ fornecido
db = db[db['CNPJ'].astype(str).str.strip() == cnpj_arg]

# Filtra apenas o ano_edit fornecido
db = db[db['Ano_edit'].astype(str).str.strip() == ano]

# Considere apenas linhas com status == "ativa"
if 'status' in db.columns:
    db = db[db['status'] == 'ativa']

pl = pd.to_numeric(db['Patrimônio Líquido'], errors='coerce')
rol = pd.to_numeric(db['Receita Operacional Líquida(ROL)'], errors='coerce')
cnpj = db['CNPJ']
ano_edit = db['Ano_edit']

# Calcula GRO normalmente, mas se PL <= 0, define GRO = 76
gro = (rol / pl).where(pl > 0, 76)

# Cria um novo DataFrame com CNPJ, Ano_edit, GRO, Data/Hora e Status
resultado = pd.DataFrame({
    'CNPJ': cnpj,
    'Ano_edit': ano_edit,
    'GRO': gro,
    'DataHora': datetime.now().strftime('%Y-%m-%d %H:%M'),
    'status': 'ativa'
})

csv_path = r'Quantitativo\GRO\gro.csv'

if os.path.exists(csv_path):
    df_existente = pd.read_csv(csv_path)
    if not resultado.empty:
        # Para cada linha do resultado, verifica se já existe CNPJ + Ano_edit
        for idx, row in resultado.iterrows():
            mask = (
                df_existente['CNPJ'].astype(str).str.strip() == str(row['CNPJ']).strip()
            ) & (
                df_existente['Ano_edit'].astype(str).str.strip() == str(row['Ano_edit']).strip()
            )
            if mask.any():
                df_existente.loc[mask, 'status'] = 'editada'
        # Salva o CSV sobrescrevendo
        df_existente.to_csv(csv_path, index=False)
        # Adiciona as novas linhas com status "ativa"
        resultado.to_csv(csv_path, mode='a', header=False, index=False)
    else:
        print("Nenhum resultado encontrado para o CNPJ informado.")
else:
    resultado.to_csv(csv_path, index=False)