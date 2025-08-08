import pandas as pd
from datetime import datetime
import os
import sys

## salvar a nota final e olhar só para ativas na alavancagem

# Recebe o CNPJ e o ano como argumento
if len(sys.argv) < 3:
    print("Erro: forneça um CNPJ e o Ano como argumento.")
    sys.exit(1)
cnpj_arg = sys.argv[1].strip()
ano = sys.argv[2].strip()

# Carrega os dados
db_gro = pd.read_csv(r'Quantitativo\GRO\gro.csv')
db_varvolume = pd.read_csv(r'Quantitativo\VarVolume\mediana_dp.csv')
db_alavancagem = pd.read_csv(r'Quantitativo\Alavancagem\alavancagem.csv')
db_tipo = pd.read_csv(r'Rating\db.csv')

# Normaliza CNPJ para garantir merge correto
for df in [db_gro, db_varvolume, db_alavancagem, db_tipo]:
    df['CNPJ'] = df['CNPJ'].astype(str).str.strip()

# Função genérica para filtrar por CNPJ, Ano e opcionalmente Status
def filtra_por_cnpj_ano(df, cnpj, ano, nome_arquivo, filtrar_status=False):
    ano_col = None
    for col in df.columns:
        if str(col).strip().lower() in ['ano_edit']:
            ano_col = col
            break

    if ano_col is not None:
        filtro = (df['CNPJ'] == cnpj) & (df[ano_col].astype(str).str.strip() == ano)

        # Aplica filtro extra para status se necessário
        if filtrar_status and 'status' in df.columns:
            filtro &= (df['status'].astype(str).str.strip().str.lower() == 'ativa')

        df_filtrado = df[filtro]
        if df_filtrado.empty:
            print(f"Aviso: Ano {ano} não encontrado para CNPJ {cnpj} em {nome_arquivo}.")
            empty_row = {col: '' for col in df.columns}
            empty_row['CNPJ'] = cnpj
            empty_row[ano_col] = ano
            return pd.DataFrame([empty_row])
        return df_filtrado
    else:
        print(f"Aviso: Coluna de ano não encontrada em {nome_arquivo}.")
        empty_row = {col: '' for col in df.columns}
        empty_row['CNPJ'] = cnpj
        return pd.DataFrame([empty_row])

# Filtragem com regra extra para alavancagem.csv
db_gro = filtra_por_cnpj_ano(db_gro, cnpj_arg, ano, 'gro.csv')
db_varvolume = filtra_por_cnpj_ano(db_varvolume, cnpj_arg, ano, 'mediana_dp.csv')
db_alavancagem = filtra_por_cnpj_ano(db_alavancagem, cnpj_arg, ano, 'alavancagem.csv', filtrar_status=True)
db_tipo = db_tipo[db_tipo['CNPJ'] == cnpj_arg]

# Junta os DataFrames pelo CNPJ
df = db_gro.merge(db_varvolume, on='CNPJ', how='outer').merge(db_alavancagem, on='CNPJ', how='outer')
df = df.merge(db_tipo[['CNPJ', 'Tipo (Perfil Empresa)']], on='CNPJ', how='left')

# Remove duplicatas de CNPJ
df = df.drop_duplicates(subset=['CNPJ'])

# --- Funções de cálculo das notas (retornam int ou None) ---
def calc_gro(valor):
    try:
        valor = float(valor)
    except (ValueError, TypeError):
        return None
    if pd.isna(valor):
        return None
    if valor < 30:
        return 5
    elif valor < 40:
        return 4
    elif valor < 50:
        return 3
    elif valor < 75:
        return 2
    else:
        return 1

def calc_varvolume(valor):
    try:
        valor = float(valor)
    except (ValueError, TypeError):
        return None
    if pd.isna(valor):
        return None
    if valor < 0.2:
        return 5
    elif valor < 0.30:
        return 4
    elif valor < 0.40:
        return 3
    elif valor < 0.60:
        return 2
    elif valor < 1:
        return 1
    else:
        return 0

def calc_alavancagem(valor):
    try:
        valor = float(valor)
    except (ValueError, TypeError):
        return None
    if pd.isna(valor):
        return None
    if valor < 10:
        return 5
    elif valor < 20:
        return 4
    elif valor < 30:
        return 3
    elif valor < 40:
        return 2
    elif valor < 50:
        return 1
    else:
        return 0

# Aplica as funções de nota
df['nota_gro'] = df['GRO'].apply(calc_gro) if 'GRO' in df.columns else None
df['nota_varvolume'] = df['Desvio/Mediana_VENDA'].apply(calc_varvolume) if 'Desvio/Mediana_VENDA' in df.columns else None
df['nota_alavancagem'] = df['grau_alavancagem'].apply(calc_alavancagem) if 'grau_alavancagem' in df.columns else None

# Garante que todas sejam inteiros
for col in ['nota_gro', 'nota_varvolume', 'nota_alavancagem']:
    df[col] = df[col].fillna(0).astype(int)

# Função para obter pesos
def get_pesos(tipo):
    if pd.isna(tipo):
        return {'Grau de Alavancagem': 0.5, 'GRO': 0, 'Var Volume': 0.5}
    tipo = str(tipo).strip().lower()
    if tipo == 'trading':
        return {'Grau de Alavancagem': 0.50, 'GRO': 0.0, 'Var Volume': 0.5}
    elif tipo == 'gerador':
        return {'Grau de Alavancagem': 0.65, 'GRO': 0.0, 'Var Volume': 0.35}
    elif tipo == 'consumidor':
        return {'Grau de Alavancagem': 0.0, 'GRO': 0, 'Var Volume': 0.0}
    elif tipo == 'fundo de investimento':
        return {'Grau de Alavancagem': 0.0, 'GRO': 0, 'Var Volume': 0.0}
    else:
        return {'Grau de Alavancagem': 0.5, 'GRO': 0, 'Var Volume': 0.5}

# Calcula nota final (agora tudo é número)
def calcula_nota_final(row):
    pesos = get_pesos(row['Tipo (Perfil Empresa)'])
    return (
        row['nota_gro'] * pesos['GRO'] +
        row['nota_varvolume'] * pesos['Var Volume'] +
        row['nota_alavancagem'] * pesos['Grau de Alavancagem']
    )

df['nota_final'] = df.apply(calcula_nota_final, axis=1)

# Adiciona informações extras
df['DataHora'] = datetime.now().strftime('%Y-%m-%d %H:%M')
df['status'] = 'ativa'
df['Ano_Edit'] = ano

# Salva no CSV de notas
csv_path = r'Quantitativo\notas_quantitativo.csv'
cols_to_save = ['CNPJ', 'Ano_Edit', 'nota_varvolume', 'nota_alavancagem', 'nota_final', 'DataHora', 'status']

if os.path.exists(csv_path):
    df_existente = pd.read_csv(csv_path)
    for idx, row in df.iterrows():
        mask = (
            df_existente['CNPJ'].astype(str).str.strip() == str(row['CNPJ']).strip()
        ) & (
            df_existente['Ano_Edit'].astype(str).str.strip() == str(row['Ano_Edit']).strip()
        )
        if mask.any():
            df_existente.loc[mask, 'status'] = 'editada'
    df_existente.to_csv(csv_path, index=False)
    df[cols_to_save].to_csv(csv_path, mode='a', header=False, index=False)
else:
    df[cols_to_save].to_csv(csv_path, index=False)