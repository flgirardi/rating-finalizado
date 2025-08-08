import pandas as pd
from datetime import datetime
import os
import sys

# Recebe o CNPJ como argumento
if len(sys.argv) <= 2:
    print("Erro: forneça um CNPJ como argumento.")
    sys.exit(1)
cnpj_arg = sys.argv[1].strip()
ano_arg = sys.argv[2].strip()
db = pd.read_csv(r'Rating\db.csv')

# Filtra apenas o CNPJ e Ano_edit fornecidos
db['CNPJ'] = db['CNPJ'].astype(str).str.strip()
if 'Ano_edit' in db.columns:
    db['Ano_edit'] = db['Ano_edit'].astype(str).str.strip()
    db = db[(db['CNPJ'] == cnpj_arg) & (db['Ano_edit'] == ano_arg)]
else:
    db = db[db['CNPJ'] == cnpj_arg]
    db['Ano_edit'] = ano_arg  # adiciona coluna se não existir

# Considere apenas linhas com status == "ativa"
if 'status' in db.columns:
    db = db[db['status'] == 'ativa']

# Verifica se existe a coluna 'Ano_edit'
if 'Ano_edit' not in db.columns:
    print("Erro: coluna 'Ano_edit' não encontrada no arquivo de origem.")
    sys.exit(1)

def calc_atuação(valor):
    if valor == "SI":
        return 0
    elif valor == "BLOQUEADA":
        return 0
    elif valor == "RJ/QUEBRADA":
        return 0
    elif valor == "TRADING PURA":
        return 1
    elif valor == "TRADING C/ GESTÃO":
        return 3
    elif valor == "TRADING C/ GERAÇÃO":
        return 4
    elif valor == "TRADING C/ GESTÃO E GERAÇÃO":
        return 5
    elif valor == "GERADORES NÃO ATIVO":
        return 0
    elif valor == "GERADORES PEQUENO ATIVO FORA MRE":
        return 2
    elif valor == "GERADORES PEQUENO ATIVO DENTRO MRE":
        return 3
    elif valor == "GERADORES MÉDIO":
        return 4
    elif valor == "GERADOR GRANDE":
        return 5
    elif valor == "GRANDES GRUPOS":
        return 5
    elif valor == "FUNDO INVESTIMENTO PEQUENO":
        return 0
    elif valor == "FUNDO INVESTIMENTO MÉDIO":
        return 3
    elif valor == "FUNDO INVESTIMENTO GRANDE":
        return 4
    elif valor == "FUNDO INVESTIMENTO MÉDIO INTERNACIONAL":
        return 2
    elif valor == "CONSUMIDORES":
        return 5
    else:
        return None  # ou algum valor padrão

def calc_experiencia_tm(valor):
    if valor == "SI":
        return 0
    elif valor == "FECHOU":
        return 0
    elif valor == "QUEBROU":
        return 0
    elif valor == "JUNIOR":
        return 0
    elif valor == "PLENO":
        return 2
    elif valor == "SENIOR":
        return 3
    elif valor == "ESPECIALISTAS":
        return 4
    elif valor == "EXPERT":
        return 5
    elif valor == "EXCELÊNCIA":
        return 5
    else:
        return None  # ou algum valor padrão   

def calc_governança(valor):
    if valor == "LIMITADA":
        return 2.5
    elif valor == "SA":
        return 5
    
def calc_tempoOp(valor):
    if valor <= 2:
        return 0
    elif valor <= 4:
        return 1
    elif valor <= 5:
        return 2
    elif valor <= 10:
        return 3
    elif valor <= 15:
        return 4
    else:
        return 5

resultados = []

for idx, row in db.iterrows():
    cnpj = row['CNPJ']
    ano_edit = row['Ano_edit']
    atuacao = str(row['Atuação Empresa']).strip().upper()
    exp_tm = str(row['Experiência Trading e Middle']).strip().upper()
    tipo_governanca = str(row['Tipo Governança']).strip().upper()
    data_contrato = pd.to_datetime(row['Data Contrato Social'], format='%d/%m/%Y').date()
    tipo_empresa = row.get('Tipo (Perfil Empresa)', '')
    currentDate = datetime.today().date()
    tempo_op = (currentDate - data_contrato).days / 360

    tipo = tipo_empresa.strip().lower() if isinstance(tipo_empresa, str) else ""

    if tipo == 'trading':
        pesos = {
            'Atuação da Empresa': 0.30,
            'Experiência Trading e Middle': 0.30,
            'Tipo de Governança': 0.10,
            'Tempo de Operação': 0.3
        }
    elif tipo == 'gerador':
        pesos = {
            'Atuação da Empresa': 0.30,
            'Experiência Trading e Middle': 0.10,
            'Tipo de Governança': 0.1,
            'Tempo de Operação': 0.30
        }
    elif tipo == 'consumidor':
        pesos = {
            'Atuação da Empresa': 0.50,
            'Experiência Trading e Middle': 0,
            'Tipo de Governança': 0.25,
            'Tempo de Operação': 0.25
        }
    elif tipo == 'fundo de investimento':
        pesos = {
            'Atuação da Empresa': 0.50,
            'Experiência Trading e Middle': 0.25,
            'Tipo de Governança': 0.00,
            'Tempo de Operação': 0.25
        }
    else:
        pesos = {
            'Atuação da Empresa': 0.30,
            'Experiência Trading e Middle': 0.2,
            'Tipo de Governança': 0.2,
            'Tempo de Operação': 0.3
        }

    nota_atuação = calc_atuação(atuacao)
    nota_exp_tm = calc_experiencia_tm(exp_tm)
    nota_tipo_governanca = calc_governança(tipo_governanca)
    nota_tempo_op = calc_tempoOp(tempo_op)

    nota_final = (
        (nota_atuação if nota_atuação is not None else 0) * pesos['Atuação da Empresa'] +
        (nota_exp_tm if nota_exp_tm is not None else 0) * pesos['Experiência Trading e Middle'] +
        (nota_tipo_governanca if nota_tipo_governanca is not None else 0) * pesos['Tipo de Governança'] +
        (nota_tempo_op if nota_tempo_op is not None else 0) * pesos['Tempo de Operação']
    )

    resultados.append({
        'CNPJ': cnpj,
        'Ano_edit': ano_arg,
        'nota_atuação': nota_atuação,
        'nota_exp_tm': nota_exp_tm,
        'nota_tipo_governanca': nota_tipo_governanca,
        'nota_tempo_op': nota_tempo_op,
        'nota_final': nota_final,
        'DataHora': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'status': 'ativa'
    })

df_resultados = pd.DataFrame(resultados)
csv_path = r'Qualitativo\notas_qualitativo.csv'
if os.path.exists(csv_path):
    df_existente = pd.read_csv(csv_path)
    for idx, row in df_resultados.iterrows():
        mask = (
            df_existente['CNPJ'].astype(str).str.strip() == str(row['CNPJ']).strip()
        ) & (
            df_existente['Ano_edit'].astype(str).str.strip() == str(row['Ano_edit']).strip()
        )
        if mask.any():
            df_existente.loc[mask, 'status'] = 'editada'
    df_existente.to_csv(csv_path, index=False)
    df_resultados.to_csv(csv_path, mode='a', header=False, index=False)
else:
    df_resultados.to_csv(csv_path, index=False)