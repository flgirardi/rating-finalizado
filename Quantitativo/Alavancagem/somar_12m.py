import pandas as pd
import sys

if len(sys.argv) <= 1:
    print("Erro: forneça um CNPJ e o Ano como argumentos.")
    sys.exit(1)
    
cnpj = sys.argv[1].strip()
CNPJ_FILTRADO = cnpj  # altere para o CNPJ desejado

# Lê o CSV com separador ';' e vírgula como decimal
df = pd.read_csv(r'Quantitativo\VarVolume\lista_geral_fmtd.csv', sep=';', decimal=',')

# Filtra apenas o CNPJ desejado
df = df[df['CNPJ'] == CNPJ_FILTRADO]

# Ordena as colunas de compra/venda cronologicamente
colunas_compra = sorted([col for col in df.columns if col.endswith('COMPRA')])
colunas_venda = sorted([col for col in df.columns if col.endswith('VENDA')])

# Seleciona as últimas 12 colunas de cada tipo (ordem cronológica)
ultimos_12_compra = colunas_compra[-12:]
ultimos_12_venda = colunas_venda[-12:]
ultimos_12_colunas = ultimos_12_compra + ultimos_12_venda

# Descobre o range dos meses (ex: 202408-202501)
def extrai_mes(col):
    return col.split('_')[0]
mes_inicio = extrai_mes(ultimos_12_compra[0])
mes_fim = extrai_mes(ultimos_12_compra[-1])
range_meses = f"{mes_inicio}-{mes_fim}"

# Agrupa por CNPJ e soma as colunas dos últimos 12 meses
somas = df.groupby('CNPJ')[ultimos_12_colunas].sum(min_count=1)

# Soma separada para COMPRA e VENDA dos últimos 12 meses
somas['TOTAL_COMPRA_ULT12'] = somas[ultimos_12_compra].sum(axis=1, skipna=True)
somas['TOTAL_VENDA_ULT12'] = somas[ultimos_12_venda].sum(axis=1, skipna=True)
somas['RANGE_MESES'] = range_meses

resultado = somas[['TOTAL_COMPRA_ULT12', 'TOTAL_VENDA_ULT12', 'RANGE_MESES']]

# Salva o resultado em um novo CSV
resultado.to_csv(r'Quantitativo\Alavancagem\soma_12m_por_cnpj.csv')
