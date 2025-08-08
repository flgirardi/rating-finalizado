import csv
import statistics
import re
import sys
from datetime import datetime

# Caminho do arquivo CSV
csv_path = r'Quantitativo\VarVolume\lista_geral_fmtd.csv'

def format_cnpj(cnpj):
    digits = re.sub(r'\D', '', cnpj)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return cnpj

def calcula_mediana(cnpj):
    resultados = {}
    cnpj_formatado = format_cnpj(cnpj)
    periodos = {}
    for tipo in ['COMPRA', 'VENDA']:
        with open(csv_path, encoding='utf-8') as f:
            reader = list(csv.DictReader(f, delimiter=';'))
            linhas_cnpj = [row for row in reader if row['CNPJ'] == cnpj_formatado]

            if not linhas_cnpj:
                print('CNPJ não encontrado.')
                resultados[tipo] = None
                periodos[tipo] = None
                continue

            colunas = [col for col in linhas_cnpj[0].keys() if col.endswith('_' + tipo)]
            colunas_ultimos6 = colunas[-6:]

            # Determina o período a partir dos nomes das colunas
            # Espera-se que o nome da coluna seja algo como '202311_COMPRA'
            if colunas_ultimos6:
                try:
                    inicio = colunas_ultimos6[0].split('_')[0]
                    fim = colunas_ultimos6[-1].split('_')[0]
                    periodo = f"{inicio}-{fim}"
                except Exception:
                    periodo = ""
            else:
                periodo = ""
            periodos[tipo] = periodo

            # Primeiro, soma os valores de todas as linhas para cada coluna dos últimos 6 meses
            valores_somados = []
            for col in colunas_ultimos6:
                soma = 0.0
                encontrou_valor = False
                for row in linhas_cnpj:
                    val = row[col].strip().replace(',', '.')
                    if val != '':
                        try:
                            soma += float(val)
                            encontrou_valor = True
                        except ValueError:
                            continue
                if encontrou_valor:
                    valores_somados.append(soma)
                else:
                    valores_somados.append(0.0)

            # Agora, verifica se algum mês ficou como 0 ou vazio após a soma
            alguma_coluna_vazia_ou_zero = any(v == 0 or v == 0.0 for v in valores_somados)

            if alguma_coluna_vazia_ou_zero:
                print(f'[{tipo}] Alguma coluna dos últimos 6 meses está vazia ou igual a zero para o CNPJ {cnpj}. Salvando valores como 0.')
                resultados[tipo] = (0, 0, 0)
                continue

            if valores_somados:
                print(f'[{tipo}] Valores utilizados para o cálculo: {valores_somados}')
                mediana = statistics.median(valores_somados)
                if len(valores_somados) > 1:
                    desvio = statistics.stdev(valores_somados)
                else:
                    desvio = 0.0
                print(f'[{tipo}] Mediana dos últimos 6 meses para {cnpj}: {mediana}')
                print(f'[{tipo}] Desvio padrão dos últimos 6 meses para {cnpj}: {desvio}')
                if mediana != 0:
                    desvio_mediana = desvio/mediana
                    print(f'[{tipo}] Desvio/Mediana: {desvio_mediana}')
                else:
                    desvio_mediana = None
                    print(f'[{tipo}] Mediana é zero, não é possível calcular Desvio/Mediana.')
                resultados[tipo] = (mediana, desvio, desvio_mediana)
            else:
                print(f'[{tipo}] Não há valores válidos para calcular a mediana.')
                resultados[tipo] = None
    # Retorna também os períodos
    return resultados, periodos

def salva_resultado_csv(cnpj, ano, resultados, periodos):
    output_path = r'Quantitativo\VarVolume\mediana_dp.csv'
    header = [
        'CNPJ', 'Ano_edit', 'Periodo', 'DataHora',
        'Mediana_COMPRA', 'Desvio_COMPRA', 'Desvio/Mediana_COMPRA',
        'Mediana_VENDA', 'Desvio_VENDA', 'Desvio/Mediana_VENDA'
    ]
    # Determina o período a ser salvo (prioriza COMPRA, senão VENDA, senão vazio)
    periodo = periodos.get('COMPRA') or periodos.get('VENDA') or ''
    datahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    row = [cnpj, ano, periodo, datahora]
    for tipo in ['COMPRA', 'VENDA']:
        if resultados.get(tipo):
            mediana, desvio, desvio_mediana = resultados[tipo]
            row.extend([mediana, desvio, desvio_mediana])
        else:
            row.extend([None, None, None])

    # Lê todas as linhas existentes (se houver)
    linhas = []
    cnpj_encontrado = False
    try:
        with open(output_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=',')
            linhas = list(reader)
    except FileNotFoundError:
        linhas = []

    # Atualiza ou adiciona a linha do CNPJ e Ano
    novas_linhas = []
    header_written = False
    for l in linhas:
        if not header_written and l and l[0] == 'CNPJ':
            novas_linhas.append(header)
            header_written = True
        elif l and l[0] == cnpj and len(l) > 1 and l[1] == ano:
            novas_linhas.append(row)
            cnpj_encontrado = True
        elif l and l[0] != 'CNPJ':
            novas_linhas.append(l)
    if not cnpj_encontrado:
        if not header_written:
            novas_linhas.insert(0, header)
        novas_linhas.append(row)

    # Escreve tudo de volta
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=',')
            for l in novas_linhas:
                writer.writerow(l)
    except Exception as e:
        print(f'Erro ao salvar CSV: {e}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Erro: forneça um CNPJ como argumento.")
        sys.exit(1)
    cnpj = sys.argv[1].strip()
    ano = sys.argv[2].strip()
    resultados, periodos = calcula_mediana(cnpj)
    salva_resultado_csv(format_cnpj(cnpj), ano, resultados, periodos)
