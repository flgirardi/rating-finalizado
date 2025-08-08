import csv
from collections import defaultdict
import re

input_files = [
    r"U:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\Quantitativo\VarVolume\lista_cv_2024.csv",
    r"U:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\Quantitativo\VarVolume\lista_cv_2025.csv"
    # Adicione aqui outros arquivos que deseja juntar, por exemplo:

]
output_file = r"U:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\Quantitativo\VarVolume\lista_geral_fmtd.csv"

# Data structure to hold processed data
data = defaultdict(lambda: defaultdict(lambda: {"compra": "", "venda": ""}))
agente_cnpj = {}

def format_cnpj(cnpj):
    # Remove non-digit characters
    digits = re.sub(r'\D', '', cnpj)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"

# Read the input CSVs
for input_file in input_files:
    with open(input_file, mode="r", encoding="latin1") as infile:
        reader = csv.DictReader(infile, delimiter=";")
        for row in reader:
            agente = row["CODIGO_PERFIL_AGENTE"]
            mes = row["MES_REFERENCIA"]
            data[agente][mes]["compra"] = row["CONTRATACAO_COMPRA"]
            data[agente][mes]["venda"] = row["CONTRATACAO_VENDA"]
            agente_cnpj[agente] = row.get("CNPJ", "")

# Prepare the header for the output CSV
header = ["AGENTE", "CNPJ"]
months = sorted({mes for agente in data for mes in data[agente]})
for mes in months:
    header.append(f"{mes}_COMPRA")
    header.append(f"{mes}_VENDA")

# Write the output CSV
with open(output_file, mode="w", encoding="utf-8", newline="") as outfile:
    writer = csv.writer(outfile, delimiter=";")
    writer.writerow(header)
    for agente, meses in data.items():
        cnpj_raw = agente_cnpj.get(agente, "")
        cnpj_fmt = format_cnpj(cnpj_raw)
        row = [agente, cnpj_fmt]
        for mes in months:
            compra = meses[mes]["compra"].replace(".", ",")
            venda = meses[mes]["venda"].replace(".", ",")
            row.append(compra)
            row.append(venda)
        writer.writerow(row)
