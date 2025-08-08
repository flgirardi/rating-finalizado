import urllib.request
import csv
import io

a2024 = "https://pda-download.ccee.org.br/TSPOxvC8S8W11OdMf6514w/content"
a2025 = "https://pda-download.ccee.org.br/xZ0sVKW0RyO8FzQehagBXw/content"

try:
    # URL ajustada para buscar os dados do recurso 'lista_perfil_2025'
    url = a2025
    with urllib.request.urlopen(url) as response:
        data = response.read().decode('latin-1')  # Decodifica usando 'latin-1' para lidar com caracteres especiais
        csv_reader = csv.reader(io.StringIO(data))  # Lê o conteúdo CSV
        
        # Salva os dados em um arquivo CSV local
        with open(r'U:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\Quantitativo\VarVolume\lista_cv_2025.csv', mode='w', newline='', encoding='latin-1') as file:
            writer = csv.writer(file)
            for row in csv_reader:
                writer.writerow(row)  # Escreve cada linha no arquivo CSV
        print("Dados salvos em 'lista_cv_2025.csv'")
except urllib.error.URLError as e:
    print(f"Erro ao acessar a URL: {e}")
except UnicodeDecodeError as e:
    print(f"Erro de codificação: {e}")