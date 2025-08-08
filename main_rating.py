import subprocess
import re
import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt  # Adicionado para uso do Qt.WaitCursor
import socket
from datetime import datetime  # Adicionado para data/hora

# Salva o IP local em um arquivo de log (mantido para compatibilidade)
try:
    ip = socket.gethostbyname(socket.gethostname())
    with open(r'u:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\ip_log.txt', 'a') as f:
        f.write(f"{ip}\n")
except Exception as e:
    pass

def log_execution(ip, cnpj, ano, pyfile):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(r'u:\RBE\_PÚBLICO (Apenas trocas de arquivos)\Rating - Shared\exec_log.txt', 'a') as f:
        f.write(f"{now} | IP: {ip} | CNPJ: {cnpj} | Ano: {ano} | Arquivo: {pyfile}\n")

class InputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calcular Nota")
        self.layout = QVBoxLayout()

        self.cnpj_label = QLabel("Digite o CNPJ:")
        self.cnpj_input = QLineEdit()
        self.cnpj_input.setInputMask("00.000.000/0000-00;_")  # Adiciona máscara ao input
        self.layout.addWidget(self.cnpj_label)
        self.layout.addWidget(self.cnpj_input)

        self.ano_label = QLabel("Digite o ano (4 dígitos):")
        self.ano_input = QLineEdit()
        self.layout.addWidget(self.ano_label)
        self.layout.addWidget(self.ano_input)

        self.button = QPushButton("Calcular")
        self.button.clicked.connect(self.on_calculate)
        self.layout.addWidget(self.button)

        # Adiciona barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(8)  # 7 arquivos
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        self.setLayout(self.layout)

    def validate_cnpj(self, cnpj):
        return re.fullmatch(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", cnpj)

    def validate_ano(self, ano):
        return re.fullmatch(r"\d{4}", ano)

    def on_calculate(self):
        cnpj = self.cnpj_input.text()
        ano = self.ano_input.text()
        if not self.validate_cnpj(cnpj):
            QMessageBox.warning(self, "Erro", "CNPJ inválido. Tente novamente.")
            return
        if not self.validate_ano(ano):
            QMessageBox.warning(self, "Erro", "Ano inválido. Tente novamente.")
            return

        self.button.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.progress_bar.setValue(0)
        try:
            ip = socket.gethostbyname(socket.gethostname())
            pyfiles = [
                r'Financeiro\calcular_nota_financeiro.py',
                r'Qualitativo\calcular_notas_qualitativo.py',
                r'Quantitativo\Alavancagem\somar_12m.py',
                r'Quantitativo\Alavancagem\calcular_alavancagem.py',
                r'Quantitativo\GRO\calcular_gro.py',
                r'Quantitativo\VarVolume\calcular_mediana.py',
                r'Quantitativo\calcular_notas_quantitativo.py',
                r'calculo_final.py'
            ]
            for idx, pyfile in enumerate(pyfiles, 1):
                subprocess.run(['python', pyfile, cnpj, ano])
                log_execution(ip, cnpj, ano, pyfile)
                self.progress_bar.setValue(idx)
                QApplication.processEvents()  # Atualiza a interface
            QMessageBox.information(self, "Sucesso", f"Cálculos finalizados para o CNPJ: {cnpj}")
        finally:
            QApplication.restoreOverrideCursor()
            self.button.setEnabled(True)
            self.progress_bar.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InputWindow()
    window.show()
    sys.exit(app.exec_())
