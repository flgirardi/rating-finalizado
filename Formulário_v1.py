
#---------------------------- -----------------------------------
#                            IMPORTS                             
#---------------------------- -----------------------------------

from PyQt5 import QtWidgets, QtCore, QtGui
import sys
from datetime import datetime
import csv
import os
import ast
import socket

#---------------------------- -----------------------------------
#                        CLASSE PRINCIPAL                        
#---------------------------- -----------------------------------
class MainWindow(QtWidgets.QMainWindow):
    #---------------------------- -----------------------------------
    #                        CONSTRUTOR E INICIALIZAÇÃO             
    #---------------------------- -----------------------------------
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Análise de Rating - Contrapartes")
        self.setGeometry(100, 100, 1200, 700)
        self.setWindowIcon(QtGui.QIcon("icone.ico"))  
        self.atualizado_em = None  
        self.acrc_opcoes = self.carregar_opcoes_acrc()
        self.init_ui()
        self.bloquear_campos_iniciais()

    #---------------------------- -----------------------------------
    #                  CARREGAMENTO DE OPÇÕES CSV                   
    #---------------------------- -----------------------------------
    def carregar_opcoes_acrc(self):
        acrc_path = os.path.join(os.path.dirname(__file__), "Rating", "listas.csv")
        opcoes = []
        if os.path.exists(acrc_path):
            with open(acrc_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row and row[0].strip():
                        opcoes.append(row[0].strip())
        opcoes = list(dict.fromkeys(opcoes))
        return opcoes

    def carregar_opcoes_acrc_coluna(self, coluna_nome):
        acrc_path = os.path.join(os.path.dirname(__file__), "Rating", "listas.csv")
        opcoes = []
        if os.path.exists(acrc_path):
            with open(acrc_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    val = row.get(coluna_nome)
                    if val is not None:
                        val = val.strip()
                        if val:
                            opcoes.append(val)
        # Remove duplicatas, mantém ordem
        return list(dict.fromkeys(opcoes))

    def carregar_opcoes_acrc_agencias(self):
        # Lê apenas a coluna "Agencias" do arquivo CSV
        acrc_path = os.path.join(os.path.dirname(__file__), "Rating", "listas.csv")
        opcoes = []
        if os.path.exists(acrc_path):
            with open(acrc_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    val = row.get("Agencias", "").strip()
                    if val:
                        opcoes.append(val)
        # Remove duplicatas, mantém ordem
        return list(dict.fromkeys(opcoes))

    def atualizar_opcoes_acrc(self, combo_to_update=None, new_value=None):
        # Atualiza ambos os combos, mantendo a seleção de cada um
        agencias_opcoes = self.carregar_opcoes_acrc_agencias()
        opcoes = agencias_opcoes + ["Cadastrar"]

        # Salva seleção atual de ambos
        empresa_val = self.empresa_publicadora_combo.currentText() if hasattr(self, "empresa_publicadora_combo") else None
        auditoria_val = self.auditoria_combo.currentText() if hasattr(self, "auditoria_combo") else None

        # Atualiza Empresa Publicadora
        if hasattr(self, "empresa_publicadora_combo"):
            self.empresa_publicadora_combo.blockSignals(True)
            self.empresa_publicadora_combo.clear()
            self.empresa_publicadora_combo.addItems(opcoes)
            if combo_to_update is self.empresa_publicadora_combo and new_value:
                self.empresa_publicadora_combo.setCurrentText(new_value)
            elif empresa_val:
                self.empresa_publicadora_combo.setCurrentText(empresa_val)
            self.empresa_publicadora_combo.blockSignals(False)

        # Atualiza Auditoria
        if hasattr(self, "auditoria_combo"):
            self.auditoria_combo.blockSignals(True)
            self.auditoria_combo.clear()
            self.auditoria_combo.addItems(opcoes)
            if combo_to_update is self.auditoria_combo and new_value:
                self.auditoria_combo.setCurrentText(new_value)
            elif auditoria_val:
                self.auditoria_combo.setCurrentText(auditoria_val)
            self.auditoria_combo.blockSignals(False)

    def cadastrar_nova_opcao_acrc(self, combo):
        text, ok = QtWidgets.QInputDialog.getText(self, "Cadastrar nova opção", "Digite o nome:")
        if ok and text.strip():
            acrc_path = os.path.join(os.path.dirname(__file__), "Rating", "listas.csv")
            # Descobre todas as colunas existentes no arquivo (para não perder dados)
            fieldnames = []
            if os.path.exists(acrc_path):
                with open(acrc_path, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    fieldnames = reader.fieldnames if reader.fieldnames else []
            # Garante que "Agencias" está nas colunas
            if "Agencias" not in fieldnames:
                fieldnames = (fieldnames or []) + ["Agencias"]
            # Cria um dicionário vazio para todas as colunas, só preenchendo "Agencias"
            row_data = {col: "" for col in fieldnames}
            row_data["Agencias"] = text.strip()
            with open(acrc_path, "a", encoding="utf-8", newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                # Escreve cabeçalho se arquivo estava vazio
                if os.stat(acrc_path).st_size == 0:
                    writer.writeheader()
                writer.writerow(row_data)
            self.atualizar_opcoes_acrc(combo_to_update=combo, new_value=text.strip())

    def cadastrar_nova_opcao_acrc_coluna(self, coluna_nome, combo):
        text, ok = QtWidgets.QInputDialog.getText(self, f"Cadastrar nova opção para {coluna_nome}", "Digite o nome:")
        if ok and text.strip():
            acrc_path = os.path.join(os.path.dirname(__file__), "Rating", "listas.csv")
            # Descobre todas as colunas existentes no arquivo (para não perder dados)
            fieldnames = []
            if os.path.exists(acrc_path):
                with open(acrc_path, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    fieldnames = reader.fieldnames if reader.fieldnames else []
            # Garante que a coluna está nas colunas
            if coluna_nome not in fieldnames:
                fieldnames = (fieldnames or []) + [coluna_nome]
            # Cria um dicionário vazio para todas as colunas, só preenchendo a desejada
            row_data = {col: "" for col in fieldnames}
            row_data[coluna_nome] = text.strip()
            with open(acrc_path, "a", encoding="utf-8", newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                # Escreve cabeçalho se arquivo estava vazio
                if os.stat(acrc_path).st_size == 0:
                    writer.writeheader()
                writer.writerow(row_data)
            # Atualiza o combo
            opcoes = self.carregar_opcoes_acrc_coluna(coluna_nome) + ["Cadastrar"]
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(opcoes)
            combo.setCurrentText(text.strip())
            combo.blockSignals(False)

    #---------------------------- -----------------------------------
    #                        INTERFACE GRÁFICA                       
    #---------------------------- -----------------------------------
    def init_ui(self):
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QtWidgets.QGridLayout(central_widget)
        main_layout.setSpacing(15)

        # Frame: Dados da Empresa
        frame_empresa = QtWidgets.QGroupBox("Dados da Empresa")
        frame_empresa.setStyleSheet("QGroupBox { font-weight: bold; }")
        empresa_layout = QtWidgets.QGridLayout(frame_empresa)
        campos_empresa = ["Razão Social", "CNPJ", "Nome Fantasia","Ano do Balanço", "Rating Externo", "Empresa Publicadora"]
        self.entradas = {}
        ano_min = 1900
        ano_max = 2025  # fixo conforme pedido
        self.ano_btn_down = None  # NOVO: referência ao botão -
        self.ano_btn_up = None    # NOVO: referência ao botão +
        for i, campo in enumerate(campos_empresa):
            label = QtWidgets.QLabel(campo)
            if campo == "CNPJ":
                cnpj_widget = QtWidgets.QWidget()
                cnpj_layout = QtWidgets.QHBoxLayout(cnpj_widget)
                cnpj_layout.setContentsMargins(0,0,0,0)
                cnpj_layout.setSpacing(2)
                entrada = QtWidgets.QLineEdit()
                entrada.setInputMask("00.000.000/0000-00;_")
                btn_lupa = QtWidgets.QPushButton()
                btn_lupa.setFixedWidth(22)
                btn_lupa.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogContentsView))
                btn_lupa.setToolTip("Buscar CNPJ")
                btn_lupa.clicked.connect(self.atualizar_data_lupa)  # NOVO: conecta função
                cnpj_layout.addWidget(entrada)
                cnpj_layout.addWidget(btn_lupa)
                entrada_widget = cnpj_widget
            elif campo == "Rating Externo":
                entrada = QtWidgets.QComboBox()
                entrada.addItems([
                    "N/C", "AAA", "AA+", "AA", "AA-", "A+", "A", "A-", 
                    "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", 
                    "B+", "B", "B-", "CCC", "CC", "C", "D", "E", "F"
                ])
                entrada_widget = entrada
                self.rating_combo = entrada  # Salva referência para reset
            elif campo == "Ano do Balanço":
                # Custom widget: QLineEdit (expands) + two fixed-size buttons
                ano_widget = QtWidgets.QWidget()
                ano_layout = QtWidgets.QHBoxLayout(ano_widget)
                ano_layout.setContentsMargins(0,0,0,0)
                ano_layout.setSpacing(2)
                ano_edit = QtWidgets.QLineEdit(str(ano_max))
                ano_edit.setAlignment(QtCore.Qt.AlignCenter)
                # NOVO: Validador para 4 dígitos entre 1900 e 2025
                ano_validator = QtGui.QIntValidator(ano_min, ano_max)
                ano_edit.setValidator(ano_validator)
                ano_edit.setMaxLength(4)
                ano_edit.setInputMask("0000")
                ano_edit.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
                btn_down = QtWidgets.QPushButton("−")
                btn_down.setFixedWidth(22)
                btn_up = QtWidgets.QPushButton("+")
                btn_up.setFixedWidth(22)
                # Funções para alterar o ano
                def set_ano(delta, edit=ano_edit):
                    try:
                        val = int(edit.text())
                    except ValueError:
                        val = ano_max
                    val = max(ano_min, min(ano_max, val + delta))
                    edit.setText(str(val).zfill(4))
                btn_down.clicked.connect(lambda: set_ano(-1))
                btn_up.clicked.connect(lambda: set_ano(1))
                ano_layout.addWidget(ano_edit)
                ano_layout.addWidget(btn_down)
                ano_layout.addWidget(btn_up)
                entrada_widget = ano_widget
                # Para acessar o valor depois, salve o QLineEdit
                self.entradas["Ano_edit"] = ano_edit
                self.ano_btn_down = btn_down  # Salva referência
                self.ano_btn_up = btn_up      # Salva referência
                self.ano_edit_widget = ano_edit  # Salva referência para reset
            elif campo == "Empresa Publicadora":
                # QComboBox com opções da coluna "Agencias" do CSV + "Cadastrar"
                combo = QtWidgets.QComboBox()
                combo.addItems(self.carregar_opcoes_acrc_agencias() + ["Cadastrar"])
                combo.setEditable(False)
                def on_empresa_publicadora_changed(idx, combo=combo):
                    if combo.currentText() == "Cadastrar":
                        self.cadastrar_nova_opcao_acrc(combo)
                combo.currentIndexChanged.connect(on_empresa_publicadora_changed)
                entrada_widget = combo
                self.empresa_publicadora_combo = combo
            else:
                entrada = QtWidgets.QLineEdit()
                entrada_widget = entrada
            empresa_layout.addWidget(label, i, 0)
            empresa_layout.addWidget(entrada_widget, i, 1)
            if campo == "CNPJ":
                self.entradas[campo] = entrada  # salva o QLineEdit, não o widget
            elif campo == "Empresa Publicadora":
                self.entradas[campo] = combo
            elif campo != "Ano do Balanço":
                self.entradas[campo] = entrada
        # Botões Importar PDF e Resetar (um em cima do outro)
        botoes_widget = QtWidgets.QWidget()
        botoes_layout = QtWidgets.QVBoxLayout(botoes_widget)
        botoes_layout.setContentsMargins(0,0,0,0)
        botoes_layout.setSpacing(4)
        botao_pdf = QtWidgets.QPushButton("Importar Dados")
        botao_pdf.setStyleSheet("font-size: 13px;")
        botao_pdf.clicked.connect(self.importar_pdf)
        botao_resetar = QtWidgets.QPushButton("Resetar")
        botao_resetar.setStyleSheet("font-size: 13px;")
        botao_resetar.clicked.connect(self.resetar_formulario)
        botoes_layout.addWidget(botao_pdf)
        botoes_layout.addWidget(botao_resetar)
        empresa_layout.addWidget(botoes_widget, 0, 2, len(campos_empresa), 1)
        empresa_layout.setRowStretch(len(campos_empresa), 1)
        # Ajusta largura máxima dos botões para não ficarem largos demais
        botao_pdf.setMaximumWidth(140)
        botao_resetar.setMaximumWidth(140)
        botao_pdf.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
        botao_resetar.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
        self.botao_pdf = botao_pdf
        self.botao_resetar = botao_resetar

        # Frame: Financeiro
        frame_modelos = QtWidgets.QGroupBox("Financeiro")
        frame_modelos.setStyleSheet("QGroupBox { font-weight: bold; }")
        modelos_layout = QtWidgets.QGridLayout(frame_modelos)
        self.modelos_inputs = []
        novos_textos = [
            "INFO",
            #"Ativo Circulante",
            "Ativo Circulante","Caixa e Equivalente de caixa","Clientes", "Estoques","Outras Contas Receber CP",
            #"Ativo Não Circulante",  
            "Ativo Não Circulante",  "Imobilizado e Investimentos","Outras Contas Receber LP",
           # "Passivo Circulante",
            "Passivo Circulante","Empréstimos CP", "Fornecedores", "Obrigações Trabalhistas", "Obrigações Fiscais", "Outras Obrigações CP",
            #"Passivo Não Circulante"
            "Passivo Não Circulante","Empréstimos LP", "Outras Obrigações LP", 
            #"Patrimônio Líquido",
            "Patrimônio Líquido","Capital Social", "Reserva de Lucros","Outros Itens do PL",
            #DRE
            "Receita Bruta", "Deduções das Receitas",
            "Receita Operacional Líquida(ROL)", "CPV", "Valor Justo Contratos Futuros",
            "Despesas Operacionais", "Resultado Financeiro", "IR+CSLL", "Lucro Líquido", "Auditoria",
            
        ]
        total_campos = len(novos_textos)
        campos_por_coluna = (total_campos + 1) // 2  # arredonda para cima

        for idx, texto in enumerate(novos_textos):
            label = QtWidgets.QLabel(texto)
            # Deixa em negrito apenas os campos principais
            if texto in [
                "Ativo Circulante",
                "Ativo Não Circulante",
                "Passivo Circulante",
                "Passivo Não Circulante",
                "Patrimônio Líquido"
            ]:
                font = label.font()
                font.setBold(True)
                label.setFont(font)
            if texto == "Auditoria":
                # QComboBox com opções da coluna "Agencias" do CSV + "Cadastrar"
                combo = QtWidgets.QComboBox()
                combo.addItems(self.carregar_opcoes_acrc_agencias() + ["Cadastrar"])
                combo.setEditable(False)
                def on_auditoria_changed(idx, combo=combo):
                    if combo.currentText() == "Cadastrar":
                        self.cadastrar_nova_opcao_acrc(combo)
                combo.currentIndexChanged.connect(on_auditoria_changed)
                entrada = combo
                self.auditoria_combo = combo
            elif texto == "INFO":
                entrada = QtWidgets.QLineEdit()
            else:
                entrada = QtWidgets.QLineEdit()
                # Removido o QDoubleValidator para permitir '+'
                entrada.setPlaceholderText("0,00")
                # Permite digitar '-' no início
                def keyPressEvent(event, edit=entrada):
                    if event.key() == QtCore.Qt.Key_Minus and edit.cursorPosition() == 0:
                        edit.insert("-")
                    else:
                        QtWidgets.QLineEdit.keyPressEvent(edit, event)
                entrada.keyPressEvent = lambda event, edit=entrada: keyPressEvent(event, edit)
                # Formata e resolve soma ao sair do campo
                def formatar_decimal(edit=entrada, self_ref=self):
                    val = edit.text().replace(",", ".")
                    try:
                        if val:
                            soma = self_ref.parse_soma(val)
                            edit.setText("{:.2f}".format(soma).replace(".", ","))
                    except Exception:
                        pass
                entrada.editingFinished.connect(formatar_decimal)
            # Novo cálculo: primeiro preenche linhas, depois colunas
            if idx < campos_por_coluna:
                row = idx
                col = 0
            else:
                row = idx - campos_por_coluna
                col = 2
            modelos_layout.addWidget(label, row, col)
            modelos_layout.addWidget(entrada, row, col + 1)
            self.modelos_inputs.append(entrada)

        # Frame: Qualitativo
        frame_qualitativo = QtWidgets.QGroupBox("Qualitativo")
        frame_qualitativo.setStyleSheet("QGroupBox { font-weight: bold; }")
        qualitativo_layout = QtWidgets.QGridLayout(frame_qualitativo)
        self.qualitativo_inputs = []
        qualitativos_textos = [
            # "Nome Fantasia",  # removido
            "Grupo Econômico",
            "Observação",
            # "Ramo Atividade", # removido
            "Atuação Empresa",
            "Data Contrato Social",
            "Tipo Governança",
            "Experiência Trading e Middle",
            "Gestor Cliente",
            "Sede",
            "Setor",
            "Controle Acionário",
            "Tipo (Perfil Empresa)"
        ]
        atuacao_opcoes = [
            "SI", "Bloqueada", "Rj / quebrada", "Trading pura", "Trading c/ gestão", "Trading c/ geração",
            "Trading c/ gestão e geração", "Geradores não ativo", "Geradores pequeno ativo fora mre",
            "Geradores pequeno ativo dentro mre", "Geradores médio", "Gerador grande", "Grandes grupos",
            "Fundo investimento pequeno", "Fundo investimento médio", "Fundo investimento grande",
            "Fundo investimento médio internacional", "Consumidores"
        ]
        experiencia_opcoes = [
            "SI", "Fechou", "Quebrou", "Junior", "Pleno", "Senior", "Especialistas", "Expert", "Excelência"
        ]
        estados_brasil = [
            "AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"
        ]
        # Carregar opções do CSV para os campos específicos
        opcoes_tipo_governanca = self.carregar_opcoes_acrc_coluna("Tipo Governança")
        opcoes_controle_acionario = self.carregar_opcoes_acrc_coluna("Controle Acionário")
        opcoes_tipo_perfil_empresa = self.carregar_opcoes_acrc_coluna("Tipo (Perfil Empresa)")
        opcoes_gestor_cliente = self.carregar_opcoes_acrc_coluna("Gestor Cliente")
        opcoes_setor = self.carregar_opcoes_acrc_coluna("Setor")

        for i, texto in enumerate(qualitativos_textos):
            label = QtWidgets.QLabel(texto)
            if texto == "Atuação Empresa":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(atuacao_opcoes)
            elif texto == "Data Contrato Social":
                entrada = QtWidgets.QDateEdit()
                entrada.setDisplayFormat("dd/MM/yyyy")
                entrada.setCalendarPopup(True)
                entrada.setDate(QtCore.QDate.currentDate())
            elif texto == "Experiência Trading e Middle":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(experiencia_opcoes)
            elif texto == "Sede":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(estados_brasil)
            elif texto == "Tipo Governança":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(opcoes_tipo_governanca)
                # Removido: opção de cadastro
            elif texto == "Controle Acionário":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(opcoes_controle_acionario)
                # Removido: opção de cadastro
            elif texto == "Tipo (Perfil Empresa)":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(opcoes_tipo_perfil_empresa)
                # Removido: opção de cadastro
            elif texto == "Gestor Cliente":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(opcoes_gestor_cliente + ["Cadastrar"])
                def on_gestor_cliente_changed(idx, combo=entrada):
                    if combo.currentText() == "Cadastrar":
                        self.cadastrar_nova_opcao_acrc_coluna("Gestor Cliente", combo)
                entrada.currentIndexChanged.connect(on_gestor_cliente_changed)
            elif texto == "Setor":
                entrada = QtWidgets.QComboBox()
                entrada.addItems(opcoes_setor + ["Cadastrar"])
                def on_setor_changed(idx, combo=entrada):
                    if combo.currentText() == "Cadastrar":
                        self.cadastrar_nova_opcao_acrc_coluna("Setor", combo)
                entrada.currentIndexChanged.connect(on_setor_changed)
            else:
                entrada = QtWidgets.QLineEdit()
            qualitativo_layout.addWidget(label, i, 0)
            qualitativo_layout.addWidget(entrada, i, 1)
            self.qualitativo_inputs.append(entrada)

        # Frame: Dados Comparativos
        frame_comp = QtWidgets.QGroupBox("Dados Comparativos")
        frame_comp.setStyleSheet("QGroupBox { font-weight: bold; }")
        comp_layout = QtWidgets.QVBoxLayout(frame_comp)

        # Substitua QTextEdit por QTableWidget
        self.dados_comparativos_table = QtWidgets.QTableWidget()
        self.dados_comparativos_table.setColumnCount(5)
        self.dados_comparativos_table.setHorizontalHeaderLabels([
            "Grupo", "Balanço", "Somatório", "Diferença", "Status"
        ])
        header = self.dados_comparativos_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        for i in range(1, 5):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.ResizeToContents)
        self.dados_comparativos_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.dados_comparativos_table.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        comp_layout.addWidget(self.dados_comparativos_table)

        # Ajusta a altura da tabela para caber exatamente 5 linhas (sem espaço extra)
        self.dados_comparativos_table.setRowCount(5)
        self.dados_comparativos_table.setMinimumHeight(1)  # força o cálculo correto
        self.dados_comparativos_table.resizeRowsToContents()
        QtWidgets.QApplication.processEvents()  # garante atualização dos tamanhos
        total_height = self.dados_comparativos_table.horizontalHeader().height()
        for row in range(5):
            total_height += self.dados_comparativos_table.rowHeight(row)
        self.dados_comparativos_table.setMaximumHeight(total_height + 60)  # +30 para garantir espaço para 5 linhas
        self.dados_comparativos_table.setRowCount(0)  # limpa para uso normal

        botao_calcular = QtWidgets.QPushButton("Calcular Comparativos")
        botao_calcular.setMinimumHeight(35)
        botao_calcular.setStyleSheet("font-size: 13px;")
        botao_calcular.clicked.connect(self.calcular_comparativos)
        comp_layout.addWidget(botao_calcular)

        # --- NOVO: Container "tools" ---
        tools_container = QtWidgets.QGroupBox("Tools")
        tools_container.setStyleSheet("QGroupBox { font-weight: bold; }")
        tools_layout = QtWidgets.QVBoxLayout(tools_container)
        tools_layout.setAlignment(QtCore.Qt.AlignCenter)  # Centraliza o conteúdo verticalmente

        # Linha 1: Usuário (linha separada)
        # REMOVIDO: label_usuario = QtWidgets.QLabel("Usuário: teste.teste@rbenergia.com.br - trading")
        # REMOVIDO: label_usuario.setStyleSheet("font-size: 13px; color: rgba(0,0,0,0.5);")
        # REMOVIDO: label_usuario.setAlignment(QtCore.Qt.AlignCenter)
        # REMOVIDO: tools_layout.addWidget(label_usuario)

        # Linha 2: Data de atualização (linha separada, destaque)
        self.label_atualizado = QtWidgets.QLabel("Atualizado em: --/--/---- --:--:--")
        self.label_atualizado.setAlignment(QtCore.Qt.AlignCenter)
        self.label_atualizado.setStyleSheet("font-weight: bold; color: #174378; font-size: 13px;")
        tools_layout.addWidget(self.label_atualizado)

        # Linha 3: Toggle milhar/milhão (linha separada, maior e estilo diferente)
        linha_toggle_layout = QtWidgets.QHBoxLayout()
        linha_toggle_layout.setAlignment(QtCore.Qt.AlignCenter)
        self.toggle_milhar = QtWidgets.QRadioButton("Milhares (R$)")
        self.toggle_milhao = QtWidgets.QRadioButton("Milhões (R$)")
        self.toggle_bruto = QtWidgets.QRadioButton("Bruto")
        # self.toggle_milhar.setChecked(True)
        self.toggle_bruto.setChecked(True)  # Default agora é "Bruto"
        # Aumenta fonte e padding dos toggles
        toggle_style = """
            QRadioButton {
                font-size: 15px;
                padding: 8px 18px;
            }
        """
        self.toggle_milhar.setStyleSheet(toggle_style)
        self.toggle_milhao.setStyleSheet(toggle_style)
        self.toggle_bruto.setStyleSheet(toggle_style)
        linha_toggle_layout.addWidget(self.toggle_milhar)
        linha_toggle_layout.addWidget(self.toggle_milhao)
        linha_toggle_layout.addWidget(self.toggle_bruto)
        tools_layout.addLayout(linha_toggle_layout)

        # Linha 4: Botão "Excluir análise" em linha separada, expandido, vermelho escuro e mesmo estilo do dashboard
        linha_excluir_layout = QtWidgets.QHBoxLayout()
        linha_excluir_layout.setAlignment(QtCore.Qt.AlignCenter)
        btn_excluir = QtWidgets.QPushButton("Excluir análise")
        btn_excluir.setMinimumHeight(35)
        btn_excluir.setStyleSheet("font-size: 13px; background-color: #a31515; color: white; font-weight: bold; border-radius: 5px;")
        btn_excluir.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        linha_excluir_layout.addWidget(btn_excluir)
        tools_layout.addLayout(linha_excluir_layout)

        # Linha 5: Botão "Finalizar" centralizado e expandido
        linha3_layout = QtWidgets.QHBoxLayout()
        linha3_layout.setAlignment(QtCore.Qt.AlignCenter)
        btn_finalizar = QtWidgets.QPushButton("Finalizar")
        btn_finalizar.setStyleSheet("background-color: rgb(23,67,120); color: white; font-size: 14px; font-weight:bold")
        btn_finalizar.setMinimumHeight(60)
        btn_finalizar.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        linha3_layout.addWidget(btn_finalizar)
        tools_layout.addLayout(linha3_layout)
        btn_finalizar.clicked.connect(self.finalizar_e_salvar_csv)  # NOVO: conecta função

        # Linha 6: Botão "Abrir dashboard" centralizado, expandido e com mesmo estilo do botão calcular comparativos
        linha4_layout = QtWidgets.QHBoxLayout()
        linha4_layout.setAlignment(QtCore.Qt.AlignCenter)
        btn_dashboard = QtWidgets.QPushButton("Abrir dashboard")
        btn_dashboard.setMinimumHeight(35)
        btn_dashboard.setStyleSheet("font-size: 13px;")
        btn_dashboard.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        linha4_layout.addWidget(btn_dashboard)
        tools_layout.addLayout(linha4_layout)

        # Adiciona o container tools ao layout de comparativos
        comp_layout.addWidget(tools_container)
        # --- FIM TOOLS ---

        # Layout vertical da esquerda (empresa + financeiro)
        esquerda_layout = QtWidgets.QVBoxLayout()
        esquerda_layout.addWidget(frame_empresa)
        esquerda_layout.addWidget(frame_modelos)

        esquerda_widget = QtWidgets.QWidget()
        esquerda_widget.setLayout(esquerda_layout)

        # Adiciona ao layout principal
        main_layout.addWidget(esquerda_widget, 0, 0, 2, 1)  # Coluna 0 ocupa duas linhas
        main_layout.addWidget(frame_qualitativo, 0, 1)
        main_layout.addWidget(frame_comp, 1, 1)

        # Ajuste de tamanhos
        frame_empresa.setMaximumHeight(frame_empresa.sizeHint().height())
        frame_empresa.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        frame_modelos.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)

        main_layout.setColumnStretch(0, 4)
        main_layout.setColumnStretch(1, 2)

        # Salve referências dos botões para bloquear/desbloquear depois
        self.btn_finalizar = btn_finalizar
        self.btn_dashboard = btn_dashboard
        self.botao_calcular = botao_calcular
        self.btn_excluir = btn_excluir

    #---------------------------- -----------------------------------
    #                        IMPORTAÇÃO PDF                         
    #---------------------------- -----------------------------------
    def importar_pdf(self):
        caminho, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if caminho:
            self.dados_comparativos_var.setPlainText("PDF processado com sucesso!\n(Comparativos gerados aqui...)")

    #---------------------------- -----------------------------------
    #                    CÁLCULO DE COMPARATIVOS                    
    #---------------------------- -----------------------------------
    def calcular_comparativos(self):
        # Lista dos campos financeiros na ordem exata dos inputs
        financeiro_labels = [
            "INFO",
            "Ativo Circulante", "Caixa e Equivalente de caixa", "Clientes", "Estoques", "Outras Contas Receber CP",
            "Ativo Não Circulante", "Imobilizado e Investimentos", "Outras Contas Receber LP",
            "Passivo Circulante", "Empréstimos CP", "Fornecedores", "Obrigações Trabalhistas", "Obrigações Fiscais", "Outras Obrigações CP",
            "Passivo Não Circulante", "Empréstimos LP", "Outras Obrigações LP",
            "Patrimônio Líquido", "Capital Social", "Reserva de Lucros", "Outros Itens do PL",
            "Receita Bruta", "Deduções das Receitas",
            "Receita Operacional Líquida(ROL)", "CPV", "Valor Justo Contratos Futuros",
            "Despesas Operacionais", "Resultado Financeiro", "IR+CSLL", "Lucro Líquido", "Auditoria",
        ]

        # Defina os grupos e os índices dos campos principais e subitens
        grupos = [
            {
                "nome": "Ativo Circulante",
                "principal": financeiro_labels.index("Ativo Circulante"),
                "subitens": [
                    financeiro_labels.index("Caixa e Equivalente de caixa"),
                    financeiro_labels.index("Clientes"),
                    financeiro_labels.index("Estoques"),
                    financeiro_labels.index("Outras Contas Receber CP"),
                ]
            },
            {
                "nome": "Ativo Não Circulante",
                "principal": financeiro_labels.index("Ativo Não Circulante"),
                "subitens": [
                    financeiro_labels.index("Imobilizado e Investimentos"),
                    financeiro_labels.index("Outras Contas Receber LP"),
                ]
            },
            {
                "nome": "Passivo Circulante",
                "principal": financeiro_labels.index("Passivo Circulante"),
                "subitens": [
                    financeiro_labels.index("Empréstimos CP"),
                    financeiro_labels.index("Fornecedores"),
                    financeiro_labels.index("Obrigações Trabalhistas"),
                    financeiro_labels.index("Obrigações Fiscais"),
                    financeiro_labels.index("Outras Obrigações CP"),
                ]
            },
            {
                "nome": "Passivo Não Circulante",
                "principal": financeiro_labels.index("Passivo Não Circulante"),
                "subitens": [
                    financeiro_labels.index("Empréstimos LP"),
                    financeiro_labels.index("Outras Obrigações LP"),
                ]
            },
            {
                "nome": "Patrimônio Líquido",
                "principal": financeiro_labels.index("Patrimônio Líquido"),
                "subitens": [
                    financeiro_labels.index("Capital Social"),
                    financeiro_labels.index("Reserva de Lucros"),
                    financeiro_labels.index("Outros Itens do PL"),
                ]
            }
        ]

        linhas = []
        for grupo in grupos:
            try:
                valor_principal = self.parse_soma(self.modelos_inputs[grupo["principal"]].text())
            except Exception:
                valor_principal = 0.0
            soma_subitens = 0.0
            for idx in grupo["subitens"]:
                try:
                    soma_subitens += self.parse_soma(self.modelos_inputs[idx].text())
                except Exception:
                    pass
            diferenca = valor_principal - soma_subitens
            linhas.append((grupo["nome"], valor_principal, soma_subitens, diferenca))

        # Preenche a tabela
        self.dados_comparativos_table.setRowCount(len(linhas))
        for row, (grupo, valor, soma, diff) in enumerate(linhas):
            grupo_item = QtWidgets.QTableWidgetItem(grupo)
            self.dados_comparativos_table.setItem(row, 0, grupo_item)

            valor_item = QtWidgets.QTableWidgetItem(f"{valor:,.2f}")
            valor_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.dados_comparativos_table.setItem(row, 1, valor_item)

            soma_item = QtWidgets.QTableWidgetItem(f"{soma:,.2f}")
            soma_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.dados_comparativos_table.setItem(row, 2, soma_item)

            diff_item = QtWidgets.QTableWidgetItem(f"{diff:,.2f}")
            diff_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.dados_comparativos_table.setItem(row, 3, diff_item)

            status = "❌"
            if valor != 0:
                perc = abs(diff) / abs(valor)
                if perc < 0.03:
                    status = "✅"
            status_item = QtWidgets.QTableWidgetItem(status)
            status_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.dados_comparativos_table.setItem(row, 4, status_item)

        self.dados_comparativos_table.resizeColumnsToContents()

    #---------------------------- -----------------------------------
    #                BUSCA E PREENCHIMENTO VIA LUPA                 
    #---------------------------- -----------------------------------
    def atualizar_data_lupa(self):
        self.atualizado_em = datetime.now()
        self.label_atualizado.setText(
            "Atualizado em: " + self.atualizado_em.strftime("%d/%m/%Y %H:%M:%S")
        )
        # --- NOVO: Busca no CSV e preenche campos ---
        cnpj = self.entradas.get("CNPJ").text().strip()
        # Ano do Balanço pode estar em "Ano_edit" ou "Ano do Balanço"
        ano = ""
        if "Ano_edit" in self.entradas:
            ano = self.entradas["Ano_edit"].text().strip()
        elif "Ano do Balanço" in self.entradas:
            ano = self.entradas["Ano do Balanço"].text().strip()
        if not cnpj or not ano:
            QtWidgets.QMessageBox.warning(self, "Aviso", "Informe o CNPJ e o Ano do Balanço para buscar.")
            return
        # Caminho do CSV
        pasta = os.path.join(os.path.dirname(__file__), "Rating")
        caminho_csv = os.path.join(pasta, "db.csv")
        if not os.path.exists(caminho_csv):
            QtWidgets.QMessageBox.information(self, "Cadastro Não Encontrado", "Realize um novo cadastro para o CNPJ e Ano")
            return
        encontrado = None
        with open(caminho_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                # CNPJ pode estar com máscara, então compara só números
                cnpj_row = row.get("CNPJ", "").replace(".", "").replace("/", "").replace("-", "")
                cnpj_cmp = cnpj.replace(".", "").replace("/", "").replace("-", "")
                ano_row = row.get("Ano_edit") or row.get("Ano do Balanço") or ""
                status_row = row.get("status", "")
                if cnpj_row == cnpj_cmp and ano_row == ano and status_row == "ativa":
                    encontrado = row
                    break
        if encontrado:
            # Preenche campos da empresa
            for campo, entrada in self.entradas.items():
                if campo in encontrado and hasattr(entrada, "setText"):
                    entrada.setText(encontrado[campo])
                elif campo in encontrado and hasattr(entrada, "setCurrentText"):
                    entrada.setCurrentText(encontrado[campo])
            # Preencher campo do ano explicitamente se existir
            if "Ano_edit" in self.entradas and ("Ano_edit" in encontrado or "Ano do Balanço" in encontrado):
                ano_valor = encontrado.get("Ano_edit") or encontrado.get("Ano do Balanço") or ""
                self.entradas["Ano_edit"].setText(ano_valor)
            # Preenche financeiro
            financeiro_labels = [
                "INFO",
                "Ativo Circulante","Caixa e Equivalente de caixa","Clientes", "Estoques","Outras Contas Receber CP",
                "Ativo Não Circulante",  "Imobilizado e Investimentos","Outras Contas Receber LP",
                "Passivo Circulante","Empréstimos CP", "Fornecedores", "Obrigações Trabalhistas", "Obrigações Fiscais", "Outras Obrigações CP",
                "Passivo Não Circulante","Empréstimos LP", "Outras Obrigações LP", 
                "Patrimônio Líquido","Capital Social", "Reserva de Lucros","Outros Itens do PL",
                "Receita Bruta", "Deduções das Receitas",
                "Receita Operacional Líquida(ROL)", "CPV", "Valor Justo Contratos Futuros",
                "Despesas Operacionais", "Resultado Financeiro", "IR+CSLL", "Lucro Líquido", "Auditoria",
            ]
            for label, entrada in zip(financeiro_labels, self.modelos_inputs):
                if label in encontrado:
                    if hasattr(entrada, "setText"):
                        entrada.setText(encontrado[label])
                    elif hasattr(entrada, "setCurrentText"):
                        entrada.setCurrentText(encontrado[label])
            # Preenche qualitativo
            qualitativos_textos = [
                "Grupo Econômico",
                "Observação",
                "Atuação Empresa",
                "Data Contrato Social",
                "Tipo Governança",
                "Experiência Trading e Middle",
                "Gestor Cliente",
                "Sede",
                "Setor",
                "Controle Acionário",
                "Tipo (Perfil Empresa)"
            ]
            for label, entrada in zip(qualitativos_textos, self.qualitativo_inputs):
                if label in ["Atuação Empresa", "Experiência Trading e Middle", "Sede", "Tipo Governança", "Controle Acionário", "Tipo (Perfil Empresa)"] and isinstance(entrada, QtWidgets.QComboBox):
                    if label in encontrado:
                        idx = entrada.findText(encontrado[label])
                        if idx >= 0:
                            entrada.setCurrentIndex(idx)
                elif label == "Gestor Cliente" and isinstance(entrada, QtWidgets.QComboBox):
                    if label in encontrado:
                        entrada.setCurrentText(encontrado[label])
                elif label == "Setor" and isinstance(entrada, QtWidgets.QComboBox):
                    if label in encontrado:
                        entrada.setCurrentText(encontrado[label])
                elif label == "Data Contrato Social" and isinstance(entrada, QtWidgets.QDateEdit):
                    if label in encontrado:
                        try:
                            data = QtCore.QDate.fromString(encontrado[label], "dd/MM/yyyy")
                            if data.isValid():
                                entrada.setDate(data)
                        except Exception:
                            pass
                elif isinstance(entrada, QtWidgets.QLineEdit) and label in encontrado:
                    entrada.setText(encontrado[label])

            self.toggle_milhar.setChecked(False)
            self.toggle_milhao.setChecked(False)
            self.toggle_bruto.setChecked(True)
            self.liberar_campos_apos_lupa()  # NOVO: libera campos após busca
        else:
            QtWidgets.QMessageBox.information(self, "Cadastro Não Encontrado", "Realize um novo cadastro para o CNPJ e Ano")
            self.liberar_campos_apos_lupa()  # NOVO: libera campos mesmo se não encontrar

    #---------------------------- -----------------------------------
    #                  FINALIZAÇÃO E SALVAMENTO CSV                  
    #---------------------------- -----------------------------------
    def finalizar_e_salvar_csv(self):
        # Coleta os dados dos inputs
        dados = {}
        # 1. Dados da Empresa
        for campo, entrada in self.entradas.items():
            if isinstance(entrada, QtWidgets.QComboBox):
                dados[campo] = entrada.currentText()
            elif hasattr(entrada, "text"):
                dados[campo] = entrada.text()
            else:
                dados[campo] = str(entrada)
        # 2. Financeiro
        financeiro_labels = [
            "INFO",
            "Ativo Circulante","Caixa e Equivalente de caixa","Clientes", "Estoques","Outras Contas Receber CP",
            "Ativo Não Circulante",  "Imobilizado e Investimentos","Outras Contas Receber LP",
            "Passivo Circulante","Empréstimos CP", "Fornecedores", "Obrigações Trabalhistas", "Obrigações Fiscais", "Outras Obrigações CP",
            "Passivo Não Circulante","Empréstimos LP", "Outras Obrigações LP", 
            "Patrimônio Líquido","Capital Social", "Reserva de Lucros","Outros Itens do PL",
            "Receita Bruta", "Deduções das Receitas",
            "Receita Operacional Líquida(ROL)", "CPV", "Valor Justo Contratos Futuros",
            "Despesas Operacionais", "Resultado Financeiro", "IR+CSLL", "Lucro Líquido", "Auditoria",
        ]
        # Determina o fator de multiplicação conforme unidade
        if self.toggle_milhar.isChecked():
            fator = 1000
        elif self.toggle_milhao.isChecked():
            fator = 1000000
        else:  # Bruto
            fator = 1
        for idx, (label, entrada) in enumerate(zip(financeiro_labels, self.modelos_inputs)):
            if label in ["INFO", "Auditoria"]:
                # Não converter, apenas salvar como está
                if isinstance(entrada, QtWidgets.QComboBox):
                    dados[label] = entrada.currentText()
                elif hasattr(entrada, "text"):
                    dados[label] = entrada.text()
                else:
                    dados[label] = str(entrada)
            else:
                # Converter valor conforme unidade
                valor = ""
                if isinstance(entrada, QtWidgets.QComboBox):
                    valor = entrada.currentText()
                elif hasattr(entrada, "text"):
                    valor = entrada.text()
                else:
                    valor = str(entrada)
                # Se vazio, salva como "0"
                if not valor.strip():
                    dados[label] = "0"
                else:
                    # Tenta converter expressão para soma, aplica fator se possível
                    try:
                        valor_num = self.parse_soma(valor.replace(".", "").replace(",", "."))
                        valor_convertido = valor_num * fator
                        # Salva sempre com ponto como separador decimal, sem separador de milhar
                        dados[label] = "{:.2f}".format(valor_convertido)
                    except Exception:
                        # Se não for número, salva como está
                        dados[label] = valor
        # 3. Qualitativo
        qualitativos_textos = [
            # "Nome Fantasia",  # removido
            "Grupo Econômico",
            "Observação",
            # "Ramo Atividade", # removido
            "Atuação Empresa",
            "Data Contrato Social",
            "Tipo Governança",
            "Experiência Trading e Middle",
            "Gestor Cliente",
            "Sede",
            "Setor",
            "Controle Acionário",
            "Tipo (Perfil Empresa)"
        ]
        for label, entrada in zip(qualitativos_textos, self.qualitativo_inputs):
            if isinstance(entrada, QtWidgets.QComboBox):
                dados[label] = entrada.currentText()
            elif isinstance(entrada, QtWidgets.QDateEdit):
                dados[label] = entrada.date().toString("dd/MM/yyyy")
            elif hasattr(entrada, "text"):
                dados[label] = entrada.text()
            else:
                dados[label] = str(entrada)
        # 4. Usuário, Data Atualização e Unidade
        if self.atualizado_em:
            dados["Data Atualização"] = self.atualizado_em.strftime("%d/%m/%Y %H:%M:%S")
        else:
            dados["Data Atualização"] = ""
        if self.toggle_milhar.isChecked():
            dados["Unidade"] = "Milhares (R$)"
        elif self.toggle_milhao.isChecked():
            dados["Unidade"] = "Milhões (R$)"
        elif self.toggle_bruto.isChecked():
            dados["Unidade"] = "Bruto"
        else:
            dados["Unidade"] = ""

        # --- NOVO: lógica para coluna status ---
        cnpj = (self.entradas.get("CNPJ").text() if self.entradas.get("CNPJ") else "").replace(".", "").replace("/", "").replace("-", "")
        ano = ""
        if "Ano_edit" in self.entradas:
            ano = self.entradas["Ano_edit"].text().strip()
        elif "Ano do Balanço" in self.entradas:
            ano = self.entradas["Ano do Balanço"].text().strip()

        pasta = os.path.join(os.path.dirname(__file__), "Rating")
        os.makedirs(pasta, exist_ok=True)
        caminho_csv = os.path.join(pasta, "db.csv")

        # Lê todos os registros existentes
        registros = []
        existe_igual = False
        if os.path.exists(caminho_csv):
            with open(caminho_csv, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    cnpj_row = row.get("CNPJ", "").replace(".", "").replace("/", "").replace("-", "")
                    ano_row = row.get("Ano_edit") or row.get("Ano do Balanço") or ""
                    if cnpj_row == cnpj and ano_row == ano:
                        row["status"] = "editada"
                        existe_igual = True
                    registros.append(row)

        # Define status do novo registro
        dados["status"] = "ativa"

        # Garante que a coluna status está no cabeçalho
        fieldnames = list(dados.keys())
        if "status" not in fieldnames:
            fieldnames.append("status")

        # Atualiza registros antigos se necessário
        if existe_igual:
            # Atualiza todos os registros para garantir que "status" está presente
            for r in registros:
                if "status" not in r:
                    r["status"] = ""
        # Escreve tudo de volta (sobrescreve o arquivo)
        with open(caminho_csv, "w", newline='', encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in registros:
                # Garante que todos os campos estão presentes
                for k in fieldnames:
                    if k not in r:
                        r[k] = ""
                writer.writerow(r)
            # Escreve o novo registro
            writer.writerow(dados)

        QtWidgets.QMessageBox.information(self, "Salvo", "Dados salvos em db.csv com sucesso!")

        # --- LOG: salva IP, data/hora, CNPJ e Ano do Balanço ---
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except Exception:
            ip = "N/A"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_path = os.path.join(os.path.dirname(__file__), "Rating", "form_exec_log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{now} | IP: {ip} | CNPJ: {cnpj} | Ano: {ano}\n")

    #---------------------------- -----------------------------------
    #                BLOQUEIO E LIBERAÇÃO DE CAMPOS                 
    #---------------------------- -----------------------------------
    def bloquear_campos_iniciais(self):
        # Bloqueia todos os campos, exceto CNPJ e Ano, e trava botões
        for campo, entrada in self.entradas.items():
            if campo == "CNPJ":
                entrada.setEnabled(True)
                entrada.setStyleSheet("opacity: 1;")
            elif campo == "Ano_edit" or campo == "Ano do Balanço":
                entrada.setEnabled(True)
                entrada.setStyleSheet("opacity: 1;")
                # Travar/destravar botões + e -
                if self.ano_btn_down: self.ano_btn_down.setEnabled(True)
                if self.ano_btn_up: self.ano_btn_up.setEnabled(True)
            elif campo == "Empresa Publicadora":
                entrada.setEnabled(False)
            else:
                entrada.setEnabled(False)
                entrada.setStyleSheet("opacity: 0.5;")
        for entrada in self.modelos_inputs:
            if hasattr(entrada, "setEnabled"):
                entrada.setEnabled(False)
                if hasattr(entrada, "setStyleSheet"):
                    entrada.setStyleSheet("opacity: 0.5;")
        for entrada in self.qualitativo_inputs:
            entrada.setEnabled(False)
            entrada.setStyleSheet("opacity: 0.5;")
        # Trava botões
        self.botao_pdf.setEnabled(False)
        self.botao_resetar.setEnabled(True)
        self.btn_finalizar.setEnabled(False)
        self.btn_dashboard.setEnabled(False)
        self.botao_calcular.setEnabled(False)
        self.btn_excluir.setEnabled(False)
        # Travar/destravar botões + e - se não for ano
        if self.ano_btn_down and not self.entradas["Ano_edit"].isEnabled():
            self.ano_btn_down.setEnabled(False)
        if self.ano_btn_up and not self.entradas["Ano_edit"].isEnabled():
            self.ano_btn_up.setEnabled(False)

    def liberar_campos_apos_lupa(self):
        # Libera todos os campos, exceto CNPJ e Ano, e libera botões
        for campo, entrada in self.entradas.items():
            if campo == "CNPJ":
                entrada.setEnabled(False)
                entrada.setStyleSheet("opacity: 0.5;")
            elif campo == "Ano_edit" or campo == "Ano do Balanço":
                entrada.setEnabled(False)
                entrada.setStyleSheet("opacity: 0.5;")
                # Travar botões + e -
                if self.ano_btn_down: self.ano_btn_down.setEnabled(False)
                if self.ano_btn_up: self.ano_btn_up.setEnabled(False)
            elif campo == "Empresa Publicadora":
                entrada.setEnabled(True)
            else:
                entrada.setEnabled(True)
                entrada.setStyleSheet("opacity: 1;")
        for entrada in self.modelos_inputs:
            if hasattr(entrada, "setEnabled"):
                entrada.setEnabled(True)
                if hasattr(entrada, "setStyleSheet"):
                    entrada.setStyleSheet("opacity: 1;")
        for entrada in self.qualitativo_inputs:
            entrada.setEnabled(True)
            entrada.setStyleSheet("opacity: 1;")
        # Libera botões
        self.botao_pdf.setEnabled(True)
        self.botao_resetar.setEnabled(True)
        self.btn_finalizar.setEnabled(True)
        self.btn_dashboard.setEnabled(True)
        self.botao_calcular.setEnabled(True)
        self.btn_excluir.setEnabled(True)

    #---------------------------- -----------------------------------
    #                        RESET DO FORMULÁRIO                     
    #---------------------------- -----------------------------------
    def resetar_formulario(self):
        # Limpa todos os campos e bloqueia tudo exceto CNPJ e Ano, mas NÃO limpa Rating Externo nem Ano
        for campo, entrada in self.entradas.items():
            if campo == "CNPJ":
                if hasattr(entrada, "clear"):
                    entrada.clear()
            elif campo == "Ano_edit" or campo == "Ano do Balanço":
                # NÃO limpa o campo de ano
                continue
            elif campo == "Rating Externo":
                # NÃO limpa o rating
                continue
            elif campo == "Empresa Publicadora":
                if hasattr(entrada, "setCurrentIndex"):
                    entrada.setCurrentIndex(0)
            elif hasattr(entrada, "clear"):
                entrada.clear()
            elif hasattr(entrada, "setCurrentIndex"):
                entrada.setCurrentIndex(0)
        for entrada in self.modelos_inputs:
            # Corrige: só chama clear() se for QLineEdit, setCurrentIndex(0) se for QComboBox
            if isinstance(entrada, QtWidgets.QLineEdit):
                entrada.clear()
            elif isinstance(entrada, QtWidgets.QComboBox):
                entrada.setCurrentIndex(0)
        for entrada, label in zip(self.qualitativo_inputs, [
            # "Nome Fantasia",  # removido
            "Grupo Econômico",
            "Observação",
            # "Ramo Atividade", # removido
            "Atuação Empresa",
            "Data Contrato Social",
            "Tipo Governança",
            "Experiência Trading e Middle",
            "Gestor Cliente",
            "Sede",
            "Setor",
            "Controle Acionário",
            "Tipo (Perfil Empresa)"
        ]):
            if label == "Gestor Cliente" and isinstance(entrada, QtWidgets.QComboBox):
                # Recarrega opções do CSV + "Cadastrar"
                opcoes_gestor_cliente = self.carregar_opcoes_acrc_coluna("Gestor Cliente") + ["Cadastrar"]
                entrada.blockSignals(True)
                entrada.clear()
                entrada.addItems(opcoes_gestor_cliente)
                entrada.setCurrentIndex(0)
                entrada.blockSignals(False)
            elif label == "Setor" and isinstance(entrada, QtWidgets.QComboBox):
                # Recarrega opções do CSV + "Cadastrar"
                opcoes_setor = self.carregar_opcoes_acrc_coluna("Setor") + ["Cadastrar"]
                entrada.blockSignals(True)
                entrada.clear()
                entrada.addItems(opcoes_setor)
                entrada.setCurrentIndex(0)
                entrada.blockSignals(False)
            elif label in ["Atuação Empresa", "Experiência Trading e Middle", "Sede", "Tipo Governança", "Controle Acionário", "Tipo (Perfil Empresa)"] and isinstance(entrada, QtWidgets.QComboBox):
                entrada.setCurrentIndex(0)
            elif label == "Data Contrato Social" and isinstance(entrada, QtWidgets.QDateEdit):
                entrada.setDate(QtCore.QDate.currentDate())
            else:
                entrada.clear()
        self.toggle_milhar.setChecked(False)
        self.toggle_milhao.setChecked(False)
        self.toggle_bruto.setChecked(True)
        self.label_atualizado.setText("Atualizado em: --/--/---- --:--:--")
        self.atualizado_em = None
        self.bloquear_campos_iniciais()

    def parse_soma(self, texto):
        """
        Avalia expressões simples de soma/subtração, como '1000+200-300'.
        Aceita ',' ou '.' como separador decimal.
        Só permite números, + e -.
        """
        if not texto:
            return 0.0
        texto = texto.replace(',', '.')
        try:
            # Avaliação segura usando ast
            node = ast.parse(texto, mode='eval')

            def eval_node(n):
                if isinstance(n, ast.Expression):
                    return eval_node(n.body)
                elif isinstance(n, ast.BinOp) and isinstance(n.op, (ast.Add, ast.Sub)):
                    return eval_node(n.left) + (eval_node(n.right) if isinstance(n.op, ast.Add) else -eval_node(n.right))
                elif isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.USub):
                    return -eval_node(n.operand)
                elif isinstance(n, ast.Constant):  # Python 3.8+ e futuro
                    return float(n.value)
                else:
                    raise ValueError("Expressão inválida")
            return eval_node(node)
        except Exception:
            return 0.0

#---------------------------- -----------------------------------
#                        EXECUÇÃO PRINCIPAL                       
#---------------------------- -----------------------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec_())