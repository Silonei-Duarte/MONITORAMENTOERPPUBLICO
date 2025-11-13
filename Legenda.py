import os
import sys
import configparser
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt

# Variável global para os dados da legenda
dados_legenda = []

class Legendas(QWidget):
    def __init__(self):
        super().__init__()
        self.carregar_legenda()  # Atualiza a variável global dados_legenda
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Legenda")
        self.setGeometry(600, 100, 650, 800)
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        self.setWindowIcon(QIcon(icon_path))

        # Configura o estilo do widget principal
        self.setStyleSheet(""" background-color: #2e2e2e; color: #ffffff;""")

        # Layout principal
        layout = QVBoxLayout(self)

        # Cria a tabela
        tabela = QTableWidget(self)
        tabela.setColumnCount(2)
        tabela.setHorizontalHeaderLabels(["Sigla", "Descrição"])

        # Define o tamanho das colunas
        tabela.setColumnWidth(0, 50)  # Largura da coluna de Sigla
        tabela.setColumnWidth(1, 570)  # Largura da coluna de Descrição

        # Define o número de linhas na tabela
        tabela.setRowCount(len(dados_legenda))

        # Preenche a tabela com os dados da legenda
        for i, (sigla, descricao) in enumerate(dados_legenda):
            sigla_item = QTableWidgetItem(sigla)
            descricao_item = QTableWidgetItem(descricao)

            tabela.setItem(i, 0, sigla_item)
            tabela.setItem(i, 1, descricao_item)

            # Define as células como não editáveis
            sigla_item.setFlags(sigla_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            descricao_item.setFlags(descricao_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        tabela.verticalHeader().setVisible(False)

        # Configura o estilo da tabela
        tabela.setStyleSheet(
            "QTableWidget { background-color: #2e2e2e; color: #ffffff; } "
            "QHeaderView::section { background-color: #1e1e1e; color: #ffffff; font-weight: normal; } "
        )
        tabela.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        # Adiciona a tabela ao layout
        layout.addWidget(tabela)

        # Configura o layout na janela principal
        self.setLayout(layout)

        # Exibe a interface
        self.show()

    def carregar_legenda(self):
        """Carrega os dados da legenda do arquivo db.ini e atualiza a variável global"""
        global dados_legenda  # Indica que estamos modificando a variável global
        config = configparser.ConfigParser()
        caminho_arquivo_ini = os.path.join(os.path.dirname(__file__), 'db.ini')
        if not os.path.exists(caminho_arquivo_ini):
            return  # Sai se o arquivo não existir

        config.read(caminho_arquivo_ini)

        # Tenta carregar a legenda do db.ini no formato simplificado
        try:
            legenda_str = config.get('LEGENDA', 'legenda', fallback="")
            linhas = legenda_str.strip().split('\n')  # Quebra as linhas

            # Processa cada linha no formato "sigla, descrição,"
            dados_legenda = []
            for linha in linhas:
                partes = linha.split(',', 1)  # Divide em duas partes: sigla e descrição
                if len(partes) == 2:
                    sigla = partes[0].strip()
                    descricao = partes[1].strip().rstrip(',')  # Remove a vírgula final
                    dados_legenda.append((sigla, descricao))
        except Exception as e:
            print(f"Erro ao carregar a legenda: {e}")
            dados_legenda = []

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Legendas()
    sys.exit(app.exec())
