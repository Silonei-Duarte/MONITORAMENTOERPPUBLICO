from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QComboBox, QLabel, QVBoxLayout, QGridLayout, \
    QWidget, QTextEdit, QFileDialog, QMessageBox, QStackedWidget
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from datetime import datetime, timedelta
import numpy as np
import os
import configparser
from Legenda import Legendas
from Sobre import Sobre


class Dashboards(QMainWindow):
    def __init__(self):
        super().__init__()
        self.processos_lista, self.licencas = self.carregar_configuracao() 
        if self.processos_lista is None or self.licencas is None:
            QMessageBox.critical(self, "Erro", "Erro ao carregar processos e licenças.")
            return

        # Chama o método init_ui para configurar a interface gráfica
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Dashboard")
        self.setGeometry(0, 0, 1900, 1050) 

        # Defina o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        self.setWindowIcon(QIcon(icon_path))
        self.setStyleSheet("background-color: white;")

        # Botão para carregar o arquivo Excel
        self.btn_carregar = QPushButton("Selecionar")
        self.btn_carregar.setStyleSheet("background-color: Green; color: #ffffff; font-weight: bold;")
        self.btn_carregar.setFixedSize(200, 35)  # Tamanho do botão
        self.btn_carregar.clicked.connect(self.carregar_arquivo)

        # Cria a interface inicial
        self.configurar_layout()
        self.carregar_planilha_inicial()

    def carregar_configuracao(self):
        """Carrega as configurações do arquivo db.ini"""
        config = configparser.ConfigParser()
        caminho_arquivo_ini = os.path.join(os.path.dirname(__file__), 'db.ini')

        if not os.path.exists(caminho_arquivo_ini):
            QMessageBox.critical(self, "Erro", "Arquivo de configuração db.ini não encontrado.")
            return None, None

        config.read(caminho_arquivo_ini)

        # Carrega os processos e licenças da seção PROPRIETARIA
        try:
            processos_str = config.get('PROPRIETARIA', 'processos')
            licencas_str = config.get('PROPRIETARIA', 'licencas')

            processos = [proc.strip() for proc in processos_str.split(',')]
            licencas = [int(licenca.strip()) for licenca in licencas_str.split(',')]

            return processos, licencas
        except configparser.NoSectionError:
            QMessageBox.critical(self, "Erro", "Seção PROPRIETARIA não encontrada em db.ini.")
            return None, None

    def configurar_layout(self):
        """Configura o layout principal e o widget_interface."""
        # Recria o widget_interface com o mesmo estilo e tamanho
        self.widget_interface = QWidget()
        self.widget_interface.setStyleSheet("background-color: white;")
        self.widget_interface.setFixedSize(1900, 1000)  # Define tamanho da caixa branca

        # Usando QStackedWidget para sobreposição correta
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.addWidget(self.widget_interface)  
        self.stacked_widget.addWidget(self.btn_carregar)  

        # Alinhamento do botão ao canto superior esquerdo
        self.stacked_widget.setCurrentIndex(1)

        # Cria o container principal e define o layout
        self.container = QWidget()
        self.container.setLayout(QVBoxLayout())
        self.container.layout().addWidget(self.stacked_widget)
        self.setCentralWidget(self.container)
        self.showMaximized()

    def processar_arquivo(self, nome_arquivo):
        """Carrega e processa o arquivo Excel."""
        try:
            self.df = pd.read_excel(nome_arquivo, usecols="A:Z", engine='openpyxl')

            # Converte a coluna de tempo para timedelta
            for col in self.df.columns[1:]:
                self.df[col] = pd.to_timedelta(self.df[col])

            # Ordena os usuários em ordem alfabética e adiciona "Todos Usuários ou" no início
            self.usuarios = sorted(self.df.iloc[:, 0].tolist())
            self.usuarios.insert(0, "Todos Usuários ou")

            # Atualiza o texto do botão com o nome do arquivo
            nome_arquivo_curto = os.path.basename(nome_arquivo)  # Pega apenas o nome do arquivo, sem o caminho
            self.btn_carregar.setText(f"Selecionar: ({nome_arquivo_curto})")

            # Cria a interface interna agora que o arquivo foi carregado
            self.criar_interface()

            # Mostra o widget da interface
            self.widget_interface.show()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo: {e}")

    def carregar_planilha_inicial(self):
        """Carrega automaticamente o arquivo mais recente com base no nome padrão."""
        try:
            arquivos = [f for f in os.listdir() if f.startswith('Senior_') and f.endswith('.xlsx')]
            if arquivos:
                mais_recentes = sorted(arquivos, key=lambda x: os.path.getmtime(x), reverse=True)
                nome_arquivo = mais_recentes[0]

                # Atualiza o texto do botão com o nome do arquivo
                nome_arquivo_curto = os.path.basename(nome_arquivo)
                self.btn_carregar.setText(f"Selecionar: ({nome_arquivo_curto})")

                self.processar_arquivo(nome_arquivo)
            else:
                QMessageBox.warning(self, "Aviso", "Nenhum arquivo Excel encontrado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo inicial: {e}")

    def carregar_arquivo(self):
        """Abre um diálogo para selecionar o arquivo Excel e carrega a interface."""
        try:
            file_name, _ = QFileDialog.getOpenFileName(self, "Selecione o Arquivo Excel", "", "Arquivos Excel (*.xlsx)")

            if file_name:
                self.excluir_interface()
                self.processar_arquivo(file_name)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro inesperado: {e}")

    def excluir_interface(self):
        """Remove completamente a interface interna e recria a nova."""
        # Remove todos os widgets e o layout antigo da widget_interface
        if self.widget_interface:
            layout = self.widget_interface.layout()
            if layout:
                while layout.count():
                    item = layout.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()  # Marca o widget para ser excluído

            if self.container:
                self.container.layout().removeWidget(self.widget_interface)

            self.widget_interface.deleteLater()

        # Recria o layout principal e o widget_interface
        self.configurar_layout()

    def criar_interface(self):
        """Cria toda a interface interna dentro do quadrado branco com posicionamento absoluto."""
        # Cria os gráficos
        self.criar_graficos()

        # Estilo para os QComboBox
        combobox_style = (
            "QComboBox { background-color: #3C3F41; color: #FFFFFF; border: 1px solid #555555; border-radius: 4px; padding: 5px; font-size: 14px; } "
            "QComboBox::drop-down { background-color: #2D2D2D; border-left: 1px solid #555555; } "
            "QComboBox QAbstractItemView { background-color: #3C3F41; color: #FFFFFF; selection-background-color: #5C5F61; selection-color: #FFFFFF; border: 1px solid #555555; } "
            "QComboBox QAbstractItemView::item:hover { background-color: #505355; color: #FFFFFF; } "
            "QScrollBar:vertical { background-color: #2D2D2D; width: 12px; } "
            "QScrollBar::handle:vertical { background-color: #555555; border-radius: 6px; min-height: 20px; } "
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { background-color: #2D2D2D; height: 0px; } "
            "QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical { background: none; } "
            "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }"
        )

        # Adiciona os gráficos (canvas)
        self.canvas1.setParent(self.widget_interface)
        self.canvas1.setFixedSize(1700, 550) 
        self.canvas1.move(0, 0)  

        self.canvas2.setParent(self.widget_interface)
        self.canvas2.setFixedSize(1700, 500)  
        self.canvas2.move(0, 500)  
        
        # Botão para abrir sobre
        self.btn_abrir_sobre = QPushButton("Sobre", self.widget_interface)
        self.btn_abrir_sobre.setStyleSheet("background-color: #3e3e3e; color: #ffffff;")
        self.btn_abrir_sobre.setFixedSize(120, 35)
        self.btn_abrir_sobre.clicked.connect(self.abrir_sobre)
        self.btn_abrir_sobre.move(1780, 0) 

        # Botão para abrir a legenda
        self.btn_abrir_legenda = QPushButton("Legenda", self.widget_interface)
        self.btn_abrir_legenda.setStyleSheet("background-color: #3e3e3e; color: #ffffff;")
        self.btn_abrir_legenda.setFixedSize(120, 35)
        self.btn_abrir_legenda.clicked.connect(self.abrir_legenda)
        self.btn_abrir_legenda.move(1650, 0) 

        # Adiciona o combobox de seleção de processo
        self.combobox_processo = QComboBox(self.widget_interface)
        processos = self.df.columns[1:].tolist()
        self.combobox_processo.addItems(processos)
        self.combobox_processo.currentIndexChanged.connect(self.on_processo_selecionado)
        self.combobox_processo.setStyleSheet(combobox_style)
        self.combobox_processo.setFixedSize(200, 30)
        self.combobox_processo.move(1700, 55) 

        # Cria um QTextEdit para mostrar informações dos usuários
        self.text_edit_usuarios = QTextEdit(self.widget_interface)
        self.text_edit_usuarios.setReadOnly(True)
        self.text_edit_usuarios.setStyleSheet("border: 2px solid white; color: black;")
        self.text_edit_usuarios.setFixedSize(200, 850)
        self.text_edit_usuarios.move(1700, 90)  

        # Adiciona o combobox de seleção de usuário
        self.combobox_usuario = QComboBox(self.widget_interface)
        self.combobox_usuario.addItems(self.usuarios)
        self.combobox_usuario.currentIndexChanged.connect(self.on_usuario_selecionado)
        self.combobox_usuario.setStyleSheet(combobox_style)
        self.combobox_usuario.setFixedSize(200, 30)
        self.combobox_usuario.move(220, 4)  

        # Atualiza os gráficos e o conteúdo inicial
        self.atualizar_graficos("Todos Usuários ou")
        self.on_processo_selecionado()

    def criar_graficos(self):
        """Cria gráficos para exibir os dados do dashboard."""
        # Inicialmente, gráficos vazios
        self.fig1, self.ax1 = plt.subplots(figsize=(6, 4), dpi=90)
        self.ax1.set_title('Tempo de Uso dos Processos')
        self.ax1.set_ylabel('Tempo (HH:MM:SS)')

        # Ajusta manualmente o espaço ao redor do gráfico
        self.fig1.subplots_adjust(left=0.09, right=0.99, top=0.90, bottom=0.2)
        self.canvas1 = FigureCanvas(self.fig1)
        self.canvas1.setFixedSize(1700, 500)

        self.fig2, self.ax2 = plt.subplots(figsize=(6, 4), dpi=90)
        self.ax2.set_title('Aberturas no Dia')
        self.ax2.set_ylabel('Número de Aberturas')

        # Ajusta manualmente o espaço ao redor do gráfico
        self.fig2.subplots_adjust(left=0.04, right=0.99, top=0.9, bottom=0.2)
        self.canvas2 = FigureCanvas(self.fig2)
        self.canvas2.setFixedSize(1700, 500)

    def format_time(self, x, pos):
        """Formata o tempo em segundos para mostrar dias, horas, minutos e segundos, sem zero à esquerda para horas."""
        total_seconds = int(x)
        days, remainder = divmod(total_seconds, 86400)  # 86400 segundos em um dia
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)

        if days > 0:
            return f'{days}d {hours}:{minutes:02}:{seconds:02}'
        else:
            return f'{hours}:{minutes:02}:{seconds:02}'

    def atualizar_graficos(self, usuario):
        """Atualiza os gráficos com dados reais da planilha para o usuário selecionado."""
        # Filtra os dados para o usuário selecionado
        if usuario == "Todos Usuários ou":
            dados_usuario = self.df.iloc[:, 1:]
        else:
            dados_usuario = self.df[self.df.iloc[:, 0] == usuario].iloc[:, 1:]

        if dados_usuario.empty:
            return

        # Calcula o tempo total de uso de cada processo
        tempos = dados_usuario.sum()

        # Calcula a frequência de uso (vezes com tempo > 1 minuto)
        frequencias = (self.df.iloc[:, 1:] > pd.Timedelta(minutes=1)).sum()

        processos = self.df.columns[1:]

        # Define uma lista de cores (uma cor para cada processo)
        num_processos = len(processos)
        cores = plt.cm.viridis(np.linspace(0, 1, num_processos))

        # Atualiza o gráfico de Tempo de Uso
        self.ax1.clear()
        bars1 = self.ax1.bar(processos, tempos.dt.total_seconds(), color=cores)
        self.ax1.set_title('Duração do Uso dos Processos')
        self.ax1.set_ylabel('Tempo (HH:MM:SS)')
        self.ax1.set_xticks(range(len(processos)))  
        self.ax1.set_xticklabels(processos, rotation=25, ha='right')  
        self.ax1.yaxis.set_major_formatter(plt.FuncFormatter(self.format_time))

        # Utiliza a função format_time para formatar os rótulos das barras
        for i, tempo in enumerate(tempos):
            formatted_time = self.format_time(tempo.total_seconds(), None)
            self.ax1.text(i, tempos.dt.total_seconds().iloc[i], formatted_time, ha='center', va='bottom')

        self.canvas1.draw()


        # Atualiza o gráfico de Frequência de Acessos
        self.ax2.clear()
        self.ax2.bar(processos, frequencias, color=cores)
        self.ax2.set_title('Acessos Diários por Diferentes Usuários')
        self.ax2.set_ylabel('Número de Acessos')
        self.ax2.set_xticks(range(len(processos)))  
        self.ax2.set_xticklabels(processos, rotation=25, ha='right') 
        maior_licenca = max(self.licencas)
        max_frequencia = frequencias.max()
        if max_frequencia > maior_licenca:
            self.ax2.set_ylim(0, max_frequencia + 10)  # Ajusta dinamicamente se o máximo for maior que a maior licença
        else:
            self.ax2.set_ylim(0, maior_licenca + 10)

        # Dicionário para linhas de referência específicas, usando a lista carregada de licenças
        referencias = dict(zip(self.processos_lista, self.licencas))

        # Adiciona as linhas de referência específicas
        for i, processo in enumerate(processos):
            if processo in referencias:
                referencia = referencias[processo]
                self.ax2.text(i, referencia, f'{referencia}____', ha='right', va='bottom', color='red', fontsize=10,
                              fontweight='bold')

        for i, freq in enumerate(frequencias):
            self.ax2.text(i, freq, freq, ha='center', va='bottom', color='black')

        self.canvas2.draw()

    def atualizar_usuarios(self, processo):
        """Atualiza a lista de usuários e tempos de uso para o processo selecionado."""
        if processo not in self.df.columns:
            return

        tempos_uso = self.df[processo]
        usuarios = self.df.iloc[:, 0]

        # Combina os usuários e tempos de uso em um DataFrame temporário
        df_temp = pd.DataFrame({"usuario": usuarios, "tempo": tempos_uso})

        # Ordena por tempo de uso decrescente e remove usuários com tempo zero
        df_temp = df_temp[df_temp["tempo"] > pd.Timedelta(0)]
        df_temp = df_temp.sort_values(by="tempo", ascending=False)

        # Atualiza o QTextEdit com as informações dos usuários
        texto = ""
        for _, row in df_temp.iterrows():
            tempo_formatado = str(row["tempo"]).split()[2]  # Formata como HH:MM:SS
            texto += f"{row['usuario']: <20} {tempo_formatado}\n"

        self.text_edit_usuarios.setText(texto)

    def on_processo_selecionado(self):
        """Atualiza a exibição dos usuários e gráficos com base no processo selecionado."""
        processo = self.combobox_processo.currentText()
        self.atualizar_graficos(processo)
        self.atualizar_usuarios(processo)

    def on_usuario_selecionado(self):
        """Atualiza a exibição dos gráficos com base no usuário selecionado."""
        usuario = self.combobox_usuario.currentText()
        self.atualizar_graficos(usuario)

    def abrir_legenda(self):
        try:
            self.legenda = Legendas()
            self.legenda.show()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir a Legenda {e}")

    def abrir_sobre(self):
        try:
            self.sobre = Sobre()
            self.sobre.show()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir Sobre {e}")


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    window = Dashboards()
    window.show()
    sys.exit(app.exec())

