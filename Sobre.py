from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtCore import Qt
import os

class Sobre(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Sobre")
        self.setGeometry(700, 300, 200, 200)  # Define o tamanho da janela

        # Define o estilo da janela para ser completamente escura
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")

        # Define o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        self.setWindowIcon(QIcon(icon_path))

        # Layout principal
        layout = QVBoxLayout()

        # Label para mostrar as informações centralizadas
        label_texto = QLabel("Programa de Monitoramento Senior\n\n"
            "Versão: 1.0.8\n"
            "Criador: Silonei Duarte\n"
            "Data de Lançamento: Agosto 2024\n"
            "\n"
            "Este programa foi desenvolvido para monitoramento e análise de dados do Sistema Senior .\n"
            "Todos os direitos reservados."
        )
        label_texto.setStyleSheet("color: #ffffff;")  # Define a cor do texto
        label_texto.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Centraliza o texto
        label_texto.setFont(QFont("Arial", 10))  # Define a fonte e o tamanho do texto

        # Adiciona a label ao layout
        layout.addWidget(label_texto)

        # Define o layout da janela
        self.setLayout(layout)

if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    window = Sobre()
    window.show()
    sys.exit(app.exec())
