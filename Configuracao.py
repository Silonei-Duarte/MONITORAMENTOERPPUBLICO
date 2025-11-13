import os
import sys
import configparser
from cryptography.fernet import Fernet
from PyQt6.QtWidgets import QDialog, QLabel, QLineEdit, QPushButton, QVBoxLayout, QTabWidget, QWidget, QMessageBox, \
    QApplication, QFormLayout, QTextEdit, QHBoxLayout
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

class Configuracao(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configurações")
        self.setGeometry(500, 300, 500, 400)
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), 'icone.ico')))

        # Defina o estilo para remover o padding da aba "Proprietária" e seus filhos
        self.setStyleSheet("""
            QDialog, QWidget { background-color: #2e2e2e; color: #ffffff; }  
            QLabel { color: #ffffff; }  
            QLineEdit, QPushButton { background-color: #3e3e3e; color: #ffffff; border: none; padding: 5px; }  
            QPushButton:focus, QLineEdit:focus { outline: none; }
            QTabBar::tab { background-color: #3a3a3a; color: #ffffff; padding: 10px; border: 1px solid #1e1e1e; }  
            QTabBar::tab:selected { background-color: #1e1e1e; border: 1px solid #1e1e1e; }  
            QTabWidget::pane { border: 2px solid #2e2e2e; padding: 1px; }
            QTextEdit { background-color: #3e3e3e; color: #ffffff; border: 1px solid black; padding: 0px; }  
            QTextEdit:focus { outline: none; border: 1px solid black; }
        """)

        # Define o caminho para o arquivo de configuração db.ini
        self.caminho_arquivo_ini = os.path.join(os.path.dirname(__file__), 'db.ini')

        # Cria a apims de criptografia e salva (se não existir)
        self.key = self.obter_apims_criptografia()
        self.fernet = Fernet(self.key)

        # TabWidget para abas
        self.tabs = QTabWidget(self)

        # Criação das abas
        self.aba_conexao = QWidget()
        self.aba_email = QWidget()
        self.aba_monitoramento = QWidget()
        self.aba_proprietaria = QWidget()
        self.aba_legenda = QWidget()

        # Adicionando as abas ao TabWidget
        self.tabs.addTab(self.aba_conexao, "Conexão Banco de Dados")
        self.tabs.addTab(self.aba_email, "E-mail")
        self.tabs.addTab(self.aba_monitoramento, "Monitoramento")
        self.tabs.addTab(self.aba_proprietaria, "Proprietária")
        self.tabs.addTab(self.aba_legenda, "Legenda")

        # Layouts para as abas usando QFormLayout
        self.layout_conexao = QFormLayout()
        self.layout_email = QFormLayout()
        self.layout_monitoramento = QFormLayout()
        self.layout_proprietaria = QFormLayout()
        self.layout_legenda = QFormLayout()

        # Criação e adição dos elementos
        self.criar_elementos_conexao()
        self.criar_elementos_email()
        self.criar_elementos_monitoramento()
        self.criar_elementos_proprietaria()
        self.criar_elementos_legenda()

        # Botão de teste de e-mail centralizado
        self.btn_teste_email = QPushButton("Testar Envio", self)
        self.btn_teste_email.setFixedHeight(35)
        self.btn_teste_email.setFixedWidth(125)
        self.btn_teste_email.clicked.connect(self.testar_envio_email)

        # Linha vazia para espaçamento
        self.layout_email.addRow(QLabel(""))

        # Centraliza o botão dentro de um layout horizontal
        h_layout = QHBoxLayout()
        h_layout.addStretch()
        h_layout.addWidget(self.btn_teste_email)
        h_layout.addStretch()
        self.layout_email.addRow(h_layout)

        # Botão de salvar
        self.btn_salvar = QPushButton("Salvar", self)
        self.btn_salvar.setFixedHeight(35)  # Define a altura do botão
        self.btn_salvar.setFixedWidth(125)  # Define a largura do botão
        self.btn_salvar.clicked.connect(self.salvar_configuracao)

        # Adiciona os layouts às abas
        self.aba_conexao.setLayout(self.layout_conexao)
        self.aba_email.setLayout(self.layout_email)
        self.aba_monitoramento.setLayout(self.layout_monitoramento)
        self.aba_proprietaria.setLayout(self.layout_proprietaria)
        self.aba_legenda.setLayout(self.layout_legenda)

        # Layout principal
        layout_principal = QVBoxLayout()
        layout_principal.addWidget(self.tabs)
        # Adiciona o botão "Salvar" centralizado
        layout_principal.addWidget(self.btn_salvar, alignment=Qt.AlignmentFlag.AlignCenter)
        self.setLayout(layout_principal)

        # Carrega os dados se o arquivo de configuração já existir
        self.carregar_configuracao()

    def criar_elementos_conexao(self):
        """Cria os campos para a aba Conexão Banco de Dados"""
        self.label_ip = QLabel("IP do Servidor:")
        self.input_ip = QLineEdit(self)

        self.label_porta = QLabel("Porta:")
        self.input_porta = QLineEdit(self)

        self.label_service_name = QLabel("Service Name:")
        self.input_service_name = QLineEdit(self)

        self.label_usuario = QLabel("Usuário:")
        self.input_usuario = QLineEdit(self)

        self.label_senha = QLabel("Senha:")
        self.input_senha = QLineEdit(self)
        self.input_senha.setEchoMode(QLineEdit.EchoMode.Password)

        # Adicionando widgets ao layout da aba Conexão
        self.layout_conexao.addRow(self.label_ip, self.input_ip)
        self.layout_conexao.addRow(self.label_porta, self.input_porta)
        self.layout_conexao.addRow(self.label_service_name, self.input_service_name)
        self.layout_conexao.addRow(self.label_usuario, self.input_usuario)
        self.layout_conexao.addRow(self.label_senha, self.input_senha)

    def criar_elementos_email(self):
        """Cria os campos para a aba Configuração E-mail"""
        self.label_smtp_server = QLabel("Servidor SMTP:")
        self.input_smtp_server = QLineEdit(self)

        self.label_smtp_port = QLabel("Porta SMTP:")
        self.input_smtp_port = QLineEdit(self)

        self.label_smtp_user = QLabel("Email Envio SMTP:")
        self.input_smtp_user = QLineEdit(self)

        self.label_smtp_password = QLabel("Senha Email:")
        self.input_smtp_password = QLineEdit(self)
        self.input_smtp_password.setEchoMode(QLineEdit.EchoMode.Password)

        self.label_horarios_email = QLabel("Horários de E-mail (separados por vírgula):")
        self.input_horarios_email = QLineEdit(self)

        self.label_emails = QLabel("E-mails Destinatários (separados por ponto e vírgula):")
        self.input_emails = QLineEdit(self)

        # Adicionando widgets ao layout da aba E-mails
        self.layout_email.addRow(self.label_smtp_server, self.input_smtp_server)
        self.layout_email.addRow(self.label_smtp_port, self.input_smtp_port)
        self.layout_email.addRow(self.label_smtp_user, self.input_smtp_user)
        self.layout_email.addRow(self.label_smtp_password, self.input_smtp_password)
        self.layout_email.addRow(QLabel())  # Linha vazia para espaçamento
        self.layout_email.addRow(self.label_horarios_email, self.input_horarios_email)
        self.layout_email.addRow(QLabel())  # Linha vazia para espaçamento
        self.layout_email.addRow(self.label_emails)
        self.layout_email.addRow(self.input_emails)

    def criar_elementos_monitoramento(self):
        """Cria os campos para a aba Monitoramento"""
        self.label_horarios_desconexao = QLabel("Horários de Desconexão (separados por vírgula):")
        self.input_horarios_desconexao = QLineEdit(self)

        self.label_minutos_desconexao = QLabel("Minutos para Desconexão:")
        self.input_minutos_desconexao = QLineEdit(self)

        self.label_permitidos = QLabel("Usuários Ignorados (separados por vírgula):")
        self.input_permitidos = QLineEdit(self)

        self.label_destino = QLabel("Destino de Cópia do Arquivos:")
        self.input_destino = QLineEdit(self)

        # Adicionando widgets ao layout da aba Monitoramento
        self.layout_monitoramento.addRow(self.label_horarios_desconexao, self.input_horarios_desconexao)
        self.layout_monitoramento.addRow(self.label_minutos_desconexao, self.input_minutos_desconexao)
        self.layout_monitoramento.addRow(QLabel())  # Linha vazia para espaçamento
        self.layout_monitoramento.addRow(self.label_permitidos)
        self.layout_monitoramento.addRow(self.input_permitidos)
        self.layout_monitoramento.addRow(QLabel())  # Linha vazia para espaçamento
        self.layout_monitoramento.addRow(self.label_destino)
        self.layout_monitoramento.addRow(self.input_destino)

    def criar_elementos_proprietaria(self):
        """Cria os campos para a aba Proprietária"""

        self.label_processos = QLabel("Processos:")
        self.input_processos = QTextEdit(self)
        self.input_processos.setFixedHeight(75)

        self.label_licencas = QLabel("Licenças:")
        self.input_licencas = QTextEdit(self)
        self.input_licencas.setFixedHeight(50)

        self.input_vazia = QLabel(" ")
        self.input_vazia.setFixedHeight(90)

        # Adicionando os labels e os campos de entrada em linhas separadas
        self.layout_proprietaria.addRow(self.label_processos)
        self.layout_proprietaria.addRow(self.input_processos)
        self.layout_proprietaria.addRow(self.label_licencas)
        self.layout_proprietaria.addRow(self.input_licencas)
        self.layout_proprietaria.addRow(self.input_vazia)

    def criar_elementos_legenda(self):
        """Cria os campos para a aba Legenda"""
        self.label_legenda = QLabel("Legenda:")
        self.input_legenda = QTextEdit(self)
        self.input_legenda.setFixedHeight(250)

        # Adicionando o label e o campo de entrada ao layout da aba Legenda
        self.layout_legenda.addRow(self.label_legenda)
        self.layout_legenda.addRow(self.input_legenda)


    def obter_apims_criptografia(self):
        """Gera uma apims de criptografia e salva em um arquivo"""
        apims_arquivo = os.path.join(os.path.dirname(__file__), 'apims.key')
        if not os.path.exists(apims_arquivo):
            apims = Fernet.generate_key()
            with open(apims_arquivo, 'wb') as f:
                f.write(apims)
        else:
            with open(apims_arquivo, 'rb') as f:
                apims = f.read()
        return apims

    def carregar_configuracao(self):
        """Carrega ou cria o arquivo db.ini, garantindo que todas as seções estejam presentes com valores padrão."""
        config = configparser.ConfigParser()

        # Se o arquivo db.ini já existir, vamos carregá-lo
        if os.path.exists(self.caminho_arquivo_ini):
            config.read(self.caminho_arquivo_ini)
        else:
            # Se o arquivo não existir, criamos com os valores padrão fornecidos
            config['CONEXAO'] = {
                'ip': '',
                'porta': '',
                'service_name': '',
                'usuario': '',
                'senha': ''
            }

            config['PERMITIDOS'] = {
                'usuarios': 'Silonei.Duarte, Pedro, joao.souza'
            }

            config['HORARIOS DESCONEXOES'] = {
                'horarios': '24:55'
            }

            config['MINUTOS PARA DESCONEXAO'] = {
                'minutos': '180'
            }

            config['HORARIOS E-MAILS'] = {
                'horarios': '12:40, 18:30'
            }

            config['EMAILS'] = {
                'e-mails': 'email1@seuemail.com.br; email2@seuemail.com.br'
            }

            config['DESTINO'] = {                'destino': r'\\192.168.100.230\Destino_de_Copia_do_arquivo_xlsx_para_o_Dashboard.py_ler'
            }

            config['SMTP'] = {
                'smtp_server': 'mail.seusmtp.com.br',
                'smtp_port': '465',
                'smtp_user': 'ti@seuemail.com.br',
                'smtp_password': ''
            }

            # Adicionando valores padrão para a seção PROPRIETÁRIA
            config['PROPRIETARIA'] = {
                'processos': 'All, FFOR, FTCB, FTON, MCAP, MCCO, MCMM, MCRR, MCSS, MECO, MEDU, MERP, MPLP, MPMO, MPNE, MPOP, RRCA, RVOR, RVPE, SCCP, SCOC, SCSC, SECE, SERE',
                'licencas': '25, 1, 1, 1, 4, 4, 4, 2, 6, 3, 1, 1, 3, 4, 5, 6, 1, 3, 6, 8, 8, 8, 4, 5'
            }

            # Adicionando valores padrão para a seção LEGENDA
            config['LEGENDA'] = {'legenda': """FFOR, Finanças - Gestão de Plano Financeiro - Orçamentos,
            FTCB, Finanças - Gestão de Tesouraria - Caixa e Bancos,
            FTON, Finanças - Gestão de Tesouraria - Conciliação,
            MCAP, Manufatura - Gestão de Chão de Fábrica - Apontamentos de OP/OS,
            MCCO, Manufatura - Gestão de Chão de Fábrica - Componentes (OP/OS),
            MCMM, Manufatura - Gestão de Chão de Fábrica - Manutenção de Movimentos,
            MCRR, Manufatura - Gestão de Chão de Fábrica - Remessa/Retorno Serviço Terceiros,
            MCSS, Manufatura - Gestão de Chão de Fábrica - Separação de Componentes Depósito,
            MECO, Manufatura - Gestão de Engenharia de Produto/Serviço - Composição Produto/Serviço (Modelo),
            MEDU, Manufatura - Gestão de Engenharia de Produto/Serviço - Duplicação de Roteiro/Modelo,
            MERP, Manufatura - Gestão de Engenharia de Produto/Serviço - Roteiro de Produção,
            MPLP, Manufatura - Gestão de PCP - Cancelamento de Produção,
            MPMO, Manufatura - Gestão de PCP - Manutenção de OP/OS,
            MPNE, Manufatura - Gestão de PCP - Necessidades de Produção/Compra (MRP),
            MPOP, Manufatura - Gestão de PCP - Ordens de Produção/Serviço,
            RRCA, Mercado - Gestão de Relacionamento (CRM) - Controle de Atendimento,
            RVOR, Mercado - Gestão de Vendas - Orçamentos,
            RVPE, Mercado - Gestão de Vendas - Pedidos,
            SCCP, Suprimentos - Gestão de Compras - Cotação de Preço,
            SCOC, Suprimentos - Gestão de Compras - Ordens de Compra,
            SCSC, Suprimentos - Gestão de Compras - Solicitação de Compra,
            SECE, Suprimentos - Gestão de Estoques - Controle de Estoque,
            SERE, Suprimentos - Gestão de Estoques - Requisição Eletrônica"""
            }

            # Salva o arquivo db.ini com os valores padrão
            with open(self.caminho_arquivo_ini, 'w') as configfile:
                config.write(configfile)

        # Depois de carregar ou criar o arquivo, carrega os dados nos campos da interface (descriptografando-os)
        if 'CONEXAO' in config:
            self.input_ip.setText(
                self.fernet.decrypt(config.get('CONEXAO', 'ip').encode()).decode() if config.get('CONEXAO',
                                                                                                 'ip') else "")
            self.input_porta.setText(
                self.fernet.decrypt(config.get('CONEXAO', 'porta').encode()).decode() if config.get('CONEXAO',
                                                                                                    'porta') else "")
            self.input_service_name.setText(
                self.fernet.decrypt(config.get('CONEXAO', 'service_name').encode()).decode() if config.get('CONEXAO',
                                                                                                           'service_name') else "")
            self.input_usuario.setText(
                self.fernet.decrypt(config.get('CONEXAO', 'usuario').encode()).decode() if config.get('CONEXAO',
                                                                                                      'usuario') else "")
            self.input_senha.setText(
                self.fernet.decrypt(config.get('CONEXAO', 'senha').encode()).decode() if config.get('CONEXAO',
                                                                                                    'senha') else "")

        if 'HORARIOS E-MAILS' in config:
            self.input_horarios_email.setText(config.get('HORARIOS E-MAILS', 'horarios'))
        if 'EMAILS' in config:
            self.input_emails.setText(config.get('EMAILS', 'e-mails'))

        if 'PERMITIDOS' in config:
            self.input_permitidos.setText(config.get('PERMITIDOS', 'usuarios'))
        if 'HORARIOS DESCONEXOES' in config:
            self.input_horarios_desconexao.setText(config.get('HORARIOS DESCONEXOES', 'horarios'))
        if 'MINUTOS PARA DESCONEXAO' in config:
            self.input_minutos_desconexao.setText(config.get('MINUTOS PARA DESCONEXAO', 'minutos'))
        if 'DESTINO' in config:
            self.input_destino.setText(config.get('DESTINO', 'destino'))

        if 'SMTP' in config:
            self.input_smtp_server.setText(config.get('SMTP', 'smtp_server'))
            self.input_smtp_port.setText(config.get('SMTP', 'smtp_port'))
            self.input_smtp_user.setText(config.get('SMTP', 'smtp_user'))
            self.input_smtp_password.setText(
                self.fernet.decrypt(config.get('SMTP', 'smtp_password').encode()).decode() if config.get('SMTP',
                                                                                                         'smtp_password') else "")

        # Carregando os dados da aba Proprietária
        if 'PROPRIETARIA' in config:
            self.input_processos.setText(config.get('PROPRIETARIA', 'processos'))
            self.input_licencas.setText(config.get('PROPRIETARIA', 'licencas'))

        # Carregando os dados da aba Legenda
        if 'LEGENDA' in config:
            self.input_legenda.setText(config.get('LEGENDA', 'legenda'))

    def salvar_configuracao(self):
        """Salva as configurações no arquivo db.ini com criptografia"""
        ip = self.input_ip.text()
        porta = self.input_porta.text()
        service_name = self.input_service_name.text()
        usuario = self.input_usuario.text()
        senha = self.input_senha.text()
        smtp_password = self.input_smtp_password.text()

        if ip and porta and service_name and usuario and senha and smtp_password:
            try:
                ip_criptografado = self.fernet.encrypt(ip.encode()).decode()
                porta_criptografado = self.fernet.encrypt(porta.encode()).decode()
                service_name_criptografado = self.fernet.encrypt(service_name.encode()).decode()
                usuario_criptografado = self.fernet.encrypt(usuario.encode()).decode()
                senha_criptografada = self.fernet.encrypt(senha.encode()).decode()
                smtp_password_criptografada = self.fernet.encrypt(smtp_password.encode()).decode()

                config = configparser.ConfigParser()
                config.read(self.caminho_arquivo_ini)

                config['CONEXAO'] = {
                    'ip': ip_criptografado,
                    'porta': porta_criptografado,
                    'service_name': service_name_criptografado,
                    'usuario': usuario_criptografado,
                    'senha': senha_criptografada
                }

                config['HORARIOS E-MAILS'] = {
                    'Horarios': self.input_horarios_email.text()
                }
                config['EMAILS'] = {
                    'E-mails': self.input_emails.text()
                }
                config['PERMITIDOS'] = {
                    'Usuarios': self.input_permitidos.text()
                }
                config['HORARIOS DESCONEXOES'] = {
                    'Horarios': self.input_horarios_desconexao.text()
                }
                config['MINUTOS PARA DESCONEXAO'] = {
                    'Minutos': self.input_minutos_desconexao.text()
                }
                config['DESTINO'] = {
                    'Destino': self.input_destino.text()
                }

                config['SMTP'] = {
                    'smtp_server': self.input_smtp_server.text(),
                    'smtp_port': self.input_smtp_port.text(),
                    'smtp_user': self.input_smtp_user.text(),
                    'smtp_password': smtp_password_criptografada
                }

                # Salvando os dados da aba Proprietária
                config['PROPRIETARIA'] = {
                    'processos': self.input_processos.toPlainText(),  # Corrigido para QTextEdit
                    'licencas': self.input_licencas.toPlainText()  # Corrigido para QTextEdit
                }

                # Salvando os dados da aba Legenda
                config['LEGENDA'] = {
                    'legenda': self.input_legenda.toPlainText()  # Corrigido para QTextEdit
                }

                with open(self.caminho_arquivo_ini, 'w') as configfile:
                    config.write(configfile)

                # Mensagem de sucesso
                QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")
                self.accept()
            except Exception as e:
                # Mensagem de erro com detalhes da exceção
                QMessageBox.critical(self, "Erro", f"Erro ao salvar configurações: {str(e)}")
        else:
            # Mensagem de erro se os campos obrigatórios não estiverem preenchidos
            QMessageBox.warning(self, "Erro", "Todos os campos de conexão são obrigatórios.")

    def testar_envio_email(self):
        """Testa o envio de e-mail e mostra erros."""
        try:
            smtp_server = self.input_smtp_server.text().strip()
            smtp_port = int(self.input_smtp_port.text().strip())
            smtp_user = self.input_smtp_user.text().strip()
            smtp_password = self.input_smtp_password.text().strip()
            destinatarios = [e.strip() for e in self.input_emails.text().split(';') if e.strip()]

            if not all([smtp_server, smtp_port, smtp_user, smtp_password, destinatarios]):
                QMessageBox.warning(self, "Erro", "Preencha todos os campos SMTP e e-mails de destino.")
                return



            msg = MIMEMultipart()
            msg["From"] = smtp_user
            msg["To"] = "; ".join(destinatarios)
            msg["Subject"] = "Teste de Envio de E-mail - Monitor"
            corpo = f"<html><body><p>Teste realizado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}.</p></body></html>"
            msg.attach(MIMEText(corpo, "html"))

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(smtp_user, smtp_password)
                server.send_message(msg)

            QMessageBox.information(self, "Sucesso", "E-mail de teste enviado com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro no envio", f"Falha ao enviar e-mail:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    tela_inicial = Configuracao()
    tela_inicial.show()
    sys.exit(app.exec())
