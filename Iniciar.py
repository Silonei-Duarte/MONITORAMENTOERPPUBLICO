import os
import sys
import configparser
from PyQt6.QtWidgets import QApplication, QMessageBox
from Configuracao import Configuracao  # Importa a tela de configuração
from Monitor import AplicativoMonitoramento  # Importa a tela de monitoramento
from cryptography.fernet import Fernet

def verificar_configuracao_completa(caminho_arquivo_ini, apims):
    """Verifica se todos os campos obrigatórios do arquivo de configuração estão preenchidos"""
    if not os.path.exists(caminho_arquivo_ini):
        return False

    config = configparser.ConfigParser()
    config.read(caminho_arquivo_ini)
    fernet = Fernet(apims)

    try:
        # Verifica se todos os campos da seção CONEXAO estão preenchidos
        if not (config.has_section('CONEXAO') and
                config.get('CONEXAO', 'ip') and
                config.get('CONEXAO', 'porta') and
                config.get('CONEXAO', 'service_name') and
                config.get('CONEXAO', 'usuario') and
                config.get('CONEXAO', 'senha')):
            return False

        # Verifica se a senha do SMTP está preenchida e descriptografada corretamente
        if not (config.has_section('SMTP') and
                config.get('SMTP', 'smtp_server') and
                config.get('SMTP', 'smtp_port') and
                config.get('SMTP', 'smtp_user') and
                config.get('SMTP', 'smtp_password')):
            return False

        # Verifica se todos os campos foram preenchidos corretamente
        fernet.decrypt(config.get('CONEXAO', 'ip').encode())
        fernet.decrypt(config.get('CONEXAO', 'porta').encode())
        fernet.decrypt(config.get('CONEXAO', 'service_name').encode())
        fernet.decrypt(config.get('CONEXAO', 'usuario').encode())
        fernet.decrypt(config.get('CONEXAO', 'senha').encode())
        fernet.decrypt(config.get('SMTP', 'smtp_password').encode())

    except Exception as e:
        # Se ocorrer algum erro na leitura ou descriptografia, consideramos a configuração incompleta
        return False

    # Se tudo estiver ok, a configuração está completa
    return True

def main():
    # Caminho do arquivo de configuração db.ini
    caminho_arquivo_ini = os.path.join(os.path.dirname(__file__), 'db.ini')
    apims_arquivo = os.path.join(os.path.dirname(__file__), 'apims.key')

    # Verifica se a apims de criptografia existe
    if not os.path.exists(apims_arquivo):
        # Se a apims de criptografia não existir, abre a tela de configuração
        app = QApplication(sys.argv)
        tela_configuracao = Configuracao()
        if tela_configuracao.exec() == Configuracao.DialogCode.Accepted:
            janela = AplicativoMonitoramento()
            janela.show()
            sys.exit(app.exec())
        return

    # Carrega a apims de criptografia
    with open(apims_arquivo, 'rb') as f:
        apims = f.read()

    # Verifica se o arquivo de configuração está completo
    if verificar_configuracao_completa(caminho_arquivo_ini, apims):
        # Se estiver completo, abre a tela de monitoramento
        app = QApplication(sys.argv)
        janela = AplicativoMonitoramento()
        janela.show()
        sys.exit(app.exec())
    else:
        # Se não estiver completo, abre a tela de configuração
        app = QApplication(sys.argv)
        QMessageBox.warning(None, "Configuração Incompleta", "A configuração está incompleta. Por favor, preencha todos os campos.")
        tela_configuracao = Configuracao()
        if tela_configuracao.exec() == Configuracao.DialogCode.Accepted:
            janela = AplicativoMonitoramento()
            janela.show()
            sys.exit(app.exec())

if __name__ == "__main__":
    main()
