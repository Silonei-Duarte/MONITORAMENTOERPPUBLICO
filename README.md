<h1>Monitoramento Senior – Aplicação Desktop (Python + PyQt6)</h1>

<p>
Aplicação desktop completa para <strong>monitoramento, análise, dashboard e automações</strong> utilizando dados do 
<strong>ERP Senior</strong>.  
A aplicação se conecta diretamente ao banco <strong>Oracle</strong> (tabelas R911MOD, R911SEC etc.).
</p>

<hr>

<h2>📌 Funcionalidades Principais</h2>

<h3>1. Monitoramento em Tempo Real</h3>
<p>Coleta, processa e exibe conexões ativas no Senior em tempo real.</p>
<ul>
  <li>Consulta SQL automática (R911MOD + R911SEC).</li>
  <li>Exibe: Conexão, Usuário e Processos.</li>
  <li>Derruba conexões diretamente pela interface.</li>
  <li>Controle de tempo individual por processo (HH:MM:SS).</li>
  <li>Detecção e cálculo de “SEG. TELA” (usuário segurando múltiplas telas).</li>
</ul>

<h3>2. Geração Automática de Planilhas Senior_*.xlsx</h3>
<p>Diariamente é criada/atualizada uma planilha contendo:</p>
<ul>
  <li>Tempo acumulado por usuário e processo.</li>
  <li>Últimos Processos (ULTS.PROC.)</li>
  <li>Processos Pendentes (PEN.PROC.)</li>
  <li>Contador de tempo parado.</li>
</ul>
<p>As 7 últimas planilhas são mantidas automaticamente.</p>

<h3>3. Dashboard Gráfico (PyQt6 + Matplotlib)</h3>
<p>Leitura completa da planilha Senior_xx.xx.xx.xlsx:</p>
<ul>
  <li>Gráfico de <strong>Tempo de Uso</strong> por processo.</li>
  <li>Gráfico de <strong>Frequência de Acessos Diários</strong>.</li>
  <li>Filtro por processo e por usuário.</li>
  <li>Exibe usuários ordenados por tempo de uso.</li>
  <li>Linhas de referência com base nas <strong>licenças configuradas</strong>.</li>
</ul>

<h3>4. Automação de E-mails</h3>
<ul>
  <li>Envio automático em horários programados.</li>
  <li>Conteúdo em HTML com tabela de conexões + legenda.</li>
  <li>Corpo possui link direto para o diretório dos dashboards.</li>
  <li>SMTP com senha criptografada (Fernet).</li>
</ul>

<h3>5. Desconexão Automática</h3>
<ul>
  <li>Derruba usuários fora da lista de permitidos após X minutos.</li>
  <li>Derruba conexões em horários agendados.</li>
  <li>Baseado no contador acumulado (CONTADOR).</li>
</ul>

<h3>6. Sistema de Configurações Completo (db.ini + Fernet)</h3>
<p>Via interface PyQt6, é possível configurar:</p>
<ul>
  <li>IP, Porta, Service Name, Usuário, Senha (criptografada).</li>
  <li>SMTP (servidor, porta, usuário, senha criptografada).</li>
  <li>Horários de desconexão.</li>
  <li>Usuários permitidos.</li>
  <li>Horários de envio de e-mail.</li>
  <li>Diretório de destino para cópia das planilhas.</li>
  <li>Processos e Licenças utilizados no dashboard.</li>
  <li>Legenda completa com códigos + descrição.</li>
</ul>

<h3>7. Janelas Auxiliares</h3>
<ul>
  <li><strong>Legenda</strong>: tabela com siglas e descrições do Senior.</li>
  <li><strong>Sobre</strong>: versão, criador e informações do sistema.</li>
</ul>

<hr>

<h2>📁 Estrutura dos Arquivos</h2>

<pre>
├── Iniciar.py               (Ponto de entrada)
├── Monitor.py               (Coleta SQL, planilha, desconexão, e-mails)
├── Dashboard.py             (Dashboard gráfico)
├── Configuracao.py          (Interface de configurações + criptografia)
├── Legenda.py               (Tabela de legenda)
├── Sobre.py                 (Informações do programa)
├── apims.key                (Chave Fernet)
├── db.ini                   (Configurações criptografadas)
└── icone.ico                (Ícone da aplicação)
</pre>

<hr>

<h2>🔒 Segurança</h2>
<ul>
  <li>Uso de <strong>Fernet</strong> para criptografia de todas as credenciais sensíveis.</li>
  <li>db.ini armazena valores criptografados.</li>
  <li>Acesso a Oracle somente via credenciais descriptografadas em runtime.</li>
</ul>

<hr>

<h2>📌 Fluxo Geral de Funcionamento</h2>

<ol>
  <li>Usuário abre o programa pela primeira vez → abre “Configurações”.</li>
  <li>Credenciais são salvas criptografadas no db.ini.</li>
  <li>Monitoramento inicia:
    <ul>
      <li>Coleta SQL a cada 60s.</li>
      <li>Atualiza planilha.</li>
      <li>Envia e-mails no horário programado.</li>
      <li>Derruba conexões quando necessário.</li>
    </ul>
  </li>
  <li>Usuário pode abrir o Dashboard a qualquer momento.</li>
</ol>

<hr>

<h2>📦 Requisitos</h2>

<pre>
Python 3.10+
PyQt6
oracledb
openpyxl
pandas
matplotlib
cryptography
</pre>

<hr>

<h2>🚀 Execução</h2>

<h3>Windows</h3>
<pre>
python Iniciar.py
</pre>

<h3>Linux</h3>
<pre>
python3 Iniciar.py
</pre>

<hr>

<h2>👨‍💻 Autor</h2>
<p><strong>Silonei Duarte</strong></p>

<hr>
<img width="671" height="979" alt="Captura de tela 2025-11-12 230403" src="https://github.com/user-attachments/assets/9959aac6-2c27-47b8-9439-5b58bc6da281" />
<img width="1914" height="1014" alt="Captura de tela 2025-11-12 230659" src="https://github.com/user-attachments/assets/2fcf981b-4177-4d2b-bc8e-24a7aec8eeda" />



