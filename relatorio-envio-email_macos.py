import pandas as pd
import os
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from openpyxl import load_workbook

# Definir o diretório de download
download_dir = '/caminho/para/diretorio/de/download'

# Nome do arquivo a ser baixado
nome_arquivo = 'naoacessoead.xlsx'

# Caminho completo do arquivo
caminho_arquivo = os.path.join(download_dir, nome_arquivo)

# Verificar se o arquivo existe e excluí-lo, se necessário
if os.path.exists(caminho_arquivo):
    os.remove(caminho_arquivo)

# Configurar as opções do Chrome para definir o diretório de download
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,  # Para evitar a janela de diálogo de download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Inicializar o navegador com as opções configuradas
driver = webdriver.Chrome(options=chrome_options)

# Navegar até a página de login
driver.get('seu.link.com')
driver.implicitly_wait(10)

# Preencher o campo de usuário
usuario_input = driver.find_element(By.ID, "username")
usuario_input.send_keys('seu_usuario')

# Preencher o campo de senha
senha_input = driver.find_element(By.ID, "password")
senha_input.send_keys('sua_senha')

# Enviar o formulário de login
senha_input.send_keys(Keys.RETURN)

# Esperar alguns segundos para o login ser concluído
driver.implicitly_wait(10)

# Ir para página de relatório
driver.get('seu.link.com')
driver.implicitly_wait(20)

# Mostrar todos os nomes
gerar_relatorio = driver.find_element(By.XPATH, 'seu/caminho/XPath')
gerar_relatorio.click()
driver.implicitly_wait(20)

# Localizar o elemento <select> pelo ID
select_element = driver.find_element(By.ID, 'downloadtype_download')

# Criar um objeto Select
select = Select(select_element)

# Selecionar a opção pelo texto visível
select.select_by_visible_text('Microsoft Excel (.xlsx)')

# Baixar a planilha
baixar_planilha = driver.find_element(By.XPATH, 'seu/caminho/XPath')
baixar_planilha.click()
time.sleep(5)


# Função para enviar e-mail com o arquivo em anexo
def enviar_email_com_anexo(destinatario, assunto, corpo, arquivo_anexo):
    remetente = "seu_email@gmail.com"  # Insira seu e-mail aqui
    senha = "sua_senha"  # Insira sua senha aqui

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatario)
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar o arquivo à mensagem de e-mail
    with open(arquivo_anexo, 'rb') as anexo:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(arquivo_anexo)}')
    msg.attach(part)

    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()
    servidor.login(remetente, senha)
    servidor.sendmail(remetente, destinatario, msg.as_string())
    servidor.quit()

# Obter a data atual no formato dd-mm-aa
data_atual = datetime.now().strftime('%d-%m-%y')

# Enviar e-mail com a planilha como anexo
email_destinatario = ["destinatario1@example.com", "destinatario2@example.com"]
assunto_email = f"Relatório Enturmação {data_atual}"
corpo_email = """Prezados,

Encaminho o relatório de enturmação das disciplinas de XXX e XXX. O relatório está na planilha única em anexo, onde constam o CURSO, DISCIPLINAS, NOME E E-MAIL dos alunos que NUNCA acessaram as respectivas disciplinas.
Sugiro filtrar a planilha para visualização mais fácil.

Atenciosamente,
Seu Nome"""

# Enviar e-mail com o arquivo em anexo
enviar_email_com_anexo(email_destinatario, assunto_email, corpo_email, caminho_arquivo)

# Função para enviar e-mail personalizado
def enviar_email(destinatario, assunto, corpo):
    remetente = "seu_email@gmail.com"  # Insira seu e-mail aqui
    senha = "sua_senha"  # Insira sua senha aqui

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Parte HTML do corpo do e-mail
    html_part = MIMEText(corpo, 'html')
    msg.attach(html_part)

    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()
    servidor.login(remetente, senha)
    servidor.send_message(msg)
    servidor.quit()

# Carregar planilha
planilha = pd.read_excel(caminho_arquivo)

# Filtrar linhas com "Papel" diferente de "Professor"
planilha_filtrada = planilha[planilha['Papel'] != 'Professor']

# Iterar sobre cada linha da planilha filtrada e enviar e-mail personalizado
for indice, linha in planilha_filtrada.iterrows():
    nome_disc = linha['Nome completo do curso']
    nome_curso = linha['Nome da categoria']
    nome = linha['Nome completo']
    email = linha['Endereço de email']
    assunto = f"{nome_disc} - {nome_curso}"
    corpo = f"""\
<html>
  <body>
    <p>Prezado(a) {nome},</p>
    <p>Parabéns por escolher a XXX para sua graduação!</p>
    <p>Observamos que até o momento, você não acessou a Disciplina {nome_disc} do Curso {nome_curso}.</p>
    <p>Se estiver com dificuldades, acesse o <a href="seu.link.com">guia passo a passo</a> para auxiliá-lo(a) no acesso ao AVA. Persistindo o problema, solicitamos que compartilhe suas dificuldades por meio deste e-mail ou contate-nos pelo nosso <a href="seu.link.com">WhatsApp (99) 9 9999-9999</a> para que possamos prestar a assistência necessária.</p>
    <p>Agradecemos pela confiança em nossa instituição e estamos à disposição para assegurar uma experiência acadêmica de qualidade.</p>
    <p>Atenciosamente,<br>Seu Nome</p>
  </body>
</html>
"""

    enviar_email(email, assunto, corpo)

print("Processo Concluído")
