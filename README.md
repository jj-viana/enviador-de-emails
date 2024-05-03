# Script de Automação de Relatórios e Envio de E-mails

Este script Python foi desenvolvido para automatizar o processo de obtenção de relatórios de acesso e enturmação de alunos em uma plataforma de ensino online (EAD), baseado em moodle, e enviar e-mails personalizados com base nos dados do relatório.

## Requisitos

Antes de executar o script, é necessário ter instalado os seguintes requisitos:

- Python 3.x
- Bibliotecas Python:
  - pandas
  - selenium
  - openpyxl

Além disso, é preciso ter o driver do navegador Chrome compatível com a versão instalada. O driver pode ser baixado em [ChromeDriver - WebDriver for Chrome](https://sites.google.com/a/chromium.org/chromedriver/downloads).

## Funcionamento

O script realiza as seguintes etapas:

1. **Configuração Inicial**:
   - Define o diretório de download para os arquivos.
   - Verifica se existe um arquivo anterior e o exclui, se necessário.

2. **Automação do Navegador**:
   - Inicializa um navegador Chrome automatizado utilizando o Selenium.
   - Realiza login em uma plataforma EAD.
   - Navega até a página de relatório desejada.
   - Seleciona as opções necessárias para gerar e baixar o relatório.

3. **Envio de E-mail com Anexo**:
   - Utiliza a função `enviar_email_com_anexo` para enviar um e-mail com o relatório como anexo para destinatários específicos.

4. **Envio de E-mail Personalizado**:
   - Carrega o relatório baixado em um DataFrame do pandas.
   - Filtra os dados do relatório para enviar e-mails personalizados para alunos.
   - Utiliza a função `enviar_email` para enviar e-mails personalizados para cada aluno.

5. **Finalização**:
   - Exibe uma mensagem indicando a conclusão do processo.

## Configuração

Antes de executar o script, certifique-se de:

- Escolher corretamente o script para execução, dependendo de seu sistema operacional. (Windows ou macOS)
- Atualizar as credenciais de login da plataforma EAD nas variáveis `usuario_input` e `senha_input`.
- Verificar e atualizar o diretório de download na variável `download_dir`.
- Atualizar as informações de remetente, senha e destinatários nos métodos `enviar_email_com_anexo` e `enviar_email`.

## Execução

Após configurar os requisitos e as variáveis conforme necessário, execute o script Python em seu ambiente local.

**Nota**: Certifique-se de estar conectado à internet durante a execução, pois o script depende do acesso à web para interagir com a plataforma EAD e enviar e-mails.
