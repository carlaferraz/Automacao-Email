Este projeto é uma aplicação de automação para envio de e-mails em lote, desenvolvida em Python utilizando a biblioteca Tkinter para a interface gráfica. Ele integra com o Microsoft Outlook para enviar e-mails baseados em dados de uma planilha Excel, que deve estar no mesmo diretório que o executável.

A aplicação facilita o envio de e-mails personalizados para uma lista de contatos, processando os dados a partir de uma tabela Excel. O sistema também limita o envio de e-mails a lotes de 10 para evitar sobrecarregar o servidor e manter a estabilidade.

Funcionalidades
- Interface gráfica simples e intuitiva criada com Tkinter.
- Integração com Microsoft Outlook para envio de e-mails.
- Importação de uma tabela Excel contendo os dados dos destinatários.
- Envio automático de e-mails em lotes de 10 para evitar sobrecarga.
- Validação de campos e verificação da planilha antes do envio.
  
Pré-requisitos
Para rodar o projeto, é necessário ter as seguintes dependências instaladas:

- Python 3.x
- Tkinter: Biblioteca para criar a interface gráfica (já incluída no Python).
- Pandas: Para manipulação e leitura de dados do Excel.
- OpenPyXL: Para trabalhar com arquivos Excel (.xlsx).
- Win32com: Para integração com o Outlook.

   Você pode instalar as dependências com o seguinte comando:


   pip install pandas openpyxl pywin32
  

Como Funciona
1. Entrada de dados: A planilha Excel deve estar no mesmo diretório que o executável e deve seguir o formato correto (colunas como 'Nome', 'E-mail', etc.).
2. Configuração do Outlook: O Outlook precisa estar configurado e conectado à sua conta de e-mail.
3. Interface Gráfica: Ao iniciar a aplicação, a interface gráfica permite selecionar a planilha e iniciar o processo de envio.
4. Envio em Lotes: O sistema envia os e-mails em lotes de 10, garantindo que o servidor de e-mail não seja sobrecarregado.
5. Log de Erros: Caso algum e-mail falhe, a aplicação gera um log com o erro específico.

Instalação
1. Clone este repositório para o seu computador:

   
   git clone https://github.com/seu-usuario/seu-repositorio.git
   

2. Instale as dependências necessárias:

   
   pip install -r requirements.txt
   

3. Certifique-se de que o arquivo Excel que contém os dados está no mesmo diretório que o executável.

4. Execute a aplicação:

   
   python main.py
  

Formato da Planilha Excel
A planilha Excel deve conter os seguintes campos (colunas) para funcionar corretamente:


- E-mail: Endereço de e-mail do destinatário.



Uso da Interface
1. Abra a aplicação e selecione o arquivo Excel com os dados dos destinatários.
2. Clique em "Iniciar Envio" para começar o envio em lotes.
3. Acompanhe a barra de progresso na interface e o log de sucesso/falhas exibido na tela.

Possíveis Melhorias Futuras
- Adicionar suporte para anexos de e-mail.
- Permitir a personalização de modelos de e-mail (templates).
- Implementar controle de tempo entre lotes para evitar bloqueios de e-mails.
- Suporte para múltiplas contas de e-mail.

Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.

