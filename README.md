# Excel-Bot

### IDEIA PRINCIPAL
- Automatizar criação de tabelas no Excel

### FUNCIONALIDADES
- Criar tabelas de acordo com a quantidade de colunas e linhas especificados pelo sistema

### PROCESSO DE CRIAÇÃO
- Anotei a ideia principal, funcionalidades e anotei o 'corpo' do projeto em um caderno de desenho.
- Fiz a utilização do método dos 5Q's que aprendi no curso de Lógica de Programação do canal __Dev Aprender__
- Criar dois arquivos .py, um para as funcionalidades e o outro para montar o bot em si.

### 'CORPO' DO PROJETO
- 1 - Página de Bem-Vindos(as)
-   1.1. Opções de criar tabela
-   ===========================
- 2 - Criar tabela
-   2.1. Mostrar opções de criação de acordo com a quantidade de colunas e linhas disponibilizados pelo sistema.
-   ===========================  
- 3 - Finalizando
-   3.1. Salvar tabela
-   3.2. Abrir tabela.
-   ===========================
  
### EXPLICANDO O SCRIPT
- Nas três primeiras linhas fiz a importação de 5 bibliotecas, sendo elas: OS, SYS, XLSXWRITER, PYAUTOGUI & TIME

- Logo abaixo eu criei a função de __Página Principal__ que serve mais para dar um norte ao usuário para ele saber em qual parte do programa ele se encontra.

- Mais abaixo criei um mini menu com duas opções, sendo elas: Criar Planilhas e Fechar o programa.

- Mais abaixo criei outro mini menu contendo as informações de quantidade de colunas e linhas disponíveis pelo sistema (Pode ser acrescentado muito mais).

- Logo após temos a primeira opção que leva o usuário a criar sua planilha, algumas informações são solicitadas como: Nome do Arquivo, titulos das colunas e os seus respectivos conteúdos.

- Salva o arquivo e o executa por meio da automação com Pyautogui.


#### ISSO SE REPETE TAMBÉM NA SEGUNDA FUNÇÃO, A DIFERENÇA É QUE ACRESCENTAMOS MAIS LINHAS DE CÓDIGO POIS FOI SOLICITADO UMA PLANILHA MAIOR.
