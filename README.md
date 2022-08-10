# LinkedIn_automation_add Versão 1.0
Este projeto utiliza a linguagem pura em VBA Excel para automatizar e otimizar o crescimento orgânico da rede Linkedin. Possui módulo de controle admininstrativo e formulário GUI. Faz uso dos métodos "Application" e "Shell (cmd)" para manipulação do navegador Google Chrome; Toda a iteração com o site é efetuada através de seletores Java Script e Xpath no próprio console da página. Possui a opção de gerar um arquivo para execução automática, que poderá ser programado através do agendador de tarefa do Windows.

Notas: 
1- O aplicativo do excel necessita estar com a opção "Habilitar Macros VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)
2- O aplicativo do excel necessita estar com a opção "Confiar no acesso ao modelo de projeto do VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)
3- Foram utilizadas e modificadas API's do Windows para as funções de criação de GUI e alertas, estas não devem ser modificadas.

UTILIZAÇÃO EM MODO AUTOMÁTICO: 
1- Na GUI selecione automático
2- Insira login/password
3- Inicie a macro. Ela irá gerar um novo arquivo, na pasta raiz onde se encontra a macro, com o novo código automático que deverá ser programado no agendador de tarefa do Windows. 
4- No agendador de tarefa do Windows selecione as configurações de gatilho conforme desejado.
5- Na aba "Ações", selecionar "iniciar programa"; Em "Programa/Script" inserir o caminho completo de instalação do Excel entre aspas Ex: "C:\Program Files (x86)\Microsoft Office\root\Office16\excel.exe"; Em "Adicionar Argumentos" inserir o caminho completo do arquivo gerado no passo 3 EX: "D:\username\Documentos\VBA Projetos\Linkedin_add_21_day.xlsm"

Nota: Caso a conta de usuário do windows não esteja logada na hora definida no agendador de tarefas, não será possível executar a macro, visto que esta faz uso do método VBA "Application.SendKeys" para simular inputs no keyboard.
