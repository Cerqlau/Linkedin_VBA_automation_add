# LinkedIn_automation_add Versão 1.0

Este projeto utiliza a linguagem pura em VBA Excel para automatizar e otimizar o crescimento orgânico da rede Linkedin. Possui módulo de controle admininstrativo e formulário GUI. Faz uso dos métodos "Application" e "Shell (cmd)" para manipulação do navegador Google Chrome; Toda a iteração com o site é efetuada através de seletores Java Script e Xpath no próprio console da página. Possui a opção de gerar um arquivo para execução automática, que poderá ser programado através do agendador de tarefa do Windows.

## 🚀 Começando

Essas instruções permitirão que você obtenha uma cópia do projeto em operação na sua máquina local para fins de desenvolvimento e teste.

### 📋 Pré-requisitos

```
=> Excel 2019 ou superior;
=> Navegador Google Chorme instalado;
=> Para melhor compatibilidade a execução das notificações, utilizar a marco no Windows 10 ou superior. 
```

### 🔧 Pré-configurações

1- O aplicativo do excel necessita estar com a opção "Habilitar Macros VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)

2- O aplicativo do excel necessita estar com a opção "Confiar no acesso ao modelo de projeto do VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)

3- Foram utilizadas e modificadas API's do Windows para as funções de criação de GUI e alertas, estas não devem ser modificadas.

### ⚙️ Executando o programa

1- Na GUI selecione automático

2- Insira login/password

3- Inicie a macro. Ela irá gerar um novo arquivo, na pasta raiz onde se encontra a macro, com o novo código automático que deverá ser programado no agendador de tarefa do Windows. 

4- No agendador de tarefa do Windows selecione as configurações de gatilho conforme desejado.

5- Na aba "Ações", selecionar "iniciar programa"; Em "Programa/Script" inserir o caminho completo de instalação do Excel entre aspas Ex: "C:\Program Files (x86)\Microsoft Office\root\Office16\excel.exe"; Em "Adicionar Argumentos" inserir o caminho completo do arquivo gerado no passo 3 EX: "D:\username\Documentos\VBA Projetos\Linkedin_add_21_day.xlsm"

Nota²: Caso a conta de usuário do windows não esteja logada na hora definida no agendador de tarefas, não será possível executar a macro, visto que esta faz uso do método VBA "Application.SendKeys" para simular inputs no keyboard.

### 📨 Distribuição

É possivel efetuar a distribuição salvando os módulos em um pasta de trabalho habilitada para macros do vba. 

## 📦 Desenvolvimento

Lauro Cerqueira

LinkdIn: https://www.linkedin.com/in/lauro-cerqueira-70473568/

Instagram : laurorcerqueira

## 🛠️ Construído com

* [Microssoft Office Excel](https://docs.microsoft.com/pt-br/office/client-developer/excel/excel-home)
* [Visual Basic for Applications](https://docs.microsoft.com/pt-br/office/vba/api/overview/)

## 📄 Licença

Este projeto está sob a licença MIT - veja o arquivo [LICENSE.md](https://github.com/usuario/projeto/licenca) para detalhes.

## 🎁 

* Conte a outras pessoas sobre este projeto 📢
* Convide alguém da equipe para uma cerveja 🍺 
* Obrigado publicamente 🤓.
