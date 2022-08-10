Attribute VB_Name = "Main"
Option Explicit
'---------------------------------------------------------------------------------------
' Autor.....: Lauro Cerqueira
' Contato...: laurorc@hotmail.com.br - Empresa: Lauro Cerqueira - Rotina: Public Sub loginsite2(ByVal id As String, ByVal password As String, ByVal qtd As Integer)
' Data......: 10/08/2022
' Github....: https://github.com/Cerqlau
' Linkedin..: https://www.linkedin.com/in/lauro-cerqueira-70473568/
' Descricao.: Um projeto para crescimento org�nico da rede Linkedin utilizando manipula��o pura do Excel VBA
'---------------------------------------------------------------------------------------
Public Sub loginsite2(ByVal id As String, ByVal password As String, ByVal qtd As Integer)
    Dim login As String, senha As String, msg_final As String
    Dim cont  As Integer, mem  As Integer, cont_list As Integer, i As Integer, j As Integer
    Dim solicitar  As String, nome As String, status As String, list_save_position As String, automatic As String
    login = id 'REPASSE DE LOGIN VIA GUI
    senha = password 'REPASSE DE PASSWORD VIA GUI
    cont = qtd ' REPASSE DE QUANTIDADE PARA ADICIONAR VIA GUI
    mem = -1
    cont_list = 0
    msg_final = ""
    'VERIFICAR SE O NAVEGADOR ENCONTRA-SE ABERTO E FECHAR VIA CMD
    Windows_Notification "Linkedin automation", "Eliminando p�ginas do Chrome abertas" + vbCrLf + "Created by: Lauro Cerqueira", 1
    Call Kill_App
    'limpa nome de opera��es antigas na planiha de registros
    Call limpa_lista_nomes_antigos
    'ABERTURA DE P�GINA VIA CMD
    Windows_Notification "Linkedin automation", "INICIANDO A ROTINA DE EXECU��o EVITE UTILIZAR TECLADO E MOUSE AT� O T�RMINO DE EXECU��O" + vbCrLf + "Created by: Lauro Cerqueira", 2
    Call Open_linekdin_page
    'INSER��O DE DADOS DOS FORMUL�RIOS LINKEDIN ATRAV�S DO APLICTION
    With Application
        .Wait Now + TimeValue("00:00:10")
        .SendKeys login
        .SendKeys "{TAB}"
        .Wait Now + TimeValue("00:00:02")
        .SendKeys senha
        .SendKeys "{TAB 3}"
        .Wait Now + TimeValue("00:00:02")
        .SendKeys "~"
        .Wait Now + TimeValue("00:00:10")
    End With
    'ABERTURA DE ABA COM P�GINA DE SUGEST�ES DE PERFIL VIA LINKEDIN
    Call Open_linekdin_network_page
    'FECHAMENTO DE ABA CHATBOX DO LINKEDIN
    With Application
      .Wait Now + TimeValue("00:00:10")
      .SendKeys "^+(i)"
      .Wait Now + TimeValue("00:00:02")
      Clipboard_Sent ("document.getElementsByClassName('msg-overlay-bubble-header__controls display-flex')[0].childNodes[8].click()")
      .SendKeys "^(v)"
      .Wait Now + TimeValue("00:00:03")
      .SendKeys "~"
      .Wait Now + TimeValue("00:00:03")
      .SendKeys "^+(i)"
      .Wait Now + TimeValue("00:00:03")
      'ROLAGEM DE P�GINA PARA CARREGAMENTO DE PERFILS
      Windows_Notification "Linkedin automation", "Carregando p�gina de sugest�es de conex�es" + vbCrLf + "Created by: Lauro Cerqueira", 1
      For i = 1 To 10
        .SendKeys "{END}"
        .Wait Now + TimeValue("00:00:05")
      Next i
      'ABERTURA DE CONSOLE DEVTOOLS PARA INSERIR JAVA SCRIPT
      Windows_Notification "Linkedin automation", "In�cio de manipula��o do site via console e javascript" + vbCrLf + "Created by: Lauro Cerqueira", 1
      .SendKeys "^+(i)"
      .Wait Now + TimeValue("00:00:05")
      ' LOOP PARA EXECU��O DE ROTINA
      For j = 1 To cont
        mem = mem + 1
        'BLOCO RESPONS�VEL POR ENVIAR A SOLICITA��O ATRAV�S DO SELETOR XPATH
        solicitar = "+(4)x+(9)+(')//+(8){[}+(2)class='relative pb2'{]}//child+(;)+(;)span{[}text+(9)+(0)='Conectar'{]}+(')+(0){[}0{]}.click+(9)+(0)"
        .SendKeys solicitar
        .Wait Now + TimeValue("00:00:03")
        .SendKeys "~"
        .Wait Now + TimeValue("00:00:01")
        .SendKeys "~"
        .Wait Now + TimeValue("00:00:02")
        'COPIA O TEXTO DO BOT�O DE SOLICITA��O
        status = "copy+(9)+(4)x+(9)+(')//+(8){[}+(2)class='relative pb2'{]}//child+(;)+(;)footer{[}+(2)class='mt2'{]}//child+(;)+(;)span+(')+(0){[}" & CStr(mem) & "{]}.innerText+(0)"
        .SendKeys status
        .Wait Now + TimeValue("00:00:01")
        .SendKeys "~"
        .Wait Now + TimeValue("00:00:02")
        'Bloco para verifica��o dO TEXTO DO BOT�O DE SOLCITA��O e coletar nomes das solicita��es efetuadas com sucesso
        If Check_Status(Salva_status_mem) Then
            nome = "copy+(9)+(4)x+(9)+(')//+(8){[}+(2)class='relative pb2'{]}//child+(;)+(;)span{[}+(2)class='discover-person-card__name t-16 t-black t-bold'{]}+(')+(0){[}" & CStr(mem) & "{]}.innerText+(0)"
            .SendKeys nome
            .Wait Now + TimeValue("00:00:03")
            .SendKeys "~"
            .Wait Now + TimeValue("00:00:02")
            If verificanomes(clipboard_text) Then
                cont_list = cont_list + 1
            Else
                nome = "copy+(9)document.getElementsByClassName+(9)'relative pb2'+(0){[}0{]}.children{[}0{]}.children{[}0{]}.children{[}" & mem & "{]}.children{[}0{]}.children{[}0{]}.children{[}2{]}.children{[}0{]}.children{[}3{]}.innerText+(0)"
                .SendKeys nome
                .Wait Now + TimeValue("00:00:03")
                .SendKeys "~"
                .Wait Now + TimeValue("00:00:02")
                If verificanomes(clipboard_text) Then
                    cont_list = cont_list + 1
                Else
                    Clipboard_Sent ("Erro ao capturar usu�rio")
                    cont_list = cont_list + 1
                End If
            End If
            list_save_position = "C" + CStr(cont_list)
            msg_final = msg_final & cont_list & " - " & Salva_lista_nomes(list_save_position) & vbCrLf
        End If
        .Wait Now + TimeValue("00:00:03")
      Next j
      'A��ES FINAIS
      salva_dados_finais (cont_list) 'Salva data e hora da �ltima utiliza��o
      Call Kill_App  'Fecha o browser via comando do CMD
      Call gerarrelatoriotxt 'Gera um relat�rio em txt
      Application.Wait Now + TimeValue("00:00:03")
      trigger_message ("Total de pessoas adicionadas: " & CStr(cont_list) & _
                        vbCrLf & vbCrLf & "Lista pessoas adicionadas: " & vbCrLf & vbCrLf & msg_final)
      UserForm.Trigger_notification_Check 'Ativa o icone de notifica��o
      Windows_Notification "Linkedin automation", "Tarefa Conclu�da com Sucesso !!!" + vbCrLf + "Created by: Lauro Cerqueira", 2
      Call CloseBook_save
      Planilha1.Activate
      automatic = Planilha1.Range("B7").Value
      If automatic = "True" Then
         Call Kill_Excel
      End If
    End With
End Sub





