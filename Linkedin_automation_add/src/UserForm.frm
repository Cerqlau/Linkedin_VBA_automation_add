VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Useform1"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13005
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Autor.....: Lauro Cerqueira
' Contato...: laurorc@hotmail.com.br - Empresa: Lauro Cerqueira - Rotina: Public Sub loginsite2(ByVal id As String, ByVal password As String, ByVal qtd As Integer)
' Data......: 10/08/2022
' Github....: https://github.com/Cerqlau
' Linkedin..: https://www.linkedin.com/in/lauro-cerqueira-70473568/
' Descricao.: Um projeto para crescimento orgânico da rede Linkedin utilizando manipulação pura do Excel VBA
'---------------------------------------------------------------------------------------
Private Sub CheckBox_automatico_Change()
    Planilha1.Activate
    UserForm.login = Planilha1.Range("B1").Value
    UserForm.password = Planilha1.Range("B2").Value
    UserForm.search = Planilha1.Range("B6").Value
End Sub
Private Sub CheckBox1_Change()
    Planilha1.Activate
    UserForm.login = Planilha1.Range("B1").Value
    UserForm.password = Planilha1.Range("B2").Value
End Sub
Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.iniciar_antes.Visible = True
    Me.iniciar_depois.Visible = False
    Me.terminar_antes.Visible = True
    Me.terminar_depois.Visible = False
    Me.Quadro_notificacao.Visible = False
    Me.Quadro_notificacao_text.Visible = False
End Sub
Public Function Trigger_notification_Check()
    'Modifica o item de notificacao apos termino do programa
    Me.Notificacao_antes.Visible = False
    Me.Notificacao_depois.Visible = True
    Me.Notificacao_depois.Left = 335
    Me.Notificacao_depois.Top = 52
End Function
Private Sub iniciar_antes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.iniciar_antes.Visible = False
    Me.iniciar_depois.Visible = True
    Me.iniciar_depois.Left = Me.iniciar_antes.Left
    Me.iniciar_depois.Top = Me.iniciar_antes.Top
End Sub
Private Sub iniciar_depois_Click()
    'Manipulação da tela do usuário para login
    Dim login As String
    Dim senha  As String
    Dim quantidade As Integer
    Dim novonome As Integer
    login = ""
    senha = ""
    godmode = False
    login = UserForm.login
    senha = UserForm.password
    'Ativação do modo de edição da planilha
    If login = "admin" And senha = "admin" Then
       MsgBox "GOOD MOD ACTIVATED"
       ThisWorkbook.Application.Visible = True
       godmode = True
       Exit Sub
    End If
    If UserForm.search = "" Then
       quantidade = 0
    Else
        If UserForm.search > 0 Then
            quantidade = CInt(UserForm.search)
            If UserForm.search >= 150 Then
                MsgBox "EXCEDIDA A QUANTIA AUTORIZADA PELO LINKEDIN PARA ADIÇÃO SEMANAL (MÁX 150)", vbCritical
                Exit Sub
            End If
        Else
            MsgBox "CAMPO QUANTIDADE APRESENTANDO ERRO", vbCritical
            Exit Sub
        End If
    End If
    If login <> "" And senha <> "" And quantidade <> 0 Then
        
        If UserForm.CheckBox1.Value Then
            Planilha1.Activate
            Planilha1.Range("B1").ClearContents
            Planilha1.Range("B2").ClearContents
            Planilha1.Range("B1").Value = login
            Planilha1.Range("B2").Value = senha
            Planilha1.Range("B6").Value = quantidade
            Planilha2.Activate
        End If
        If Me.CheckBox_automatico.Value Then
            Planilha1.Activate
            Planilha1.Range("B1").ClearContents
            Planilha1.Range("B2").ClearContents
            Planilha1.Range("B3").ClearContents
            Planilha1.Range("B1").Value = login
            Planilha1.Range("B2").Value = senha
            Planilha1.Range("B6").Value = quantidade
            Planilha1.Range("B7").Value = "True"
            Planilha2.Activate
            ActiveWorkbook.SaveAs Application.GetSaveAsFilename & "xlsm"
            Call agendartarefa 'inicializa o agendador de tarefa
        Else
            Call loginsite2(login, senha, quantidade)
        End If
    Else
        MsgBox "TODOS OS CAMPOS DEVEM ESTAR PREENCHIDOS PARA PROSSEGUIR COM A OPERAÇÃO", vbCritical
    End If
End Sub

Private Sub Notificacao_depois_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    'ativação de notificação após conclusão da tarefa
    Quadro_notificacao_text.Text = trigger_msg ' recebe a mensagem compilada para o painel
    Me.Quadro_notificacao.Visible = True
    Me.Quadro_notificacao_text.Visible = True
End Sub
Private Sub terminar_antes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    'ações para simulação de movimento com o mouse
    Me.terminar_antes.Visible = False
    Me.terminar_depois.Visible = True
    Me.terminar_depois.Left = Me.terminar_antes.Left
    Me.terminar_depois.Top = Me.terminar_antes.Top
End Sub
Private Sub terminar_depois_Click()
    'ações para fechamento do formulário
    Call CloseUseform
    If godmode = False Then
        Application.Wait Now + TimeValue("00:00:01")
        Kill_Excel
        CloseBook_save
    Else
       ThisWorkbook.Application.Visible = True
    End If
End Sub
Private Sub UserForm_Activate()
    HideTitleBarAndBordar Me 'esconde a barra do useform
    MakeUserformTransparent Me 'esconde o fundo do useform
    Me.iniciar_depois.Visible = False
    Me.terminar_depois.Visible = False
    Me.Quadro_notificacao.Visible = False
    Me.Quadro_notificacao_text.Visible = False
    ThisWorkbook.Application.Visible = False
    Me.Notificacao_depois.Visible = False
    trigger_msg = ""
    Planilha1.Activate
    status = Planilha1.Range("B7").Value
    If status = "True" Then
        login = Planilha1.Range("B1").Value
        senha = Planilha1.Range("B2").Value
        quantidade = Planilha1.Range("B6").Value
        Windows_Notification "Linkedin automation", "Macro iniciada com sucesso!" + vbCrLf + "Created by: Lauro Cerqueira", 1
        Call loginsite2(login, senha, quantidade)
    End If
End Sub
Private Sub automtic_generate()
    'atualização do trigger do modo automático
    Planilha1.Activate
    Planilha1.Range("B7").Value = "True"
End Sub
