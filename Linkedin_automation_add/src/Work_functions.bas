Attribute VB_Name = "Work_functions"
'---------------------------------------------------------------------------------------
' Autor.....: Lauro Cerqueira
' Contato...: laurorc@hotmail.com.br - Empresa: Lauro Cerqueira - Rotina: Public Sub loginsite2(ByVal id As String, ByVal password As String, ByVal qtd As Integer)
' Data......: 10/08/2022
' Github....: https://github.com/Cerqlau
' Linkedin..: https://www.linkedin.com/in/lauro-cerqueira-70473568/
' Descricao.: Um projeto para crescimento orgânico da rede Linkedin utilizando manipulação pura do Excel VBA
'---------------------------------------------------------------------------------------
Option Explicit
Global quantidade               As String
Global inicio                   As String
Global godmode                  As Boolean
Global trigger_msg              As String
Global trigger_notification     As Boolean

Public Function limpa_lista_nomes_antigos()
    Planilha1.Activate
    Planilha1.Range("C1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Function
Public Function Salva_lista_nomes(ByVal position As String)
    Planilha1.Range(position).ClearContents
    Planilha1.Range(position).NumberFormat = "@"
    Planilha1.Range(position).Value = clipboard_text
    Salva_lista_nomes = clipboard_text
    Planilha1.Range(position).Offset(0, 1).ClearContents
    Planilha1.Range(position).Offset(0, 1).Value = Now
End Function
Public Function Salva_status_mem()
    Planilha1.Range("B5").ClearContents
    Planilha1.Range("B5").NumberFormat = "@"
    Planilha1.Range("B5").Value = clipboard_text
    Salva_status_mem = clipboard_text
End Function
Public Function Check_Status(ByVal status As String)
    Select Case status
    Case "Pendente": Check_Status = True
    Case Is <> "Pendente": Check_Status = False
    End Select
End Function
Public Function salva_dados_finais(ByVal qtd As Integer)
    Planilha1.Range("B3").Value = Now
    Planilha1.Range("B4").Value = qtd
End Function
Public Function backup_solicitacoes()
    Dim qtd_linhas As Integer
    Application.ScreenUpdating = False
    Planilha1.Activate
    Planilha1.Range("C1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    Planilha4.Activate
    Planilha4.Range("A1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).PasteSpecial
    Application.ScreenUpdating = True
End Function
Public Function verificanomes(ByVal nome As String)
    If nome = "" Or nome = "Cargo do usuário" Then
        verificanomes = False
    Else
        verificanomes = True
    End If
End Function
Public Function clipboard_text()
    Dim transfer As New DataObject
    Dim texto As String
    On Error GoTo erro
        transfer.GetFromClipboard
        texto = transfer.GetText
        clipboard_text = texto
erro:
    If IsEmpty(clipboard_text) Then
        clipboard_text = ""
    End If
End Function
Public Function Clipboard_Sent(msg As String)
    Dim transfer As New DataObject
    transfer.SetText msg
    transfer.PutInClipboard
End Function
Public Function trigger_message(msg As String)
trigger_msg = msg
End Function
Public Sub Kill_App()
    Shell "cmd /c taskkill /f /im chrome.exe"
    With Application
    .Wait Now + TimeValue("00:00:03")
    End With
End Sub
Public Sub Kill_Excel()
    Shell "cmd /c taskkill /f /im EXCEL.EXE"
End Sub
Public Sub CloseBook_save()
    Application.DisplayAlerts = False
    Planilha2.Activate
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Public Sub CloseUseform()
Unload UserForm
End Sub
Public Sub Open_linekdin_page()
Shell ("cmd /c start chrome --start-maximized /incognito www.linkedin.com\login  ")
End Sub
Public Function Open_linekdin_network_page()
 Shell ("cmd /c start chrome --start-maximized /incognito www.linkedin.com/mynetwork ")
End Function
Public Sub Retângulo1_Clique()
UserForm.Show
End Sub
Public Function agendartarefa()
Windows_Notification "Linkedin automation", "Realize o Agendamento para execução do arquivo gerado!" + vbCrLf + "Created by: Lauro Cerqueira", 2
Shell ("cmd /c taskschd.msc")
End Function
Public Function gerarrelatoriotxt()
    Dim data As String
    data = CStr(Format(Date, "DD.MM.YYYY"))
    Application.ScreenUpdating = False
    Open ThisWorkbook.Path & "\Log_de_execução_" & data & ".txt" For Append As 1
    Planilha1.Activate
    Planilha1.Range("C1").Select
    Do While Selection.Value <> ""
        Print #1, Selection.Value & " | " & Selection.Offset(0, 1).Value
        Selection.Offset(1, 0).Select
    Loop
    Close 1
    Application.ScreenUpdating = True
End Function
Private Sub unhide_planilha()
    ThisWorkbook.Application.Visible = True
End Sub
