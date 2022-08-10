Attribute VB_Name = "API_Transparency_Form"
Option Explicit
#If VBA7 Then
'// 64 Bits
'// Declarações DLL para alterar ou aparência do UserForm
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

#Else
'// 32 Bits
'// Declarações DLL para alterar ou aparência do UserForm
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

#End If

'// Constantes windows para barra de título
Private Const GWL_STYLE As Long = (-16)           '//The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         '//The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       '//Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   '//Controls if the window has an icon
 
'// Constantes windows para transparência
Private Const WS_EX_LAYERED = &H80000             '//cor
Private Const LWA_COLORKEY = &H1                  '//Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     '//Only needed if you want to fade the entire userform


Function HideTitleBarAndBordar(frm As Object)

'// Ocultar barra de título e borda em torno do formulário
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
'// Build window and set window until you remove the caption, title bar and frame around the window
'// Cria a janela e define a janela até remover a legenda, a barra de título e o quadro ao redor da janela
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl

End Function

Function MakeUserformTransparent(frm As Object, Optional Color As Variant)

'//set transparencies on userform***********************************
Dim formhandle As Long
Dim bytOpacity As Byte

formhandle = FindWindow(vbNullString, frm.Caption)
If IsMissing(Color) Then Color = &H800000    '&H8000& //rgbWhite
bytOpacity = 0

SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED

frm.BackColor = Color
SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY

End Function
