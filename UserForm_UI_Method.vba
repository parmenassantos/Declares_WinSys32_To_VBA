
Option Explicit

Function uHideTitleBarAndBorder(frm As Object)
'// Hide title bar and border around the userform.
'// Ocultar barra de título e borda em torno do userform
Dim lngWindow As LongPtr
Dim lFrmHdl As LongPtr
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
Function uMakeUserformTransparent(frm As Object, Optional Color As Variant)
'// seleciona transparência no useform
'//set transparencies on userform
Dim formhandle As LongPtr
Dim bytOpacity As Byte
formhandle = FindWindow(vbNullString, frm.Caption)
If IsMissing(Color) Then Color = &H8000&        '//rgbWhite
bytOpacity = 0
SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
frm.BackColor = Color
SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Copy and paste data below in your UserForm end enjoy.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public ID As String
Public Password As String

Private Sub UserForm_Activate()
uHideTitleBarAndBorder Me
uMakeUserformTransparent Me
Me.yellow.Visible = False
Me.LogonOn.Visible = False
End Sub

Private Sub LabelPrimus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.LogonOff.Visible = True
Me.yellow.Visible = False
Me.LogonOn.Visible = False
End Sub

Private Sub LogonOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.LogonOff.Visible = False
Me.yellow.Visible = True
Me.yellow.Left = Me.LogonOff.Left
Me.yellow.Top = Me.LogonOff.Top
End Sub

Private Sub Login_Change()
End Sub
Private Sub Senha_Change()
End Sub

Private Sub yellow_Click()
Me.yellow.Visible = False
Me.LogonOn.Visible = True
Me.LogonOn.Left = Me.yellow.Left
Me.LogonOn.Top = Me.yellow.Top
Call LogonOn_Click
End Sub

Private Sub LogonOn_Click()
If Login.Text = Empty And Senha.Text = Empty Then
MsgBox "Campos vazios. Por gentileza, insira os dados", vbExclamation, "Sem Credenciais"
Exit Sub
ElseIf Senha.Text = Empty Then
MsgBox "Campo de senha está vazio. Por gentileza, insira a senha", vbExclamation, "Sem SENHA"
Exit Sub
ElseIf Login.Text = Empty Then
MsgBox "Campo de login está vazio. Por gentileza, insira o login", vbExclamation, "Sem LOGIN"
Exit Sub
End If
MsgBox "Login: " & Login.Text & " e Senha: " & Senha.Text, vbInformation, "SUCESS"
ID = Login.Text
Password = Senha.Text
Call sair_Click
End Sub

Private Sub sair_Click()
Me.Hide
End Sub

