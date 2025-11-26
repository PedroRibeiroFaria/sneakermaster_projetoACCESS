VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_login_Click()


If IsNull(txt_login) Or _
   IsNull(txt_senha) = True Then
   
   MsgBox ("Preencha todos os campos."), vbExclamation + vbOKOnly, "ATENÇÃO"
   Exit Sub
   
   
   
ElseIf txt_login = "111.111.111-11" And txt_senha = "admin" Then
    MsgBox ("Conta de ADM logada com sucesso"), vbInformation + vbOKOnly, "AVISO"
    
     DoCmd.Close
     DoCmd.OpenForm ("homeadm"), acNormal
     
     
     Else
     
sql = "select * from tb_cadastro where cpf='" & txt_login & "' and senha ='" & txt_senha & "'"
validar_leitura
If rs.EOF = False Then
    MsgBox ("Login realizado com sucesso"), vbInformation + vbOKOnly, "AVISO"
     DoCmd.Close
        DoCmd.OpenForm ("home"), acNormal
    
Else
    MsgBox ("Conta não cadastrada ou Dados inseridos incorretos"), vbExclamation + vbOKOnly, "ATENÇÃO"
    txt_login = ""
    txt_senha = ""
    txt_login.SetFocus
End If
End If

End Sub

Private Sub ckb_visualizar_Click()
If Me.ckb_visualizar.Value = True Then
        Me.txt_senha.InputMask = ""
Else
        Me.txt_senha.InputMask = "Password"
End If
End Sub

Private Sub Comando14_Click()
resp = MsgBox("Deseja realmente sair do sistema?", vbQuestion + vbYesNo, "ATENÇÃO")
If resp = vbYes Then
   DoCmd.Quit acQuitSaveAll
End If
End Sub

Private Sub Form_Load()
Call conectar_banco
DoCmd.ShowToolbar "Ribbon", acToolbarNo

End Sub




Private Sub txt_login_LostFocus()
txt_usuario = txt_login

End Sub
