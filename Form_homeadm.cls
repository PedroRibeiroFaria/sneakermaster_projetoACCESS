VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_homeadm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub btn_registro_Click()
 DoCmd.Close
     DoCmd.OpenForm ("registrartenis"), acNormal
     
End Sub

Private Sub btn_relatórios_Click()
 DoCmd.Close
     DoCmd.OpenForm ("relatorios"), acNormal
End Sub

Private Sub btn_sair_Click()
resp = MsgBox("Deseja realmente sair da conta?", vbQuestion + vbYesNo, "ATENÇÃO")
If resp = vbYes Then
 DoCmd.Close
     DoCmd.OpenForm ("login"), acNormal

End If
End Sub
