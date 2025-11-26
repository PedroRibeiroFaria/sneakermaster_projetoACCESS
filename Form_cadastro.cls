VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_cadastrar_Click()
If IsNull(txt_cpf) Or _
   IsNull(txt_nome) Or _
   IsNull(txt_data_nasc) Or _
   IsNull(txt_email) Or _
   IsNull(txt_telefone) Or _
   IsNull(txt_senha) Or _
   IsNull(txt_csenha) = True Then
   MsgBox ("Preencha todos os campos."), vbExclamation + vbOKOnly, "ATENÇÃO"
   Exit Sub
   
  ElseIf txt_senha <> txt_csenha Then
MsgBox ("Senhas são diferentes"), vbExclamation + vbOKOnly, "ATENÇÃO"
Exit Sub


Else

sql = "select * from tb_cadastro where cpf ='" & txt_cpf & "'"
validar_leitura
If rs.EOF = False Then
    MsgBox ("Este CPF já está sendo utilizado!"), vbExclamation + vbOKOnly, "ATENÇÃO"
limpar_cadastro
Exit Sub
End If

    sql = "insert into tb_cadastro values('" & txt_cpf & "', " & _
                                        "'" & txt_nome & "'," & _
                                        "'" & txt_data_nasc & "'," & _
                                        "'" & txt_telefone & "'," & _
                                        "'" & txt_email & "'," & _
                                        "'" & txt_senha & "')"
                                        
    db.Execute (UCase(sql))
    MsgBox ("Conta criada com sucesso."), vbInformation + vbOKOnly, "AVISO"
    Call limpar_cadastro
    DoCmd.Close
    DoCmd.OpenForm ("login"), acNormal
End If

End Sub


Private Sub ckb_visualizar_Click()
If Me.ckb_visualizar.Value = True Then
        Me.txt_senha.InputMask = ""
        Me.txt_csenha.InputMask = ""
Else
        Me.txt_senha.InputMask = "Password"
        Me.txt_csenha.InputMask = "Password"
End If
End Sub

Private Sub Form_Load()
Call conectar_banco
End Sub

