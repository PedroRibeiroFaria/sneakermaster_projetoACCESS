VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_registrartenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_registrar_Click()
If IsNull(txt_tenis) Or _
   IsNull(cmb_marca) Or _
   IsNull(cmb_categoria) Or _
   IsNull(txt_preco) = True Then
   MsgBox ("Preencha todos os campos."), vbExclamation + vbOKOnly, "ATENÇÃO"
   Exit Sub
Else
    sql = "insert into tb_tenis (tenis,marca,categoria,preco) values('" & txt_tenis & "', " & _
                                        "'" & cmb_marca & "', " & _
                                        "'" & cmb_categoria & "', " & _
                                        "'" & txt_preco & "')"
                                                                    
                                    
 db.Execute (UCase(sql))
 resp = MsgBox("Tênis registrado com sucesso, Deseja cadastrar outro?", vbQuestion + vbYesNo, "AVISO")
 If resp = vbYes Then
 Call limpar_registro
 Else
 DoCmd.Close
 DoCmd.OpenForm ("homeadm"), acNormal
End If
End If

End Sub

Private Sub btn_voltar_Click()
DoCmd.Close
        DoCmd.OpenForm ("homeadm"), acNormal
End Sub

Private Sub Form_Load()

Call conectar_banco

    cmb_categoria.AddItem "Casual"
    cmb_categoria.AddItem "Corrida"
    cmb_categoria.AddItem "Treino"
    cmb_categoria.AddItem "Skateboarding"
    
    cmb_marca.AddItem "Nike"
    cmb_marca.AddItem "Adidas"
    cmb_marca.AddItem "Puma"
    cmb_marca.AddItem "Vans"
    cmb_marca.AddItem "Converse"
    cmb_marca.AddItem "Mizuno"
    cmb_marca.AddItem "Asics"
    




    
End Sub
