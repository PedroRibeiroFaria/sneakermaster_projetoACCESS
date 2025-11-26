VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_carrinho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_comprar_Click()
If IsNull(txt_tenis) Or _
   IsNull(txt_marca) Or _
   IsNull(txt_categoria) Or _
   IsNull(txt_preco) Or _
   IsNull(cmb_tamanho) Or _
   IsNull(txt_qtde) Or _
   IsNull(txt_cep) Or _
   IsNull(txt_num_comp) Or _
   IsNull(txt_total) = True Then
   MsgBox ("Preencha todos os campos."), vbExclamation + vbOKOnly, "ATENÇÃO"
   Exit Sub
   

Else


     sql = "insert into tb_pedidos (cpf, tenis, marca, categoria, tamanho, preco, quantidade, cep, endereco, bairro, cidade, numero, uf, total_pedido) values ('" & txt_cpf & "', " & _
                                        "'" & txt_tenis & "'," & _
                                        "'" & txt_marca & "'," & _
                                        "'" & txt_categoria & "'," & _
                                        "'" & cmb_tamanho & "'," & _
                                        "'" & txt_preco & "'," & _
                                        "'" & txt_qtde & "'," & _
                                        "'" & txt_cep & "'," & _
                                        "'" & txt_endereco & "'," & _
                                        "'" & txt_bairro & "'," & _
                                        "'" & txt_cidade & "'," & _
                                        "'" & txt_num_comp & "'," & _
                                        "'" & txt_uf & "'," & _
                                        "'" & txt_total & "')"
                                        
                                        
 
    db.Execute (UCase(sql))
     resp = MsgBox("Compra registrada com sucesso, Deseja comprar novamente?", vbQuestion + vbYesNo, "AVISO")
 If resp = vbYes Then
 Call limpar_compra
 Else
 DoCmd.Close
 DoCmd.OpenForm ("home"), acNormal
End If
        
        
End If



End Sub



Private Sub btn_frete_Click()
If txt_uf.Value = "SP" Then
  txt_frete.Value = "Frete Grátis"
  txt_total = txt_subtotal
 Else
 txt_frete.Value = 10
 txt_total = txt_subtotal + 10
 End If
 

 
End Sub

Private Sub btn_voltar_Click()
DoCmd.Close
        DoCmd.OpenForm ("home"), acNormal
End Sub

Private Sub cmb_tenis_AfterUpdate()
    On Error GoTo TratarErro
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim tenisSelecionado As Variant


    tenisSelecionado = Me.cmb_tenis.Value


    If Not IsNull(tenisSelecionado) Then
        Set db = CurrentDb
        
        strSQL = "SELECT Marca, Categoria, Preco FROM tb_tenis WHERE ID_Tenis = " & tenisSelecionado & ";"
        
        Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
        If Not rs.EOF Then
            
            Me.txt_marca.Value = rs!Marca
            Me.txt_categoria.Value = rs!Categoria
            Me.txt_preco.Value = rs!Preco
        Else
            
            Me.txt_marca.Value = ""
            Me.txt_categoria.Value = ""
            Me.txt_preco.Value = ""
            MsgBox "Nenhum registro encontrado para o Tênis selecionado.", vbExclamation
        End If
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    Else
        MsgBox "Por favor, selecione um Tênis na lista.", vbInformation
    End If
    
    Exit Sub

TratarErro:
    MsgBox "Ocorreu um erro ao buscar os dados: " & Err.Description, vbCritical
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing


        
End Sub

Private Sub cmb_tenis_Click()
Me.txt_tenis.Value = cmb_tenis.Column(1)
End Sub

Private Sub Form_Load()




Call conectar_banco


    cmb_tamanho.AddItem "34"
    cmb_tamanho.AddItem "35"
    cmb_tamanho.AddItem "36"
    cmb_tamanho.AddItem "37"
    cmb_tamanho.AddItem "38"
    cmb_tamanho.AddItem "39"
    cmb_tamanho.AddItem "40"
    cmb_tamanho.AddItem "41"
    cmb_tamanho.AddItem "42"
    cmb_tamanho.AddItem "43"
    cmb_tamanho.AddItem "44"
    cmb_tamanho.AddItem "45"
    cmb_tamanho.AddItem "46"
    cmb_tamanho.AddItem "47"
    cmb_tamanho.AddItem "48"


End Sub


Private Sub txt_cep_LostFocus()
sql = "select * from tb_cep where cep ='" & txt_cep & "'"
validar_leitura
If rs.EOF = False Then
    txt_endereco = rs.Fields(1).Value
    txt_cidade = rs.Fields(2).Value
    txt_bairro = rs.Fields(3).Value
    txt_uf = rs.Fields(4).Value
    txt_num_comp.SetFocus
    txt_endereco.Enabled = False
    txt_cidade.Enabled = False
    txt_bairro.Enabled = False
    txt_uf.Enabled = False
Else
    MsgBox ("CEP não localizado!"), vbInformation + vbOKOnly, "AVISO"
End If
End Sub



Private Sub txt_qtde_AfterUpdate()
If Not IsNull(txt_qtde) Then
    txt_subtotal = txt_preco * txt_qtde
Else
    txt_subtotal = ""
End If
End Sub



