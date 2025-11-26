Attribute VB_Name = "mod_global"
Option Compare Database

Public db As Database
Public rs As Recordset
Public sql, nome_relat, resp As String




Function conectar_banco()
Set db = CurrentDb
End Function

Function validar_leitura()
Set rs = db.OpenRecordset(sql, dbOpenDynaset)
End Function

Function limpar_cadastro()
With Form_cadastro
    .txt_cpf = Empty
    .txt_nome = Empty
    .txt_data_nasc = Empty
    .txt_email = Empty
    .txt_telefone = Empty
    .txt_senha = Empty
    .txt_csenha = Empty
    .txt_cpf.SetFocus

End With
End Function

Function limpar_registro()
With Form_registrartenis
    .txt_tenis = Empty
    .cmb_marca = Empty
    .cmb_categoria = Empty
    .txt_preco = Empty
    
End With
End Function

Function limpar_compra()
With Form_carrinho
.cmb_tenis = Empty
    .txt_marca = Empty
    .txt_categoria = Empty
    .txt_preco = Empty
    .cmb_tamanho = Empty
    .txt_qtde = Empty
    .txt_subtotal = Empty
     .txt_endereco = Empty
     .txt_cep = Empty
     .txt_num_comp = Empty
     .txt_bairro = Empty
     .txt_cidade = Empty
     .txt_uf = Empty
     .txt_total = Empty
     .txt_frete = Empty
     
End With
     
    
End Function
