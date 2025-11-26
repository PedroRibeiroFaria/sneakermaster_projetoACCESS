VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_gerar_Click()
If IsNull(cmb_relatorios) = True Then
    MsgBox ("Selecione um relatório"), vbExclamation + vbOKOnly, "ATENÇÃO"
    Exit Sub
Else
    If cmb_relatorios = "Listagem de Clientes" Then
        nome_relat = "relat_clientes_geral"
End If

    If cmb_relatorios = "Listagem de Pedidos" Then
        nome_relat = "relat_pedidos"
        End If

    If cmb_relatorios = "Listagem de Compras por CPF" Then
        nome_relat = "relat_compra_por_cpf"
        End If

    If cmb_relatorios = "Listagem de Tênis" Then
        nome_relat = "relat_produto"
        End If

    If cmb_relatorios = "Listagem de Tênis Por Marca" Then
        nome_relat = "relat_produtos_marca"
        End If

    If cmb_relatorios = "Listagem de Tênis por Categoria" Then
        nome_relat = "tb_produtos_categoria"
    End If
    DoCmd.OpenReport (nome_relat), acViewReport
End If
End Sub

Private Sub btn_voltar_Click()
DoCmd.Close
        DoCmd.OpenForm ("homeadm"), acNormal
End Sub


Private Sub Form_Load()
With cmb_relatorios
    .AddItem ("Listagem de Clientes")
    .AddItem ("Listagem de Pedidos")
    .AddItem ("Listagem de Compras por CPF")
    .AddItem ("Listagem de Tênis")
    .AddItem ("Listagem de Tênis Por Marca")
    .AddItem ("Listagem de Tênis por Categoria")
    
End With
End Sub
