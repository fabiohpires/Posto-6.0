VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPgAntecipadoConfirma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirma Pagamento Antecipado"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmPgAntecipadoConfirma.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dbProdutosHistorico 
      Caption         =   "dbProdutosHistorico"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProdutosHistorico"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbEntradaProdutos 
      Caption         =   "dbEntradaProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProdutosEntrada2"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Produtos"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbNotasCorpo 
      Caption         =   "dbNotasCorpo"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProdutosNotasCorpo"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Data dbTanques 
      Caption         =   "dbTanques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tanques"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   2100
      TabIndex        =   8
      Top             =   840
      Width           =   795
   End
   Begin VB.TextBox txtQuantidade 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   900
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtTanque 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Data dbNotas 
      Caption         =   "dbNotas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Contas order by descri"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbPedidosConfirma 
      Caption         =   "dbPedidosConfirma"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PedidosConfirma"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPgAntecipadoConfirma.frx":0442
      Height          =   2295
      Left            =   570
      OleObjectBlob   =   "frmPgAntecipadoConfirma.frx":0462
      TabIndex        =   10
      Top             =   1320
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   1380
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37862
   End
   Begin VB.TextBox txtNrNota 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tanque:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Data da Nota:"
      Height          =   255
      Left            =   1380
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número da Nota:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPgAntecipadoConfirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Totaliza()
Dim Total As Double
With dbPedidosConfirma
  .RecordSource = "select *from pedidosconfirma where codigopedido=" & frmPgAntecipado.dbPedidos.Recordset!codigopedido & " order by tanque"
  .Refresh
  lblTotal.Caption = ""
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    Total = Total + .Recordset!Quantidade
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  lblTotal.Caption = Total
End With
End Sub

Private Sub cmdCancela_Click()
Unload Me
End Sub

Private Sub cmdConfirmar_Click()
Dim VariaEstoque As Currency, CodigoEntrada As Double, Total As Double
Dim Resposta As Double

With frmPgAntecipado.dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If DateDiff("d", Date, txtData.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode confirmar nota com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -15 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode confirmar nota com data anterior a 15 dias!"
    Exit Sub
  End If
End If

If txtNrNota.Text = "" Then
  MsgBox "Informe um número de nota!"
  txtNrNota.SetFocus
  Exit Sub
End If
If CDbl(lblTotal.Caption) <> frmPgAntecipado.dbPedidos.Recordset!Quantidade Then
  MsgBox "O valor total não confere!"
  Exit Sub
End If
Resposta = MsgBox("Deseja confirmar o recebimento atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Erro na tabela de produtos! Tabela de produtos vazia!"
    Exit Sub
  End If
  .Recordset.FindFirst "codigoproduto=" & frmPgAntecipado.dbPedidos.Recordset!CodigoProduto
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de produtos! Produto não encontrado!"
    Exit Sub
  End If
End With

'Verifica se vai caber no tanque!
With dbPedidosConfirma
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Precisa lançar alguma entrada!"
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    dbTanques.Recordset.FindFirst "tanque=" & .Recordset!Tanque
    If dbTanques.Recordset.NoMatch = False Then
      If dbTanques.Recordset!Estoque + .Recordset!Quantidade > dbTanques.Recordset!estoquefisico + 1000 Then
        MsgBox "O Tanque " & .Recordset!Tanque & " ficará além da sua capacidade física! Corrija o lançamento!"
        Exit Sub
      End If
    End If
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
End With

With dbNotas
  .RecordSource = "select *from produtosnotas"
  .Refresh
  .Recordset.AddNew
  CodigoEntrada = .Recordset!CodigoEntrada
  .Recordset!codigofornecedor = frmPgAntecipado.dbPedidos.Recordset!codigofornecedor
  .Recordset!fornecedor = frmPgAntecipado.dbPedidos.Recordset!fornecedor
  .Recordset!NrNota = txtNrNota.Text
  .Recordset!datalancada = Now
  .Recordset!datanota = txtData.Value
  .Recordset!Vencimento = Date
  .Recordset!Origem = "Pg. Antecipado"
  .Recordset!Confirmado = True
  .Recordset!codigoPosto = 1
  .Recordset.Update
End With
dbPedidosConfirma.Refresh
dbPedidosConfirma.Recordset.MoveFirst
VariaEstoque = frmPgAntecipado.dbPedidos.Recordset!valorUnitario - dbProdutos.Recordset!precocompra
Do While dbPedidosConfirma.Recordset.EOF = False
  With dbNotasCorpo
    .Recordset.AddNew
    .Recordset!codigoprodutonota = CodigoEntrada
    .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
    .Recordset!Codigo = dbProdutos.Recordset!Codigo
    .Recordset!Descri = dbProdutos.Recordset!Descri
    .Recordset!valorUnitario = frmPgAntecipado.dbPedidos.Recordset!valorUnitario
    .Recordset!Quantidade = dbPedidosConfirma.Recordset!Quantidade
    .Recordset!Total = .Recordset!valorUnitario * .Recordset!Quantidade
    .Recordset!Tanque = dbPedidosConfirma.Recordset!Tanque
    .Recordset.Update
  End With
  With dbEntradaProdutos
    .Recordset.AddNew
    .Recordset!CodigoFechamento = 0
    .Recordset!Data = txtData.Value
    .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
    .Recordset!Codigo = dbProdutos.Recordset!Codigo
    .Recordset!Descri = dbProdutos.Recordset!Descri
    .Recordset!PrecoAntigo = dbProdutos.Recordset!precocompra
    .Recordset!PrecoNovo = frmPgAntecipado.dbPedidos.Recordset!valorUnitario
    .Recordset!VariaEstoque = VariaEstoque
    Total = Total + dbPedidosConfirma.Recordset!Quantidade
    .Recordset!Quantidade = dbPedidosConfirma.Recordset!Quantidade
    .Recordset!valornota = dbPedidosConfirma.Recordset!Quantidade * frmPgAntecipado.dbPedidos.Recordset!valorUnitario
    .Recordset!Tanque = dbPedidosConfirma.Recordset!Tanque
    .Recordset.Update
  End With
  With dbProdutosHistorico
    .Recordset.AddNew
    .Recordset!lancadoem = Now
    .Recordset!dataalteracao = txtData.Value
    .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
    .Recordset!Codigo = dbProdutos.Recordset!Codigo
    .Recordset!descriproduto = dbProdutos.Recordset!Descri
    .Recordset!descrioperacao = "Entrada de nota:  " & CodigoNota
    .Recordset!precocompra = dbProdutos.Recordset!precocompra
    .Recordset!PrecoVenda = dbProdutos.Recordset!PrecoVenda
    .Recordset!EstoqueAnterior = dbProdutos.Recordset!Estoque
    .Recordset!Quantidade = dbPedidosConfirma.Recordset!Quantidade
    .Recordset.Update
  End With
  
  With dbTanques
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
    End If
    .Recordset.FindFirst "Tanque=" & dbPedidosConfirma.Recordset!Tanque
    If .Recordset.NoMatch = False Then
      .Recordset.Edit
      .Recordset!Estoque = .Recordset!Estoque + dbPedidosConfirma.Recordset!Quantidade
      .Recordset.Update
    Else
      MsgBox "Erro para encontrar o tanque " & dbPedidosConfirma.Recordset!Tanque
    End If
  End With
  dbPedidosConfirma.Recordset.MoveNext
Loop
With dbProdutos
  .Recordset.Edit
  
  If IsNull(.Recordset!ValorEstoque) = True Then
    .Recordset!ValorEstoque = .Recordset!precocompra * .Recordset!Estoque
  End If
  If IsNull(dbProdutos.Recordset!PrecoMedio) = True Then
    .Recordset!PrecoMedio = .Recordset!precocompra
  End If
  If IsNull(.Recordset!DifEstoque) = True Then
    .Recordset!DifEstoque = 0
  End If
  If IsNull(dbProdutos.Recordset!valordifestoque) = True Then
    .Recordset!valordifestoque = 0
  End If
  If IsNull(dbProdutos.Recordset!LucroMedio) = True Then
    .Recordset!LucroMedio = 0
  End If
  
  ValorProduto = (Total * frmPgAntecipado.dbPedidos.Recordset!valorUnitario)
  .Recordset!ValorEstoque = .Recordset!ValorEstoque + ValorProduto
  
  VariaEstoque = (.Recordset!Estoque * frmPgAntecipado.dbPedidos.Recordset!valorUnitario) - (.Recordset!Estoque * .Recordset!precocompra)
  .Recordset!Variacao = .Recordset!Variacao + VariaEstoque
  .Recordset!Estoque = .Recordset!Estoque + Total
  .Recordset!precocompra = frmPgAntecipado.dbPedidos.Recordset!valorUnitario
  If IsNull(.Recordset!qtdcomprado) = True Then
    .Recordset!qtdcomprado = 0
  End If
  If IsNull(.Recordset!valorcomprado) = True Then
    .Recordset!valorcomprado = 0
  End If
  If IsNull(.Recordset!qtdcomprado) = True Then .Recordset!qtdcomprado = 0
  .Recordset!qtdcomprado = .Recordset!qtdcomprado + Total
  .Recordset!valorcomprado = .Recordset!valorcomprado + (Total * frmPgAntecipado.dbPedidos.Recordset!valorUnitario)
  .Recordset.Update
End With
With frmPgAntecipado.dbPedidos
  .Recordset.Edit
  .Recordset!recebido = True
  .Recordset!dataentrega = Date
  .Recordset.Update
  .Refresh
End With
Unload Me
End Sub

Private Sub cmdIncluir_Click()

If IsNumeric(txtTanque.Text) = False Then
  MsgBox "Informe um Tanque válido!"
  txtTanque.SetFocus
  Exit Sub
End If
If IsNumeric(txtQuantidade.Text) = False Then
  MsgBox "Informe uma quantidade correta!"
  txtQuantidade.SetFocus
  Exit Sub
End If
With dbTanques
  .RecordSource = "select *from tanques where codigoproduto=" & frmPgAntecipado.dbPedidos.Recordset!CodigoProduto
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe tanque cadastrado para este produto!"
    Exit Sub
  Else
    .Recordset.FindFirst "tanque=" & txtTanque.Text
    If .Recordset.NoMatch = True Then
      MsgBox "Tanque inválido!"
      txtTanque.SetFocus
      Exit Sub
    End If
  End If
End With
With dbPedidosConfirma
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "tanque=" & txtTanque.Text
    If .Recordset.NoMatch = False Then
      .Recordset.Edit
    Else
      .Recordset.AddNew
    End If
  Else
    .Recordset.AddNew
  End If
  .Recordset!codigopedido = frmPgAntecipado.dbPedidos.Recordset!codigopedido
  .Recordset!Tanque = txtTanque.Text
  .Recordset!Quantidade = txtQuantidade.Text
  .Recordset.Update
End With
Totaliza
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbPedidosConfirma
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With
Totaliza
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
txtData.Value = Date
With dbNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbPedidosConfirma
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTanques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbEntradaProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutosHistorico
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Totaliza
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyPress(KeyCode)
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtQuantidade_GotFocus()
With txtQuantidade
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTanque_GotFocus()
With txtTanque
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub
