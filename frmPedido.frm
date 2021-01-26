VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPgAntecipado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Combustível"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmPedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   3360
      Top             =   3960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from bloqueiafechamento"
      Caption         =   "dbBloqueiaFechamento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from concilianova"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Contas order by descri"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Não Recebidos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Recebidos"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Data dbPedidos 
      Caption         =   "dbPedidos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from pedidos where recebido=0 order by datalanc"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Produtos where combustivel=-1 order by descri"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbFornecedor 
      Caption         =   "dbFornecedor"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from fornecedores order by nome"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBCtls.DBCombo cboFornecedor 
      Bindings        =   "frmPedido.frx":0442
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPedido.frx":045D
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "frmPedido.frx":0475
      TabIndex        =   18
      Top             =   1320
      Width           =   9495
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtValTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtValUnitario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtQuantidade 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCodFornecedor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtCodProduto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Bindings        =   "frmPedido.frx":1A24
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmPedido.frx":1A3D
      Height          =   315
      Left            =   4320
      TabIndex        =   15
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   "Descri"
      Text            =   ""
   End
   Begin VB.Label Label8 
      Caption         =   "Conta:"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Val. Total:"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Val. Un.:"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Qtd.:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Fornecedor:"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cod.:"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPgAntecipado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFiltro As String

Private Sub CboConta_LostFocus()
With dbContas
  .Refresh
  If cboConta.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboConta.Text & "'"
  If .Recordset.NoMatch = False Then
    cboConta.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboFornecedor_LostFocus()
With dbFornecedor
  .Refresh
  If cboFornecedor.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboFornecedor.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFornecedor.Text = .Recordset!Nome
    txtCodFornecedor.Text = .Recordset!codigofornecedor
  End If
End With
End Sub

Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProduto.Text & "'"
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
  End If
End With
End Sub

Private Sub cmdConfirma_Click()
If dbPedidos.Recordset.EOF = True Then
  MsgBox "Selecione um pedido a ser confirmado!"
  Exit Sub
End If
If dbPedidos.Recordset!recebido = True Then
  MsgBox "Pedido já confirmado!"
  Exit Sub
End If
Load frmPgAntecipadoConfirma
With frmPgAntecipadoConfirma.dbPedidosConfirma
  .RecordSource = "Select *from pedidosconfirma where codigopedido=" & dbPedidos.Recordset!codigopedido
  .Refresh
End With
frmPgAntecipadoConfirma.Show vbModal
dbPedidos.Refresh
End Sub

Private Sub cmdIncluir_Click()

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If cboFornecedor.Text <> dbFornecedor.Recordset!Nome Then
  MsgBox "Escolha um fornecedor!"
  txtCodFornecedor.SetFocus
  Exit Sub
End If
If cboProduto.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Escolha um produto!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If IsNumeric(txtQuantidade.Text) = False Then
  MsgBox "Informe uma quantidade correta!"
  txtQuantidade.SetFocus
  Exit Sub
End If
If IsNumeric(txtValUnitario.Text) = False Then
  MsgBox "Informe um valor válido!"
  txtValUnitario.SetFocus
  Exit Sub
End If
If CCur(txtValUnitario.Text) <= 0 Then
  MsgBox "Informe um valor válido!"
  txtValUnitario.SetFocus
  Exit Sub
End If
If IsNumeric(txtValTotal.Text) = False Then
  MsgBox "Informe um valor válido!"
  txtValTotal.SetFocus
  Exit Sub
End If
If cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Selecione uma conta!"
  cboConta.SetFocus
  Exit Sub
End If
With dbPedidos
  .Recordset.AddNew
  .Recordset!DataLanc = Now
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!CodProduto = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!Quantidade = txtQuantidade.Text
  .Recordset!valorUnitario = txtValUnitario.Text
  .Recordset!ValorTotal = txtValTotal.Text
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!descriconta = dbContas.Recordset!Descri
  .Recordset!codigofornecedor = dbFornecedor.Recordset!codigofornecedor
  .Recordset!fornecedor = dbFornecedor.Recordset!Nome
  .Recordset.Update
End With
With dbContas
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo - CCur(txtValTotal.Text)
  .Recordset.Update
End With
With dbConciliaNova
  .Recordset.AddNew
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!DataLanc = Date
  If dbContas.Recordset!temcpmf = False Then
    .Recordset!compensado = True
    .Recordset!Data = Date
  Else
    .Recordset!compensado = False
  End If
  .Recordset!Tipo = "Pg Antecipado"
  .Recordset!Codigo = "999999991"
  .Recordset!Descri = Left("Pg. Ant. - " & cboProduto.Text & " - qtd " & txtQuantidade.Text & " - " & txtValTotal.Text, 50)
  .Recordset!NrDocumento = "222222222"
  .Recordset!Valor = -CCur(txtValTotal.Text)
  .Recordset.Update
End With
txtCodFornecedor.Text = ""
cboFornecedor.Text = ""
txtCodProduto.Text = ""
cboProduto.Text = ""
txtQuantidade.Text = ""
txtValUnitario.Text = ""
txtValTotal.Text = ""
cboConta.Text = ""
txtCodFornecedor.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
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
strFiltro = ""
With dbFornecedor
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbPedidos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
  End If
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With


Select Case Usuarios.Grupo.ControlePgAntecipado
  Case 1 'Somente leitura
    cmdIncluir.Enabled = False
    cmdConfirma.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
  Case 0
    strFiltro = ""
  Case 1
    strFiltro = " where recebido=0 "
  Case 2
    strFiltro = " where recebido=-1 "
End Select
With dbPedidos
  .RecordSource = "Select *from pedidos" & strFiltro & " order by datalanc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
  End If
End With
End Sub

Private Sub txtCodFornecedor_GotFocus()
With txtCodFornecedor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodFornecedor_LostFocus()
With dbFornecedor
  .Refresh
  If txtCodFornecedor.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigofornecedor=" & txtCodFornecedor.Text
  If .Recordset.NoMatch = False Then
    cboFornecedor.Text = .Recordset!Nome
    txtCodFornecedor.Text = .Recordset!codigofornecedor
  End If
End With
End Sub

Private Sub txtCodProduto_GotFocus()
With txtCodProduto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos
  .Refresh
  If txtCodProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
  End If
End With
End Sub

Private Sub txtQuantidade_GotFocus()
With txtQuantidade
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValTotal_GotFocus()
With txtValTotal
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValTotal_LostFocus()
With txtValTotal
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.0000")
  Total = CCur(txtValTotal.Text) / CDbl(txtQuantidade.Text)
  txtValUnitario.Text = Format(Total, "#,##0.0000")
End With
End Sub

Private Sub txtValUnitario_GotFocus()
With txtValUnitario
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValUnitario_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValUnitario_LostFocus()
With txtValUnitario
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.0000")
  Total = CDbl(txtQuantidade.Text) * CCur(txtValUnitario.Text)
  txtValTotal.Text = Format(Total, "#,##0.0000")
End With
End Sub
