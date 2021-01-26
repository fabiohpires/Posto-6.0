VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPedidoDeCompra2 
   Caption         =   "Entrada de Notas"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPgAntecipado 
      Caption         =   "Pagamento Antecipado"
      Height          =   255
      Left            =   1680
      TabIndex        =   38
      Top             =   960
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3375
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Width           =   9135
      Begin VB.TextBox txtValorUnitario 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtTanque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remover"
         Height          =   375
         Left            =   8040
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7320
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCodProduto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin MSDBCtls.DBCombo cboTanque 
         Bindings        =   "frmPedidoDeCompra2.frx":0000
         Height          =   315
         Left            =   5160
         TabIndex        =   18
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Tanque"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo cboProdutos 
         Bindings        =   "frmPedidoDeCompra2.frx":0017
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPedidoDeCompra2.frx":0030
         Height          =   1695
         Left            =   120
         OleObjectBlob   =   "frmPedidoDeCompra2.frx":004B
         TabIndex        =   20
         Top             =   1200
         Width           =   8775
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Left            =   4440
         TabIndex        =   34
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   6840
         TabIndex        =   32
         Top             =   3000
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2400
         TabIndex        =   31
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Qtd.:"
         Height          =   195
         Left            =   6840
         TabIndex        =   30
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdNovaNota 
      Caption         =   "Nova"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtCodFornecedor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtNrNota 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtDiasParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   4560
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Data dbPosto 
         Caption         =   "dbPosto"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Postos"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbStatus 
         Caption         =   "dbStatus"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from status"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbTanque 
         Caption         =   "dbTanque"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from tanques"
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbMovimento 
         Caption         =   "dbMovimento"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ProdutosEntrada2"
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbDespesaLanc 
         Caption         =   "dbDespesaLanc"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from despesaslanc2"
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data qNota 
         Caption         =   "qNota"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=0"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbProdutosEntrada 
         Caption         =   "dbProdutosEntrada"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtosentrada2"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbNotasCorpo 
         Caption         =   "dbNotasCorpo"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ProdutosNotasCorpo where codigoProdutoNota=0 order by codigo"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbNotas 
         Caption         =   "dbNotas"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ProdutosNotas where confirmado=0 order by codigoentrada"
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtos order by descri"
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbFornecedores 
         Caption         =   "dbFornecedores"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Fornecedores order by nome"
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data dbProdutosHistorico 
         Caption         =   "dbProdutosHistorico"
         Connect         =   "Access"
         DatabaseName    =   "posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ProdutosHistorico"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
         Height          =   330
         Left            =   2760
         Top             =   2400
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
      Begin MSAdodcLib.Adodc dbTurnos 
         Height          =   330
         Left            =   120
         Top             =   2400
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   "Select *from turnos order by horaini"
         Caption         =   "dbTurnos"
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
   End
   Begin VB.ComboBox cboFormaDePg 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8640
      Picture         =   "frmPedidoDeCompra2.frx":10E2
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Imprimir"
      Top             =   480
      Width           =   735
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmPedidoDeCompra2.frx":1B64
      Height          =   1575
      Left            =   240
      OleObjectBlob   =   "frmPedidoDeCompra2.frx":1B7A
      TabIndex        =   10
      Top             =   1320
      Width           =   9135
   End
   Begin MSDBCtls.DBCombo cboFornecedor 
      Bindings        =   "frmPedidoDeCompra2.frx":2DF5
      Height          =   315
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtEmissao 
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   78643201
      CurrentDate     =   37680
   End
   Begin MSComCtl2.DTPicker txtVencimento 
      Height          =   315
      Left            =   5640
      TabIndex        =   16
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   78643201
      CurrentDate     =   37680
   End
   Begin MSComCtl2.DTPicker txtRecebida 
      Height          =   315
      Left            =   1440
      TabIndex        =   36
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   78643201
      CurrentDate     =   37680
   End
   Begin MSComCtl2.DTPicker txtEntrega 
      Height          =   315
      Left            =   7080
      TabIndex        =   37
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   78643201
      CurrentDate     =   37680
   End
   Begin MSDataListLib.DataCombo cboTurnos 
      Bindings        =   "frmPedidoDeCompra2.frx":2E12
      Height          =   315
      Left            =   2880
      TabIndex        =   39
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento:"
      Height          =   195
      Left            =   5640
      TabIndex        =   50
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      Height          =   195
      Left            =   4200
      TabIndex        =   49
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      Height          =   195
      Left            =   960
      TabIndex        =   48
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   240
      TabIndex        =   47
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Recebida:"
      Height          =   195
      Left            =   1440
      TabIndex        =   46
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Nota:"
      Height          =   195
      Left            =   240
      TabIndex        =   45
      Top             =   6360
      Width           =   645
   End
   Begin VB.Label Label15 
      Caption         =   "Dias:"
      Height          =   255
      Left            =   1080
      TabIndex        =   44
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Entrega em:"
      Height          =   195
      Left            =   7080
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   2880
      TabIndex        =   41
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "Forma de Pg.:"
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frmPedidoDeCompra2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoNota As Double

Private Sub cboFornecedor_LostFocus()
With dbFornecedores
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFornecedor.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboFornecedor.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFornecedor.Text = .Recordset!Nome
    txtCodFornecedor.Text = .Recordset!codigofornecedor
  End If
End With
End Sub

Private Sub cboProdutos_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProdutos.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProdutos.Text & "'"
  If .Recordset.NoMatch = False Then
    cboProdutos.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
    txtValorUnitario.Text = Format(.Recordset!precocompra, "#,##0.0000")
    If .Recordset!Combustivel = True Then
      txtQtd.Text = "5000"
      With dbTanque
        .RecordSource = "select *from tanques where codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " order by tanque"
        .Refresh
      End With
    Else
      txtQtd.Text = ""
    End If
  End If
End With
End Sub

Private Sub cboTurnos_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboTurnos.Text & "'"
  If .Recordset.EOF = False Then
    cboTurnos.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdConfirma_Click()
Dim db As New ADODB.Connection
Dim dbCaixas As New ADODB.Recordset
Dim Dia As Date, Resposta As Integer
Dim dbProdutos As New ADODB.Recordset


db.Open CaminhoADO
dbCaixas.CursorLocation = adUseClient
dbCaixas.Open "select datacaixa+horaini as Dia, turno from fechamentodecaixa where fechado=-1 order by datacaixa desc, horaini desc", db, adOpenForwardOnly, adLockReadOnly

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select codigoproduto, codigo, descri, precocompra, precovenda, lucrominimo from produtos", db, adOpenKeyset, adLockOptimistic

Call cboTurnos_LostFocus

If dbTurnos.Recordset.RecordCount <> 0 Then
  If cboTurnos.Text <> dbTurnos.Recordset!Descri Then
    MsgBox "Truno não encontrado!"
    Exit Sub
  End If
End If

'If dbCaixas.RecordCount <> 0 Then
'  dbCaixas.MoveFirst
'  Dia = txtRecebida.Value + dbTurnos.Recordset!HoraIni
'  If dbCaixas!Dia >= Dia Then
'    If Usuarios.Grupo.AdmEstatus = 2 Then
'      Resposta = MsgBox("Já foi confirmado caixa " & Format(dbCaixas!Dia, "short date") & " turno " & dbCaixas!Turno & ". Deseja continuar?", vbYesNo + vbDefaultButton2)
'      If Resposta = vbNo Then Exit Sub
'      If ConfirmaNota(dbNotas.Recordset!CodigoEntrada, txtRecebida.Value, dbTurnos.Recordset!CodigoTurno, cboFormaDePg.Text, txtNrNota.Text) = False Then
'        MsgBox "Não foi possível confirmar a nota atual!"
'        Exit Sub
'      End If
'    Else
'      MsgBox "Já foi confirmado caixa igual ou inferior à entrada de nota."
'      Exit Sub
'    End If
'  End If
'End If

With dbNotasCorpo
    .Refresh
    If .Recordset.RecordCount = 0 Then
        MsgBox "A nota precisa ter pelo menos um produto lançado!"
        Exit Sub
    Else
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutos.MoveFirst
        dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
        If dbProdutos.EOF = False Then
          If IsNull(dbProdutos!lucrominimo) = False Then
            If .Recordset!valorUnitario + (.Recordset!valorUnitario * (dbProdutos!lucrominimo / 100)) > dbProdutos!PrecoVenda Then
              MsgBox "O Produto " & dbProdutos!Codigo & " - " & dbProdutos!Descri & " está com o preço de venda abaixo do sugerido!"
            End If
          End If
        End If
        .Recordset.MoveNext
      Loop
    End If
End With

With dbNotas
  .Recordset.Edit
  .Recordset!gravado = True
  .Recordset!NrNota = txtNrNota.Text
  .Recordset!datanota = txtRecebida.Value
  .Recordset!formadepg = cboFormaDePg.Text
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset.Update
  .Refresh
End With
txtNrNota.Text = ""
cboFormaDePg.Text = ""
cboTurnos.Text = ""
txtNrNota.SetFocus

End Sub

Private Sub cmdEditar_Click()
Frame1.Enabled = True
txtCodProduto.SetFocus
End Sub

Private Sub cmdExcluir_Click()
Dim Resposta As Integer
With dbNotas
  If .Recordset.EOF = True Then
    MsgBox "Selecione um pedido primeiro!"
    Exit Sub
  End If
End With
Resposta = MsgBox("Deseja excluir todo o pedido atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbNotasCorpo
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.RecordCount <> 0
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
With dbNotas
  .Recordset.Delete
  .Refresh
End With
End Sub

Private Sub cmdGravar_Click()
Frame1.Enabled = False
End Sub

Private Sub cmdGravarNota_Click()

End Sub

Private Sub cmdImprime_Click()

frmRelatPedido.Show

End Sub

Private Sub cmdIncluir_Click()
Dim vUnitario As Currency, Quantidade As Double, Total As Currency
Dim Tanque As Double

If cboProdutos.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Produto inválido!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If IsNumeric(txtQtd.Text) = False Then
  MsgBox "Informe uma quantidade válida!"
  txtQtd.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If
If IsNumeric(txtValorUnitario.Text) = False Then
  MsgBox "O valor unitário está inválido!"
  txtValorUnitario.SetFocus
  Exit Sub
End If
If CCur(txtValorUnitario.Text) <= 0 Then
  If Usuarios.Nome = "Usuário Master" Then
    Resposta = MsgBox("O valor unitário deve ser positivo! Deseja continuar?", vbYesNo)
    If Resposta = vbNo Then Exit Sub
  Else
    MsgBox "O valor unitário deve ser positivo!"
    txtValorUnitario.SetFocus
    Exit Sub
  End If
End If
Tanque = 0
If dbProdutos.Recordset!Combustivel = True Then
  If IsNumeric(cboTanque.Text) = False Then
    MsgBox "Tanque inválido!"
    cboTanque.SetFocus
    Exit Sub
  Else
    dbTanque.Recordset.FindFirst "tanque=" & cboTanque.Text
    If dbTanque.Recordset.NoMatch = True Then
      MsgBox "Tanque inválido!"
      cboTanque.SetFocus
      Exit Sub
    Else
      Tanque = CDbl(cboTanque.Text)
    End If
  End If
End If
Total = CCur(txtValor.Text)
Quantidade = CDbl(txtQtd.Text)
vUnitario = Total / Quantidade

With dbNotasCorpo
  .Refresh
  .Recordset.AddNew
  .Recordset!codigoprodutonota = CodigoNota
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!Codigo = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!valorUnitario = vUnitario
  .Recordset!Quantidade = Quantidade
  .Recordset!Total = Total
  .Recordset!Tanque = Tanque
  .Recordset!lmc = dbProdutos.Recordset!lmc
  .Recordset!codbarras = "*" & dbProdutos.Recordset!Codigo & "*"
  .Recordset.Update
End With
Call dbNotas_Reposition

txtCodProduto.Text = ""
cboProdutos.Text = ""
txtQtd.Text = ""
cboTanque.Text = ""
txtValor.Text = ""
txtValorUnitario.Text = ""

txtCodProduto.SetFocus

End Sub

Private Sub cmdNovaNota_Click()
Call cboFornecedor_LostFocus

'If DateDiff("d", Date, txtEmissao.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido com data futura!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtEmissao.Value) <= -5 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido com data anterior a 5 dias!"
'    Exit Sub
'  End If
'End If
'
'If DateDiff("d", Date, txtVencimento.Value) >= 90 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido com vencimento acima de 90 dias!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtVencimento.Value) <= -1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido já vencido!"
'    Exit Sub
'  End If
'End If
'
'If DateDiff("d", Date, txtEntrega.Value) >= 90 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido com entrega acima de 90 dias!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtEntrega.Value) <= -10 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar pedido entregue com mais de 10 dias!"
'    Exit Sub
'  End If
'End If


If dbFornecedores.Recordset.EOF = True Then
  MsgBox "Tabela de fornecedores está vazia!"
  Exit Sub
End If
If dbFornecedores.Recordset!Nome <> cboFornecedor.Text Then
  MsgBox "Fornecedor inválido!"
  cboFornecedor.SetFocus
  Exit Sub
End If
dbPosto.Refresh
If dbPosto.Recordset.EOF = True Then
  MsgBox "Erro no cadastro do posto!"
  Exit Sub
End If
If IsNumeric(txtParcelas.Text) = False Then
  MsgBox "Informe o número de parcelas!"
  txtParcelas.SetFocus
  Exit Sub
End If
If CDbl(txtParcelas.Text) > 1 Then
  If IsNumeric(txtDiasParcelas.Text) = False Then
    MsgBox "Informe quantos dias entre as parcelas!"
    txtDiasParcelas.SetFocus
    Exit Sub
  End If
Else
  txtDiasParcelas.Text = "0"
End If

With dbNotas
  .Recordset.AddNew
  CodigoNota = .Recordset!CodigoEntrada
  .Recordset!codigofornecedor = dbFornecedores.Recordset!codigofornecedor
  .Recordset!fornecedor = dbFornecedores.Recordset!Nome
  .Recordset!NrNota = " "
  .Recordset!datalancada = Now
  .Recordset!Vencimento = txtVencimento.Value
  .Recordset!datanota = txtEmissao.Value
  .Recordset!dataentrega = txtEntrega.Value
  .Recordset!Origem = "Entrada de Produtos"
  .Recordset!codigoPosto = dbPosto.Recordset!codigoPosto
  .Recordset!Parcelas = CDbl(txtParcelas.Text)
  .Recordset!Dias = CDbl(txtDiasParcelas.Text)
  .Recordset!pgantecipado = chkPgAntecipado.Value
  .Recordset.Update
End With
With dbNotasCorpo
  .RecordSource = "select *from produtosnotascorpo where codigoprodutoNota=" & CodigoNota & " order by codigo"
  .Refresh
End With
With qNota
  .RecordSource = "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=" & CodigoNota
  .Refresh
  If IsNull(.Recordset!VTotal) = False Then
    lblTotal.Caption = Format(.Recordset!VTotal, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With dbNotas
  .Refresh
  .Recordset.MoveLast
End With
txtCodFornecedor.Text = ""
cboFornecedor.Text = ""
txtNrNota.Text = ""
txtParcelas.Text = ""
txtDiasParcelas.Text = ""
Frame1.Enabled = True
txtCodProduto.SetFocus
End Sub

Private Sub cmdRemove_Click()
Dim Resposta As Integer
With dbNotasCorpo
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o produto atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With
Call dbNotas_Reposition
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbNotas_Reposition()
If dbNotas.Recordset.EOF = False And dbNotas.Recordset.BOF = False Then
  CodigoNota = dbNotas.Recordset!CodigoEntrada
  If IsNull(dbNotas.Recordset!datanota) = False Then
    txtRecebida.Value = dbNotas.Recordset!dataentrega
  End If
Else
  CodigoNota = 0
End If
With dbNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from produtosnotascorpo where codigoprodutoNota=" & CodigoNota & " order by codigo"
  .Refresh
End With
With qNota
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=" & CodigoNota
  .Refresh
  If IsNull(.Recordset!VTotal) = False Then
    lblTotal.Caption = Format(.Recordset!VTotal, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
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

With cboFormaDePg
  .Clear
  .AddItem "A Vista"
  .AddItem "Boleto"
End With

With dbNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from ProdutosNotas where confirmado=0 and gravado=0 order by codigoentrada"
  .Refresh
End With

With dbFornecedores
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutosEntrada
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qNota
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=0"
  .Refresh
End With
With dbPosto
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaLanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbMovimento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTanque
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutosHistorico
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from turnos order by horaini"
  .Refresh
End With

txtEmissao.Value = Date
txtVencimento.Value = Date
txtEntrega.Value = Date
Select Case Usuarios.Grupo.ControleNotas
  Case 1 'Somente leitura
    cmdNovaNota.Enabled = False
    cmdIncluir.Enabled = False
    cmdRemove.Enabled = False
    cmdExcluir.Enabled = False
    cmdEditar.Enabled = False
    cmdGravar.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtCodFornecedor_LostFocus()
With dbFornecedores
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodFornecedor.Text = "" Then Exit Sub
  .Recordset.FindFirst "codigofornecedor=" & txtCodFornecedor.Text
  If .Recordset.NoMatch = False Then
    cboFornecedor.Text = .Recordset!Nome
    txtCodFornecedor.Text = .Recordset!codigofornecedor
  End If
End With
End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodProduto.Text = "" Then Exit Sub
  On Error Resume Next
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    cboProdutos.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
    txtValorUnitario.Text = Format(.Recordset!precocompra, "#,##0.0000")
    If .Recordset!Combustivel = True Then
      txtQtd.Text = "5000"
    Else
      txtQtd.Text = ""
    End If
  End If
End With
End Sub

Private Sub txtEmissao_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtEmissao_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtQtd_GotFocus()
With txtQtd
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRecebida_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtRecebida_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtRecebida_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
txtValor.SelStart = 0
txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
If IsNumeric(txtValor.Text) = False Then Exit Sub
txtValor.Text = Format(txtValor.Text, "#,##0.000")
If IsNumeric(txtQtd.Text) = False Then Exit Sub
A = CCur(txtValor.Text) / CDbl(txtQtd)
txtValorUnitario.Text = Format(A, "#,##0.0000")
End Sub

Private Sub txtValorUnitario_GotFocus()
txtValorUnitario.SelStart = 0
txtValorUnitario.SelLength = Len(txtValorUnitario.Text)
End Sub

Private Sub txtValorUnitario_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValorUnitario_LostFocus()
With txtValorUnitario
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.0000")
  If IsNumeric(txtQtd.Text) = False Then Exit Sub
  A = CDbl(txtQtd.Text) * CCur(.Text)
  txtValor.Text = Format(A, "#,##0.000")
End With
End Sub
Private Sub txtVencimento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtVencimento_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtEntrega_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtEntrega_LostFocus()
Me.KeyPreview = True
End Sub

