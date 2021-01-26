VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadProdutosFornecedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Produtos e Fornecedores"
   ClientHeight    =   6660
   ClientLeft      =   855
   ClientTop       =   750
   ClientWidth     =   10035
   Icon            =   "frmCadProdutosFornecedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbProdutosFornecedores 
      Height          =   330
      Left            =   3720
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ProdutosFornecedores"
      Caption         =   "dbProdutosFornecedores"
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
   Begin MSAdodcLib.Adodc qProdutosFornecedores 
      Height          =   330
      Left            =   3720
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qProdutosFornecedores"
      Caption         =   "qProdutosFornecedores"
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
   Begin MSDataListLib.DataCombo cboFornecedor 
      Bindings        =   "frmCadProdutosFornecedor.frx":0442
      Height          =   315
      Left            =   5160
      TabIndex        =   23
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbFornecedores 
      Height          =   330
      Left            =   480
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Fornecedores"
      Caption         =   "dbFornecedores"
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
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   480
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Produtos"
      Caption         =   "dbProdutos"
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
      TabIndex        =   10
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtDataOrcamento 
      Height          =   300
      Left            =   8520
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   38673
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhes do produto"
      Height          =   1455
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   5415
      Begin VB.TextBox txtUnCaixa 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtEstoque 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCompra 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtVenda 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkCombustivel 
         Caption         =   "Combustível"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Un./Caixa:"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   21
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Estoque:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "$ Compra:"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "$ Venda:"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   17
         Top             =   600
         Width           =   645
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadProdutosFornecedor.frx":045F
      Height          =   2775
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Descri"
         Caption         =   "Descrição"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2984,882
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCadProdutosFornecedor.frx":0478
      Height          =   2775
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "fornecedores.CodigoFornecedor"
         Caption         =   "Cod"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Telefone"
         Caption         =   "Telefone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Contato"
         Caption         =   "Contato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "OrcadoEm"
         Caption         =   "Orcado Em"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Preco"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1920,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1035,213
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Unitário:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Dt. Orçamento:"
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Código:"
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor:"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Produtos:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCadProdutosFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub cboFornecedor_LostFocus()
With dbFornecedores
  .Refresh
  If cboFornecedor.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboFornecedor.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!codigofornecedor
    cboFornecedor.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub cmdAlterar_Click()
Dim Resposta As Integer

If qProdutosFornecedores.Recordset.EOF = True Then
  MsgBox "Selecione um item a ser alterado!"
  Exit Sub
End If
If dbProdutos.Recordset.EOF = True Then
  MsgBox "Selecione um produto primeiro!"
  Exit Sub
End If
If dbFornecedores.Recordset.EOF = True Then
  MsgBox "Erro na tabela de fornecedores!"
  Exit Sub
End If
If cboFornecedor.Text <> dbFornecedores.Recordset!Nome Then
  MsgBox "Selecione um fornecedor primeiro!"
  txtCodigo.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um preço correto!"
  txtValor.SetFocus
  Exit Sub
End If

Resposta = MsgBox("Deseja alterar o registro atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbProdutosFornecedores
  .Refresh
  .Recordset.Find "codigoprodutosfornecedores=" & qProdutosFornecedores.Recordset!codigoprodutosfornecedores
  If .Recordset.EOF = False Then
    .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
    .Recordset!codigofornecedor = dbFornecedores.Recordset!codigofornecedor
    .Recordset!orcadoem = txtDataOrcamento.Value
    .Recordset!Preco = CCur(txtValor.Text)
    .Recordset.Update
  Else
    MsgBox "Erro na tabela de produtos/fornecedores!"
    Exit Sub
  End If
End With
A = qProdutosFornecedores.Recordset.AbsolutePosition
qProdutosFornecedores.Refresh
qProdutosFornecedores.Recordset.AbsolutePosition = A
End Sub

Private Sub cmdIncluir_Click()
If dbProdutos.Recordset.EOF = True Then
  MsgBox "Selecione um produto primeiro!"
  Exit Sub
End If
If dbFornecedores.Recordset.EOF = True Then
  MsgBox "Erro na tabela de fornecedores!"
  Exit Sub
End If
If cboFornecedor.Text <> dbFornecedores.Recordset!Nome Then
  MsgBox "Selecione um fornecedor primeiro!"
  txtCodigo.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um preço correto!"
  txtValor.SetFocus
  Exit Sub
End If
With dbProdutosFornecedores
  .Recordset.AddNew
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!codigofornecedor = dbFornecedores.Recordset!codigofornecedor
  .Recordset!orcadoem = txtDataOrcamento.Value
  .Recordset!Preco = CCur(txtValor.Text)
  .Recordset.Update
End With
qProdutosFornecedores.Recordset.Requery
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer

If qProdutosFornecedores.Recordset.EOF = True Then
  MsgBox "Selecione um item a ser alterado!"
  Exit Sub
End If

Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With dbProdutosFornecedores
  .Refresh
  .Recordset.Find "codigoprodutosfornecedores=" & qProdutosFornecedores.Recordset!codigoprodutosfornecedores
  If .Recordset.EOF = False Then
    .Recordset.Delete
  Else
    MsgBox "Erro na tabela de produtos/fornecedores!"
    Exit Sub
  End If
End With
qProdutosFornecedores.Refresh
qProdutosFornecedores.Recordset.Filter = "produtos.codigoproduto=" & dbProdutos.Recordset!CodigoProduto
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbProdutos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
With dbProdutos
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  On Error Resume Next
  If .Recordset!Combustivel = True Then
    chkCombustivel.Value = vbChecked
  Else
    chkCombustivel.Value = vbUnchecked
  End If
  txtEstoque.Text = Format(.Recordset!Estoque, "#,##0")
  txtCompra.Text = Format(.Recordset!precocompra, "Currency")
  txtVenda.Text = Format(.Recordset!PrecoVenda, "Currency")
  If IsNull(.Recordset!uncaixa) = False Then
    txtUnCaixa.Text = Format(.Recordset!uncaixa, "#,##0")
  Else
    txtUnCaixa.Text = Format(0, "#,##0")
  End If
End With
With qProdutosFornecedores
  .Recordset.Filter = "produtos.codigoproduto=" & dbProdutos.Recordset!CodigoProduto
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
Dim Ws As Workspace, db As Database
txtDataOrcamento.Value = Date

strOrdem = "Codigo"
With qProdutosFornecedores
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from qProdutosFornecedores"
  .Refresh
  .Recordset.Sort = "nome, orcadoem"
End With
With dbProdutosFornecedores
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from produtosfornecedores"
  .Refresh
End With
With dbFornecedores
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from fornecedores order by nome"
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos"
  .Refresh
  .Recordset.Sort = strOrdem
End With
Select Case Usuarios.Grupo.CadProdutosFornecedores
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdRemover.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub qProdutosFornecedores_Reposition()
With qProdutosFornecedores
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  txtCodigo.Text = .Recordset("fornecedores.codigofornecedor")
  cboFornecedor.Text = .Recordset!Nome
  txtValor.Text = .Recordset!Preco
  Call cboFornecedor_LostFocus
End With
End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbFornecedores
  .Refresh
  If txtCodigo.Text = "" Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "codigofornecedor=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!codigofornecedor
    cboFornecedor.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtDataOrcamento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataOrcamento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataOrcamento_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
End With
End Sub

