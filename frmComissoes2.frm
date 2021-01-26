VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmComissoes2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comissões não pagas"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   120
      Picture         =   "frmComissoes2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Imprimir"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Data qVendas 
      Caption         =   "qVendas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "QVendaProdutos"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbComissoes 
      Caption         =   "dbComissoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Comissao"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmComissoes2.frx":0A82
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmComissoes2.frx":0A9C
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmComissoes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String

Private Sub cmdImprime_Click()
On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  ImprimeGrid DBGrid1, Printer, dbComissoes, 5, True, , , 7, , "Comissoes não pagas", Format(Now, "long Date")
  Printer.EndDoc
  
naoImprime:
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdVerificar_Click()
With dbComissoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from comissao" & StrOrdem
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
With qVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  On Error GoTo 0
  .RecordSource = "Select venda2.*, Produtos.* from venda2, produtos where produtos.codigoproduto=venda2.codigoproduto and produtos.comissao<>0 and codigovendedor=0"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbComissoes.Recordset.AddNew
      dbComissoes.Recordset!codfun = .Recordset!CodigoVendedor
      dbComissoes.Recordset!CodProduto = .Recordset!Codigo
      dbComissoes.Recordset!codcaixa = .Recordset!CodigoFechamento
      dbComissoes.Recordset!DataCaixa = .Recordset!Data
      dbComissoes.Recordset!Qtd = .Recordset!Quantidade
      dbComissoes.Recordset!valorUnitario = .Recordset!valorUnitario
      dbComissoes.Recordset!ValorTotal = .Recordset!ValorTotal
      dbComissoes.Recordset!porcentocomissao = .Recordset("produtos.comissao")
      dbComissoes.Recordset!ValorComissao = .Recordset!ValorTotal * .Recordset("produtos.comissao")
      dbComissoes.Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
dbComissoes.Refresh

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " Desc"
Else
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
With dbComissoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from comissao" & StrOrdem
  .Refresh
End With
End Sub

Private Sub Form_Load()

StrOrdem = " order by Data"

With dbComissoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from venda2" & StrOrdem
  .Refresh
End With

End Sub
