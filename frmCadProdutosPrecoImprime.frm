VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadProdutosPrecoImprime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprime tabela de Preço"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "frmCadProdutosPrecoImprime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dbProdutosCategoria 
      Height          =   330
      Left            =   4200
      Top             =   3600
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
      RecordSource    =   "ProdutosSubCategoria"
      Caption         =   "dbProdutosCategoria"
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
      Left            =   4200
      Top             =   3240
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
      RecordSource    =   $"frmCadProdutosPrecoImprime.frx":0442
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
   Begin VB.CommandButton Command1 
      Caption         =   "Preços Alterados"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadProdutosPrecoImprime.frx":0546
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   33
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "produtosalteradetalhe.Codigo"
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
         DataField       =   "produtosalteradetalhe.Descri"
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
      BeginProperty Column02 
         DataField       =   "produtosalteradetalhe.PrecoVenda"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7200
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadProdutosPrecoImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
With dbProdutosCategoria
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    On Error GoTo TrataErro:
    If ShowPrinter(Me) = 0 Then Exit Sub
    
    Do While .Recordset.EOF = False
      With dbProdutos
        .ConnectionString = CaminhoADO
        .RecordSource = "select ProdutosAlteraDetalhe.*, produtos.*, produtosSubCategoria.* from produtosalteradetalhe, produtos, produtossubcategoria where produtosalteradetalhe.codigoproduto=produtos.codigoproduto and produtos.subcategoria=produtossubcategoria.codigosubcategoria and codigoprodutoaltera=" & frmCadProdutosPreco.dbProdutosAltera.Recordset!codigoprodutoaltera & " and subcategoria=" & dbProdutosCategoria.Recordset!codigosubcategoria & " order by produtosalteradetalhe.descri"
        .Refresh
        If .Recordset.RecordCount <> 0 Then
          ImprimeADOGrid DataGrid1, Printer, dbProdutos, , True, 2, , , , "Tabela de Preços - " & NomePosto, dbProdutosCategoria.Recordset!Descri, "Data da alteração: " & frmCadProdutosPreco.dbProdutosAltera.Recordset!DataCaixa & " - Turno: " & frmCadProdutosPreco.dbProdutosAltera.Recordset!Turno & Chr(vbKeyReturn) & "Impresso em: " & Format(Now, "Long Date")
          Printer.NewPage
        End If
      End With

      .Recordset.MoveNext
    Loop
  End If
End With
Printer.EndDoc

TrataErro:
Unload Me
End Sub

Private Sub Command1_Click()
Dim strTemp As String
On Error GoTo TrataErro:
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select ProdutosAlteraDetalhe.*, produtos.*, produtosSubCategoria.* from produtosalteradetalhe, produtos, produtossubcategoria where produtosalteradetalhe.codigoproduto=produtos.codigoproduto and produtos.subcategoria=produtossubcategoria.codigosubcategoria and codigoprodutoaltera=" & frmCadProdutosPreco.dbProdutosAltera.Recordset!codigoprodutoaltera & " and ProdutosAlteraDetalhe.precovenda<>ProdutosAlteraDetalhe.precoantigo order by ProdutosAlteraDetalhe.descri"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    DataGrid1.Font.Size = 12
    ImprimeADOGrid DataGrid1, Printer, dbProdutos, , True, 2, , , , "Preços Alterados - " & NomePosto, , "Data da alteração: " & frmCadProdutosPreco.dbProdutosAltera.Recordset!DataCaixa & " - Turno: " & frmCadProdutosPreco.dbProdutosAltera.Recordset!Turno & Chr(vbKeyReturn) & "Impresso em: " & Format(Now, "Long Date")
  Else
    MsgBox "Não existe produto com preço alterado"
    Exit Sub
  End If
End With

Printer.FontSize = 10
Printer.CurrentX = 0
Printer.Print "Data Recebido: ______/______/____________"
Printer.Print ""
Printer.Print "Assinatura do Gerente: ______________________________________________________"
Printer.EndDoc

TrataErro:
Unload Me
End Sub

Private Sub Form_Load()
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select ProdutosAlteraDetalhe.*, produtos.*, produtosSubCategoria.* from produtosalteradetalhe, produtos, produtossubcategoria where produtosalteradetalhe.codigoproduto=produtos.codigoproduto and produtos.subcategoria=produtossubcategoria.codigosubcategoria and codigoprodutoaltera=" & frmCadProdutosPreco.dbProdutosAltera.Recordset!codigoprodutoaltera & " order by produtosalteradetalhe.descri"
  .Refresh
End With
With dbProdutosCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With
End Sub
