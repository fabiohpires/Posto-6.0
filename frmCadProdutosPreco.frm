VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadProdutosPreco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de preços"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "frmCadProdutosPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboRelatorio 
      Height          =   315
      ItemData        =   "frmCadProdutosPreco.frx":0442
      Left            =   120
      List            =   "frmCadProdutosPreco.frx":044C
      TabIndex        =   10
      Text            =   "Tabela de Preços"
      Top             =   5640
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc dbProdutosAltera2 
      Height          =   330
      Left            =   4440
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
      RecordSource    =   "select *from ProdutosAltera order by datacaixa desc, turno desc"
      Caption         =   "dbProdutosAltera2"
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
   Begin MSAdodcLib.Adodc qProdutosAltera 
      Height          =   330
      Left            =   4440
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
      RecordSource    =   "QProdutosAltera"
      Caption         =   "qProdutosAltera"
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
   Begin MSAdodcLib.Adodc dbProdutosAlteraDetalhe 
      Height          =   330
      Left            =   4440
      Top             =   2880
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
      RecordSource    =   "select *from ProdutosAlteraDetalhe order by codigo"
      Caption         =   "dbProdutosAlteraDetalhe"
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
   Begin MSAdodcLib.Adodc dbProdutosAltera 
      Height          =   330
      Left            =   4440
      Top             =   2520
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
      RecordSource    =   "select *from ProdutosAltera order by codigoprodutoaltera"
      Caption         =   "dbProdutosAltera"
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
      Left            =   4440
      Top             =   2160
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
      RecordSource    =   "select *from Produtos where combustivel=0 order by codigo"
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
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmCadProdutosPreco.frx":0476
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbTurnos 
      Height          =   330
      Left            =   4440
      Top             =   1800
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
      RecordSource    =   "select *from Turnos order by horaini"
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
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAtualiza 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtDataCaixa 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39140
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadProdutosPreco.frx":048D
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8705
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "DataCaixa"
         Caption         =   "Data Caixa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Turno"
         Caption         =   "Turno"
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
         DataField       =   "Alterado"
         Caption         =   "Alterado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sim"
            FalseValue      =   "Não"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   764,787
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCadProdutosPreco.frx":04AC
      Height          =   5415
      Left            =   3600
      TabIndex        =   9
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column02 
         DataField       =   "PrecoCompra"
         Caption         =   "Compra"
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
      BeginProperty Column03 
         DataField       =   "PrecoAntigo"
         Caption         =   "Antigo"
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
      BeginProperty Column04 
         DataField       =   "PrecoVenda"
         Caption         =   "Venda"
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
         BeginProperty Column00 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1934,929
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Turno da Alteração:"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Dia da alteração:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCadProdutosPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub IncluirProdutosNaLista()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If dbProdutosAlteraDetalhe.Recordset.RecordCount = 0 Then Exit Do
      dbProdutosAlteraDetalhe.Recordset.MoveFirst
      dbProdutosAlteraDetalhe.Recordset.Find "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutosAlteraDetalhe.Recordset.EOF = True Then
        dbProdutosAlteraDetalhe.Recordset.AddNew
        dbProdutosAlteraDetalhe.Recordset!codigoprodutoaltera = dbProdutosAltera.Recordset!codigoprodutoaltera
        dbProdutosAlteraDetalhe.Recordset!CodigoProduto = .Recordset!CodigoProduto
        dbProdutosAlteraDetalhe.Recordset!Codigo = .Recordset!Codigo
        dbProdutosAlteraDetalhe.Recordset!Descri = "*" & .Recordset!Descri
        dbProdutosAlteraDetalhe.Recordset!precocompra = .Recordset!precocompra
        dbProdutosAlteraDetalhe.Recordset!PrecoVenda = .Recordset!PrecoVenda
        dbProdutosAlteraDetalhe.Recordset!PrecoAntigo = .Recordset!PrecoVenda
        dbProdutosAlteraDetalhe.Recordset!Combustivel = .Recordset!Combustivel
        dbProdutosAlteraDetalhe.Recordset.Update
      End If
      
      .Recordset.MoveNext
    Loop
  End If
End With
End Sub

Private Sub AtualizaLista()
Dim CodigoAltera As Double
CodigoAltera = 0

DataGrid2.AllowUpdate = False

With dbProdutosAltera
  If .Recordset.EOF = False Then
    If .Recordset.BOF = False Then
      If IsNull(.Recordset!codigoprodutoaltera) = False Then
        CodigoAltera = .Recordset!codigoprodutoaltera
      End If
      If Usuarios.Nome = "Usuário Master" Then
        DataGrid2.AllowUpdate = True
        DataGrid2.Columns(3).Locked = False
      Else
        If .Recordset.AbsolutePosition <> .Recordset.RecordCount - 1 Then
          DataGrid2.AllowUpdate = False
        Else
          DataGrid2.AllowUpdate = True
        End If
        If .Recordset!Alterado = True Then
          DataGrid2.AllowUpdate = False
        Else
          DataGrid2.AllowUpdate = True
        End If
      End If
    End If
  End If
  
End With
With dbProdutosAlteraDetalhe
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtosalteradetalhe where codigoprodutoaltera=" & CodigoAltera & strOrdem
  .Refresh
End With
End Sub

Private Sub VerificaFechamentos()
Dim Ws As Workspace, db As Database, dbFechamento As Recordset, dbAlteracoes As Recordset
Dim TempData As Date, TempHora As Date

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , False)

Set dbFechamento = db.OpenRecordset("select *from Fechamentodecaixa where fechado=-1 order by datacaixa desc, horaini desc")
If dbFechamento.RecordCount = 0 Then
  Exit Sub
End If
TempData = dbFechamento!DataCaixa
TempHora = dbFechamento!HoraIni

Set dbAlteracoes = db.OpenRecordset("select *from produtosaltera order by datacaixa desc, horaini desc")

If dbAlteracoes.RecordCount <> 0 Then
  dbAlteracoes.MoveLast
  dbAlteracoes.MoveFirst
  Do While dbAlteracoes.EOF = False
    If dbAlteracoes!DataCaixa <= TempData Then
      If dbAlteracoes!HoraIni < TempHora Then
        dbAlteracoes.Edit
        dbAlteracoes!Alterado = True
        dbAlteracoes.Update
      ElseIf dbAlteracoes!DataCaixa < TempData Then
        dbAlteracoes.Edit
        dbAlteracoes!Alterado = True
        dbAlteracoes.Update
      Else
        If dbAlteracoes.AbsolutePosition = 0 Then
          dbAlteracoes.Edit
          dbAlteracoes!Alterado = False
          dbAlteracoes.Update
        Else
          dbAlteracoes.Edit
          dbAlteracoes!Alterado = True
          dbAlteracoes.Update
        End If
      End If
    Else
      If dbAlteracoes.AbsolutePosition = 0 Then
        dbAlteracoes.Edit
        dbAlteracoes!Alterado = False
        dbAlteracoes.Update
      Else
        dbAlteracoes.Edit
        dbAlteracoes!Alterado = True
        dbAlteracoes.Update
      End If
    End If
    dbAlteracoes.MoveNext
  Loop
End If
End Sub

Private Sub VerificaPrecos(ByVal CodigoAltera As Double)
Dim CodigoAnterior As Double

With dbProdutosAltera2
  .RecordSource = "Select *from produtosaltera order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.Find "codigoprodutoaltera=" & CodigoAltera
    If .Recordset.EOF = False Then
      .Recordset.MoveNext
      If .Recordset.EOF = False Then
        CodigoAnterior = .Recordset!codigoprodutoaltera
      Else
        CodigoAnterior = 0
      End If
      .RecordSource = "select *from qprodutosaltera where produtosaltera.codigoprodutoaltera=" & CodigoAnterior & " order by codigo"
      .Refresh
      Do While .Recordset.EOF = False
        dbProdutosAlteraDetalhe.RecordSource = "select *from produtosalteradetalhe where codigoprodutoaltera=" & CodigoAltera & strOrdem
        dbProdutosAlteraDetalhe.Refresh
        dbProdutosAlteraDetalhe.Recordset.Find "codigoproduto=" & .Recordset!CodigoProduto
        If dbProdutosAlteraDetalhe.Recordset.EOF = False Then
          dbProdutosAlteraDetalhe.Recordset!PrecoAntigo = .Recordset!PrecoVenda
          dbProdutosAlteraDetalhe.Recordset!PrecoVenda = .Recordset!PrecoVenda
          dbProdutosAlteraDetalhe.Recordset.Update
        End If
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "Select *from produtosaltera order by datacaixa desc, horaini desc"
    .Refresh
    .Recordset.Find "codigoprodutoaltera=" & CodigoAltera
    If .Recordset.EOF = False Then
      .Recordset.MovePrevious
      Do While .Recordset.BOF = False
        MsgBox "Atenção! O dia " & .Recordset!DataCaixa & " - Turno " & .Recordset!Turno & " pode ficar dezatualizado!"
        .Recordset.MovePrevious
      Loop
    End If
  End If
End With
End Sub

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = False Then cboTurno.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdAtualiza_Click()

VerificaFechamentos

With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutosAltera
  .ConnectionString = CaminhoADO
  .Refresh
End With
With qProdutosAltera
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutosAltera2
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub

Private Sub cmdImprimir_Click()
Select Case cboRelatorio.Text
  Case "Tabela de Preços"
    frmCadProdutosPrecoImprime.Show vbModal
  Case "Visualizar em Tela"
    frmRelatTabelaDePrecos.Show
End Select

End Sub

Private Sub cmdNovo_Click()
Dim CodigoAltera As Double

If dbTurnos.Recordset.EOF = True Then
  MsgBox "Erro na tabela de turnos!"
  Exit Sub
End If
If dbTurnos.Recordset!Descri <> cboTurno.Text Then
  MsgBox "Escolha um turno válido!"
  cboTurno.SetFocus
  Exit Sub
End If
With qProdutosAltera
  .RecordSource = "select *from qprodutosaltera where datacaixa>=#" & DataInglesa(txtDataCaixa.Value) & "# order by datacaixa desc, horaini desc, codigo"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If .Recordset!DataCaixa > txtDataCaixa.Value Then
      If Usuarios.Nome = "Usuário Master" Then
        Resposta = MsgBox("Já existe alteração posterior! Deseja continuar?", vbYesNo)
        If Resposta = vbNo Then
          Exit Sub
        End If
      Else
        MsgBox "Já existe alteração posterior!"
        Exit Sub
      End If
    Else
      If .Recordset!HoraIni >= dbTurnos.Recordset!HoraIni Then
        If Usuarios.Nome = "Usuário Master" Then
          Resposta = MsgBox("Já existe alteração posterior! Deseja continuar?", vbYesNo)
          If Resposta = vbNo Then
            Exit Sub
          End If
        Else
          MsgBox "Já existe alteração posterior!"
          Exit Sub
        End If
      End If
    End If
  End If
  .RecordSource = "select *from qprodutosaltera where datacaixa<#" & DataInglesa(txtDataCaixa.Value) & "# order by datacaixa desc, horaini desc, codigo"
  .Refresh
End With
With dbProdutosAltera
  .Recordset.AddNew
  .Recordset!DataCaixa = txtDataCaixa.Value
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset!Turno = dbTurnos.Recordset!Descri
  .Recordset!HoraIni = dbTurnos.Recordset!HoraIni
  .Recordset.Update
  .Refresh
  .Recordset.MoveLast
  CodigoAltera = .Recordset!codigoprodutoaltera
End With
With dbProdutos
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbProdutosAlteraDetalhe.Recordset.AddNew
      dbProdutosAlteraDetalhe.Recordset!codigoprodutoaltera = CodigoAltera
      dbProdutosAlteraDetalhe.Recordset!CodigoProduto = .Recordset!CodigoProduto
      dbProdutosAlteraDetalhe.Recordset!Codigo = .Recordset!Codigo
      dbProdutosAlteraDetalhe.Recordset!Descri = .Recordset!Descri
      dbProdutosAlteraDetalhe.Recordset!precocompra = .Recordset!precocompra
      dbProdutosAlteraDetalhe.Recordset!PrecoVenda = .Recordset!PrecoVenda
      dbProdutosAlteraDetalhe.Recordset!PrecoAntigo = .Recordset!PrecoVenda
      dbProdutosAlteraDetalhe.Recordset!Combustivel = .Recordset!Combustivel
      dbProdutosAlteraDetalhe.Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
VerificaPrecos CodigoAltera
VerificaFechamentos
dbProdutosAltera.Refresh
dbProdutosAltera.Recordset.Find "codigoprodutoaltera=" & CodigoAltera

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
Dim Resposta As Integer, CodigoAltera As Double
Dim db As New ADODB.Connection

Cancel = True
On Error GoTo 0
On Error Resume Next
If dbProdutosAltera.Recordset!Alterado = True Then
  Cancel = True
  MsgBox "Já está marcado como alterado! Não poderá ser removido!"
  Exit Sub
End If
Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then
  Cancel = True
  Exit Sub
End If
CodigoAltera = dbProdutosAltera.Recordset!codigoprodutoaltera

db.Open CaminhoADO

db.Execute "delete *from produtosalteradetalhe where codigoprodutoaltera=" & CodigoAltera

If Err.Number = 0 Then Cancel = False

End Sub

Private Sub DataGrid2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField
Else
  strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField
End If
AtualizaLista
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF5
    If Usuarios.Grupo.AdmEstatus = 2 Then
      IncluirProdutosNaLista
    End If
End Select
End Sub

Private Sub DataGrid2_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub dbProdutosAltera_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
AtualizaLista
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyF5
    If Usuarios.Grupo.AdmEstatus = 2 Then
      IncluirProdutosNaLista
    End If
End Select
End Sub

Private Sub Form_Load()
strOrdem = " order by Codigo"
txtDataCaixa.Value = Date
Call cmdAtualiza_Click
If dbProdutosAltera.Recordset.RecordCount <> 0 Then
  dbProdutosAltera.Recordset.MoveLast
End If
End Sub
