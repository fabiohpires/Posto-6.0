VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFechamentoDeCaixaComissoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comissoes no Caixa"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   Icon            =   "frmFechamentoDeCaixaComissoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton optProdutos 
         Caption         =   "Produtos"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optCombustiveis 
         Caption         =   "Combustíveis"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc dbComissoes 
      Height          =   375
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      RecordSource    =   "select *from venda2 where codigofechamento=0"
      Caption         =   "dbComissoes"
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFechamentoDeCaixaComissoes.frx":0442
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Bico"
         Caption         =   "Bico"
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
         DataField       =   "CodProduto"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Quantidade"
         Caption         =   "Qtd."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ValorUnitario"
         Caption         =   "Unitário"
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
      BeginProperty Column05 
         DataField       =   "ValorTotal"
         Caption         =   "Total"
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
      BeginProperty Column06 
         DataField       =   "CodigoVendedor"
         Caption         =   "Func."
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
      BeginProperty Column07 
         DataField       =   "ValorComissao"
         Caption         =   "Comissão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Pago"
         Caption         =   "Pago"
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
            ColumnWidth     =   524,976
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   540,284
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmFechamentoDeCaixaComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoFechamento As Double
Dim strFiltro As String, strOrdem As String

Public Sub Filtra()
Dim Total As Currency

If optTodos.Value = True Then
  strFiltro = ""
ElseIf optCombustiveis.Value = True Then
  strFiltro = " and combustivel=-1"
Else
  strFiltro = " and combustivel=0"
End If

Total = 0
With dbComissoes
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from venda2 where codigofechamento=" & CodigoFechamento & strFiltro & strOrdem
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Total = Total + .Recordset!ValorComissao
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

lblTotal.Caption = "R$ " & Format(Total, "0.0000")

End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " Order by " & DataGrid1.Columns(ColIndex).DataField Then
  strOrdem = " Order by " & DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " Order by " & DataGrid1.Columns(ColIndex).DataField
End If
Filtra
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
strOrdem = " Order by Bico"

strFiltro = ""

Filtra

End Sub

Private Sub optCombustiveis_Click()
Filtra
End Sub

Private Sub optProdutos_Click()
Filtra
End Sub

Private Sub optTodos_Click()
Filtra
End Sub
