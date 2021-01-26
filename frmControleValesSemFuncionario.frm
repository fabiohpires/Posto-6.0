VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmControleValesSemFuncionario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comissões não identificadas"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmControleValesSemFuncionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAltera 
      Caption         =   "Altera"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin MSDataListLib.DataCombo cboVendedor 
      Bindings        =   "frmControleValesSemFuncionario.frx":0442
      Height          =   315
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin MSAdodcLib.Adodc dbVendedores 
      Height          =   375
      Left            =   2400
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select *from vendedores"
      Caption         =   "dbVendedores"
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
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optCombustiveis 
         Caption         =   "Combustíveis"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optProdutos 
         Caption         =   "Produtos"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc dbComissoes 
      Height          =   375
      Left            =   120
      Top             =   5160
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmControleValesSemFuncionario.frx":045D
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
            Locked          =   -1  'True
            ColumnWidth     =   524,976
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   540,284
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmControleValesSemFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoFechamento As Double
Dim strFiltro As String, strOrdem As String
Dim Alterado As Boolean

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
  .RecordSource = "select *from venda2 where pago=0 and valorcomissao<>0 and fechamentodiario=-1" & strFiltro & strOrdem
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

With dbVendedores
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub

Private Sub cboVendedor_LostFocus()
With dbVendedores
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboVendedor.Text = "" Then
    Exit Sub
  End If
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & cboVendedor.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!Codigo
  End If
End With

End Sub

Private Sub cmdAltera_Click()
Call cboVendedor_LostFocus
With dbVendedores
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboVendedor.Text <> .Recordset!Nome Then
    MsgBox "Funcionário não localizado!"
    Exit Sub
  End If
  If dbComissoes.Recordset.RecordCount = 0 Then Exit Sub
  If dbComissoes.Recordset.EOF = True Or dbComissoes.Recordset.BOF = True Then Exit Sub
  dbComissoes.Recordset!codigovendedor = .Recordset!codigovendedor
  dbComissoes.Recordset!CodigoPagamento = .Recordset!Codigo
  dbComissoes.Recordset.Update
  
  Alterado = True
End With

Filtra

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

Private Sub txtCodigo_LostFocus()
With dbVendedores
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodigo.Text = "" Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then
    MsgBox "Código Inválido!"
    Exit Sub
  End If
  .Recordset.MoveFirst
  .Recordset.Find "codigo=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    cboVendedor.Text = .Recordset!Nome
  End If
End With
End Sub
