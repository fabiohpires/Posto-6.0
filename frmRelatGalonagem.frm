VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatGalonagem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Galonagem"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmRelatGalonagem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbTurnos 
      Height          =   330
      Left            =   600
      Top             =   3240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from turnos order by horaini"
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
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6120
      Picture         =   "frmRelatGalonagem.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "Imprimir"
      Top             =   3360
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gasolina / Álcool / Diesel"
      Height          =   1095
      Left            =   4440
      TabIndex        =   22
      Top             =   4200
      Width           =   2415
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Média:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblMediaGeral 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc qTemp 
      Height          =   330
      Left            =   600
      Top             =   2880
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from qgalonagemfechamento where qgalonagem.codigoproduto=0 order by data, fechamentodiario.codigofechamento"
      Caption         =   "qTemp"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60227585
      CurrentDate     =   37665
   End
   Begin MSAdodcLib.Adodc qGalonagem 
      Height          =   330
      Left            =   600
      Top             =   2520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from qgalonagemfechamento where codigoproduto=0 order by datacaixa, horaini"
      Caption         =   "qGalonagem"
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
   Begin MSDataListLib.DataCombo cboProduto 
      Bindings        =   "frmRelatGalonagem.frx":0EC4
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   600
      Top             =   2160
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatGalonagem.frx":0EDD
      Height          =   4095
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      _Version        =   393216
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
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd MMMM yyyy dddd"
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
         DataField       =   "Vendido"
         Caption         =   "Vendido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60227585
      CurrentDate     =   37665
   End
   Begin MSDataListLib.DataCombo cboTurnos 
      Bindings        =   "frmRelatGalonagem.frx":0EF6
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   "Descri"
      Text            =   ""
   End
   Begin VB.Label Label7 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblEstoque 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Estoque Atual:"
      Height          =   195
      Left            =   4200
      TabIndex        =   20
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label lblDias 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Previsão de Períodos:"
      Height          =   195
      Left            =   4320
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblMedia 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Média:"
      Height          =   195
      Left            =   4680
      TabIndex        =   16
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Vendido:"
      Height          =   195
      Left            =   4200
      TabIndex        =   14
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmRelatGalonagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubCabeca(ByVal TempValor As Double)
X1 = 0 + TempValor
Y1 = Printer.CurrentY
X2 = 20 + TempValor
Y2 = Printer.CurrentY + 5

Printer.FillStyle = vbFSSolid
Printer.FillColor = RGB(175, 175, 175)
Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B

Printer.FillStyle = vbFSTransparent
StrTemp = "Data"
Printer.CurrentX = X1 + 1
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(StrTemp) / 2)
Printer.Print StrTemp;

X1 = 20 + TempValor
X2 = 40 + TempValor

Printer.FillStyle = vbFSSolid
Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B

Printer.FillStyle = vbFSTransparent
StrTemp = "Turno"
Printer.CurrentX = X1 + 1
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(StrTemp) / 2)
Printer.Print StrTemp;

X1 = 40 + TempValor
X2 = 60 + TempValor

Printer.FillStyle = vbFSSolid
Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B

Printer.FillStyle = vbFSTransparent
StrTemp = "Venda"
Printer.CurrentX = X2 - 1 - (Printer.TextWidth(StrTemp))
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(StrTemp) / 2)
Printer.Print StrTemp

End Sub

Private Sub Corpo(ByVal Coluna As Double, ByVal Dia As String, ByVal Turno As String, ByVal Venda As String)
TempValor = Coluna
X1 = 0 + TempValor
Y1 = Printer.CurrentY
X2 = 20 + TempValor
Y2 = Printer.CurrentY + 5

Printer.FillStyle = vbTransparent
Printer.FillColor = RGB(255, 255, 255)
Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B



Printer.CurrentX = X1 + 1
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(Dia) / 2)
Printer.Print Dia;

X1 = 20 + TempValor
X2 = 40 + TempValor

Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B

Printer.CurrentX = X1 + 1
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(Turno) / 2)
Printer.Print Turno;

X1 = 40 + TempValor
X2 = 60 + TempValor

Printer.CurrentX = 0 + TempValor
Printer.Line (X1, Y1)-(X2, Y2), , B


Printer.CurrentX = X2 - 1 - (Printer.TextWidth(Venda))
Printer.CurrentY = (Y1 + ((Y2 - Y1) / 2)) - (Printer.TextHeight(Venda) / 2)
Printer.Print Venda

End Sub

Private Sub Cabeca(ByVal Largura As Double, ByVal Dia As Date)
Dim StrTemp As String, TempValor As Double

Printer.ScaleMode = vbMillimeters
Printer.FontSize = 14
Printer.FontBold = True
Printer.FontName = "Arial"

Printer.FillColor = RGB(175, 175, 175)
Printer.FillStyle = vbFSSolid
Printer.Line (0, 0)-(Largura, 10), , B

Printer.FillStyle = vbFSTransparent

StrTemp = "Relatório de Galonagem"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.CurrentY = 5 - (Printer.TextHeight(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
Printer.FontBold = False

StrTemp = "Data: " & Dia
Printer.CurrentY = 13
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Produto: " & dbProdutos.Recordset!Descri
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short Date")
Printer.CurrentX = 90
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1.5

End Sub

Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  cboProduto.Text = .Recordset!Descri
  txtCodigo.Text = .Recordset!Codigo
End With
End Sub

Private Sub cboTurnos_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboTurnos.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  cboTurnos.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdExibir_Click()
Dim db As New ADODB.Connection
Dim dbPdvs As New ADODB.Recordset

Dim TempValor As Double
Dim PDV As Double

Screen.MousePointer = vbHourglass
'On Error Resume Next
If dbProdutos.Recordset.EOF = True Then
  MsgBox "Produto inválido!"
  cboProduto.SetFocus
  Exit Sub
End If
If cboProduto.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Produto inválido!"
  cboProduto.SetFocus
  Exit Sub
End If
If txtDataIni.Value > txtDataFim.Value Then
  MsgBox "A data inicial deve ser menor que a data final!"
  txtDataIni.SetFocus
  Exit Sub
End If

dbPdvs.CursorLocation = adUseClient
db.Open CaminhoADO
dbPdvs.Open "select *from pdvs where intermitente=0", db, adOpenDynamic, adLockOptimistic

If dbPdvs.RecordCount <> 0 Then
  PDV = dbPdvs!CodigoPdv
End If

With qGalonagem
  .ConnectionString = CaminhoADO
  If cboTurnos.Text = "" Then
    .RecordSource = "select *from qgalonagemfechamento where codigopdv=" & PDV & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# order by datacaixa, horaini"
  Else
    .RecordSource = "select *from qgalonagemfechamento where codigopdv=" & PDV & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " and codigoturno=" & dbTurnos.Recordset!CodigoTurno & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# order by datacaixa, horaini"
  End If
  .Refresh
End With
With qTemp
  If cboTurnos.Text = "" Then
    .RecordSource = "select sum(vendido) as Total from qgalonagemfechamento where codigopdv=" & PDV & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  Else
    .RecordSource = "select sum(vendido) as Total from qgalonagemfechamento where codigopdv=" & PDV & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " and codigoturno=" & dbTurnos.Recordset!CodigoTurno & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  End If
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblTotal.Caption = Format(TempValor, "#,##0")
  
  If qGalonagem.Recordset.RecordCount <> 0 Then
    TempValor = TempValor / qGalonagem.Recordset.RecordCount
  Else
    TempValor = 1
  End If
  lblMedia.Caption = Format(TempValor, "#,##0")
  
  lblEstoque.Caption = Format(dbProdutos.Recordset!Estoque, "#,##0")
  
  TempValor = CInt(dbProdutos.Recordset!Estoque / TempValor)
  lblDias.Caption = Format(TempValor, "#,##0")
  
  If dbProdutos.Recordset!Combustivel = True Then
    If cboTurnos.Text = "" Then
      .RecordSource = "select sum(vendido) as Total from qgalonagemfechamento where codigopdv=" & PDV & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
    Else
      .RecordSource = "select sum(vendido) as Total from qgalonagemfechamento where codigopdv=" & PDV & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# and codigoturno=" & dbTurnos.Recordset!CodigoTurno
    End If
    .Refresh
    If IsNull(.Recordset!Total) = False Then
      TempValor = .Recordset!Total
    Else
      TempValor = 0
    End If
    lblTotalGeral.Caption = Format(TempValor, "#,##0")
    
    If qGalonagem.Recordset.RecordCount <> 0 Then
      TempValor = TempValor / qGalonagem.Recordset.RecordCount
    Else
      TempValor = 0
    End If
    lblMediaGeral.Caption = Format(TempValor, "#,##0")
  Else
    lblTotalGeral.Caption = ""
    lblMediaGeral.Caption = ""
  End If
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Coluna As Double

If qGalonagem.Recordset.RecordCount = 0 Then Exit Sub
If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
If dbProdutos.Recordset.EOF = True Then Exit Sub
If cboProduto.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Indique um Produto!"
  Exit Sub
End If
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

With qGalonagem
  Largura = 190
  Coluna = 0
  Dia = Now
  Printer.FontName = "Arial"
  .Recordset.MoveFirst
  Cabeca Largura, Dia
  
  Inicio = Printer.CurrentY
  SubCabeca Coluna
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      If Coluna >= 130 Then
        Printer.CurrentY = Printer.CurrentY + 2
        Printer.Print "Página: " & Printer.Page
        Coluna = 0
        Printer.NewPage
        Cabeca Largura, Dia
      Else
        Coluna = Coluna + 65
      End If
      Printer.CurrentY = Inicio
      SubCabeca Coluna
    End If
    Corpo Coluna, Format(.Recordset!DataCaixa, "Short date"), .Recordset!Turno, Format(.Recordset!Vendido, "#,##0")
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.ScaleHeight - 20
  Printer.CurrentX = 0
  Printer.FontBold = True
  StrTemp = "Total: " & lblTotal.Caption
  Printer.Print StrTemp;
  
  Printer.CurrentX = 40
  Printer.FontBold = True
  StrTemp = "Média: " & lblMedia.Caption
  Printer.Print StrTemp;
  
  Printer.CurrentX = 80
  Printer.FontBold = True
  StrTemp = "Estoque: " & lblEstoque.Caption
  Printer.Print StrTemp;
  
  Printer.CurrentX = 120
  Printer.FontBold = True
  StrTemp = "Previsão de Turnos: " & lblDias.Caption
  Printer.Print StrTemp
  
  Printer.FontBold = False
  StrTemp = "Página: " & Printer.Page
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataIni.Value = Date
txtDataFim.Value = Date
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos where combustivel=-1 order by descri"
  .Refresh
End With
With qGalonagem
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from qgalonagemfechamento where codigoproduto=0 order by datacaixa, horaini"
  .Refresh
End With
With qTemp
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(vendido) as Total from qgalonagemfechamento where codigoproduto=" & dbProdutos.Recordset!CodigoProduto & " and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from turnos order by horaini"
  .Refresh
End With


End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbProdutos
  If txtCodigo.Text = "" Then Exit Sub
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.Find "codigo=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodigo.Text = .Recordset!Codigo
  End If
End With
End Sub
