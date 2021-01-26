VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatPosicaoDoEstoque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posição do Estoque por Caixa"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   7800
      Picture         =   "frmRelatPosicaoDoEstoque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbEstoque 
      Height          =   375
      Left            =   4440
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      RecordSource    =   "select codigo, descri, estoquefisico from produtos where combustivel=0 order by codigo"
      Caption         =   "dbEstoque"
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
      Bindings        =   "frmRelatPosicaoDoEstoque.frx":0A82
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
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
         DataField       =   "codigo"
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
         DataField       =   "descri"
         Caption         =   "Produto"
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
         DataField       =   "estoquefisico"
         Caption         =   "Estoque"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5325,166
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dbTurnos 
      Height          =   375
      Left            =   4440
      Top             =   240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
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
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmRelatPosicaoDoEstoque.frx":0A9A
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   126353409
      CurrentDate     =   37600
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      Caption         =   "&Turno:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmRelatPosicaoDoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String

Private Sub Ativado(ByVal Ativa As Boolean)
txtData.Enabled = Ativa
cboTurno.Enabled = Ativa
cmdExibir.Enabled = Ativa
End Sub
Private Sub Exibir()
Dim StrTemp As String, Sequencia As Double
Dim db As New ADODB.Connection
Dim dbFechamento As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbVendas As New ADODB.Recordset
Dim dbEntradas As New ADODB.Recordset
Dim dbEntradas2 As New ADODB.Recordset
Dim SequenciaFinalizado As Double

Call cboTurno_LostFocus

db.Open CaminhoADO
dbFechamento.CursorLocation = adUseClient
dbFechamento.Open "Select fechado, datacaixa, horaini, codigoturno, sequencia from fechamentodecaixa where datacaixa<=#" & DataInglesa(txtData.Value) & "# order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic
dbEntradas.CursorLocation = adUseClient
dbEntradas.Open "select datanota, codigoproduto, quantidade from qprodutosnotas where datanota>#" & DataInglesa(txtData.Value) & "# and tanque=0 and confirmado=-1 order by codigoproduto", db, adOpenKeyset, adLockOptimistic
dbEntradas2.CursorLocation = adUseClient
dbEntradas2.Open "select datanota, codigoproduto, quantidade from qprodutosnotas where dataentrega<=#" & DataInglesa(txtData.Value) & "# and tanque=0 and confirmado=0 order by codigoproduto", db, adOpenKeyset, adLockOptimistic

If dbFechamento.RecordCount <> 0 Then
  dbFechamento.Find "fechado=-1"
  If dbFechamento.EOF = False Then
    SequenciaFinalizado = dbFechamento!Sequencia
  Else
    SequenciaFinalizado = 1
  End If
  dbFechamento.MoveFirst
  If dbFechamento!DataCaixa >= txtData.Value Then
    If dbFechamento!DataCaixa > txtData.Value Then
      dbFechamento.Find "datacaixa=#" & DataInglesa(TempData) & "#"
    End If
    If dbFechamento!HoraIni <= dbTurnos.Recordset!HoraIni Then
      Sequencia = dbFechamento!Sequencia
    Else
      dbFechamento.Find "horaini<=#" & dbTurnos.Recordset!HoraIni & "#"
      If dbFechamento!DataCaixa < txtData.Value Then
        TempData = dbFechamento!DataCaixa
        dbFechamento.MoveFirst
        dbFechamento.Find "datacaixa=#" & DataInglesa(TempData) & "#"
        Sequencia = dbFechamento!Sequencia
      Else
        Sequencia = dbFechamento!Sequencia
      End If
    End If
  Else
    Sequencia = dbFechamento!Sequencia
  End If
Else
  Sequencia = 0
End If


dbProdutos.Open "Select codigoproduto, estoque, estoquefisico from produtos order by codigoproduto", db, adOpenKeyset, adLockOptimistic
If dbProdutos.RecordCount <> 0 Then
  dbProdutos.MoveFirst
  Do While dbProdutos.EOF = False
    dbProdutos!estoquefisico = dbProdutos!Estoque
    dbProdutos.Update
    dbProdutos.MoveNext
  Loop
End If

StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where fechado=0 and sequencia>" & SequenciaFinalizado & " and sequencia<=" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbVendas!CodigoProduto
    If dbProdutos.EOF = False Then
      dbProdutos!estoquefisico = dbProdutos!estoquefisico - dbVendas!estoquedia
      dbProdutos.Update
    End If
    dbVendas.MoveNext
  Loop
End If

dbVendas.Close
StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where fechado=-1 and sequencia>" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbVendas!CodigoProduto
    If dbProdutos.EOF = False Then
      dbProdutos!estoquefisico = dbProdutos!estoquefisico + dbVendas!estoquedia
      dbProdutos.Update
    End If
    dbVendas.MoveNext
  Loop
End If


If dbEntradas.RecordCount <> 0 Then
  Do While dbEntradas.EOF = False
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbEntradas!CodigoProduto
    If dbProdutos.EOF = False Then
      dbProdutos!estoquefisico = dbProdutos!estoquefisico - dbEntradas!Quantidade
      dbProdutos.Update
    End If
    dbEntradas.MoveNext
  Loop
End If
If dbEntradas2.RecordCount <> 0 Then
  Do While dbEntradas2.EOF = False
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbEntradas2!CodigoProduto
    If dbProdutos.EOF = False Then
      dbProdutos!estoquefisico = dbProdutos!estoquefisico + dbEntradas2!Quantidade
      dbProdutos.Update
    End If
    dbEntradas2.MoveNext
  Loop
End If

With dbEstoque
  .RecordSource = "select codigo, descri, estoquefisico from produtos where combustivel=0" & StrOrdem
  .Refresh
End With
DataGrid1.Refresh
dbFechamento.Close
dbProdutos.Close
dbVendas.Close
db.Close
End Sub

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.Find "descri='" & cboTurno.Text & "'"
    If .Recordset.EOF = False Then
      cboTurno.Text = .Recordset!Descri
    End If
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Ativado False
Exibir
Ativado True
dbEstoque.Refresh

End Sub

Private Sub cmdImprime_Click()
Dim Titulo1 As String, Titulo2 As String, Titulo3 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Titulo1 = NomePosto & " - Posição de estoque por turno"
Titulo2 = "Estoque até o caixa: " & txtData.Value & " - Turno: " & cboTurno.Text
Titulo3 = Titulo3 & Chr(vbKeyReturn) & "Impresso em: " & Format(Date, "long Date")
Coluna = 1000
With DataGrid1
  .Columns.Add 3
  .Columns.Add 4
  .Columns.Add 5
  .Columns.Add 6
  .Columns(3).Width = Coluna
  .Columns(4).Width = Coluna
  .Columns(5).Width = Coluna
  .Columns(6).Width = Coluna
End With

ImprimeADOGrid DataGrid1, Printer, dbEstoque, , , , , , , Titulo1, Titulo2, Titulo3

DataGrid1.Columns.Remove 6
DataGrid1.Columns.Remove 5
DataGrid1.Columns.Remove 4
DataGrid1.Columns.Remove 3


Printer.EndDoc
NaoImprime:

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField Then
  StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField
End If
Call cmdExibir_Click
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
Dim db As New ADODB.Connection
Dim dbFechamentos As New ADODB.Recordset

StrOrdem = " order by codigo"

With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
With dbEstoque
  .ConnectionString = CaminhoADO
  .Refresh
End With

db.Open CaminhoADO
dbFechamentos.CursorLocation = adUseClient
dbFechamentos.Open "Select datacaixa, turno from fechamentodecaixa order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic
If dbFechamentos.RecordCount <> 0 Then
  txtData.Value = dbFechamentos!DataCaixa
  cboTurno.Text = dbFechamentos!Turno
Else
  txtData.Value = Date
  If dbTurnos.Recordset.RecordCount <> 0 Then
    cboTurno.Text = dbTurnos.Recordset!Descri
  End If
End If

dbFechamentos.Close
db.Close
Exibir
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub
