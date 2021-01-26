VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatFaturaCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faturamento de Cheque"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "frmRelatFaturaCheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8790
   Begin VB.CheckBox chkBomPara 
      Caption         =   "Pela data de Bom Para"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.Data qChequesPre 
      Caption         =   "qChequesPre"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatFaturaCheque.frx":0442
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data qChequesAvista 
      Caption         =   "qChequesAvista"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatFaturaCheque.frx":04D0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox chkPre 
      Caption         =   "Pré-Datado"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkAVista 
      Caption         =   "À vista"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Data qCheques 
      Caption         =   "qCheques"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatFaturaCheque.frx":055E
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmRelatFaturaCheque.frx":05EC
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "Imprimir"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdExibe 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *from qchequescx"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatFaturaCheque.frx":106E
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frmRelatFaturaCheque.frx":1086
      TabIndex        =   6
      Top             =   840
      Width           =   8535
   End
   Begin MSComCtl2.DTPicker txtDataini 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37683
   End
   Begin MSComCtl2.DTPicker txtDatafim 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37683
   End
   Begin VB.Label lblJuros 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Juros:"
      Height          =   195
      Left            =   6720
      TabIndex        =   19
      Top             =   5640
      Width           =   420
   End
   Begin VB.Label lblCheques 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de Cheques:"
      Height          =   195
      Left            =   1365
      TabIndex        =   17
      Top             =   5640
      Width           =   1770
   End
   Begin VB.Label lblPre 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Pré-Datados"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblAvista 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "À Vista"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   5640
      Width           =   405
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   240
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatFaturaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub cmdExibe_Click()
Dim StrTemp As String
If chkAVista.Value = vbUnchecked And chkPre.Value = vbUnchecked Then
  MsgBox "É preciso selecionar À vista, Pré-datado ou ambos!"
  Exit Sub
End If

If chkAVista.Value = vbChecked And chkPre.Value = vbUnchecked Then
  StrTemp = " and datacaixa=datacheque"
Else
  If chkAVista.Value = vbUnchecked And chkPre.Value = vbChecked Then
    StrTemp = " and datacaixa<>datacheque"
  End If
End If
With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  If chkBomPara.Value = vbChecked Then
    .RecordSource = "select *from qchequescx where datacheque between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp & strOrdem
  Else
    .RecordSource = "select *from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp & strOrdem
  End If
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
  lblCheques.Caption = Format(.Recordset.RecordCount, "#,##0")
End With

With qCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total, sum(cheques.juros) as totaljuros from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Totaljuros) = False Then
    lblJuros.Caption = Format(.Recordset!Totaljuros, "Currency")
  Else
    lblJuros.Caption = Format(0, "Currency")
  End If
End With
With qChequesAvista
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and datacaixa=datacheque"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblAvista.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblAvista.Caption = Format(0, "Currency")
  End If
End With
With qChequesPre
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and datacaixa<>datacheque"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblPre.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblPre.Caption = Format(0, "Currency")
  End If
End With

End Sub

Private Sub cmdImprime_Click()
Dim StrTitulo2 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTitulo2 = Chr(vbKeyReturn) & "Período: " & txtDataIni.Value & " a " & txtDataFim.Value
If chkAVista.Value = vbChecked And chkPre.Value = vbUnchecked Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Cheques À Vista"
End If
If chkAVista.Value = vbUnchecked And chkPre.Value = vbChecked Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Cheques Pré-Datados"
End If
If chkAVista.Value = vbChecked And chkPre.Value = vbChecked Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Cheques À Vista e Pré-Datados"
End If
If chkBomPara.Value = vbChecked Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Pela data de Bom Para"
End If

ImprimeGrid DBGrid1, Printer, dbCheques, 7, , , , 5, 8, "Faturamento de Cheques", StrTitulo2, Chr(vbKeyReturn) & Format(Date, "Long Date") & " - " & Format(Time, "short time")

StrTemp = "Cheques À Vista: " & lblAvista.Caption & "     Pré-Datados: " & lblPre.Caption
Printer.Print ""
Printer.Print StrTemp

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
Call cmdExibe_Click
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

strOrdem = " order by DataCaixa"
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date

With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "SELECT *from qchequescx WHERE datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & strOrdem
  .Refresh
End With
With qCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  .Refresh
End With
With qChequesAvista
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  .Refresh
End With
With qChequesPre
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  .Refresh
End With
Call cmdExibe_Click
End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataFim_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
