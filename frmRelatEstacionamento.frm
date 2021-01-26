VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatEstacionamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Estacionamento"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   780
   ClientWidth     =   10200
   Icon            =   "frmRelatEstacionamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   9120
      Picture         =   "frmRelatEstacionamento.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "Imprimir"
      Top             =   240
      Width           =   735
   End
   Begin VB.Data qEstacionamentoTotal 
      Caption         =   "qEstacionamentoTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatEstacionamento.frx":0EC4
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data qEstacionamento 
      Caption         =   "qEstacionamento"
      Connect         =   "Access 2000;"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from qEstacionamento where fechamentodecaixa.codigofechamento=0 order by datacaixa, horaini"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbFuncionarios 
      Caption         =   "dbFuncionarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from vendedores order by nome"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbTurnos 
      Caption         =   "dbTurnos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Turnos order by horaini"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatEstacionamento.frx":0F59
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "frmRelatEstacionamento.frx":0F77
      TabIndex        =   9
      Top             =   840
      Width           =   9855
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo cboTurno 
      Bindings        =   "frmRelatEstacionamento.frx":2516
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataini 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37683
   End
   Begin MSComCtl2.DTPicker txtDatafim 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37683
   End
   Begin MSDBCtls.DBCombo cboResponsavel 
      Bindings        =   "frmRelatEstacionamento.frx":252D
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lblTotalCanc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Canc.:"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblTotalUn 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Un.:"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Responsável:"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatEstacionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub cboResponsavel_LostFocus()
With dbFuncionarios
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboResponsavel.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboResponsavel.Text & "'"
  If .Recordset.NoMatch = False Then
    cboResponsavel.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboTurno.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboTurno.Text & "'"
  If .Recordset.NoMatch = False Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String
StrTemp = " where datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
If cboTurno.Text <> "" Then
  If dbTurnos.Recordset.EOF = False Then
    If cboTurno.Text = dbTurnos.Recordset!Descri Then
      StrTemp = StrTemp & " and codigoturno=" & dbTurnos.Recordset!CodigoTurno
    End If
  End If
End If
If cboResponsavel.Text <> "" Then
  If dbFuncionarios.Recordset.EOF = False Then
    If cboResponsavel.Text = dbFuncionarios.Recordset!Nome Then
      StrTemp = StrTemp & " and codigoresponsavel=" & dbFuncionarios.Recordset!codigovendedor
    End If
  End If
End If
With qEstacionamento
  .RecordSource = "select *from qEstacionamento" & StrTemp & strOrdem
  .Refresh
End With
With qEstacionamentoTotal
  .RecordSource = "select sum(TotalUn) as unidades, sum(cancelados) as cancela, sum(total) as totais from qEstacionamento" & StrTemp
  .Refresh
  If IsNull(.Recordset!Unidades) = False Then
    lblTotalUn.Caption = Format(.Recordset!Unidades, "#,##0")
  Else
    lblTotalUn.Caption = "0"
  End If
  If IsNull(.Recordset!cancela) = False Then
    lblTotalCanc.Caption = Format(.Recordset!cancela, "#,##0")
  Else
    lblTotalCanc.Caption = "0"
  End If
  If IsNull(.Recordset!Totais) = False Then
    lblTotal.Caption = Format(.Recordset!Totais, "Currency")
  Else
    lblTotal.Caption = "0"
  End If
  
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTitulo2 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTitulo2 = Chr(vbKeyReturn) & "Período: " & txtDataIni.Value & " a " & txtDataFim.Value
If cboTurno.Text <> "" Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Turno: " & cboTurno.Text
End If
If cboResponsavel.Text <> "" Then
  StrTitulo2 = StrTitulo2 & Chr(vbKeyReturn) & "Responsável: " & cboResponsavel.Text
End If

ImprimeGrid DBGrid1, Printer, qEstacionamento, 5, , , , 7, 8, "Relatório de Estacionamento" & Chr(vbKeyReturn) & NomePosto, StrTitulo2, Chr(vbKeyReturn) & Format(Date, "Long Date") & " - " & Format(Time, "short time")

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", HoraIni" Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc, HoraIni desc"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", HoraIni"
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
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
strOrdem = " order by DataCaixa, HoraIni"
With dbTurnos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbFuncionarios
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qEstacionamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qEstacionamentoTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Call cmdExibir_Click
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
