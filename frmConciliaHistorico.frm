VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaHistorico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico da Conta"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmConciliaHistorico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   120
      Picture         =   "frmConciliaHistorico.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Imprimir"
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select concilianova.*, Contas.* from concilianova, contas where concilianova.codigoconta=contas.codigoconta"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmConciliaHistorico.frx":0EC4
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   37257
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   37257
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaHistorico.frx":0EDB
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmConciliaHistorico.frx":0EF4
      TabIndex        =   7
      Top             =   840
      Width           =   9735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Conta:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmConciliaHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboConta_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConta_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub CboConta_LostFocus()
Me.KeyPreview = True
With dbContas
  .Refresh
  If cboConta.Text = "" Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.FindFirst "descri='" & cboConta.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboConta.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdExibir_Click()
Dim strTemp As String
strTemp = "select concilianova.*, Contas.* from concilianova, contas where concilianova.codigoconta=contas.codigoconta and datalanc between #" & DataInglesa(txtDataIni.Value) & " 00:00:00# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59#"
If cboConta.Text <> "" Then
  If dbContas.Recordset.EOF = False Then
    If cboConta.Text = dbContas.Recordset!Descri Then
      strTemp = strTemp & " and contas.codigoconta=" & dbContas.Recordset!codigoconta
    End If
  End If
End If
With dbConcilia
  .RecordSource = strTemp & " order by codigoconciliaconta"
  .Refresh
End With
End Sub

Private Sub cmdImprime_Click()
Dim strTemp As String
On Error GoTo Pendencias
If ShowPrinter(Me) = 0 Then GoTo Pendencias
On Error GoTo 0
strTemp = NomePosto

strTemp = strTemp & Chr(vbKeyReturn) & "Impresso em: " & Format(Now, "long date") & " - " & Format(Now, "short time")
strTemp = strTemp & Chr(vbKeyReturn) & "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "short date")

If cboConta.Text <> "" Then
  strTemp = strTemp & Chr(vbKeyReturn) & "Conta: " & cboConta.Text
End If

ImprimeGrid DBGrid1, Printer, dbConcilia, , False, , , , , "Histórico de Contas", , strTemp

Printer.EndDoc

Pendencias:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataIni.Value = Date
txtDataFim.Value = Date

With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
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
