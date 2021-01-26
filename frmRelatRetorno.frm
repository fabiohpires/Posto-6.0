VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatRetorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Retornos"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9345
   Begin VB.Data qTotal 
      Caption         =   "qTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "qbicoencerrantes"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6240
      Picture         =   "frmRelatRetorno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Data QBicoEncerrantes 
      Caption         =   "QBicoEncerrantes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "qbicoencerrantes"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatRetorno.frx":0A82
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frmRelatRetorno.frx":0AA1
      TabIndex        =   7
      Top             =   840
      Width           =   9135
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Retorno:"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String, IntQuebra As Integer

Private Sub cmdExibir_Click()
With QBicoEncerrantes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from qbicoencerrantes where datacaixa between #" & DataInglesa(txtDataIni.Value) & " 00:00:00# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59# and retorno <>0 order by " & strOrdem
  .Refresh
End With
With QTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select sum(retorno) as total from qbicoencerrantes where datacaixa between #" & DataInglesa(txtDataIni.Value) & " 00:00:00# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59# and retorno <>0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "#,##0.00")
  Else
    lblTotal.Caption = Format(0, "#,##0.00")
  End If
End With
End Sub

Private Sub cmdImprime_Click()
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then GoTo NaoImprime
  On Error GoTo 0
  
  ImprimeGrid DBGrid1, Printer, QBicoEncerrantes, 5, True, , IntQuebra, , , "Retorno de Combustível", NomePosto, "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "Short date") & Chr(vbKeyReturn) & Format(Now, "long date")
  
  Printer.EndDoc
  
NaoImprime:
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DBGrid1.Columns(ColIndex).DataField & ", datacaixa, horaini" Then
  strOrdem = DBGrid1.Columns(ColIndex).DataField & " desc, datacaixa, horaini"
Else
  strOrdem = DBGrid1.Columns(ColIndex).DataField & ", datacaixa, horaini"
End If
IntQuebra = ColIndex
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
strOrdem = " datacaixa, horaini" & ", datacaixa, horaini"
IntQuebra = 0
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With QBicoEncerrantes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from qbicoencerrantes where datacaixa between #" & DataInglesa(txtDataIni.Value) & " 00:00:00# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59# and retorno <>0 order by " & strOrdem
  .Refresh
End With
With QTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select sum(retorno) as total from qbicoencerrantes where datacaixa between #" & DataInglesa(txtDataIni.Value) & " 00:00:00# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59# and retorno <>0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "#,##0.00")
  Else
    lblTotal.Caption = Format(0, "#,##0.00")
  End If
End With
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
