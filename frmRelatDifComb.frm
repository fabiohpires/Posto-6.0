VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatDifComb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Diferença de Combustível"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.Data qDifComb 
      Caption         =   "qDifComb"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from qDifCombustivel order by datacaixa, HoraIni, tanquenr"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtTanque 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8400
      Picture         =   "frmRelatDifComb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37620
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37620
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatDifComb.frx":0A82
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "frmRelatDifComb.frx":0A99
      TabIndex        =   9
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "Tanque:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatDifComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Diferença de Combustível"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = "Impresso em: " & Format(Dia, "Short date") & " - " & Format(Dia, "Short Time")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "Short date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp


Printer.Print ""

StrTemp = "Data Caixa"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 25
Printer.Print StrTemp;

StrTemp = "Combustível"
Printer.CurrentX = 50
Printer.Print StrTemp;

StrTemp = "Tq."
Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Estoque"
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Posto"
Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Vendido"
Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Diferença"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5


End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String
'If chkAgrupar.Value = vbUnchecked Then
  With qDifComb
    StrTemp = "select *from qDifCombustivel where datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
    If IsNumeric(txtTanque.Text) = True Then
      StrTemp = StrTemp & " and tanquenr=" & txtTanque.Text
    End If
    StrTemp = StrTemp & " order by datacaixa, HoraIni, tanquenr"
    .RecordSource = StrTemp
    .Refresh
  End With
'Else
  
'End If
End Sub

Private Sub cmdImprime_Click()
Dim Dia As Date, StrTemp As String, Largura As Double

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0


With qDifComb
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  Dia = Now
  Largura = 190
  Cabeca Dia, Largura
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      Printer.NewPage
      Cabeca Dia, Largura
    End If
    StrTemp = .Recordset!DataCaixa
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 25
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 50
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!tanquenr
    Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Estoque
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Tanque
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Vendido
    Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset("diferencacombustivel.diferenca")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
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
With qDifComb
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
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
