VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLavagem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Lavagem"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8760
      Picture         =   "frmLavagem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8760
      TabIndex        =   24
      Top             =   6480
      Width           =   855
   End
   Begin VB.Data dbTotal 
      Caption         =   "dbTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(Totalbruto) as bruto, sum(totalliquido) as liquido from lavagem"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.Data dbTurno 
      Caption         =   "dbTurno"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Turnos order by HoraIni"
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data dbLavagem 
      Caption         =   "dbLavagem"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Lavagem order by dia, horaturno"
      Top             =   3120
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtPosto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5520
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtLavadores 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin MSDBCtls.DBCombo cboTurno 
      Bindings        =   "frmLavagem.frx":0A82
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
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
      Format          =   72351745
      CurrentDate     =   38026
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmLavagem.frx":0A98
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmLavagem.frx":0AB0
      TabIndex        =   0
      Top             =   1320
      Width           =   9495
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   5640
      TabIndex        =   15
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   38026
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   7320
      TabIndex        =   17
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   38026
   End
   Begin VB.Label lblLiquido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Valor Posto:"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lblBruto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Total:"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "a"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "Período"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Observações:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Valor do Posto:"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total:"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Lavadores:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Dia da Lavagem:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLavagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

Printer.FontName = "Arial"
Printer.FontSize = 16
Printer.ScaleMode = vbMillimeters

StrTemp = "Relatório de Lavagem"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short date")
Printer.Print StrTemp

StrTemp = "Dia"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 35
Printer.Print StrTemp;

StrTemp = "Lav."
Printer.CurrentX = 55
Printer.Print StrTemp;

StrTemp = "V. Total"
Printer.CurrentX = 85 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "V. Posto"
Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Observações"
Printer.CurrentX = 106
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

End Sub

Private Sub cboTurno_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboTurno_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboTurno_LostFocus()
Me.KeyPreview = True
With dbTurno
  .Refresh
  If cboTurno.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboTurno.Text & "'"
  If .Recordset.NoMatch = False Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
With dbLavagem
  .RecordSource = "Select *from lavagem where dia between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# order by dia, horaturno"
  .Refresh
End With

With dbTotal
  .RecordSource = "select sum(Totalbruto) as bruto, sum(totalliquido) as liquido from lavagem where dia between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh

  lblBruto.Caption = Format("0", "Currency")
  If IsNull(.Recordset!Bruto) = False Then
    lblBruto.Caption = Format(.Recordset!Bruto, "Currency")
  End If
  lblLiquido.Caption = Format("0", "Currency")
  If IsNull(.Recordset!Liquido) = False Then
    lblLiquido.Caption = Format(.Recordset!Liquido, "Currency")
  End If
End With

End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date

With dbLavagem
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Largura = 190
  Printer.FontSize = 10
  Printer.FontName = "Arial"
  Printer.ScaleMode = vbMillimeters
  Dia = Now
  
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    
    StrTemp = .Recordset!Dia
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 35
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!lavadores
    Printer.CurrentX = 55
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!TotalBruto, "Currency")
    Printer.CurrentX = 85 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Totalliquido, "Currency")
    Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Obs
    Printer.CurrentX = 106
    Printer.Print StrTemp

    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  StrTemp = "Total: " & lblBruto.Caption & "   Valor Posto: " & lblLiquido.Caption
  Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub cmdIncluir_Click()
If cboTurno.Text <> dbTurno.Recordset!Descri Then
  MsgBox "Turno inválido!"
  cboTurno.SetFocus
  Exit Sub
End If
If IsNumeric(txtLavadores.Text) = False Then
  MsgBox "Informe um valor numérico para lavadores!"
  txtLavadores.SetFocus
  Exit Sub
End If
If IsNumeric(txtTotal.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtTotal.SetFocus
  Exit Sub
End If
If IsNumeric(txtPosto.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtPosto.SetFocus
  Exit Sub
End If

With dbLavagem
  .Recordset.AddNew
  .Recordset!Dia = txtData.Value
  .Recordset!CodigoTurno = dbTurno.Recordset!CodigoTurno
  .Recordset!horaturno = dbTurno.Recordset!HoraIni
  .Recordset!Turno = dbTurno.Recordset!Descri
  .Recordset!lavadores = txtLavadores.Text
  .Recordset!TotalBruto = txtTotal.Text
  .Recordset!Totalliquido = txtPosto.Text
  .Recordset!Obs = txtObs.Text
  .Recordset.Update
  .Refresh
  .Recordset.MoveLast
End With
cboTurno.Text = ""
txtLavadores.Text = ""
txtTotal.Text = ""
txtPosto.Text = ""
txtObs.Text = ""
txtData.SetFocus
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbLavagem
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With
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
txtData.Value = Date
txtDataIni.Value = DateAdd("d", -1, Date)
txtDataFim.Value = Date

With dbLavagem
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTurno
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  lblBruto.Caption = Format("0", "Currency")
  If IsNull(.Recordset!Bruto) = False Then
    lblBruto.Caption = Format(.Recordset!Bruto, "Currency")
  End If
  lblLiquido.Caption = Format("0", "Currency")
  If IsNull(.Recordset!Liquido) = False Then
    lblLiquido.Caption = Format(.Recordset!Liquido, "Currency")
  End If
End With
Select Case Usuarios.Grupo.ControleLavagem
  Case 1 'Somente leitura
    cmdIncluir.Enabled = False
    cmdRemover.Enabled = False
  Case 2 'Liberado
    
End Select

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

Private Sub txtPosto_GotFocus()
With txtPosto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotal_GotFocus()
With txtTotal
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotal_LostFocus()
With txtTotal
  If IsNumeric(.Text) = False Then Exit Sub
  TempValor = .Text / 2
  txtPosto.Text = Format(TempValor, "currency")
  .Text = Format(.Text, "Currency")
End With
End Sub
