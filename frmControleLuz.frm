VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControleLuz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Luz"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdConfigura 
      Caption         =   "Configurar"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmControleLuz.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Imprimir"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtFator 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.Data dbAgua 
      Caption         =   "dbAgua"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Controleluz order by data"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmControleLuz.frx":0A82
      Height          =   3375
      Left            =   240
      OleObjectBlob   =   "frmControleLuz.frx":0A97
      TabIndex        =   8
      Top             =   840
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37869
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37869
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37869
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Período para Cálculo:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Previsto:"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fator:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Data da Leitura:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Número Lido:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmControleLuz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

StrTemp = "Relatório de Controle de Luz"
Printer.FontSize = 14
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.FontSize = 14
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = Format(Dia, "Short date") & " - " & Format(Dia, "Short Time")
Printer.FontSize = 10
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Data"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Leitura"
Printer.CurrentX = 70 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Média Apurada"
Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1


End Sub
Private Sub Recalcular()
Dim Dias As Double, ValorIni As Double, ValorFinal As Double, Media As Double
Dim DataFim As Date, DataIni As Date
With dbAgua
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  DataIni = .Recordset!Data
  ValorInicial = .Recordset!Valor
  
  Do While .Recordset.EOF = False
    DataFim = .Recordset!Data
    ValorFinal = .Recordset!Valor
    Dias = DateDiff("d", DataIni, DataFim)
    If Dias < 0 Then Dias = Dias * -1
    If Dias = 0 Then Dias = 1
    Media = (ValorFinal - ValorInicial) / Dias
    .Recordset.Edit
    .Recordset!Media = Media
    .Recordset.Update
    DataIni = .Recordset!Data
    ValorInicial = .Recordset!Valor
    .Recordset.MoveNext
  Loop
End With
End Sub

Private Sub Calcular(ByVal Codigo As Double)
Dim Dias As Double, ValorIni As Double, ValorFinal As Double, Media As Double
Dim DataFim As Date, DataIni As Date
With dbAgua
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  DataFim = .Recordset!Data
  ValorFinal = .Recordset!Valor
  .Recordset.FindFirst "CodigoAgua=" & Codigo
  If .Recordset!Data < DataFim Then
    Recalcular
    Exit Sub
  End If
  .Recordset.MovePrevious
  If .Recordset.BOF = True Then Exit Sub
  DataIni = .Recordset!Data
  ValorInicial = .Recordset!Valor
  Dias = DateDiff("d", DataIni, DataFim)
  If Dias < 0 Then Dias = Dias * -1
  If Dias = 0 Then Dias = 1
  Media = (ValorFinal - ValorInicial) / Dias
  .Recordset.FindFirst "CodigoAgua=" & Codigo
  .Recordset.Edit
  .Recordset!Media = Media
  .Recordset.Update
  .Recordset.MoveNext
End With
End Sub

Private Sub cmdConfigura_Click()
frmControleLuzConfig.Show vbModal
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double

With dbAgua
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Dia = Now
  Largura = 115
  Cabeca Dia, Largura
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.NewPage
      Cabeca Dia, Largura
      
    End If
    Printer.FontSize = 10
    
    StrTemp = Format(.Recordset!Data, "short date")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Valor, "#,##0.00")
    Printer.CurrentX = 70 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Media, "#,##0.00")
    Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.EndDoc
End With
NaoImprime:
End Sub

Private Sub cmdIncluir_Click()
Dim Codigo As Double
If IsNumeric(txtNumero.Text) = False Then
  MsgBox "Informe um número válido"
  txtNumero.SetFocus
  Exit Sub
End If
If IsNumeric(txtFator.Text) = False Then
  MsgBox "Informe um fator Válido!"
  txtFator.SetFocus
  Exit Sub
End If
With dbAgua
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "data>=#" & DataInglesa(txtData.Value) & "#"
    If .Recordset.NoMatch = False Then
      MsgBox "Número já lançado na data!"
      txtData.SetFocus
      Exit Sub
    End If
  End If
  .Recordset.AddNew
  Codigo = .Recordset!codigoagua
  .Recordset!Data = txtData.Value
  .Recordset!Valor = CDbl(txtNumero.Text) * CDbl(txtFator.Text)
  .Recordset.Update
  Calcular Codigo
  .Refresh
  .Recordset.MoveLast
  txtNumero.Text = ""
  txtData.SetFocus
End With
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbAgua
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o registro atual!", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF5
    Recalcular
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
  Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
txtData.Value = Date
With dbAgua
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Select Case Usuarios.Grupo.ControleLuz
  Case 1 'Somente leitura
    cmdIncluir.Enabled = False
    cmdRemover.Enabled = False
    cmdRecalcular.Enabled = False
    cmdConfigura.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyPress(KeyCode)
End Sub

Private Sub txtFator_GotFocus()
With txtFator
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNumero_GotFocus()
With txtNumero
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

