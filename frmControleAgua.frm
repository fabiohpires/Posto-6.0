VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControleAgua 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Água"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmControleAgua.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdConfigura 
      Caption         =   "Configurar"
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   240
      Picture         =   "frmControleAgua.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Imprimir"
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4560
      Width           =   855
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
      RecordSource    =   "select *from ControleAgua order by data"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmControleAgua.frx":0EC4
      Height          =   3375
      Left            =   240
      OleObjectBlob   =   "frmControleAgua.frx":0ED9
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   240
      TabIndex        =   0
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
      Left            =   240
      TabIndex        =   10
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
      Left            =   1920
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37869
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Previsto:"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1680
      TabIndex        =   12
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Período para Cálculo:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Número Lido:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Data da Leitura:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmControleAgua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

StrTemp = "Relatório de Controle de Água"
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
With dbAgua
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "data>=#" & DataInglesa(txtData.Value) & "#"
    If .Recordset.NoMatch = False Then
      MsgBox "Número lançado na data!"
      txtData.SetFocus
      Exit Sub
    End If
  End If
  .Recordset.AddNew
  Codigo = .Recordset!codigoagua
  .Recordset!Data = txtData.Value
  .Recordset!Valor = CDbl(txtNumero.Text)
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
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbAgua
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Select Case Usuarios.Grupo.ControleAgua
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

Private Sub txtNumero_GotFocus()
With txtNumero
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub
