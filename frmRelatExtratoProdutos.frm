VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatExtratoProdutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extrato de Produtos"
   ClientHeight    =   7410
   ClientLeft      =   240
   ClientTop       =   720
   ClientWidth     =   11790
   Icon            =   "frmRelatExtratoProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Produtos order by Descri"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data dbExtrato 
      Caption         =   "dbExtrato"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ProdutosHistorico order by codigohistorico"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatExtratoProdutos.frx":0442
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "frmRelatExtratoProdutos.frx":045A
      TabIndex        =   10
      Top             =   840
      Width           =   11535
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   10440
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   9480
      Picture         =   "frmRelatExtratoProdutos.frx":16D9
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Imprimir"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Bindings        =   "frmRelatExtratoProdutos.frx":215B
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   91488257
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   91488257
      CurrentDate     =   37678
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmRelatExtratoProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboProduto_LostFocus()
Me.KeyPreview = True
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String
StrTemp = "select *from produtoshistorico where dataalteracao between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
If dbProdutos.Recordset.EOF = False Then
  If cboProduto.Text <> "" Then
    If dbProdutos.Recordset!Descri = cboProduto.Text Then
      StrTemp = StrTemp & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
    End If
  End If
End If
StrTemp = StrTemp & " order by codigohistorico"
With dbExtrato
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp
  .Refresh
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String
StrTemp = ""
If dbProdutos.Recordset.EOF = False Then
  If cboProduto.Text <> "" Then
    If dbProdutos.Recordset!Descri = cboProduto.Text Then
      StrTemp = "Código: " & dbProdutos.Recordset!Codigo & "    Produto: " & dbProdutos.Recordset!Descri
    End If
  End If
End If
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

ImprimeGrid DBGrid1, Printer, dbExtrato, , , , , , , "Extrato de Produtos", NomePosto & Chr(vbKeyReturn) & "Período: " & txtDataIni.Value & " a " & txtDataFim.Value & Chr(vbKeyReturn) & StrTemp

Printer.EndDoc

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
With dbExtrato
  .Connect = Conectar
  .DatabaseName = Caminho
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Call cmdExibir_Click
End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyAscii = 0
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
    KeyAscii = 0
    On Error Resume Next
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
