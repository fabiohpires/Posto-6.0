VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatVendasDetalhado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório detalhado de Vendas"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   7920
      Picture         =   "frmRelatVendasDetalhado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Data QVendasDia2 
      Caption         =   "QVendasDia2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatVendasDetalhado.frx":0A82
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from produtos order by descri"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   37678
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmRelatVendasDetalhado.frx":0B0B
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "frmRelatVendasDetalhado.frx":0B25
      TabIndex        =   6
      Top             =   840
      Width           =   8535
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblVendido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatVendasDetalhado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontSize = 14
Printer.FontName = "Arial"
StrTemp = "Relatório de Venda Detalhado"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = "Impresso em: " & Format(Dia, "short date") & " - " & Format(Dia, "short time")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Período de: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Data Venda"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Código"
Printer.CurrentX = 42 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Produto"
Printer.CurrentX = 45
Printer.Print StrTemp;

StrTemp = "Quantidade"
Printer.CurrentX = 149 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total"
Printer.CurrentX = 180 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(180, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

End Sub

Private Sub cboProdutos_LostFocus()
With dbProdutos
  .Refresh
  If cboProdutos.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProdutos.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodProduto.Text = .Recordset!Codigo
    cboProdutos.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp1 As String, StrTemp2 As String, strTemp3 As String
Dim Total As Currency, Vendido As Double

StrTemp1 = " and data between #" & DataInglesa(Str(txtDataIni.Value)) & "# and #" & DataInglesa(Str(txtDataFim.Value)) & "#"
StrTemp1 = FiltraProdutos(StrTemp1)
Vendido = 0
Total = 0

With QVendasDia2
  .RecordSource = "select qvendadia2.*, produtos.*  from qvendadia2, produtos where qvendadia2.codigoproduto=produtos.codigoproduto" & StrTemp1 & " order by descri,data"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Vendido = Vendido + .Recordset!Vendido
      Total = Total + .Recordset!Valor
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

lblTotal.Caption = Format(Total, "Currency")
lblVendido.Caption = Format(Vendido, "General Number")

End Sub

Private Sub cmdImprime_Click()
Dim Dia As Date, Largura As Double

On Error GoTo TrataErro
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Dia = Now
Largura = 180
Cabeca Dia, Largura

If QVendasDia2.Recordset.RecordCount <> 0 Then
  With QVendasDia2
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.Line (0, Printer.CurrentY)-(180, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.5
        
        StrTemp = "Página: " & Printer.Page
        Printer.CurrentX = 0
        Printer.Print StrTemp;
        
        Printer.NewPage
        
        Cabeca Dia, Largura
      End If
      StrTemp = Format(.Recordset!Data, "Short date")
      Printer.CurrentX = 0
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Codigo
      Printer.CurrentX = 42 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 45
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Vendido, "General Number")
      Printer.CurrentX = 149 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Valor, "Currency")
      Printer.CurrentX = 180 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End With
End If

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(180, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Total:"
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = lblVendido.Caption
Printer.CurrentX = 149 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = lblTotal.Caption
Printer.CurrentX = 180 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.EndDoc
TrataErro:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QVendasDia2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select qvendadia2.*, produtos.*  from qvendadia2, produtos where qvendadia2.codigoproduto=produtos.codigoproduto and qvendadia2.codigoproduto=0 order by descri,data"
  .Refresh
End With
End Sub

Private Sub txtCodProduto_GotFocus()
With txtCodProduto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos
  .Refresh
  If txtCodProduto.Text = "" Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    txtCodProduto.Text = .Recordset!Codigo
    cboProdutos.Text = .Recordset!Descri
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
  Case vbKeyEscape
    Call cmdSair_Click
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
  Case vbKeyEscape
    Call cmdSair_Click
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
