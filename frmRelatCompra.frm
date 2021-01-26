VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatCompra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório para Compras"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   13365
   Icon            =   "frmRelatCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbTemp4 
      Caption         =   "dbTemp4"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(estoque*precocompra) as total from produtos where combustivel=0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbTemp3 
      Caption         =   "dbTemp3"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from VendasTemp where codigoproduto=0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbTemp2 
      Caption         =   "dbTemp2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from VendasTemp where codigoproduto=0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbTemp 
      Caption         =   "dbTemp"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from VendasTemp where codigoproduto=0"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbVendas 
      Caption         =   "dbVendas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from VendasTemp where codigoproduto=0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
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
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   137232385
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   137232385
      CurrentDate     =   37678
   End
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   6135
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   13215
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label lblTotalEstoque 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Total de custo do estoque:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   1935
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRReport As New CRAXDRT.Report
Dim CRApp As New CRAXDRT.Application
Dim strOrdem As String

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Compra/Venda"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = Format(Dia, "short date") & " - " & Format(Dia, "Short Time")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

Printer.FontSize = 8

StrTemp = "Cod."
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Produto"
Printer.CurrentX = 10
Printer.Print StrTemp;

StrTemp = "Estoque"
Printer.CurrentX = 75 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Custo"
Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Venda"
Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Vendas"
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Pedido"
Printer.CurrentX = 139 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Forn."
Printer.CurrentX = 159 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Fornecedor"
Printer.CurrentX = 160
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

Printer.Print

End Sub

Private Sub cmdExibir_Click()
Dim Dias As Double, Sugerido As Double, Vendido As Double, Comprado As Double
Dim DBAdo As New ADODB.Connection
Dim DbVendasNaoFinalizadas As New ADODB.Recordset
Dim DbNotasNaoConfirmadas As New ADODB.Recordset
Dim Comissao As Double, ComissaoValor As Double

Dias = DateDiff("d", txtDataIni.Value, txtDataFim.Value)
If Dias = 0 Then Dias = 1
If Dias < 0 Then Dias = Dias * -1

If DBAdo.State = adStateOpen Then
  DBAdo.Close
End If
DBAdo.Open CaminhoADO
DbVendasNaoFinalizadas.CursorLocation = adUseClient
DbVendasNaoFinalizadas.Open "SELECT Venda2.CodigoProduto, sum(Venda2.Quantidade) AS vendido, sum(venda2.valortotal) AS valor FROM Venda2 where fechamentodiario=0 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by venda2.codigoproduto", DBAdo, adOpenKeyset, adLockOptimistic
DbNotasNaoConfirmadas.CursorLocation = adUseClient
DbNotasNaoConfirmadas.Open "SELECT qprodutosnotas.codigoproduto, sum(qprodutosnotas.Quantidade) AS qtd FROM qprodutosnotas where confirmado=0 and dataentrega between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by qprodutosnotas.codigoproduto", DBAdo, adOpenKeyset, adLockOptimistic


With dbTemp3
  .RecordSource = "Select codigoproduto, sum(vendido) as vendas from qvendadia2 where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by codigoproduto"
  .Refresh
End With
With dbVendas
  .RecordSource = "select *from vendastemp"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With dbTemp
  .RecordSource = "select *from produtos where combustivel=0"
  .Refresh
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If dbVendas.Recordset.EOF = True Then
      dbVendas.Recordset.AddNew
      dbVendas.Recordset!CodigoProduto = .Recordset!CodigoProduto
      dbVendas.Recordset!CodProduto = .Recordset!Codigo
      dbVendas.Recordset!Descri = Left(.Recordset!Descri, 30)
    Else
      dbVendas.Refresh
      dbVendas.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbVendas.Recordset.NoMatch = True Then
        dbVendas.Recordset.AddNew
        dbVendas.Recordset!CodigoProduto = .Recordset!CodigoProduto
        dbVendas.Recordset!CodProduto = .Recordset!Codigo
        dbVendas.Recordset!Descri = Left(.Recordset!Descri, 30)
      Else
        dbVendas.Recordset.Edit
      End If
    End If
    Vendido = 0
    Comprado = 0
    If DbVendasNaoFinalizadas.RecordCount <> 0 Then
      DbVendasNaoFinalizadas.MoveFirst
      DbVendasNaoFinalizadas.Find "codigoproduto=" & .Recordset!CodigoProduto
      If DbVendasNaoFinalizadas.EOF = False Then
        Vendido = DbVendasNaoFinalizadas!Vendido
      End If
    End If
    If DbNotasNaoConfirmadas.RecordCount <> 0 Then
      DbNotasNaoConfirmadas.MoveFirst
      DbNotasNaoConfirmadas.Find "codigoproduto=" & .Recordset!CodigoProduto
      If DbNotasNaoConfirmadas.EOF = False Then
        Comprado = DbNotasNaoConfirmadas!Qtd
      End If
    End If
    dbVendas.Recordset!Comissao = .Recordset!Comissao
    dbVendas.Recordset!ComissaoValor = .Recordset!ComissaoValor
    dbVendas.Recordset!Estoque = .Recordset!Estoque - Vendido + Comprado
    dbVendas.Recordset!precocompra = .Recordset!precocompra
    dbVendas.Recordset!PrecoVenda = .Recordset!PrecoVenda
    
    If IsNull(.Recordset!duracaoestoque) = False Then
      dbVendas.Recordset!duracaoestoque = .Recordset!duracaoestoque
    Else
      dbVendas.Recordset!duracaoestoque = 0
    End If
    If IsNull(.Recordset!estoqueideal) = False Then
      dbVendas.Recordset!estoqueideal = .Recordset!estoqueideal
    Else
      dbVendas.Recordset!estoqueideal = 0
    End If
    dbTemp3.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
    
    If dbTemp3.Recordset.NoMatch = False Then
      dbVendas.Recordset!TotalVendido = dbTemp3.Recordset!Vendas
    Else
      dbVendas.Recordset!TotalVendido = 0
    End If
    
    
    Sugerido = 0
    If dbVendas.Recordset!TotalVendido <> 0 Then
      Sugerido = ((dbVendas.Recordset!TotalVendido / Dias) * dbVendas.Recordset!duracaoestoque) - dbVendas.Recordset!Estoque
    End If
    If Sugerido < 0 Then Sugerido = 0
    If IsNull(.Recordset!lucrominimo) = False Then
      dbVendas.Recordset!lucrominimo = .Recordset!lucrominimo
    End If
    
    dbVendas.Recordset!Sugerido = Sugerido
    
    Sugerido = 0
    
    If .Recordset!lucrominimo <> 0 Then
      Comissao = 0
      ComissaoValor = 0
      Sugerido = .Recordset!precocompra + (.Recordset!precocompra * (.Recordset!lucrominimo / 100))
      If .Recordset!ComissaoValor <> 0 Then
        Sugerido = Sugerido + .Recordset!ComissaoValor
      End If
      If .Recordset!Comissao <> 0 Then
        Comissao = .Recordset!Comissao
        Sugerido = Sugerido / (1 - (Comissao))
      End If
    End If
    If .Recordset!Sugerido <> Sugerido Then
      .Recordset.Edit
      .Recordset!Sugerido = Sugerido
      .Recordset.Update
    End If
    dbVendas.Recordset!precosugerido = Sugerido
    dbVendas.Recordset!uncaixa = .Recordset!uncaixa
    dbVendas.Recordset.Update
    .Recordset.MoveNext
  Loop
End With

dbVendas.RecordSource = "select *from vendastemp" & strOrdem
dbVendas.Refresh
With dbTemp4
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalEstoque.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalEstoque.Caption = Format(0, "Currency")
  End If
End With

DbVendasNaoFinalizadas.Close
DbNotasNaoConfirmadas.Close
DBAdo.Close



Set CRReport = CRApp.OpenReport("Relatório de Compra e Venda .rpt")
'On Error GoTo ExitLabel
    With CRReport
        For i = 1 To .Database.Tables.Count
            .Database.Tables(i).Location = Caminho
        Next i
        .ParameterFields.GetItemByName("NomePosto").AddCurrentValue NomePosto
        .ParameterFields.GetItemByName("DataIni").AddCurrentValue txtDataIni.Value
        .ParameterFields.GetItemByName("DataFim").AddCurrentValue txtDataFim.Value
        .ParameterFields.GetItemByName("Dias").AddCurrentValue "Vend. " & Dias
    End With
    With CR
        .ReportSource = CRReport
        .EnablePopupMenu = True
        .ViewReport
    End With
    CRApp.CanClose
    Exit Sub
'ExitLabel:
'    MsgBox "DungTran:" & Err.Description



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
txtDataIni.Value = Date
txtDataFim.Value = Date

strOrdem = " order by CodProduto"

Dim Ws As Workspace, db As Database

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
db.Execute "delete *from vendastemp"


With dbVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  .Recordset.Sort = strOrdem
End With
With dbTemp
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTemp2
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTemp3
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTemp4
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalEstoque.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalEstoque.Caption = Format(0, "Currency")
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
