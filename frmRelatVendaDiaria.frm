VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatVendaDiaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venda diária de Combustível"
   ClientHeight    =   6225
   ClientLeft      =   270
   ClientTop       =   330
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Data qVendaDiaria2 
      Caption         =   "qVendaDiaria2"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Produtos where combustivel=-1 order by descri"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qVendaDiaria 
      Caption         =   "qVendaDiaria"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Produtos where combustivel=-1 order by descri"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo cboProdutos 
      Bindings        =   "frmRelatVendaDiaria.frx":0000
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Produtos where combustivel=-1 order by descri"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data qBicoEncerrante 
      Caption         =   "qBicoEncerrante"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "QLMCBicos"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbVendaDiaria 
      Caption         =   "dbVendaDiaria"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VendaDiaria"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatVendaDiaria.frx":0019
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmRelatVendaDiaria.frx":0035
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   4800
      Picture         =   "frmRelatVendaDiaria.frx":10D8
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Imprimir"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37665
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37665
   End
   Begin VB.Label Label4 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total em Valor:"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblLitros 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total em Litros:"
      Height          =   255
      Left            =   -120
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatVendaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Dia As Date)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Vendas Diárias de Combustível"
Printer.CurrentX = (190 / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (190 / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

StrTemp = "Impresso em:" & Format(Dia, "Short Date") & " - " & Format(Dia, "short time")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página " & Printer.Page
Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período:" & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short Date")
Printer.CurrentX = 0
Printer.Print StrTemp

End Sub

Private Sub Cabeca2(ByVal Dia As Date)

Printer.FontSize = 8

StrTemp = "Dia:" & Format(Dia, "Short Date")
Printer.CurrentX = 0
Printer.Print StrTemp


StrTemp = "Produto"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Volume"
Printer.CurrentX = 50 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Preço Venda"
Printer.CurrentX = 98 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Retorno"
Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total Venda"
Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(190, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

End Sub

Private Sub cboProdutos_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  If cboProdutos.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProdutos.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboProdutos.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdExibir_Click()
Dim Retorno As Double, Produto As String, lmcProduto As String
Dim Abertura As Double, Encerrante As Double, Volume As Double
Dim PrecoVenda As Currency
Dim Ws As Workspace, db As Database

If txtDataIni.Value > txtDataFim.Value Then
  MsgBox "A data final deve ser maior que a inicial!"
  Exit Sub
End If
DBGrid1.Visible = False

ProgressBar1.Value = 0
ProgressBar1.Visible = True
If cboProdutos.Text <> "" Then
  If dbProdutos.Recordset.EOF = False Then
    If dbProdutos.Recordset!Descri = cboProdutos.Text Then
      Produto = " and codproduto=" & dbProdutos.Recordset!CodigoProduto
      lmcProduto = " and codcombustivel=" & dbProdutos.Recordset!CodigoProduto
    End If
  End If
End If
Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
db.Execute "delete *from vendadiaria where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#" & Produto & ""

With dbVendaDiaria
  .RecordSource = "select *from vendadiaria where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#" & Produto & " order by data, bico"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.RecordCount <> 0
      db.Execute "delete *from vendadiaria where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#" & Produto & ""
      DoEvents
      .Refresh
    Loop
  End If
End With


With qBicoEncerrante
  .RecordSource = "Select codcombustivel, dia, sum(vendas) as venda, sum(precovenda) as total, sum(afericoes) as retorno from qlmcbicos where dia between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & lmcProduto & " group by codcombustivel, dia order by dia, codcombustivel"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe LMC gerado para este período!"
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  ProgressBar1.Value = .Recordset.PercentPosition
  Do While .Recordset.EOF = False
    dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!Codcombustivel
    ProgressBar1.Value = .Recordset.PercentPosition
    dbVendaDiaria.Recordset.AddNew
    dbVendaDiaria.Recordset!Data = .Recordset!Dia
    dbVendaDiaria.Recordset!CodProduto = .Recordset!Codcombustivel
    dbVendaDiaria.Recordset!Produto = dbProdutos.Recordset!Descri
    If IsNull(.Recordset!Total) = False Then
      dbVendaDiaria.Recordset!TotalVenda = .Recordset!Total
    Else
      dbVendaDiaria.Recordset!TotalVenda = 0
    End If
    On Error Resume Next
    dbVendaDiaria.Recordset!Preco = .Recordset!Total / .Recordset!Venda
    On Error GoTo 0
    If IsNull(.Recordset!Retorno) = False Then
      dbVendaDiaria.Recordset!Retorno = .Recordset!Retorno
    End If
    If IsNull(.Recordset!Venda) = False Then
      dbVendaDiaria.Recordset!Volume = .Recordset!Venda
    End If
    dbVendaDiaria.Recordset.Update
    .Recordset.MoveNext
  Loop
End With

With qVendaDiaria
  .RecordSource = "Select sum(volume) as litros, sum(totalvenda) as total from vendadiaria where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & Produto
  .Refresh
  If IsNull(.Recordset!Litros) = False Then
    lblLitros.Caption = Format(.Recordset!Litros, "#,##0")
  Else
    lblLitros.Caption = "0"
  End If
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format("0", "Currency")
  End If
  
End With
If dbVendaDiaria.Recordset.RecordCount <> 0 Then dbVendaDiaria.Recordset.MoveFirst

ProgressBar1.Visible = False
DBGrid1.Visible = True
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Dia2 As Date
Dim SubTotal As Currency, Total As Currency
Dim SubVolume As Double, Volume As Double
Dim PrecoMedio As Double, Litros As Double, PrecoTotal As Double

With dbVendaDiaria
  If .Recordset.RecordCount = 0 Then Exit Sub
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then GoTo NaoImprime
  On Error GoTo 0
  Printer.ScaleMode = vbMillimeters
  Dia = Now
  Cabeca Dia
  .Recordset.MoveFirst
  Dia2 = .Recordset!Data
  Cabeca2 Dia2
  Do While .Recordset.EOF = False
    If Dia2 <> .Recordset!Data Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(190, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      StrTemp = Format(SubVolume, "#,##0.0")
      Printer.CurrentX = 50 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = "Sub Total=" & Format(SubTotal, "#,##0.0000")
      Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      SubVolume = 0
      SubTotal = 0
      Dia2 = .Recordset!Data
      Cabeca2 Dia2
    End If
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(190, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      Printer.NewPage
      Cabeca Dia
      Cabeca2 Dia2
    End If
    StrTemp = .Recordset!Produto
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    Volume = Volume + .Recordset!Volume
    SubVolume = SubVolume + .Recordset!Volume
    StrTemp = Format(.Recordset!Volume, "#,###")
    Printer.CurrentX = 50 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Preco, "#,##0.0000")
    Printer.CurrentX = 98 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Retorno, "#,###")
    Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    SubTotal = SubTotal + .Recordset!TotalVenda
    Total = Total + .Recordset!TotalVenda
    StrTemp = Format(.Recordset!TotalVenda, "#,##0.0000")
    Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
End With
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(190, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = Format(SubVolume, "#,##0.0")
Printer.CurrentX = 50 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Sub Total=" & Format(SubTotal, "#,##0.0000")
Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = Format(Volume, "#,##0.0")
Printer.CurrentX = 50 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total=" & Format(Total, "#,##0.0000")
Printer.CurrentX = 190 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp



With qVendaDiaria2
  .RecordSource = "Select produto, sum(volume) as litros, sum(totalvenda) as total from vendadiaria where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & Produto & " group by produto"
  .Refresh
  
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    StrTemp = "Produto"
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = "Volume"
    Printer.CurrentX = 60 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Total"
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Preço médio"
    Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(120, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Total = 0
    Volume = 0
    Do While .Recordset.EOF = False
      
      StrTemp = .Recordset!Produto
      Printer.CurrentX = 0
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Litros) = False Then
        StrTemp = Format(.Recordset!Litros, "#,##0")
        Volume = Volume + .Recordset!Litros
        Litros = .Recordset!Litros
      Else
        StrTemp = "0"
      End If
      Printer.CurrentX = 60 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Total) = False Then
        StrTemp = Format(.Recordset!Total, "Currency")
        Total = Total + .Recordset!Total
        PrecoTotal = .Recordset!Total
      Else
        StrTemp = Format("0", "Currency")
      End If
      Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      PrecoMedio = 0
      If PrecoTotal <> 0 And Litros <> 0 Then
        PrecoMedio = PrecoTotal / Litros
      End If
      StrTemp = Format(PrecoMedio, "#,##0.000")
      Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(120, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      .Recordset.MoveNext
    Loop
    
    StrTemp = Format(Volume, "#,##0")
    Printer.CurrentX = 60 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Total, "Currency")
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
  End If
End With




Printer.EndDoc
NaoImprime:
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_DblClick()
dbVendaDiaria.Refresh
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

With dbVendaDiaria
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from vendadiaria where codigovendadiaria=0 order by data, bico, notanr"
  .Refresh
End With
With qBicoEncerrante
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from qlmcbicos order by dia, bico"
  .Refresh
End With
With dbProdutos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With qVendaDiaria
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With qVendaDiaria2
  .DatabaseName = Caminho
  .Connect = Conectar
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
