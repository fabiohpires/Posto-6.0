VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatCompras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Compras"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmRelatCompras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbTurnos 
      Height          =   330
      Left            =   4920
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from Turnos order by horaini"
      Caption         =   "dbTurnos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmRelatCompras.frx":0442
      Height          =   315
      Left            =   3720
      TabIndex        =   24
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.TextBox txtNrNota2 
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.Data dbLmc 
      Caption         =   "dbLmc"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMC"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbLmcNotas 
      Caption         =   "dbLmcNotas"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMCNotas"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbProdutosNotasCorpo 
      Caption         =   "dbProdutosNotasCorpo"
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
      RecordSource    =   "ProdutosNotasCorpo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbProdutosEntrada2 
      Caption         =   "dbProdutosEntrada2"
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
      RecordSource    =   "ProdutosEntrada2"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbProdutosNotas 
      Caption         =   "dbProdutosNotas"
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
      RecordSource    =   "ProdutosNotas"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdAlteraData 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdFonte 
      Caption         =   "Fonte"
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "MS Sans Serif"
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Bindings        =   "frmRelatCompras.frx":0459
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
   Begin VB.Data qEntrada 
      Caption         =   "qEntrada"
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
      RecordSource    =   "select sum(quantidade) as qtd, sum(ValorNota) as total1 from produtosEntrada2"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
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
      RecordSource    =   "select *from produtos order by descri"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbEntrada 
      Caption         =   "dbEntrada"
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
      RecordSource    =   $"frmRelatCompras.frx":0472
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtNrNota 
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton optAmbos 
      Caption         =   "Ambos"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   840
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optNaoCombustivel 
      Caption         =   "Não Combustível"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optCombustivel 
      Caption         =   "Combustível"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   6600
      Picture         =   "frmRelatCompras.frx":05C5
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   840
      Width           =   975
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
      Format          =   93192193
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
      Format          =   93192193
      CurrentDate     =   37678
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatCompras.frx":1047
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frmRelatCompras.frx":105F
      TabIndex        =   15
      Top             =   1320
      Width           =   10095
   End
   Begin MSComCtl2.DTPicker txtDataAlterar 
      Height          =   300
      Left            =   2280
      TabIndex        =   22
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   93192193
      CurrentDate     =   37678
   End
   Begin VB.Label Label9 
      Caption         =   "Truno:"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Data:"
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Nr. Nota:"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Nr. Nota:"
      Height          =   255
      Left            =   8400
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   19
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Label lblQtd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   5280
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   6720
      TabIndex        =   17
      Top             =   5280
      Width           =   405
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   540
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String, ColunaQuebra As Integer

Private Sub Filtrar()
Dim StrTemp As String, StrTemp2 As String
StrTemp = "SELECT produtosnotas.*, produtosnotascorpo.*, Turnos.*, Produtos.* FROM (produtosnotas LEFT JOIN Turnos ON produtosnotas.codigoTurno = Turnos.CodigoTurno) LEFT JOIN (produtosnotascorpo LEFT JOIN Produtos ON produtosnotascorpo.CodigoProduto = Produtos.CodigoProduto) ON produtosnotas.CodigoEntrada = produtosnotascorpo.CodigoProdutoNota where datanota between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
StrTemp2 = "Select sum(quantidade) as qtd, sum(Total) as total1 from qprodutosnotas where datanota between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
If cboProduto.Text = dbProdutos.Recordset!Descri Then
  StrTemp = StrTemp & " and produtos.codigo=" & dbProdutos.Recordset!Codigo
  StrTemp2 = StrTemp2 & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
Else
  If optAmbos.Value = False Then
    If optCombustivel.Value = True Then
      StrTemp = StrTemp & " and tanque > 0"
      StrTemp2 = StrTemp2 & " and tanque > 0"
    Else
      StrTemp = StrTemp & " and tanque=0"
      StrTemp2 = StrTemp2 & " and tanque=0"
    End If
  End If
End If

If txtNrNota.Text <> "" Then
  StrTemp = StrTemp & " and nrnota='" & txtNrNota.Text & "'"
  StrTemp2 = StrTemp2 & " and nrnota='" & txtNrNota.Text & "'"
End If

With dbEntrada
  .RecordSource = StrTemp & strOrdem
  .Refresh
End With
With qEntrada
  .RecordSource = StrTemp2
  .Refresh
  If IsNull(.Recordset!Qtd) = False Then
    lblQtd.Caption = Format(.Recordset!Qtd, "#,###")
  Else
    lblQtd.Caption = Format(0, "#,###")
  End If
  If IsNull(.Recordset!Total1) = False Then
    lblTotal.Caption = Format(.Recordset!Total1, "#,##0.0000")
  Else
    lblTotal.Caption = Format(0, "#,##0.0000")
  End If
End With
End Sub
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

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "descri='" & cboTurno.Text & "'"
End With
End Sub

Private Sub cmdAlteraData_Click()
Dim CodigoNota As Double
If dbEntrada.Recordset.EOF = True Then
  MsgBox "Selecione uma nota primeiro!"
  Exit Sub
End If
If txtNrNota2.Text = "" Then
  MsgBox "Número da nota inválido!"
  txtNrNota.SetFocus
  Exit Sub
End If
CodigoNota = dbEntrada.Recordset!CodigoEntrada
With dbProdutosNotas
  .Recordset.FindFirst "codigoentrada=" & CodigoNota
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de notas! Nota não encontrada!"
    Exit Sub
  End If
End With
Call cboTurno_LostFocus
If dbTurnos.Recordset.EOF = True Then
  MsgBox "Turno incorreto!"
  Exit Sub
End If

With dbProdutosEntrada2
  .RecordSource = "select *from produtosentrada2 where codigonota=" & CodigoNota
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .RecordSource = "select *from produtosentrada2 where data=#" & DataInglesa(dbEntrada.Recordset!datanota) & "# and codigoproduto=" & dbEntrada.Recordset("produtos.CodigoProduto")
    .Refresh
    With dbProdutosNotasCorpo
      .RecordSource = "Select *from produtosnotascorpo where codigoprodutonota=" & CodigoNota
      .Refresh
      If .Recordset.RecordCount = 0 Then
        MsgBox "Esta nota não possue lançamentos"
        Exit Sub
      End If
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutosEntrada2.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto & " and quantidade=" & .Recordset!Quantidade & " and preconovo=" & NumeroIngles(.Recordset!valorUnitario) & " and tanque=" & .Recordset!Tanque
        If dbProdutosEntrada2.Recordset.NoMatch = False Then
          If IsNull(dbProdutosEntrada2.Recordset!CodigoNota) = True Then
            dbProdutosEntrada2.Recordset.Edit
            dbProdutosEntrada2.Recordset!CodigoNota = CodigoNota
            dbProdutosEntrada2.Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
            dbProdutosEntrada2.Recordset.Update
          End If
        End If
        .Recordset.MoveNext
      Loop
    End With
  End If
  .RecordSource = "select *from produtosentrada2 where codigonota=" & CodigoNota
  .Refresh
  If .Recordset.RecordCount > 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!Data = txtDataAlterar.Value
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  
End With
With dbProdutosNotas
  .Recordset.Edit
  .Recordset!dataentrega = txtDataAlterar.Value
  .Recordset!datanota = txtDataAlterar.Value
  .Recordset!NrNota = txtNrNota2.Text
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset.Update
End With

With dbLmc
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "dia=#" & DataInglesa(txtDataAlterar.Value) & "# and codcombustivel=" & dbEntrada.Recordset("produtos.codigoproduto")
    If .Recordset.NoMatch = False Then
      CodigoLMC = .Recordset!CodLMC
    Else
      CodigoLMC = 0
    End If
  End If
End With
If CodigoLMC <> 0 Then
  With dbLmcNotas
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.FindFirst "codnota=" & CodigoNota
      If .Recordset.NoMatch = False Then
        .Recordset.Edit
        .Recordset!datanota = txtDataAlterar.Value
        .Recordset!CodLMC = CodigoLMC
        .Recordset.Update
      End If
    End If
  End With
End If
Call cmdExibir_Click

dbEntrada.Recordset.FindFirst "codigoentrada=" & CodigoNota

End Sub

Private Sub cmdExibir_Click()
Filtrar
End Sub

Private Sub cmdFonte_Click()
On Error GoTo semFonte
CommonDialog1.flags = cdlCFBoth
CommonDialog1.ShowFont
With CommonDialog1
  DBGrid1.Font.Name = .FontName
  DBGrid1.Font.Size = .FontSize
  DBGrid1.Font.Bold = .FontBold
  DBGrid1.Font.Italic = .FontItalic
End With
semFonte:
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date
Dim Total1 As Double, Total2 As Currency


  If dbEntrada.Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  Dia = Now
  StrTemp = StrTemp & Chr(13) & "Data: " & Format(Dia, "Long Date")
  StrTemp = StrTemp & Chr(13) & "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short Date")
  If cboProduto.Text = dbProdutos.Recordset!Descri Then
    StrTemp = StrTemp & Chr(13) & "Código: " & txtCodigo.Text & " - Produto: " & cboProduto.Text
  End If
  
  ImprimeGrid DBGrid1, Printer, dbEntrada, 7, True, , ColunaQuebra, 8, , "Relatório de Produtos Comprados", NomePosto, StrTemp
  
  Total1 = 0
  For i = 0 To 6
    Total1 = Total1 + DBGrid1.Columns(i).Width
  Next i
  Total2 = 0
  For i = 0 To 7
    Total2 = Total2 + DBGrid1.Columns(i).Width
  Next i
  
  Printer.EndDoc
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbEntrada_Reposition()
On Error Resume Next
txtDataAlterar.Value = dbEntrada.Recordset!datanota
txtNrNota2.Text = dbEntrada.Recordset!NrNota
cboTurno.Text = dbEntrada.Recordset("turnos.descri")
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", datanota" Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc, datanota"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", datanota"
End If
ColunaQuebra = ColIndex
Filtrar
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
ColunaQuebra = 1
strOrdem = " order by produtos.codigo, datanota"
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbEntrada
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select qprodutosnotas.*, produtos.*from qprodutosnotas, produtos where produtos.codigoproduto=qprodutosnotas.codigoproduto"
  .Refresh
End With
With dbProdutos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With qEntrada
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbProdutosEntrada2
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbProdutosNotas
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbProdutosNotasCorpo
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbLmcNotas
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbLmc
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With

Filtrar
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

Private Sub txtDataAlterar_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataAlterar_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataAlterar_LostFocus()
Me.KeyPreview = True
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
