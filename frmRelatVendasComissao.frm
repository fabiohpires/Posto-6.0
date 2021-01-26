VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatVendasComissao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Lucro de Vendas por Funcionário"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from produtos order by descri"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Height          =   315
      Left            =   4080
      TabIndex        =   15
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   9480
      Picture         =   "frmRelatVendasComissao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Data qVendas2 
      Caption         =   "qVendas2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from qcomissoes2 where venda2.codigofechamento=0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbVendas2 
      Caption         =   "dbVendas2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatVendasComissao.frx":0A82
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatVendasComissao.frx":0B5C
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frmRelatVendasComissao.frx":0B74
      TabIndex        =   0
      Top             =   960
      Width           =   10095
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin VB.Label lblVBruto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   5760
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Bruto:"
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   5760
      Width           =   1725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblComissao 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   5760
      Width           =   1365
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Comissao:"
      Height          =   195
      Left            =   5040
      TabIndex        =   9
      Top             =   5760
      Width           =   1125
   End
   Begin VB.Label lblVendido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   5760
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Lucro:"
      Height          =   195
      Left            =   7920
      TabIndex        =   7
      Top             =   5760
      Width           =   840
   End
End
Attribute VB_Name = "frmRelatVendasComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & txtCodigo.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String, StrTemp2 As String

StrTemp = "select vendedores.codigo, nome, sum(valortotal) as valorbruto, sum(valorcomissao) as comissoes, sum((quantidade*precocompra)-valorcomissao) as lucro from qVendasComissoes where fechamentodiario=-1 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
StrTemp2 = "select sum(valortotal) as valorbruto, sum(valorcomissao) as comissoes, sum((quantidade*precocompra)-valorcomissao) as lucro from qVendasComissoes where fechamentodiario=-1 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"

If txtCodigo.Text = dbProdutos.Recordset!Codigo Then
  StrTemp = StrTemp & " and produtos.codigo=" & dbProdutos.Recordset!Codigo
  StrTemp2 = StrTemp2 & " and produtos.codigo=" & dbProdutos.Recordset!Codigo
End If

StrTemp = StrTemp & " group by vendedores.codigo, nome"
With dbVendas2
  .RecordSource = StrTemp
  .Refresh
End With
With qVendas2
  .RecordSource = StrTemp2
  .Refresh
  If IsNull(.Recordset!ValorBruto) = False Then
    lblVBruto.Caption = Format(.Recordset!ValorBruto, "Currency")
  Else
    lblVBruto.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Comissoes) = False Then
    lblVendido.Caption = Format(.Recordset!Comissoes, "Currency")
  Else
    lblVendido.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Lucro) = False Then
    lblComissao.Caption = Format(.Recordset!Lucro, "Currency")
  Else
    lblComissao.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double

With dbVendas2
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  On Error GoTo TrataErro
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Largura = 190
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = "Lucro / Comissões de Produtos Comissionados"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontSize = 10
  
  StrTemp = "Impresso em: " & Format(Now, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
  
  StrTemp = "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
  
  If txtCodigo.Text = dbProdutos.Recordset!Codigo Then
    StrTemp = "Codigo: " & dbProdutos.Recordset!Codigo & "   Produto:" & dbProdutos.Recordset!Descri
    Printer.CurrentX = 0
    Printer.Print StrTemp
  End If
  
  StrTemp = "Codigo"
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Funcionário"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  
  StrTemp = "Total Bruto"
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Total Comissão"
  Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Lucro"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Do While .Recordset.EOF = False
    StrTemp = .Recordset!Codigo
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 15
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!ValorBruto, "Currency")
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Comissoes, "Currency")
    Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Lucro, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(qVendas2.Recordset!ValorBruto, "Currency")
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(qVendas2.Recordset!Comissoes, "Currency")
  Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(qVendas2.Recordset!Lucro, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
TrataErro:

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
With dbVendas2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select vendedores.codigo, nome, sum(valortotal) as valorbruto, sum(valorcomissao) as comissoes, sum(valortotal-valorcomissao) as lucro from qVendasComissoes where codigofechamento=0 group by nome, vendedores.codigo"
  .Refresh
End With
With qVendas2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valortotal) as valorbruto, sum(valorcomissao) as comissoes, sum(valortotal-valorcomissao) as lucro from qVendasComissoes where codigofechamento=0"
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbProdutos
  .Refresh
  If txtCodigo.Text = "" Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodigo.Text
  If .Recordset.NoMatch = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub
