VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatDifCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Diferença de Caixa"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmRelatDifCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.Data qFechamento 
      Caption         =   "qFechamento"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from FechamentoDeCaixa order by datacaixa, HoraIni"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data dbFechamento 
      Caption         =   "dbFechamento"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from FechamentoDeCaixa order by datacaixa, HoraIni"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   9000
      Picture         =   "frmRelatDifCaixa.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin MSAdodcLib.Adodc dbVendedores 
      Height          =   330
      Left            =   3720
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from vendedores order by nome"
      Caption         =   "dbVendedores"
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
   Begin MSDataListLib.DataCombo cboVendedores 
      Bindings        =   "frmRelatDifCaixa.frx":0EC4
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   240
      Width           =   855
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
   Begin MSAdodcLib.Adodc dbPosto 
      Height          =   330
      Left            =   3720
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from postos order by nome"
      Caption         =   "dbPosto"
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
      Bindings        =   "frmRelatDifCaixa.frx":0EDF
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "frmRelatDifCaixa.frx":0EFA
      TabIndex        =   12
      Top             =   840
      Width           =   9615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblTotalDif 
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
      Left            =   8040
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   7560
      TabIndex        =   10
      Top             =   5040
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Responsável:"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   975
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatDifCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String, Coluna As Integer

Private Sub Imprime2()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Quebra As String, Altera As String
Dim TotalParcial As Currency, Tota As Currency

With dbFechamento
  Quebra = DBGrid1.Columns(Coluna).DataField
  StrTemp = qFechamentos.Recordset.Sort
  Call cmdExibir_Click
  qFechamentos.Recordset.Sort = StrTemp
  
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe lançamento no período informado!"
    txtDataIni.SetFocus
    Exit Sub
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.ScaleMode = vbMillimeters
  
  Dia = Now
  Largura = 190
  
  Cabeca2 Dia, Largura
  Altera = .Recordset(Quebra)
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = Format(SubTotal, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.NewPage
      Printer.CurrentY = 0
      Cabeca2 Dia, Largura
    End If
    If Quebra <> "data" Then
      If Altera <> .Recordset(Quebra) Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        StrTemp = Format(SubTotal, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.NewPage
        Printer.CurrentY = 0
        Cabeca2 Dia, Largura
        Altera = .Recordset(Quebra)
        SubTotal = 0
      End If
    End If
    
    StrTemp = .Recordset!DataCaixa
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!responsavel
    Printer.CurrentX = 40
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!TotalCombustivel, "Currency")
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!TotalProdutos, "Currency")
    Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    TempValor = 0
    TempValor = .Recordset!Juros
    StrTemp = Format(TempValor, "Currency")
    Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    TempValor = 0
    TempValor = .Recordset!TotalDespesas
    StrTemp = Format(TempValor, "Currency")
    Printer.CurrentX = 175 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    SubTotal = SubTotal + .Recordset!Diferenca
    StrTemp = Format(.Recordset!Diferenca, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(SubTotal, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
End With
Printer.EndDoc
NaoImprime:

End Sub

Private Sub Cabeca2(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.ScaleMode = vbMillimeters

StrTemp = "Relatório de Diferença de Caixa"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.CurrentY = 0
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

StrTemp = "Página " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Data: " & Format(Dia, "long date")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboVendedores.Text <> "" Then
  StrTemp = "Código: " & txtCodigo.Text & "   Funcionário: " & cboVendedores.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Printer.CurrentY = Printer.CurrentY + 1


StrTemp = "Data"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Responsavel"
Printer.CurrentX = 40
Printer.Print StrTemp;

StrTemp = "T. Comb."
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "T. Prod."
Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Juros"
Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Despesas"
Printer.CurrentX = 175 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Diferença"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Dim StrTemp As String

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.ScaleMode = vbMillimeters

StrTemp = "Relatório de Diferença de Caixa"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.CurrentY = 0
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

StrTemp = "Página " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Data: " & Format(Dia, "long date")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboVendedores.Text <> "" Then
  StrTemp = "Código: " & txtCodigo.Text & "   Funcionário: " & cboVendedores.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Printer.CurrentY = Printer.CurrentY + 1


StrTemp = "Data"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 25
Printer.Print StrTemp;

StrTemp = "Cod."
Printer.CurrentX = 49 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Nome"
Printer.CurrentX = 50
Printer.Print StrTemp;

StrTemp = "T. Vendas"
Printer.CurrentX = 135 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "T. Despesa"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "T. Rec."
Printer.CurrentX = 175 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Diferença"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub cboVendedores_LostFocus()
With dbVendedores
  .Refresh
  If cboVendedores.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboVendedores.Text & "'"
  If .Recordset.EOF = False Then
    cboVendedores.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!Codigo
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim strNovo As String, strNovo2 As String

strNovo = "select *from fechamentodecaixa where distribuido=-1 and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
strNovo2 = "select sum(diferenca) as total from fechamentodecaixa where distribuido=-1 and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"


If cboVendedores.Text <> "" Then
  Call cboVendedores_LostFocus
  strNovo = strNovo & " and codigoresponsavel=" & dbVendedores.Recordset!codigovendedor
  strNovo2 = strNovo2 & " and codigoresponsavel=" & dbVendedores.Recordset!codigovendedor
End If

strNovo = strNovo & strOrdem

dbFechamento.RecordSource = strNovo
dbFechamento.Refresh
qFechamento.RecordSource = strNovo2
qFechamento.Refresh

TempValor = 0
If IsNull(qFechamento.Recordset!Total) = False Then
  TempValor = qFechamento.Recordset!Total
End If

lblTotalDif.Caption = Format(TempValor, "Currency")

End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Quebra As String, Altera As String
Dim TotalParcial As Currency, Tota As Currency

With qFechamentos
  Quebra = .Recordset.Sort
  Call cmdExibir_Click
  .Recordset.Sort = Quebra
  
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe lançamento no período informado!"
    txtDataIni.SetFocus
    GoTo NaoImprime
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.ScaleMode = vbMillimeters
  
  If Right(Quebra, 4) = "desc" Then
    Quebra = Trim(Mid(Quebra, 1, Len(Quebra) - 4))
  End If
  
  Dia = Now
  Largura = 190
  
  Cabeca Dia, Largura
  Altera = .Recordset(Quebra)
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = Format(SubTotal, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.NewPage
      Printer.CurrentY = 0
      Cabeca Dia, Largura
    End If
    If Quebra <> "data" Then
      If Altera <> .Recordset(Quebra) Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        StrTemp = Format(SubTotal, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.NewPage
        Printer.CurrentY = 0
        Cabeca Dia, Largura
        Altera = .Recordset(Quebra)
        SubTotal = 0
      End If
    End If
    
    StrTemp = .Recordset!Data
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 25
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Codigoresponsavel
    Printer.CurrentX = 49 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 50
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!TotalVendas, "Currency")
    Printer.CurrentX = 135 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!TotalDespesa, "Currency")
    Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    TempValor = 0
    TempValor = .Recordset!TotalRecebimento + .Recordset!chequeavista
    StrTemp = Format(TempValor, "Currency")
    Printer.CurrentX = 175 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    SubTotal = SubTotal + .Recordset!Diferenca
    StrTemp = Format(.Recordset!Diferenca, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(SubTotal, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
End With
Printer.EndDoc

NaoImprime:
Imprime2
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If qFechamentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  qFechamentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  qFechamentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ",datacaixa, horaini" Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc,datacaixa, horaini"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ",datacaixa, horaini"
End If
Coluna = ColIndex
StrTemp = qFechamentos.Recordset.Sort
Call cmdExibir_Click
qFechamentos.Recordset.Sort = StrTemp
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
txtDataFim.Value = Date
txtDataIni.Value = DateAdd("m", -1, Date)
With dbPosto
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbVendedores
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

Private Sub txtCodigo_Change()
With dbVendedores
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.Find "codigo=" & txtCodigo.Text
    If .Recordset.EOF = False Then
      txtCodigo.Text = .Recordset!Codigo
      cboVendedores.Text = .Recordset!Nome
    End If
  End If
End With
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    txtCodigo.SetFocus
End Select
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    txtDataFim.SetFocus
End Select
End Sub
