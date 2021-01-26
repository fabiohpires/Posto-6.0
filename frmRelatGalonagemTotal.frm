VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmRelatGalonagemTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Galonagem"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmRelatGalonagemTotal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOrdemTurno 
      Caption         =   "Ordem por turno"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin MSDBCtls.DBCombo cboTurnos 
      Bindings        =   "frmRelatGalonagemTotal.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "dbfs"
      Height          =   2295
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Data dbTurnos 
         Caption         =   "dbTurnos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Turnos order by descri"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Data qVendasCaixa 
         Caption         =   "qVendasCaixa"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(vendas) as total from BicoEncerrantes group by codigofechamento, codigoproduto"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Data dbFechamento 
         Caption         =   "dbFechamento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FechamentoDeCaixa order by datacaixa, horaini"
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADODC"
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   3015
      Begin MSAdodcLib.Adodc dbProdutos 
         Height          =   330
         Left            =   120
         Top             =   360
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
         RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
         Caption         =   "dbProdutos"
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
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   3720
      Picture         =   "frmRelatGalonagemTotal.frx":0459
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55181313
      CurrentDate     =   37665
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55181313
      CurrentDate     =   37665
   End
   Begin MSComctlLib.ProgressBar Porcento 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label7 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatGalonagemTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PosTotal As Double

Private Sub Galonagem()
Dim StrTemp As String, Largura As Double, Dia As Date
Dim CodigoFechamento As Double, SubTotal As Double, Total As Double
Dim Totais() As Double, Produto As Integer, Colunas As Integer
Dim StrTurno As String, StrQuebra As String
Dim SubTotal2() As Double, Totais2() As Double


If cboTurnos.Text <> "" Then
  With dbTurnos
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.FindFirst "descri='" & cboTurnos.Text & "'"
      If .Recordset.NoMatch = False Then
        StrTurno = " and codigoturno=" & .Recordset!codigoturno
      Else
        StrTurno = ""
      End If
    End If
  End With
End If

With dbProdutos
  .RecordSource = "Select *from produtos where combustivel=-1"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe nenhum tipo de combustível cadastrado!"
    Exit Sub
  End If
  .Recordset.Sort = "codigoproduto"
  
  ReDim Totais(.Recordset.RecordCount)
  ReDim SubTotal2(.Recordset.RecordCount)
  ReDim Totais2(.Recordset.RecordCount)
  
  Colunas = .Recordset.RecordCount
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  dbFechamento.RecordSource = "select *from fechamentodecaixa where datacaixa between #" & DataInglesa(Trim(Str(txtDataini.Value))) & "# and #" & DataInglesa(Trim(Str(txtDatafim.Value))) & "# and fechado=-1" & StrTurno
  If chkOrdemTurno.Value = vbChecked Then
    dbFechamento.RecordSource = dbFechamento.RecordSource & " order by turno, datacaixa, horaini"
  Else
    dbFechamento.RecordSource = dbFechamento.RecordSource & " order by datacaixa, horaini"
  End If
  
  dbFechamento.Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "O período retornou um resultado vazio!"
    txtDataini.SetFocus
    Exit Sub
  End If
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  If .Recordset.RecordCount > 6 Then
    Printer.Orientation = vbPRORLandscape
  Else
    Printer.Orientation = vbPRORPortrait
  End If
  Dia = Now
  Largura = Printer.ScaleWidth
End With

Cabeca Largura, Dia
Y1 = Printer.CurrentY

qVendasCaixa.Refresh

For i = 0 To Colunas
  Totais2(i) = 0
Next i

With dbFechamento
  If .Recordset.RecordCount = 0 Then
    MsgBox "O período retornou um resultado vazio!"
    txtDataini.SetFocus
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Porcento.Max = .Recordset.RecordCount
  Porcento.Min = 0
  Porcento.Value = .Recordset.AbsolutePosition
  With qVendasCaixa
    If txtDataini.Value > txtDatafim.Value Then
      MsgBox "A data inicial deve ser menor que a data final!"
      txtDataini.SetFocus
      Exit Sub
    End If
    .Refresh
  End With
  
  For i = 0 To Colunas
    SubTotal2(i) = 0
  Next i
  
  Do While .Recordset.EOF = False
    Porcento.Value = .Recordset.AbsolutePosition
    Y1 = Printer.CurrentY
    Y2 = Y1 + 4
    qVendasCaixa.RecordSource = "select sum(vendas) as total, codigofechamento, codigoproduto from BicoEncerrantes where CodigoFechamento=" & .Recordset!CodigoFechamento & " group by codigofechamento, codigoproduto"
    qVendasCaixa.Refresh
    Printer.FillStyle = vbFSTransparent
    Printer.Line (0, Y1)-(35, Y2), , B
    Printer.Line (35, Y1)-(50, Y2), , B
    X1 = 50
    For i = 1 To dbProdutos.Recordset.RecordCount + 1
      Printer.Line (X1, Y1)-(X1 + 21, Y2), , B
      X1 = X1 + 21
    Next i
    
    Printer.CurrentY = Y1 + 0.5
    
    StrTemp = Format(.Recordset!DataCaixa, "dd/mm/yy - ddd")
    Printer.CurrentX = 1
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 36
    Printer.Print StrTemp;
    X1 = 70
    With dbProdutos
      .Recordset.MoveFirst
      Produto = 1
      Do While .Recordset.EOF = False
        qVendasCaixa.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
        If qVendasCaixa.Recordset.NoMatch = False Then
          SubTotal = SubTotal + qVendasCaixa.Recordset!Total
          Totais(Produto) = Totais(Produto) + qVendasCaixa.Recordset!Total
          Totais2(Produto) = Totais2(Produto) + qVendasCaixa.Recordset!Total
          SubTotal2(Produto) = SubTotal2(Produto) + qVendasCaixa.Recordset!Total
          StrTemp = Format(qVendasCaixa.Recordset!Total, "#,##0.0")
          Printer.CurrentX = X1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
        End If
        .Recordset.MoveNext
        X1 = X1 + 21
        Produto = Produto + 1
      Loop
    End With
    StrTemp = Format(SubTotal, "#,##0.0")
    Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    SubTotal = 0
    Y1 = Y1 + 4
    Printer.CurrentY = Y1
    If Printer.CurrentY > Printer.ScaleHeight - 20 Then
      X1 = 70
      For i = 1 To Colunas
        Total = Total + SubTotal2(i)
        StrTemp = Format(SubTotal2(i), "#,##0.0")
        Printer.CurrentX = X1 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        X1 = X1 + 21
      Next i
      StrTemp = Format(Total, "#,##0.0")
      Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.NewPage
      Y1 = 0
      Printer.CurrentY = 0
      Cabeca Largura, Dia
    End If
    
    If chkOrdemTurno.Value = vbChecked Then
      StrQuebra = .Recordset!Turno
    End If
    
    .Recordset.MoveNext
    If .Recordset.EOF = False Then
      If chkOrdemTurno.Value = vbChecked Then
        If StrQuebra <> .Recordset!Turno Then
          X1 = 70
          For i = 1 To Colunas
            Total = Total + SubTotal2(i)
            StrTemp = Format(SubTotal2(i), "#,##0.0")
            Printer.CurrentX = X1 - Printer.TextWidth(StrTemp)
            Printer.Print StrTemp;
            X1 = X1 + 21
          Next i
          StrTemp = Format(Total, "#,##0.0")
          Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp
          
          For i = 0 To Colunas
            SubTotal2(i) = 0
          Next i
          Cabeca2 Largura, Dia
        End If
      End If
    End If
  Loop
End With

If chkOrdemTurno.Value = vbChecked Then
  X1 = 70
  Total = 0
  For i = 1 To Colunas
    Total = Total + SubTotal2(i)
    StrTemp = Format(SubTotal2(i), "#,##0.0")
    Printer.CurrentX = X1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    X1 = X1 + 21
  Next i
  StrTemp = Format(Total, "#,##0.0")
  Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
End If

Total = 0
X1 = 70
For i = 1 To Colunas
  Total = Total + Totais(i)
  StrTemp = Format(Totais(i), "#,##0.0")
  Printer.CurrentX = X1 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  X1 = X1 + 21
Next i
StrTemp = Format(Total, "#,##0.0")
Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.EndDoc
Porcento.Value = 0
naoImprime:

End Sub

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String, X1 As Double, X2 As Double, Y1 As Double, Y2 As Double

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

Printer.FillColor = RGB(175, 175, 175)
Printer.FillStyle = vbFSSolid
Printer.Line (0, 0)-(Largura, 10), , B

Printer.FillStyle = vbFSTransparent

Printer.CurrentY = 2

Printer.FontSize = 8
Printer.FontBold = False
StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
Printer.Print StrTemp;

Printer.FontSize = 14
StrTemp = "Relatório de Galonagem Total"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

Printer.CurrentY = 12

StrTemp = "Posto: " & NomePosto
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataini.Value, "short date") & " a " & Format(txtDatafim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboTurnos.Text <> "" Then
  StrTemp = "Turno: " & cboTurnos.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Y1 = Printer.CurrentY + 0.5

Y2 = Printer.CurrentY + 4

Printer.FillColor = RGB(175, 175, 175)
Printer.FillStyle = vbFSSolid
Printer.Line (0, Y1)-(35, Y2), , B
Printer.Line (35, Y1)-(50, Y2), , B
X1 = 50
For i = 1 To dbProdutos.Recordset.RecordCount + 1
  Printer.Line (X1, Y1)-(X1 + 21, Y2), , B
  X1 = X1 + 21
Next i

PosTotal = X1

Printer.FillStyle = vbFSTransparent

Printer.CurrentY = Y1 + 0.5


StrTemp = "Data"
Printer.CurrentX = 1
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 36
Printer.Print StrTemp;
x = 70
With dbProdutos
  .Refresh
  .Recordset.Sort = "codigoproduto"
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If IsNull(.Recordset!DescriAbreviada) = False Then
      If .Recordset!DescriAbreviada <> "" Then
        StrTemp = .Recordset!DescriAbreviada
      Else
        StrTemp = .Recordset!Descri
      End If
    Else
      StrTemp = .Recordset!Descri
    End If
    Printer.CurrentX = x - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    .Recordset.MoveNext
    x = x + 21
  Loop
End With
StrTemp = "Total"
Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
Printer.CurrentY = Printer.CurrentY + 4


End Sub

Private Sub Cabeca2(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String, X1 As Double, X2 As Double, Y1 As Double, Y2 As Double

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

Y1 = Printer.CurrentY + 0.5

Y2 = Printer.CurrentY + 4

Printer.FillColor = RGB(175, 175, 175)
Printer.FillStyle = vbFSSolid
Printer.Line (0, Y1)-(35, Y2), , B
Printer.Line (35, Y1)-(50, Y2), , B
X1 = 50
For i = 1 To dbProdutos.Recordset.RecordCount + 1
  Printer.Line (X1, Y1)-(X1 + 21, Y2), , B
  X1 = X1 + 21
Next i

PosTotal = X1

Printer.FillStyle = vbFSTransparent

Printer.CurrentY = Y1 + 0.5


StrTemp = "Data"
Printer.CurrentX = 1
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 36
Printer.Print StrTemp;
x = 70
With dbProdutos
  .Refresh
  .Recordset.Sort = "codigoproduto"
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    StrTemp = .Recordset!Descri
    Printer.CurrentX = x - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    .Recordset.MoveNext
    x = x + 21
  Loop
End With
StrTemp = "Total"
Printer.CurrentX = PosTotal - 1 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
Printer.CurrentY = Printer.CurrentY + 4


End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImprime_Click()

Galonagem

End Sub

Private Sub Form_Load()
txtDataini.Value = DateAdd("m", -1, Date)
txtDatafim.Value = Date

With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos where combustivel=-1 order by descri"
  .Refresh
End With
With dbFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendasCaixa
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTurnos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

