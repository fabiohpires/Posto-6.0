VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatResgatados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valores Resgatados"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
   Icon            =   "frmRelatResgatados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbJurosBoleto 
      Caption         =   "dbJurosBoleto"
      Connect         =   "Access"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from JurosBoleto order by inicio, final"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtCodClientesNotas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtCodChequesClientes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from clientes order by nome"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbChequesClientes 
      Caption         =   "dbChequesClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from chequesclientes order by nome"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBCtls.DBCombo cboChequesClientes 
      Bindings        =   "frmRelatResgatados.frx":0442
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   10560
      TabIndex        =   34
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   9120
      Picture         =   "frmRelatResgatados.frx":0462
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin VB.Data qPendencias 
      Caption         =   "qPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from clientescobranca where pago=0 order by datafechamento"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data qCheques 
      Caption         =   "qCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clientescobranca"
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatResgatados.frx":0EE4
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmRelatResgatados.frx":0EFC
      TabIndex        =   14
      Top             =   1320
      Width           =   9855
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   38286
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   38286
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmRelatResgatados.frx":264F
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frmRelatResgatados.frx":266A
      TabIndex        =   15
      Top             =   4080
      Width           =   11415
   End
   Begin MSDBCtls.DBCombo cboClientesNotas 
      Bindings        =   "frmRelatResgatados.frx":38F5
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.Label lblJurosDevido 
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
      Left            =   9840
      TabIndex        =   36
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Juros A Cobrar:"
      Height          =   195
      Left            =   8640
      TabIndex        =   35
      Top             =   6360
      Width           =   1080
   End
   Begin VB.Label Label10 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Clientes de Nota:"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Código:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes de Cheques:"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Juros:"
      Height          =   255
      Left            =   5760
      TabIndex        =   33
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label txtTotalJuros 
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
      Left            =   6240
      TabIndex        =   32
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Total Recebido:"
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label txtTotalRecebido 
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
      Left            =   4320
      TabIndex        =   30
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Total Geral:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label txtTotal 
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
      Left            =   1680
      TabIndex        =   28
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Juros Cobrados:"
      Height          =   195
      Left            =   5835
      TabIndex        =   27
      Top             =   6360
      Width           =   1140
   End
   Begin VB.Label txtCobrancaJuros 
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
      Left            =   7080
      TabIndex        =   26
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Total Recebido:"
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label txtCobrancaRecebido 
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
      Left            =   4320
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Total em Cobrança:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label txtCobrancaTotal 
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
      Left            =   1680
      TabIndex        =   22
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Juros:"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label txtChequesJuros 
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
      Left            =   6360
      TabIndex        =   20
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Total Recebido:"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label txtChequesRecebido 
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
      Left            =   4440
      TabIndex        =   18
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Total em cheques:"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label txtChequesTotal 
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
      Left            =   1800
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmRelatResgatados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdemCheques As String, StrOrdemCobranca As String

Private Sub ExibeCheques()
Dim StrTemp As String
If cboChequesClientes.Text <> "" Then
  If dbChequesClientes.Recordset.EOF = False Then
    If cboChequesClientes.Text = dbChequesClientes.Recordset!Nome Then
      StrTemp = " and codigocliente=" & dbChequesClientes.Recordset!codigochequecliente
    End If
  End If
End If
With dbCheques
  .RecordSource = "select *from cheques where datapgto<>null and datapgto between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp & StrOrdemCheques
  .Refresh
End With
With qCheques
  .RecordSource = "select sum(valor) as Total, sum(valorpgto) as Recebido, sum(valorpgto - valor) as juros from cheques where datapgto<>null and datapgto between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    txtChequesTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    txtChequesTotal.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!recebido) = False Then
    txtChequesRecebido.Caption = Format(.Recordset!recebido, "Currency")
  Else
    txtChequesRecebido.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Juros) = False Then
    txtChequesJuros.Caption = Format(.Recordset!Juros, "Currency")
  Else
    txtChequesJuros.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub ExibeCobranca()
Dim StrTemp As String

If cboClientesNotas.Text <> "" Then
  If dbClientes.Recordset.EOF = False Then
    If cboClientesNotas.Text = dbClientes.Recordset!Nome Then
      StrTemp = " and codigocliente=" & dbClientes.Recordset!CodigoCliente
    End If
  End If
End If

With dbPendencias
  .RecordSource = "select *from clientescobranca where pago=-1 and datapagamento between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp & StrOrdemCobranca
  .Refresh
End With
With QPendencias
  .RecordSource = "select sum(valor) as Total, sum(valorpago) as Recebido, sum(valorpago - valor) as juros, sum(jurosDevido) as JurosValor from clientescobranca where pago=-1 and datapagamento between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrTemp
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    txtCobrancaTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    txtCobrancaTotal.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!recebido) = False Then
    txtCobrancaRecebido.Caption = Format(.Recordset!recebido, "Currency")
  Else
    txtCobrancaRecebido.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Juros) = False Then
    txtCobrancaJuros.Caption = Format(.Recordset!Juros, "Currency")
  Else
    txtCobrancaJuros.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!JurosValor) = False Then
    lblJurosDevido.Caption = Format(.Recordset!JurosValor, "Currency")
  Else
    lblJurosDevido.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cboChequesClientes_LostFocus()
With dbChequesClientes
  .Refresh
  If cboChequesClientes.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboChequesClientes.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodChequesClientes.Text = .Recordset!codigochequecliente
  End If
End With

End Sub

Private Sub cboClientesNotas_LostFocus()
With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboClientesNotas.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboClientesNotas.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodClientesNotas.Text = .Recordset!CodigoCliente
  End If
End With

End Sub

Private Sub cmdExibir_Click()

ExibeCheques
ExibeCobranca

txtTotal.Caption = Format(CCur(txtChequesTotal.Caption) + CCur(txtCobrancaTotal.Caption), "Currency")
txtTotalRecebido.Caption = Format(CCur(txtChequesRecebido.Caption) + CCur(txtCobrancaRecebido.Caption), "Currency")
txtTotalJuros.Caption = Format(CCur(txtChequesJuros.Caption) + CCur(txtCobrancaJuros.Caption), "Currency")

End Sub

Private Sub cmdImprime_Click()
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

Dim StrTemp As String
If cboChequesClientes.Text <> "" Then
  If dbChequesClientes.Recordset.EOF = False Then
    If cboChequesClientes.Text = dbChequesClientes.Recordset!Nome Then
      StrTemp = Chr(vbKeyReturn) & "Código:" & dbChequesClientes.Recordset!codigochequecliente & "   Nome: " & dbChequesClientes.Recordset!Nome
    End If
  End If
End If

If dbCheques.Recordset.RecordCount <> 0 Then
  ImprimeGrid DBGrid1, Printer, dbCheques, 7, False, , , 8, , "Cheques Resgatados" & Chr(vbKeyReturn) & NomePosto, "Período: " & txtDataIni.Value & " a " & txtDataFim.Value, "Impresso em : " & Format(Now, "long date") & " - " & Format(Now, "short time")
  Printer.CurrentX = 0
  Printer.Print "Juros: " & txtChequesJuros.Caption
  Printer.EndDoc
End If

StrTemp = ""
If cboClientesNotas.Text <> "" Then
  If dbClientes.Recordset.EOF = False Then
    If cboClientesNotas.Text = dbClientes.Recordset!Nome Then
      StrTemp = Chr(vbKeyReturn) & "Código:" & dbClientes.Recordset!CodigoCliente & "   Nome: " & dbClientes.Recordset!Nome
    End If
  End If
End If

If dbPendencias.Recordset.RecordCount <> 0 Then
  ImprimeGrid DBGrid2, Printer, dbPendencias, 3, False, , , 4, 5, "Cobranças Recebidas" & Chr(vbKeyReturn) & NomePosto, "Período: " & txtDataIni.Value & " a " & txtDataFim.Value, "Impresso em : " & Format(Now, "long date") & " - " & Format(Now, "short time"), 6
  Printer.CurrentX = 0
  Printer.Print "Juros: " & txtCobrancaJuros.Caption
  Printer.EndDoc
End If

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
Dim Ws As Workspace, db As Database, dbTemp As Recordset
Dim Juros As Currency, JurosValor As Currency
Dim DiasVencidos As Double, DiaDaSemana As Integer, FimDeSemana As Integer

StrOrdemCheques = " order by " & DBGrid1.Columns(0).DataField
StrOrdemCobranca = " order by " & DBGrid2.Columns(0).DataField
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbChequesClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbJurosBoleto
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
Set dbTemp = db.OpenRecordset("select *from clientescobranca where jurosdevido=null and pago=-1 and protestado=0")
If dbJurosBoleto.Recordset.RecordCount <> 0 Then
  If dbTemp.RecordCount <> 0 Then
    dbTemp.MoveLast
    dbTemp.MoveFirst
    Do While dbTemp.EOF = False
      JurosValor = 0
      Juros = 0
      FimDeSemana = 0
      DiasVencidos = DateDiff("d", dbTemp!DataFechamento, dbTemp!DataPagamento)
      DiaDaSemana = Weekday(dbTemp!DataFechamento)
      Select Case DiaDaSemana
        Case vbSunday 'domingo
          FimDeSemana = 1
        Case vbSaturday 'sabado
          FimDeSemana = 2
      End Select
      If DiasVencidos - FimDeSemana > 1 Then
        With dbJurosBoleto
          .Refresh
          If .Recordset.RecordCount <> 0 Then
            .Recordset.FindFirst "inicio<=" & DiasVencidos & " and final>=" & DiasVencidos
            If .Recordset.NoMatch = False Then
              If IsNull(.Recordset!JurosValor) = False Then
                JurosValor = .Recordset!JurosValor
              End If
              If .Recordset!Juros > 0 Then
                Juros = (.Recordset!Juros * dbTemp!Valor) * DiasVencidos
              End If
            End If
          End If
        End With
      End If
      JurosValor = JurosValor + Juros
      dbTemp.Edit
      dbTemp!JurosDevido = JurosValor
      dbTemp.Update
      dbTemp.MoveNext
    Loop
  End If
End If
Set dbTemp = Nothing
Set db = Nothing
Set Ws = Nothing

Call cmdExibir_Click
End Sub

Private Sub txtCodChequesClientes_GotFocus()
With txtCodChequesClientes
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodChequesClientes_LostFocus()
With dbChequesClientes
  .Refresh
  If IsNumeric(txtCodChequesClientes.Text) = False Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigochequecliente=" & txtCodChequesClientes.Text
  If .Recordset.NoMatch = False Then
    cboChequesClientes.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtCodClientesNotas_GotFocus()
With txtCodClientesNotas
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodClientesNotas_LostFocus()
With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodClientesNotas.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigocliente=" & txtCodClientesNotas.Text
  If .Recordset.NoMatch = False Then
    cboClientesNotas.Text = .Recordset!Nome
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
