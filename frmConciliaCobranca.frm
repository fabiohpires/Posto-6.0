VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaCobranca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobrança"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmConciliaCobranca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbConcilia 
      Height          =   330
      Left            =   3000
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select *from concilianova"
      Caption         =   "dbConcilia"
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
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   3000
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select *from contas"
      Caption         =   "dbContas"
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
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   3000
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "Select *from bloqueiafechamento"
      Caption         =   "dbBloqueiaFechamento"
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
   Begin VB.Data dbCobranca 
      Caption         =   "dbCobranca"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ClientesCobranca"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data dbStatus 
      Caption         =   "dbStatus"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Status"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecebe 
      Caption         =   "Receber"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from clientescobranca where pago=0 order by datafechamento"
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data QPendencias 
      Caption         =   "QPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from clientescobranca where pago=0"
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaCobranca.frx":0442
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "frmConciliaCobranca.frx":045D
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin MSComCtl2.DTPicker txtDtPagamento 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58064897
      CurrentDate     =   37664
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblTotalPendente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data do Pagamento"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1425
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   4560
      Width           =   405
   End
End
Attribute VB_Name = "frmConciliaCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoConta As Double, ValorAPagar As Currency

Private Sub cmdRecebe_Click()
Dim TempValor As Currency, Taxa As Double, Juros As Currency
Dim ValorRecebido As Currency, ValorDesconto As Currency
Dim Resposta As Integer, Diferenca As Currency
Dim CodigoCliente As Double, Valor As Currency, Obs As String

With dbBloqueiafechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtDtPagamento.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With
ValorAPagar = CalculaJurosBoleto(dbPendencias.Recordset!DataFechamento, txtDtPagamento.Value, dbPendencias.Recordset!Valor)
TempValor = ValorAPagar - CCur(TxtValor.Text)
If txtDtPagamento.Value > dbPendencias.Recordset!DataFechamento Then
  If ValorAPagar > CCur(TxtValor.Text) Then
    Obs = ""
    Do While Obs = ""
      Obs = InputBox("Valor recebido está abaixo do valor com juros calculado. Justifique o motivo!", "Juros Calculado")
      If Obs = "" Then
        Resposta = MsgBox("Você deve descrever o motivo! Deseja continuar?", vbYesNo)
        If Resposta = vbNo Then Exit Sub
      End If
    Loop
  End If
End If
If dbPendencias.Recordset.RecordCount = 0 Then Exit Sub
If dbPendencias.Recordset.EOF = True Then
  MsgBox "Escolha uma cobrança primeiro!"
  Exit Sub
End If

'If DateDiff("d", Date, txtDtPagamento.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode receber boleto com data futura!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtDtPagamento.Value) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode receber boleto com data anterior a 10 dias!"
'    Exit Sub
'  End If
'End If
Resposta = MsgBox("Deseja receber o valor atual?", vbYesNo + vbDefaultButton2, App.Title)
If Resposta = vbNo Then Exit Sub

If IsNumeric(TxtValor.Text) = False Then
  MsgBox "Valor inválido!"
  TxtValor.SetFocus
  Exit Sub
End If
ValorRecebido = CCur(TxtValor.Text)
Diferenca = dbPendencias.Recordset!Valor - ValorRecebido


If ValorRecebido < CCur(dbPendencias.Recordset!Valor) Then
  MsgBox "Valor inferior!"
  Permissao = False
  frmPermissao.Show vbModal
  If Permissao = False Then Exit Sub
End If

With DbClientes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codigocliente=" & dbPendencias.Recordset!CodigoCliente
    If .Recordset.NoMatch = False Then
      If .Recordset!mensalista = False Then
        CodigoCliente = .Recordset!CodigoCliente
      Else
        CodigoCliente = 0
      End If
    End If
  End If
End With

With dbConcilia
  .Recordset.AddNew
  .Recordset!CodigoConta = CodigoConta
  .Recordset!DataLanc = Now
  .Recordset!Data = txtDtPagamento.Value
  .Recordset!compensado = True
  .Recordset!Tipo = "Cobranca"
  .Recordset!Codigo = 999999997
  .Recordset!Descri = Left(dbPendencias.Recordset!Cliente, 50)
  .Recordset!NrDocumento = dbPendencias.Recordset!CodigoCobranca
  .Recordset!Valor = ValorRecebido
  .Recordset.Update
End With

With dbContas
  .Refresh
  If .Recordset.EOF = False And .Recordset.BOF = False Then
    .Recordset.MoveFirst
    .Recordset.Find "codigoconta=" & CodigoConta
    .Recordset!Saldo = .Recordset!Saldo + ValorRecebido
    .Recordset.Update
  Else
    MsgBox "Erro na tabela de contas"
  End If
End With

'With dbStatus
'  .Refresh
'  .Recordset.Edit
'  .Recordset!difcheques = .Recordset!difcheques + Diferenca
'  .Recordset.Update
'  .Refresh
'End With

With dbPendencias
  .Recordset.Edit
  .Recordset!Pago = True
  .Recordset!valorpago = ValorRecebido
  .Recordset!DataPagamento = txtDtPagamento.Value
  .Recordset!CodigoFormadePg = CodigoConta
  .Recordset!Descri = dbContas.Recordset!Descri
  Juros = ValorRecebido - .Recordset!Valor
  Valor = .Recordset!Valor
  .Recordset!Juros = ValorRecebido - .Recordset!Valor
  If IsNull(.Recordset!Obs) = False Then
    .Recordset!Obs = .Recordset!Obs & Obs
  Else
    .Recordset!Obs = Obs
  End If
  .Recordset!fechames = False
  .Recordset.Update
  .Refresh
End With
With DbClientes
  .Recordset.Edit
  .Recordset!TotalBoleto = .Recordset!TotalBoleto - Valor
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With
dbPendencias.Refresh
qPendencias.Refresh
If CodigoCliente <> 0 Then
  With dbCobranca
    .RecordSource = "select *from clientescobranca where codigocliente=" & CodigoCliente & " and pago=0 and datafechamento<#" & DataInglesa(Date) & "#"
    .Refresh
    If .Recordset.RecordCount = 0 Then
      DbClientes.Recordset.Edit
      DbClientes.Recordset!mensalista = True
      DbClientes.Recordset.Update
    End If
  End With
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbPendencias_Reposition()
On Error Resume Next
TxtValor.Text = Format(dbPendencias.Recordset!Valor, "Currency")
txtDtPagamento.Value = dbPendencias.Recordset!DataFechamento
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case keycascii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
txtDtPagamento.Value = Date
With dbPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With DbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCobranca
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiafechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbConcilia
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub

Private Sub txtDtPagamento_Change()
If dbPendencias.Recordset.EOF = False And dbPendencias.Recordset.BOF = False Then
  ValorAPagar = CalculaJurosBoleto(dbPendencias.Recordset!DataFechamento, txtDtPagamento.Value, dbPendencias.Recordset!Valor)
  TxtValor.Text = Format(ValorAPagar, "Currency")
End If
End Sub

