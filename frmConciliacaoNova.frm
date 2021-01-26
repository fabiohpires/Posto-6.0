VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliacaoNova 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliação Bancária"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   Icon            =   "frmConciliacaoNova.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelaBaixa 
      Caption         =   "Cancela Baixa"
      Height          =   495
      Left            =   2640
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   1680
      Top             =   4440
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
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   3720
      TabIndex        =   31
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCobranca 
      Caption         =   "Cobrança"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Data qSaldo 
      Caption         =   "qSaldo"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DespesasLanc2"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "Histórico da Conta"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCartoesRecebidos 
      Caption         =   "Cartões Recebidos"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheques 
      Caption         =   "Cheques"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdCustodia 
      Caption         =   "Custódia"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Data dbDespesalanc 
      Caption         =   "dbDespesalanc"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DespesasLanc2"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   2880
      Picture         =   "frmConciliacaoNova.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "Imprimir"
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdCartao 
      Caption         =   "Cartões"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtNrDocumento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.Data qSaldoBanco 
      Caption         =   "qSaldoBanco"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as saldo from conciliaNova where compensado=-1 and codigoconta=0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo cboDespesa 
      Bindings        =   "frmConciliacaoNova.frx":0EC4
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   855
   End
   Begin VB.Data dbDespesa 
      Caption         =   "dbDespesa"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contasdespesas order by descri"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ConciliaNova where codigoconta=0 order by compensado, data, datalanc"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmConciliacaoNova.frx":0EDC
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliacaoNova.frx":0EF3
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frmConciliacaoNova.frx":0F0C
      TabIndex        =   18
      Top             =   1440
      Width           =   12375
   End
   Begin MSComCtl2.DTPicker txtDataLanc 
      Height          =   315
      Left            =   4920
      TabIndex        =   14
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   117899265
      CurrentDate     =   37257
   End
   Begin MSComCtl2.DTPicker txtDataImprime 
      Height          =   315
      Left            =   1320
      TabIndex        =   20
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   117899265
      CurrentDate     =   37257
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo no Dia selecionado:"
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label lblSaldoDia 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   28
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Imprimir a partir de:"
      Height          =   195
      Left            =   1320
      TabIndex        =   19
      Top             =   6000
      Width           =   1320
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Documento:"
      Height          =   195
      Left            =   6240
      TabIndex        =   15
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lblSaldoBanco 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo do Banco:"
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblSaldoSistema 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo do Sistema:"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Lançamento:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qtd.:"
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lançamento:"
      Height          =   195
      Left            =   4920
      TabIndex        =   13
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Conta:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmConciliacaoNova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Extrato de Conta"
Printer.CurrentY = 0
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = Format(Dia, "Short date") & " - " & Format(Dia, "Short time")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = cboConta.Text
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "A partir de: " & Format(txtDataImprime.Value, "Short date")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Lançado"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Compensa"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 40
Printer.Print StrTemp;

StrTemp = "Nr. Documento"
Printer.CurrentX = 120
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1
End Sub

Private Sub cboConta_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConta_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub CboConta_LostFocus()
Me.KeyPreview = True
With dbContas
  .Refresh
  If cboConta.Text = "" Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.FindFirst "descri='" & cboConta.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboConta.Text = .Recordset!Descri
End With
End Sub

Private Sub cboDespesa_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboDespesa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboDespesa_LostFocus()
Me.KeyPreview = True
With dbDespesa
  .Refresh
  If cboDespesa.Text = "" Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.FindFirst "descri='" & cboDespesa.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboDespesa.Text = .Recordset!Descri
End With

End Sub

Private Sub cmdCancelaBaixa_Click()
Select Case dbConcilia.Recordset!Codigo
    Case "999999997" 'cobrança de clientes
        ExtornaCobranca dbConcilia.Recordset!codigoconciliaconta
End Select
End Sub

Private Sub cmdCartao_Click()
Load frmConciliacaoCartao
frmConciliacaoCartao.CodigoConta = dbContas.Recordset!CodigoConta
frmConciliacaoCartao.Atualiza
frmConciliacaoCartao.Show vbModal
Load frmEstatus2
Unload frmEstatus2
Call CboConta_LostFocus
Call cmdExibir_Click
End Sub

Private Sub cmdCartoesRecebidos_Click()
Load frmCartoesRecebidos
frmCartoesRecebidos.CodigoConta = dbContas.Recordset!CodigoConta
frmCartoesRecebidos.ExibeCartoes
frmCartoesRecebidos.Show vbModal
End Sub

Private Sub cmdCheques_Click()
Load frmConciliaNovaCheques
frmConciliaNovaCheques.CodigoConta = dbContas.Recordset!CodigoConta
frmConciliaNovaCheques.AbreDados
frmConciliaNovaCheques.Show vbModal
Load frmEstatus2
Unload frmEstatus2
Call cmdExibir_Click
End Sub

Private Sub cmdCobranca_Click()
Load frmConciliaCobranca
frmConciliaCobranca.CodigoConta = dbContas.Recordset!CodigoConta
frmConciliaCobranca.Show vbModal
Load frmEstatus2
Unload frmEstatus2
Call cmdExibir_Click
End Sub

Private Sub cmdCustodia_Click()
Load frmConciliacaoCustodia
frmConciliacaoCustodia.CodigoConta = dbContas.Recordset!CodigoConta
frmConciliacaoCustodia.Show vbModal
Load frmEstatus2
Unload frmEstatus2
Call cmdExibir_Click
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String

If dbContas.Recordset.EOF = True Then
  MsgBox "Escolha uma conta!"
  Exit Sub
End If
If cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Escolha uma conta!"
  cboConta.SetFocus
  Exit Sub
End If
If IsNull(dbContas.Recordset!Saldo) = False Then
  lblSaldoSistema.Caption = Format(dbContas.Recordset!Saldo, "Currency")
Else
  lblSaldoSistema.Caption = Format(0, "Currency")
End If
With dbConcilia
  .RecordSource = "select *from concilianova where codigoconta=" & dbContas.Recordset!CodigoConta & " and compensado=-1 order by compensado, data, datalanc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "compensado=0"
    If .Recordset.NoMatch = True Then
      .Recordset.MoveLast
    End If
  End If
End With

With qSaldoBanco
  .RecordSource = "select sum(valor) as saldo from conciliaNova where compensado=-1 and codigoconta=" & dbContas.Recordset!CodigoConta
  .Refresh
  If IsNull(.Recordset!Saldo) = False Then
    lblSaldoBanco.Caption = Format(.Recordset!Saldo, "Currency")
  Else
    lblSaldoBanco.Caption = Format(0, "Currency")
  End If
End With

End Sub

Private Sub cmdExportar_Click()
Dim StrTemp As String

With dbConcilia
  .RecordSource = "Select *from concilianova where codigoconta=" & dbContas.Recordset!CodigoConta & " and data>=#" & DataInglesa(Trim(Str(txtDataImprime.Value))) & "# and compensado=-1 order by data"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Open "Concilia.txt" For Output As #1
    StrTemp = "Lancamento|Compensado|Descrição|Nr. Documento|Valor"
    Print #1, StrTemp
    Do While .Recordset.EOF = False
      StrTemp = .Recordset!DataLanc & "|" & .Recordset!Data & "|" & .Recordset!Descri & "|" & .Recordset!NrDocumento & "|" & .Recordset!Valor
      Print #1, StrTemp
      .Recordset.MoveNext
    Loop
    Close #1
  End If
End With
Call cmdExibir_Click
End Sub

Private Sub cmdHistorico_Click()
frmConciliaHistorico.Show vbModal
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim SqlInicial As String, Total As Currency

If dbContas.Recordset.EOF = True Then Exit Sub
If cboConta.Text = "" Then Exit Sub


With dbConcilia
  SqlInicial = .RecordSource
  .RecordSource = "Select *from concilianova where codigoconta=" & dbContas.Recordset!CodigoConta & " and data>=#" & DataInglesa(Trim(Str(txtDataImprime.Value))) & "# and compensado=-1 order by data"
  .Refresh
  
  If .Recordset.RecordCount = 0 Then
    GoTo Pendencias
  End If
  
  On Error GoTo Pendencias
  If ShowPrinter(Me) = 0 Then GoTo Pendencias
  On Error GoTo 0
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Página: " & Printer.Page
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    
    If IsNull(.Recordset!DataLanc) = False Then
      StrTemp = Format(.Recordset!DataLanc, "Short date")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!Data) = False Then
      StrTemp = Format(.Recordset!Data, "Short date")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!Descri) = False Then
      StrTemp = .Recordset!Descri
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 40
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!NrDocumento) = False Then
      StrTemp = .Recordset!NrDocumento
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 120
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!Valor) = False Then
      StrTemp = Format(.Recordset!Valor, "Currency")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Saldo do Banco: " & lblSaldoBanco.Caption
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Saldo do Sistema: " & lblSaldoSistema.Caption
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
  
Pendencias:

  On Error GoTo NaoImprime
  Resposta = MsgBox("Deseja imprimir as pendências?", vbYesNo)
  If Resposta = vbNo Then GoTo NaoImprime
  
  .RecordSource = "Select *from concilianova where codigoconta=" & dbContas.Recordset!CodigoConta & " and compensado=0 order by datalanc"
  .Refresh
  
  If .Recordset.RecordCount = 0 Then GoTo NaoImprime
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then GoTo NaoImprime
  On Error GoTo 0
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Página: " & Printer.Page
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    
    If IsNull(.Recordset!DataLanc) = False Then
      StrTemp = Format(.Recordset!DataLanc, "Short date")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!Descri) = False Then
      StrTemp = .Recordset!Descri
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 40
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!NrDocumento) = False Then
      StrTemp = .Recordset!NrDocumento
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 120
    Printer.Print StrTemp;
    
    Total = Total + .Recordset!Valor
    If IsNull(.Recordset!Valor) = False Then
      StrTemp = Format(.Recordset!Valor, "Currency")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With


NaoImprime:
  dbConcilia.RecordSource = SqlInicial
  Call cmdExibir_Click
End Sub

Private Sub cmdIncluir_Click()

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtDataLanc.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

'If DateDiff("d", Date, txtDataLanc.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar despesa bancária com data futura!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtDataLanc.Value) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar despesa bancária com data anterior a 10 dias!"
'    Exit Sub
'  End If
'End If

If cboDespesa.Text = "" Then
  MsgBox "Selecione um tipo de lançamento!"
  cboDespesa.SetFocus
  Exit Sub
End If
If dbDespesa.Recordset.EOF = True Then
  MsgBox "Selecione um tipo de lançamento!"
  cboDespesa.SetFocus
  Exit Sub
End If
If cboDespesa.Text <> dbDespesa.Recordset!Descri Then
  MsgBox "Selecione um tipo de lançamento!"
  cboDespesa.SetFocus
  Exit Sub
End If
If IsNumeric(lblTotal.Caption) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If

With dbConcilia
  .Recordset.AddNew
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  .Recordset!compensado = True
  .Recordset!Data = txtDataLanc.Value
  .Recordset!Tipo = "Conciliação"
  .Recordset!Codigo = 999999999
  .Recordset!Descri = cboDespesa.Text
  .Recordset!NrDocumento = txtNrDocumento.Text & " "
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset.Update
End With
With dbDespesaLanc
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Conciliação"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!Conta = cboConta.Text
  .Recordset!CodigoDespesa = dbDespesa.Recordset!codigolancamento
  .Recordset!NrDocumento = txtNrDocumento.Text
  .Recordset!Descri = Left(dbDespesa.Recordset!Descri, 50)
  .Recordset!Obs = txtNrDocumento.Text
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset!valorpago = CCur(lblTotal.Caption)
  .Recordset!compensado = True
  .Recordset!fechamentodiario = True
  .Recordset!Produto = False
  .Recordset!codigoenviar = "1"
  .Recordset.Update
End With
With dbContas
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo + CCur(lblTotal.Caption)
  .Recordset.Update
End With
Call CboConta_LostFocus
Call cmdExibir_Click
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbConcilia_Reposition()
Dim CurSaldo As Currency
If dbConcilia.Recordset.EOF = True Or dbConcilia.Recordset.BOF = True Then
  lblSaldoDia.Caption = Format("0", "currency")
  Exit Sub
End If
With qSaldo
  On Error Resume Next
  .RecordSource = "select sum(valor) as saldo from concilianova where compensado=-1 and codigoconta=" & dbConcilia.Recordset!CodigoConta & " and data<=#" & DataInglesa(dbConcilia.Recordset!Data) & " 23:59:59#"
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Saldo) = False Then
    lblSaldoDia.Caption = Format(.Recordset!Saldo, "currency")
  Else
    lblSaldoDia.Caption = Format("0", "currency")
  End If
End With
End Sub

Private Sub DBGrid1_DblClick()
With dbConcilia
  If .Recordset!Descri = "Depósito de cheques!" Or .Recordset!Descri = "Custódia de cheques!" Then
    Load frmConciliaChequesDepositados
    With frmConciliaChequesDepositados
      .Tipo = dbConcilia.Recordset!Descri
      .Dia = dbConcilia.Recordset!Data
      .Valor = dbConcilia.Recordset!Valor
      .NrDocumento = dbConcilia.Recordset!NrDocumento
      .Exibe
      .Show vbModal
    End With
  End If
End With
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
txtDataLanc.Value = Date
txtDataImprime.Value = Date

With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesa
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qSaldoBanco
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaLanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qSaldo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

Select Case Usuarios.Grupo.BancoConcilia
  Case 1 'Somente leitura
    cmdCartao.Enabled = False
    cmdCustodia.Enabled = False
    cmdIncluir.Enabled = False
    cmdCompensar.Enabled = False
  Case 2 'Liberado
    
End Select
If Usuarios.Nome = "Usuário Master" Then
    cmdCancelaBaixa.Visible = True
End If
End Sub

Private Sub txtDataImprime_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataImprime_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataImprime_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataLanc_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataLanc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataLanc_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_LostFocus()
Dim Total As Currency
lblTotal.Caption = ""
With dbDespesa
  If .Recordset.EOF = True Then Exit Sub
  If cboDespesa.Text <> .Recordset!Descri Then Exit Sub
  If IsNumeric(txtValor.Text) = False Then Exit Sub
  Total = CDbl(txtValor.Text) * .Recordset!Valor
  lblTotal.Caption = Format(Total, "Currency")
End With
End Sub
