VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControlePagamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento de Funcionários"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNaoIdentificadas 
      Caption         =   "Comissões não identificadas"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdImprimeFolha 
      Caption         =   "Imprime Folha de Pagamento"
      Height          =   255
      Left            =   7320
      TabIndex        =   37
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   6000
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtFechadoAte 
      Height          =   300
      Left            =   7440
      TabIndex        =   9
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   39435
   End
   Begin VB.TextBox txtCodFuncionario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtMesAno 
      Height          =   300
      Left            =   4560
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MM/yyyy"
      Format          =   72941571
      CurrentDate     =   39427
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   8160
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Data dbDespesaTipo 
         Caption         =   "dbDespesaTipo"
         Connect         =   "Access"
         DatabaseName    =   "d:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesaTipo"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Data dbDespesaLanc 
         Caption         =   "dbDespesaLanc"
         Connect         =   "Access 2000;"
         DatabaseName    =   "d:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesasLanc2"
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Data qVendasTotal 
         Caption         =   "qVendasTotal"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QVendas"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Data qValesTotal 
         Caption         =   "qValesTotal"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select QVales.*, fechamentodecaixa.* from qvales, fechamentodecaixa where qvales.codigocaixa=fechamentodecaixa.codigofechamento"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Data dbVales 
         Caption         =   "dbVales"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Vales"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Data qVales 
         Caption         =   "qVales"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QValesCaixa"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Data dbVenda 
         Caption         =   "dbVenda"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Venda2"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Data qVendas 
         Caption         =   "qVendas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "d:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno"
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Data dbPagamentos 
         Caption         =   "dbPagamentos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from vendedorespagamento"
         Top             =   720
         Width           =   3255
      End
      Begin VB.Data dbFuncionarios 
         Caption         =   "dbFuncionarios"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Vendedores order by nome"
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.OptionButton optAdiantamento 
      Caption         =   "Adiantamento"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton optSalario 
      Caption         =   "Salário"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Pagamento"
      TabPicture(0)   =   "frmControlePagamento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "framePagamento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdConfirma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelaPagamento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdImprime2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "grdFolha"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRemover"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Vales"
      TabPicture(1)   =   "frmControlePagamento.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdVales"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "lblValesTotal"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Comissões"
      TabPicture(2)   =   "frmControlePagamento.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(1)=   "lblTotalComissao"
      Tab(2).Control(2)=   "lblTotalQtd"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "grdComissoes"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdRemover 
         Caption         =   "Remover"
         Height          =   375
         Left            =   6240
         TabIndex        =   41
         Top             =   4200
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid grdFolha 
         Bindings        =   "frmControlePagamento.frx":0054
         Height          =   3255
         Left            =   4080
         OleObjectBlob   =   "frmControlePagamento.frx":006F
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.CommandButton cmdImprime2 
         Height          =   615
         Left            =   240
         Picture         =   "frmControlePagamento.frx":146F
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "Imprimir"
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelaPagamento 
         Caption         =   "Cancela Pagamento"
         Height          =   375
         Left            =   8040
         TabIndex        =   35
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Sair"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Frame framePagamento 
         Height          =   2655
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   3855
         Begin VB.TextBox txtSalario 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtVR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtVT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   16
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Salário / Adiantamento:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Vales:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblVales 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Comissões:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblComissoes 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   26
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Vale Refeição:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Vale Transporte:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo a Pagar:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblSaldoAPagar 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            TabIndex        =   24
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   3720
            Y1              =   2040
            Y2              =   2040
         End
      End
      Begin MSDBGrid.DBGrid grdVales 
         Bindings        =   "frmControlePagamento.frx":1EF1
         Height          =   3735
         Left            =   -74880
         OleObjectBlob   =   "frmControlePagamento.frx":1F06
         TabIndex        =   19
         Top             =   480
         Width           =   9855
      End
      Begin MSDBGrid.DBGrid grdComissoes 
         Bindings        =   "frmControlePagamento.frx":2E02
         Height          =   3735
         Left            =   -74880
         OleObjectBlob   =   "frmControlePagamento.frx":2E18
         TabIndex        =   20
         Top             =   480
         Width           =   9855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Qtd.:"
         Height          =   255
         Left            =   -68400
         TabIndex        =   39
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label lblTotalQtd 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67800
         TabIndex        =   38
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblTotalComissao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -66480
         TabIndex        =   33
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -67080
         TabIndex        =   32
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -68280
         TabIndex        =   31
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblValesTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67200
         TabIndex        =   30
         Top             =   4320
         Width           =   2055
      End
   End
   Begin MSDBCtls.DBCombo cboFuncionario 
      Bindings        =   "frmControlePagamento.frx":43D0
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "Vales e Comissões até:"
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Cod:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Funcionário:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Mês de Referência:"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmControlePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdemVale As String, StrOrdemComissao As String

Private Sub PagarComissoes()
Dim Resposta As Integer, Total As Currency
Total = 0

On Error GoTo 0
With qVendas
  If .Recordset.RecordCount = 0 Then
'    MsgBox "Não existe comissão para ser paga!"
    Exit Sub
  Else
    .Recordset.MoveLast
    .Recordset.MoveFirst
    dbVenda.Refresh
    If dbVenda.Recordset.RecordCount = 0 Then
      MsgBox "Erro na tabela de Venda!"
      Exit Sub
    End If
    Do While .Recordset.EOF = False
      dbVenda.Recordset.FindFirst "codigovenda=" & .Recordset!codigovenda
      If dbVenda.Recordset.NoMatch = False Then
        dbVenda.Recordset.Edit
        dbVenda.Recordset!Pago = True
        dbVenda.Recordset!idpagamento = dbPagamentos.Recordset!CodigoPagamento
        Total = Total - dbVenda.Recordset!ValorComissao
        dbVenda.Recordset.Update
      Else
        MsgBox "Erro na tabela de comissões!"
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With dbDespesalanc
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Pg Funcionários"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoConta = 0
  .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
  .Recordset!Descri = DbDespesaTipo.Recordset!Descri
  .Recordset!Obs = "Comissões - até " & Format(txtFechadoAte.Value, "short date")
  .Recordset!Valor = Total
  .Recordset!valorpago = Total
  .Recordset!Fechamento = True
  .Recordset!compensado = True
  .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
  .Recordset.Update
End With

End Sub


Private Sub RecebeVales()
Dim Valor As Currency, Descri As String
qVales.Refresh
If qVales.Recordset.EOF = True Then
  'MsgBox "Selecione um vale para ser lançado na despesa!"
  Exit Sub
End If
If qVales.Recordset!Cobrado = True Then
  MsgBox "Esse vale já foi cobrado!"
  Exit Sub
End If

Valor = 0

If DbDespesaTipo.Recordset.RecordCount = 0 Then
  MsgBox "Cadastro de despesas não possue registro!", vbCritical, "Erro!"
  Exit Sub
End If
With DbDespesaTipo
  .Refresh
  .Recordset.FindFirst "codigodespesa=" & dbFuncionarios.Recordset!CodigoDespesa
  If .Recordset.NoMatch = True Then
    MsgBox "A despesa vinculada foi removida!"
    Exit Sub
  End If
End With
With qVales
  .Refresh
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    With dbVales
      .Refresh
      If .Recordset.RecordCount = 0 Then
        MsgBox "Erro na tabela de vales!"
        Exit Sub
      End If
      .Recordset.FindFirst "codigovale=" & qVales.Recordset!codigovale
      If .Recordset.NoMatch = True Then
        GoTo tentaOutro
      End If
      .Recordset.Edit
      Valor = Valor + .Recordset!Valor
      .Recordset!Cobrado = True
      .Recordset!cobradoem = Now
      .Recordset!CodigoPagamento = dbPagamentos.Recordset!CodigoPagamento
      .Recordset.Update
    End With
tentaOutro:
    .Recordset.MoveNext
  Loop
End With
With dbDespesalanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = 0
  .Recordset!Origem = "Pg Funcionários"
  .Recordset("data") = Date
  .Recordset!Vencimento = Date
  .Recordset("hora") = Now
  .Recordset("codigoconta") = 0
  .Recordset("codigodespesa") = DbDespesaTipo.Recordset!CodigoDespesa
  .Recordset("descri") = DbDespesaTipo.Recordset!Descri
  .Recordset("obs") = "Vales até - " & Format(txtFechadoAte.Value, "Short date")
  .Recordset!compensado = True
  .Recordset("valor") = Valor
  .Recordset!valorpago = Valor
  .Recordset!compensado = -1
  .Recordset!Fechamento = True
  .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
  .Recordset.Update
  .Refresh
End With

End Sub

Private Sub AtivaAbrir(ByVal Sim As Boolean)
txtCodFuncionario.Enabled = Sim
cboFuncionario.Enabled = Sim
txtMesAno.Enabled = Sim
optSalario.Enabled = Sim
optAdiantamento.Enabled = Sim
txtFechadoAte.Enabled = Sim
cmdAbrir.Enabled = Sim
If Sim = True Then
  SSTab1.Visible = False
  txtCodFuncionario.SetFocus
  cmdSair.Cancel = True
Else
  SSTab1.Visible = True
  If txtSalario.Enabled = True Then txtSalario.SetFocus
  cmdCancelar.Cancel = True
End If
End Sub

Private Sub Calcular()
Dim Saldo As Currency, Salario As Currency, Vales As Currency
Dim Comissao As Currency, VR As Currency, VT As Currency

Saldo = 0
Salario = 0
Vales = 0
Comissao = 0
VR = 0
VT = 0

If IsNumeric(txtSalario.Text) = True Then
  Salario = CCur(txtSalario.Text)
End If
If IsNumeric(lblVales.Caption) = True Then
  Vales = CCur(lblVales.Caption)
End If
If IsNumeric(lblComissoes.Caption) = True Then
  Comissao = CCur(lblComissoes.Caption)
End If
If IsNumeric(txtVR.Text) = True Then
  VR = CCur(txtVR.Text)
End If
If IsNumeric(txtVT.Text) = True Then
  VT = CCur(txtVT.Text)
End If

Saldo = Salario + Vales + Comissao + VR + VT

lblSaldoAPagar.Caption = Format(Saldo, "Currency")

End Sub

Private Sub CarregaVales(ByVal Confirmado As Boolean, Optional CodigoPagamento As Double = 0)
Dim StrTemp As String, StrTempTotal As String
Dim Cobrado As Integer


Cobrado = Confirmado
If Confirmado = True Then
  StrTemp = "select *from QValesCaixa where qvales.fechado=-1 and codigopagamento=" & CodigoPagamento
  StrTempTotal = "select sum(valor) as total from qvalescaixa where qvales.fechado=-1 and codigopagamento=" & CodigoPagamento
Else
  StrTemp = "select *from QValesCaixa where qvales.fechado=-1 and cobrado=" & Cobrado
  StrTempTotal = "select sum(valor) as total from qvalescaixa where qvales.fechado=-1 and cobrado=" & Cobrado
  StrTemp = StrTemp & " and datacaixa <=#" & DataInglesa(txtFechadoAte.Value) & "#"
  StrTempTotal = StrTempTotal & " and datacaixa <=#" & DataInglesa(txtFechadoAte.Value) & "#"
End If

StrTemp = StrTemp & " and codfun=" & dbFuncionarios.Recordset!codigovendedor
StrTempTotal = StrTempTotal & " and codfun=" & dbFuncionarios.Recordset!codigovendedor



With qVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp & StrOrdemVale
  .Refresh
End With
With qValesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTempTotal
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblValesTotal.Caption = Format(-.Recordset!Total, "Currency")
  Else
    lblValesTotal.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub CarregaComissoes(ByVal Confirmado As Boolean, Optional CodigoPagamento As Double = 0)
Dim strComissao As String, StrComissaoTotal As String
Dim Pago As Integer, Total As Currency


Pago = Confirmado

If Confirmado = True Then
  strComissao = "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno and fechamentodiario=-1 and IdPagamento=" & CodigoPagamento
  StrComissaoTotal = "select sum(valorcomissao) as total, sum(quantidade) as qtd from qvendascomissoes where fechamentodiario=-1 and idpagamento=" & CodigoPagamento
Else
  strComissao = "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno and fechamentodiario=-1 and pago=" & Pago
  StrComissaoTotal = "select sum(valorcomissao) as total, sum(quantidade) as qtd from qvendascomissoes where fechamentodiario=-1 and pago=" & Pago
  strComissao = strComissao & " and data<=#" & DataInglesa(txtFechadoAte.Value) & "#"
  StrComissaoTotal = StrComissaoTotal & " and data<=#" & DataInglesa(txtFechadoAte.Value) & "#"
End If

strComissao = strComissao & " and codigopagamento=" & dbFuncionarios.Recordset!codigovendedor
StrComissaoTotal = StrComissaoTotal & " and codigopagamento=" & dbFuncionarios.Recordset!codigovendedor

With qVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = strComissao & StrOrdemComissao
  .Refresh
End With
With qVendasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrComissaoTotal
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalComissao.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalComissao.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Qtd) = False Then
    lblTotalQtd.Caption = .Recordset!Qtd
  Else
    lblTotalQtd.Caption = 0
  End If
End With
End Sub

Private Sub cboFuncionario_LostFocus()
With dbFuncionarios
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFuncionario.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboFuncionario.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodFuncionario.Text = .Recordset!Codigo
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub cmdAbrir_Click()
Dim intSalario As Integer
If dbFuncionarios.Recordset.EOF = True Then Exit Sub

If cboFuncionario.Text <> dbFuncionarios.Recordset!Nome Then
  MsgBox "Funcionário inválido!"
  txtCodFuncionario.SetFocus
  Exit Sub
End If
If IsNull(dbFuncionarios.Recordset!CodigoDespesa) = True Then
  MsgBox "Este funcionário não possue despesa vinculada cadastrada!"
  frmCadVendedor.Show
  frmCadVendedor.SetFocus
  Screen.MousePointer = vbDefault
  Exit Sub
End If
If dbFuncionarios.Recordset!CodigoDespesa = 0 Then
  MsgBox "Este funcionário não possue despesa vinculada cadastrada!"
  frmCadVendedor.Show
  frmCadVendedor.SetFocus
  Screen.MousePointer = vbDefault
  Exit Sub
End If
If optSalario.Value = True Then
  intSalario = -1
Else
  intSalario = 0
End If
With dbPagamentos
  .RecordSource = "Select *from vendedorespagamento where codigovendedor=" & dbFuncionarios.Recordset!codigovendedor & " and mes=" & txtMesAno.Month & " and ano=" & txtMesAno.Year & " and salario=" & intSalario
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!codigovendedor = dbFuncionarios.Recordset!codigovendedor
    .Recordset!Codigo = dbFuncionarios.Recordset!Codigo
    .Recordset!Funcionario = dbFuncionarios.Recordset!Nome
    .Recordset!Mes = txtMesAno.Month
    .Recordset!Ano = txtMesAno.Year
    .Recordset!Salario = intSalario
    If intSalario = -1 Then
      .Recordset!adiantamento = 0
    Else
      .Recordset!adiantamento = -1
    End If
    .Recordset!valorbase = 0
    .Recordset!Vales = 0
    .Recordset!Comissoes = 0
    .Recordset!VR = 0
    .Recordset!VT = 0
    .Recordset!saldoapagar = 0
    .Recordset!Pago = False
    .Recordset!fechadoate = txtFechadoAte.Value
    .Recordset!CodigoCaixa = 0
    .Recordset.Update
  End If
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Erro ao criar o pagamento!"
    Exit Sub
  End If
  
  If optSalario.Value = True Then
    
    txtVT.Enabled = True
    txtVR.Enabled = True
    
    CarregaVales .Recordset!Pago, .Recordset!CodigoPagamento
    CarregaComissoes .Recordset!Pago, .Recordset!CodigoPagamento
    
    lblVales.Caption = lblValesTotal.Caption
    lblComissoes.Caption = lblTotalComissao.Caption
  Else
    With qVales
      .RecordSource = "select *from QValesCaixa where codfun=0"
      .Refresh
    End With
    With qVendas
      .RecordSource = "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno and fechamentodiario=-1 and codigopagamento=-1"
      .Refresh
    End With
    lblValesTotal.Caption = Format(0, "Currency")
    lblTotalComissao.Caption = Format(0, "Currency")
    txtVT.Enabled = False
    txtVR.Enabled = False
    txtVT.Text = Format(0, "Currency")
    txtVR.Text = Format(0, "Currency")
  End If
  
  lblVales.Caption = lblValesTotal.Caption
  lblComissoes.Caption = lblTotalComissao.Caption
  
  txtSalario.Text = Format(.Recordset!valorbase, "Currency")
  txtVR.Text = Format(.Recordset!VR, "Currency")
  txtVT.Text = Format(.Recordset!VT, "Currency")
  lblSaldoAPagar.Caption = Format(.Recordset!saldoapagar, "Currency")
  
  If .Recordset!Pago = -1 Then
    cmdConfirma.Enabled = False
    txtSalario.Enabled = False
    txtVT.Enabled = False
    txtVR.Enabled = False
  Else
    cmdConfirma.Enabled = True
    txtSalario.Enabled = True
  End If
  
End With


Calcular

AtivaAbrir False

End Sub

Private Sub cmdCancelaPagamento_Click()
Dim Resposta As Integer, Ws As Workspace, db As Database
Dim CodigoPagamento As Integer

With dbPagamentos
  If .Recordset!CodigoCaixa <> 0 Then
    MsgBox "Não pode ser cancelado pois já foi lançado no caixa!"
    Exit Sub
  End If
  CodigoPagamento = .Recordset!CodigoPagamento
End With

With dbDespesalanc
  .Refresh
  .Recordset.FindFirst "NrDocumento='" & CodigoPagamento & "' and origem='Pg Funcionários'"
  If .Recordset.NoMatch = False Then
    If .Recordset!Fechamento = True Then
      MsgBox "Não pode ser cancelado pois já pertence a fechamento anterior!"
      Exit Sub
    End If
  Else
    MsgBox "Erro na tabela de despesas!"
    Exit Sub
  End If
End With

Resposta = MsgBox("Deseja cancelar o pagamento atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
If dbPagamentos.Recordset.EOF = True Then Exit Sub
If dbPagamentos.Recordset.BOF = True Then Exit Sub



Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)

db.Execute "delete from despesaslanc2 where NrDocumento='" & CodigoPagamento & "' and origem='Pg Funcionários'"

db.Execute "update vales set cobrado=0 where codigopagamento=" & CodigoPagamento
db.Execute "update vales set cobradoem=null where codigopagamento=" & CodigoPagamento
db.Execute "update vales set codigopagamento=0 where codigopagamento=" & CodigoPagamento

db.Execute "update venda2 set pago=0 where idpagamento=" & CodigoPagamento
db.Execute "update venda2 set idpagamento=null where idpagamento=" & CodigoPagamento

With dbPagamentos
  .Recordset.Edit
  .Recordset!Pago = 0
  .Recordset.Update
End With

Call cmdAbrir_Click

DoEvents

Dim Estatus As New frmEstatus2
Load Estatus
Unload Estatus

End Sub

Private Sub cmdCancelar_Click()
AtivaAbrir True
End Sub

Private Sub cmdConfirma_Click()
Dim Resposta As Integer


Resposta = MsgBox("Deseja confirmar o pagamento atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

With DbDespesaTipo
  .Refresh
  .Recordset.FindFirst "codigodespesa=" & dbFuncionarios.Recordset!CodigoDespesa
  If .Recordset.NoMatch = True Then
    MsgBox "Este funcionário não possue despesa vinculada cadastrada!"
    frmCadVendedor.Show
    frmCadVendedor.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
End With


If optSalario.Value = True Then
  With dbDespesalanc
    'Salario
    'If CCur(txtSalario.Text) <> 0 Then
      .Recordset.AddNew
      .Recordset!CodigoFechamento = 0
      .Recordset!Origem = "Pg Funcionários"
      .Recordset!Data = Date
      .Recordset!Hora = Now
      .Recordset!Vencimento = Date
      .Recordset!CodigoConta = 0
      .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
      .Recordset!Descri = DbDespesaTipo.Recordset!Descri
      .Recordset!Obs = "Salário - ref " & Format(txtMesAno.Value, "MM/YYYY")
      .Recordset!Valor = -CCur(txtSalario.Text)
      .Recordset!valorpago = -CCur(txtSalario.Text)
      .Recordset!compensado = -1
      .Recordset!fechamentodiario = -1
      .Recordset!Fechamento = 0
      .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
      .Recordset!codigoenviar = "1"
      .Recordset.Update
    'End If
    PagarComissoes
    
    RecebeVales
    
    'Vale Transporte
    If CCur(txtVT.Text) <> 0 Then
      .Recordset.AddNew
      .Recordset!CodigoFechamento = 0
      .Recordset!Origem = "Pg Funcionários"
      .Recordset!Data = Date
      .Recordset!Hora = Now
      .Recordset!Vencimento = Date
      .Recordset!CodigoConta = 0
      .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
      .Recordset!Descri = DbDespesaTipo.Recordset!Descri
      .Recordset!Obs = "Vale Transporte - ref " & Format(txtMesAno.Value, "MM/YYYY")
      .Recordset!Valor = -CCur(txtVT.Text)
      .Recordset!valorpago = -CCur(txtVT.Text)
      .Recordset!compensado = -1
      .Recordset!fechamentodiario = -1
      .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
      .Recordset!Fechamento = 0
      .Recordset!codigoenviar = "1"
      .Recordset.Update
    End If
    
    'Vale Refeição
    If CCur(txtVR.Text) <> 0 Then
      .Recordset.AddNew
      .Recordset!CodigoFechamento = 0
      .Recordset!Origem = "Pg Funcionários"
      .Recordset!Data = Date
      .Recordset!Hora = Now
      .Recordset!Vencimento = Date
      .Recordset!CodigoConta = 0
      .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
      .Recordset!Descri = DbDespesaTipo.Recordset!Descri
      .Recordset!Obs = "Vale Refeição - ref " & Format(txtMesAno.Value, "MM/YYYY")
      .Recordset!Valor = -CCur(txtVR.Text)
      .Recordset!valorpago = -CCur(txtVR.Text)
      .Recordset!compensado = -1
      .Recordset!fechamentodiario = -1
      .Recordset!Fechamento = 0
      .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
      .Recordset!codigoenviar = "1"
      .Recordset.Update
    End If
  End With
  
  
Else
  
  With dbDespesalanc
    'Salario
    If CCur(txtSalario.Text) <> 0 Then
      .Recordset.AddNew
      .Recordset!CodigoFechamento = 0
      .Recordset!Origem = "Pg Funcionários"
      .Recordset!Data = Date
      .Recordset!Hora = Now
      .Recordset!Vencimento = Date
      .Recordset!CodigoConta = 0
      .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
      .Recordset!Descri = DbDespesaTipo.Recordset!Descri
      .Recordset!Obs = "Adiantamento - ref " & Format(txtMesAno.Value, "MM/YYYY")
      .Recordset!Valor = -CCur(txtSalario.Text)
      .Recordset!valorpago = -CCur(txtSalario.Text)
      .Recordset!compensado = -1
      .Recordset!fechamentodiario = -1
      .Recordset!Fechamento = 0
      .Recordset!NrDocumento = dbPagamentos.Recordset!CodigoPagamento
      .Recordset!codigoenviar = "1"
      .Recordset.Update
    End If
  End With
  
End If

With dbPagamentos
  .Recordset.Edit
  .Recordset!valorbase = CCur(txtSalario.Text)
  .Recordset!Vales = CCur(lblVales.Caption)
  .Recordset!Comissoes = CCur(lblComissoes.Caption)
  .Recordset!VR = CCur(txtVR.Text)
  .Recordset!VT = CCur(txtVT.Text)
  .Recordset!saldoapagar = CCur(lblSaldoAPagar.Caption)
  .Recordset!Pago = True
  .Recordset!usuariocriou = Usuarios.Nome
  .Recordset.Update
End With

Call cmdAbrir_Click

Dim Estatus As New frmEstatus2
Load Estatus
Unload Estatus

End Sub

Private Sub cmdImprime2_Click()

Dim StrTemp As String, StrTemp2 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = Format(Now, "long date") & " " & Format(Now, "short time")
StrTemp2 = "Finalizado até: " & txtFechadoAte.Value
If cboFuncionario.Text <> "" Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Funcionário: " & cboFuncionario.Text
End If

ImprimeGrid grdVales, Printer, qVales, 3, , , , , , "Vales de Funcionários - " & NomePosto, StrTemp, StrTemp2

Printer.NewPage

StrTemp = Format(Now, "long date") & " " & Format(Now, "short time")
StrTemp2 = "Finalizado até: " & txtFechadoAte.Value
If cboFuncionario.Text <> "" Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Funcionário: " & cboFuncionario.Text
End If

ImprimeGrid grdComissoes, Printer, qVendas, 4, , , , 6, 7, "Comissões - " & NomePosto, StrTemp, StrTemp2

Printer.NewPage

Load frmControlePagamentoImprimir
Unload frmControlePagamentoImprimir

Printer.EndDoc
NaoImprime:

End Sub

Private Sub cmdImprimeFolha_Click()
Dim StrTemp As String, StrTemp2 As String, intSalario As Integer

If optSalario.Value = True Then
  intSalario = -1
Else
  intSalario = 0
End If
With dbPagamentos
  .RecordSource = "Select *from vendedorespagamento where mes=" & txtMesAno.Month & " and ano=" & txtMesAno.Year & " and salario=" & intSalario & " order by funcionario"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    StrTemp = Format(Now, "long date") & " " & Format(Now, "short time")
    
    If optSalario.Value = True Then
      StrTemp2 = "Referente: Salário - " & txtMesAno.Month & "/" & txtMesAno.Year
    Else
      StrTemp2 = "Referente: Adiantamento - " & txtMesAno.Month & "/" & txtMesAno.Year
    End If
    
    
    ImprimeGrid grdFolha, Printer, dbPagamentos, 2, , , , 3, 4, "Pagamento - " & NomePosto, StrTemp, StrTemp2, 5, 6, 7
    
    Printer.EndDoc
  End If
End With
NaoImprime:
End Sub

Private Sub cmdNaoIdentificadas_Click()
frmControleValesSemFuncionario.Show
frmControleValesSemFuncionario.SetFocus
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja remover o pagamento atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbPagamentos
  If .Recordset.EOF = True Then
    MsgBox "Erro na tabela de pagamentos"
    Exit Sub
  End If
  If .Recordset.BOF = True Then
    MsgBox "Erro na tabela de pagamentos"
    Exit Sub
  End If
  If .Recordset!Pago = True Then
    MsgBox "Não pode ser removido pagamento confirmado!"
    Exit Sub
  End If
  .Recordset.Delete
  AtivaAbrir True
End With

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
StrOrdemVale = " order by Datacaixa, HoraIni"
StrOrdemComissao = " order by Data, HoraIni"

txtFechadoAte.Value = Date
txtMesAno.Value = Date

With dbFuncionarios
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVenda
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qValesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesalanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With DbDespesaTipo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With


Select Case Usuarios.Grupo.ControleVales
  Case 1 'Somente leitura
    
  Case 2 'Liberado
    
End Select

End Sub

Private Sub grdComissoes_HeadClick(ByVal ColIndex As Integer)
If StrOrdemComissao = " order by " & grdComissoes.Columns(ColIndex).DataField & ", HoraIni" Then
  StrOrdemComissao = " order by " & grdComissoes.Columns(ColIndex).DataField & " desc, HoraIni"
Else
  StrOrdemComissao = " order by " & grdComissoes.Columns(ColIndex).DataField & ", HoraIni"
End If
CarregaComissoes dbPagamentos.Recordset!Pago, dbPagamentos.Recordset!CodigoPagamento

End Sub

Private Sub grdVales_HeadClick(ByVal ColIndex As Integer)
If StrOrdemVale = " order by " & grdVales.Columns(ColIndex).DataField & ", HoraIni" Then
  StrOrdemVale = " order by " & grdVales.Columns(ColIndex).DataField & " desc, HoraIni"
Else
  StrOrdemVale = " order by " & grdVales.Columns(ColIndex).DataField & ", HoraIni"
End If
CarregaVales dbPagamentos.Recordset!Pago, dbPagamentos.Recordset!CodigoPagamento
End Sub

Private Sub txtCodFuncionario_LostFocus()
With dbFuncionarios
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodFuncionario.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodFuncionario.Text
  If .Recordset.NoMatch = False Then
    txtCodFuncionario.Text = .Recordset!Codigo
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtFechadoAte_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtFechadoAte_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtFechadoAte_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtMesAno_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtMesAno_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtMesAno_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtSalario_GotFocus()
With txtSalario
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtSalario_LostFocus()
With txtSalario
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "Currency")
    Calcular
  End If
End With
End Sub

Private Sub txtVR_GotFocus()
With txtVR
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVR_LostFocus()
With txtVR
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "Currency")
    Calcular
  End If
End With
End Sub

Private Sub txtVT_GotFocus()
With txtVT
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVT_LostFocus()
With txtVT
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "Currency")
    Calcular
  End If
End With
End Sub
