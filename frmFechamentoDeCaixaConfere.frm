VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFechamentoDeCaixaConfere 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conferência de Valores dos Caixas"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   Icon            =   "frmFechamentoDeCaixaConfere.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "DBFs"
      Height          =   5895
      Left            =   5400
      TabIndex        =   122
      Top             =   6120
      Visible         =   0   'False
      Width           =   10095
      Begin VB.Data dbClientesNota2 
         Caption         =   "dbClientesNota2"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientesnota2 where codigofechamento=0 order by Nome"
         Top             =   4920
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbClientes2 
         Caption         =   "dbClientes2"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Clientes order by Nome"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbFechamento2 
         Caption         =   "dbFechamento2"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select *from fechamentodecaixa where codigofechamento=0"
         Top             =   4560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbVendas 
         Caption         =   "dbVendas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(valorcomissao) as total from vendas2 where codigofechamento=0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbProdutos2 
         Caption         =   "dbProdutos2"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Produtos order by descri"
         Top             =   2760
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbPagamentosCaixa 
         Caption         =   "dbPagamentosCaixa"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select *from vendedorespagamento where codigocaixa=0"
         Top             =   4200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbPagamentos 
         Caption         =   "dbPagamentos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "VendedoresPagamento"
         Top             =   3840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbBicosEncerrantes 
         Caption         =   "dbBicosEncerrantes"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from BicoEncerrantes where codigofechamento=0"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbProdutosAltera 
         Caption         =   "dbProdutosAltera"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbClientesProdutos 
         Caption         =   "dbClientesProdutos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesProdutos order by codigocliente, codigoproduto, validade, horaini"
         Top             =   2400
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbDespesaTipoGrupo 
         Caption         =   "dbDespesaTipoGrupo"
         Connect         =   "Access"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from despesatiposubgrupo order by descri"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data dbBloqueiaFechamento 
         Caption         =   "dbBloqueiaFechamento"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "BloqueiaFechamento"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Data dbClientesNota2Temp 
         Caption         =   "dbClientesNota2Temp"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Clientesnota2 order by codigocliente, codigocarro, km"
         Top             =   1320
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbClientesCarros 
         Caption         =   "dbClientesCarros"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientescarros where codigocliente=0"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data qValesTotal 
         Caption         =   "qValesTotal"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from qVales where codigocaixa=0 order by nome"
         Top             =   2760
         Width           =   3180
      End
      Begin VB.Data qVales 
         Caption         =   "qVales"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from qVales where codigocaixa=0 order by nome"
         Top             =   2400
         Width           =   3180
      End
      Begin VB.Data dbVales 
         Caption         =   "dbVales"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Vales"
         Top             =   2040
         Width           =   3180
      End
      Begin VB.Data dbVendedores 
         Caption         =   "dbVendedores"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Vendedores order by nome"
         Top             =   960
         Width           =   3180
      End
      Begin VB.Data dbPosto 
         Caption         =   "dbPosto"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Postos"
         Top             =   600
         Width           =   3180
      End
      Begin VB.Data dbProdutosNotasCorpo 
         Caption         =   "dbProdutosNotasCorpo"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ProdutosNotasCorpo"
         Top             =   240
         Width           =   3180
      End
      Begin VB.Data dbProdutosNotas 
         Caption         =   "dbProdutosNotas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ProdutosNotas"
         Top             =   5280
         Width           =   3180
      End
      Begin VB.Data dbTanques 
         Caption         =   "dbTanques"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tanques"
         Top             =   4920
         Width           =   3180
      End
      Begin VB.Data dbStatus 
         Caption         =   "dbStatus"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from status"
         Top             =   4560
         Width           =   3180
      End
      Begin VB.Data dbConciliaNova 
         Caption         =   "dbConciliaNova"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from conciliaNova"
         Top             =   4200
         Width           =   3180
      End
      Begin VB.Data dbContas 
         Caption         =   "dbContas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from contas order by descri"
         Top             =   3840
         Width           =   3180
      End
      Begin VB.Data dbCartoes 
         Caption         =   "dbCartoes"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Cartoes"
         Top             =   3120
         Width           =   3180
      End
      Begin VB.Data dbJuros 
         Caption         =   "dbJuros"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from juros order by inicio, final"
         Top             =   2760
         Width           =   3180
      End
      Begin VB.Data dbCheques 
         Caption         =   "dbCheques"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from cheques where codigofechamento=0"
         Top             =   1680
         Width           =   3180
      End
      Begin VB.Data dbClientesCheques 
         Caption         =   "dbClientesCheques"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from chequesclientes order by nome"
         Top             =   2040
         Width           =   3180
      End
      Begin VB.Data dbChequesContas 
         Caption         =   "dbChequesContas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from chequescontas"
         Top             =   2400
         Width           =   3180
      End
      Begin VB.Data dbProdutoEntra 
         Caption         =   "dbProdutoEntra"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtosentrada where codigofechamento=-1"
         Top             =   1320
         Width           =   3180
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Produtos where permitenocaixa=-1 order by descri"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbClientesNota 
         Caption         =   "dbClientesNota"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientesnota2 where codigofechamento=0 order by Nome"
         Top             =   960
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Clientes order by Nome"
         Top             =   1320
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbDespesas 
         Caption         =   "dbDespesas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DespesaTipo order by descri"
         Top             =   960
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbDespesasLanc 
         Caption         =   "dbDespesasLanc"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DespesasLanc where codigofechamento=0 order by descri"
         Top             =   600
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbFormaDePg 
         Caption         =   "dbFormaDePg"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FormaDePagamento order by descri"
         Top             =   600
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Data dbFormaDePgRecebido 
         Caption         =   "dbFormaDePgRecebido"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FormaDePagamentoRecebido where codigofechamento=0 order by descri"
         Top             =   240
         Visible         =   0   'False
         Width           =   3180
      End
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Data dbFechamento 
      Caption         =   "dbFechamento"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmFechamentoDeCaixaConfere.frx":0442
      Top             =   3840
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancelaFinaliza 
      Caption         =   "Cancela Finalização"
      Height          =   375
      Left            =   6240
      TabIndex        =   152
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdAtualiza 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   9240
      TabIndex        =   142
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   375
      Left            =   4920
      TabIndex        =   141
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1200
      Top             =   1200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   10680
      TabIndex        =   143
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Im&primir"
      Height          =   375
      Left            =   3600
      TabIndex        =   140
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Resumo "
      Height          =   2055
      Left            =   120
      TabIndex        =   127
      Top             =   4200
      Width           =   3375
      Begin VB.Label lblJurosNota 
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
         Left            =   1320
         TabIndex        =   154
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Juros Notas:"
         Height          =   195
         Left            =   120
         TabIndex        =   153
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblTotalValeResumo 
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
         Left            =   1320
         TabIndex        =   148
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Vales:"
         Height          =   195
         Left            =   120
         TabIndex        =   147
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lblJurosResumo 
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
         Left            =   1320
         TabIndex        =   137
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Juros Cheques:"
         Height          =   195
         Left            =   120
         TabIndex        =   136
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblDiferenca 
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
         Left            =   1320
         TabIndex        =   134
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Recebimentos:"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblRecebimentos 
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
         Left            =   1320
         TabIndex        =   132
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblTotalDespesas 
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
         Left            =   1320
         TabIndex        =   131
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Despesas:"
         Height          =   195
         Left            =   120
         TabIndex        =   130
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblTotalVendas 
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
         Left            =   1320
         TabIndex        =   129
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Vendas:"
         Height          =   195
         Left            =   120
         TabIndex        =   128
         Top             =   240
         Width           =   585
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   8
      Tab             =   3
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "&Recebidos"
      TabPicture(0)   =   "frmFechamentoDeCaixaConfere.frx":04D5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblTotalRecebido"
      Tab(0).Control(1)=   "Label41"
      Tab(0).Control(2)=   "Tela(0)"
      Tab(0).Control(3)=   "DBGrid2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Despesas"
      TabPicture(1)   =   "frmFechamentoDeCaixaConfere.frx":04F1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label40"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDespesas"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Tela(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DBGrid3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdGravaDespesas"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Notas"
      TabPicture(2)   =   "frmFechamentoDeCaixaConfere.frx":050D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Tela(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Compras"
      TabPicture(3)   =   "frmFechamentoDeCaixaConfere.frx":0529
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblProdutoEntraTotal"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label55"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Tela(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DBGrid5"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdGravaCompra"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "C&heques"
      TabPicture(4)   =   "frmFechamentoDeCaixaConfere.frx":0545
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Tela(4)"
      Tab(4).Control(1)=   "DBGrid6"
      Tab(4).Control(2)=   "Image1"
      Tab(4).Control(3)=   "Label70"
      Tab(4).Control(4)=   "Label54"
      Tab(4).Control(5)=   "Label53"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Vales"
      TabPicture(5)   =   "frmFechamentoDeCaixaConfere.frx":0561
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdGravaVales"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Tela(5)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "DBGrid7"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblTotalVale"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label8"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Pag. Func."
      TabPicture(6)   =   "frmFechamentoDeCaixaConfere.frx":057D
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblTotalPagamentos"
      Tab(6).Control(1)=   "Label27"
      Tab(6).Control(2)=   "DBGrid9"
      Tab(6).Control(3)=   "DBGrid8"
      Tab(6).Control(4)=   "cmdSomar"
      Tab(6).Control(5)=   "cmdSubtrair"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Microcredito"
      TabPicture(7)   =   "frmFechamentoDeCaixaConfere.frx":0599
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "DBGrid10"
      Tab(7).Control(1)=   "Tela(6)"
      Tab(7).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGrid10 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":05B5
         Height          =   2655
         Left            =   -74760
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":05D2
         TabIndex        =   204
         Top             =   2280
         Width           =   8535
      End
      Begin VB.Frame Tela 
         Height          =   5175
         Index           =   6
         Left            =   -74880
         TabIndex        =   175
         Top             =   360
         Width           =   8775
         Begin VB.TextBox txtValorMicrocredito 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   300
            Left            =   5400
            TabIndex        =   186
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdIncluirMicrocredito 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   120
            TabIndex        =   185
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtCupomMicrocredito 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5880
            TabIndex        =   184
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtKmMicrocredito 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            TabIndex        =   183
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtQtdMicrocredito 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            TabIndex        =   182
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdRemoverMicrocredito 
            Caption         =   "Remover"
            Height          =   375
            Left            =   1440
            TabIndex        =   181
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtCodMicrocredito 
            Height          =   285
            Left            =   840
            TabIndex        =   180
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdMudarAutorizaMicrocredito 
            Caption         =   "Mudar Autorização"
            Height          =   375
            Left            =   5760
            TabIndex        =   179
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmdImprimeAutorizacaoMicrocredito 
            Caption         =   "Imprimir Autorização"
            Height          =   375
            Left            =   120
            TabIndex        =   178
            Top             =   4680
            Width           =   1815
         End
         Begin VB.CheckBox chkConferidoMicrocredito 
            Caption         =   "Notas já conferidas"
            DataField       =   "NotaConferida2"
            DataSource      =   "dbFechamento2"
            Height          =   255
            Left            =   2760
            TabIndex        =   177
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton cmdImportarMicrocredito 
            Caption         =   "Importar"
            Height          =   375
            Left            =   6720
            TabIndex        =   176
            Top             =   960
            Width           =   975
         End
         Begin MSDBCtls.DBCombo cboMicrocredito 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1D2A
            Height          =   315
            Left            =   120
            TabIndex        =   187
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboPlacaMicrocredito 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1D44
            Height          =   315
            Left            =   3120
            TabIndex        =   188
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Placa"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboProdutoMicrocredito 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1D63
            Height          =   315
            Left            =   1560
            TabIndex        =   189
            Top             =   1080
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboBicoMicrocredito 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1D7D
            Height          =   315
            Left            =   4560
            TabIndex        =   190
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Bico"
            Text            =   ""
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   5400
            TabIndex        =   203
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   120
            TabIndex        =   202
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Cupom:"
            Height          =   195
            Left            =   5880
            TabIndex        =   201
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
            Height          =   195
            Left            =   3120
            TabIndex        =   200
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Km:"
            Height          =   195
            Left            =   4680
            TabIndex        =   199
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Qtd:"
            Height          =   195
            Left            =   120
            TabIndex        =   198
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label45 
            Caption         =   "Cod:"
            Height          =   255
            Left            =   840
            TabIndex        =   197
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Left            =   1560
            TabIndex        =   196
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Bico:"
            Height          =   195
            Left            =   4560
            TabIndex        =   195
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label39 
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
            Left            =   6000
            TabIndex        =   194
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Total na Bomba:"
            Height          =   195
            Left            =   4800
            TabIndex        =   193
            Top             =   4680
            Width           =   1170
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   2400
            TabIndex        =   192
            Top             =   4680
            Width           =   405
         End
         Begin VB.Label Label35 
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
            Left            =   3000
            TabIndex        =   191
            Top             =   4680
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdSubtrair 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   158
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdSomar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   157
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdGravaCompra 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   120
         TabIndex        =   155
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdGravaVales 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   -74880
         TabIndex        =   151
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdGravaDespesas 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   -74880
         TabIndex        =   150
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Frame Tela 
         Height          =   1455
         Index           =   5
         Left            =   -74880
         TabIndex        =   144
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdRemoverVale 
            Caption         =   "Remover"
            Height          =   375
            Left            =   2520
            TabIndex        =   108
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdIncluirVale 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   1320
            TabIndex        =   107
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtValeValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   106
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox cboMotivo 
            Height          =   315
            ItemData        =   "frmFechamentoDeCaixaConfere.frx":1D9E
            Left            =   3960
            List            =   "frmFechamentoDeCaixaConfere.frx":1DC0
            TabIndex        =   104
            Top             =   360
            Width           =   2655
         End
         Begin MSDBCtls.DBCombo cboFuncionario 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1E3A
            Height          =   315
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
         End
         Begin VB.Label Label7 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   3960
            TabIndex        =   103
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Funcionário:"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Tela 
         Height          =   2535
         Index           =   4
         Left            =   -74880
         TabIndex        =   138
         Top             =   360
         Width           =   7815
         Begin VB.CommandButton cmdChequeMudaAutoriza 
            Caption         =   "Mudar Autorização"
            Height          =   375
            Left            =   5880
            TabIndex        =   168
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtCMC7 
            Height          =   285
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtCodCliente 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   68
            Top             =   420
            Width           =   615
         End
         Begin MSComCtl2.DTPicker txtBomPara 
            Height          =   300
            Left            =   120
            TabIndex        =   92
            Top             =   2160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   112066561
            CurrentDate     =   38426
         End
         Begin VB.CommandButton cmdRelaciona 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   4560
            TabIndex        =   99
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   94
            Top             =   2160
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   0
            Left            =   2280
            TabIndex        =   82
            Top             =   1560
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   1
            Left            =   2880
            TabIndex        =   84
            Top             =   1560
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   2
            Left            =   3480
            TabIndex        =   86
            Top             =   1560
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   3
            Left            =   4200
            TabIndex        =   88
            Top             =   1560
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Mask            =   "999999-9"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   4
            Left            =   5160
            TabIndex        =   90
            Top             =   1560
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   6
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSDBCtls.DBCombo cboClienteCheque 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":1E55
            Height          =   315
            Left            =   840
            TabIndex        =   70
            Top             =   420
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
         End
         Begin VB.Label Label26 
            Caption         =   "Leitor de Código de barras:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   840
            TabIndex        =   69
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label83 
            Caption         =   "Código:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   1200
            TabIndex        =   79
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label lblSaldo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   80
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Pré-Datados:"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label lblPre 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Limite:"
            Height          =   195
            Left            =   5640
            TabIndex        =   73
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5640
            TabIndex        =   74
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   1920
            Width           =   390
         End
         Begin VB.Label Label47 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   1560
            TabIndex        =   93
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            Height          =   195
            Left            =   5160
            TabIndex        =   89
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   195
            Left            =   4200
            TabIndex        =   87
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
            Height          =   195
            Left            =   3480
            TabIndex        =   85
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   2880
            TabIndex        =   83
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Comp:"
            Height          =   195
            Left            =   2280
            TabIndex        =   81
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "Na Bomba:"
            Height          =   195
            Left            =   2640
            TabIndex        =   95
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label lblValorNaBomba 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2640
            TabIndex        =   96
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
            Height          =   195
            Left            =   3600
            TabIndex        =   97
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label lblJurosTabelado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3600
            TabIndex        =   98
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label lblStatus 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4800
            TabIndex        =   72
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
            Height          =   195
            Left            =   4800
            TabIndex        =   71
            Top             =   240
            Width           =   495
         End
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":1E75
         Height          =   2175
         Left            =   -74760
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":1E97
         TabIndex        =   14
         Top             =   2760
         Width           =   6255
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":2DA6
         Height          =   2775
         Left            =   -74760
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":2DC3
         TabIndex        =   31
         Top             =   1800
         Width           =   7455
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":394E
         Height          =   2655
         Left            =   -74760
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":396B
         TabIndex        =   52
         Top             =   2280
         Width           =   8535
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":50C2
         Height          =   3015
         Left            =   240
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":50DF
         TabIndex        =   66
         Top             =   1800
         Width           =   6375
      End
      Begin VB.Frame Tela 
         Height          =   4695
         Index           =   0
         Left            =   -74880
         TabIndex        =   123
         Top             =   360
         Width           =   7695
         Begin VB.CommandButton cmdGravarPagamentos 
            Caption         =   "Gravar"
            Height          =   375
            Left            =   5400
            TabIndex        =   149
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtChequeJuros 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   4
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtChequeBomba 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtOperacoes 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Left            =   2880
            TabIndex        =   12
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox txtValorRecebe 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   10
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdIncluirRecebimento 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   3840
            TabIndex        =   13
            Top             =   1920
            Width           =   975
         End
         Begin MSDBCtls.DBCombo cboRecebimento 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":5FD2
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1440
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtDataBordero 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   161284097
            CurrentDate     =   37600
         End
         Begin VB.Label lblComissoesCombustiveis 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6240
            TabIndex        =   174
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comissões de Combustíveis:"
            Height          =   195
            Left            =   5520
            TabIndex        =   173
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comissões de Produtos:"
            Height          =   195
            Left            =   5880
            TabIndex        =   171
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblComissoesPagas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6240
            TabIndex        =   170
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblJuros 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2880
            TabIndex        =   113
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros  de Cheque + Juros de Notas:"
            Height          =   195
            Left            =   285
            TabIndex        =   112
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total de Cheque C/ Juros:"
            Height          =   195
            Left            =   960
            TabIndex        =   111
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total de Cheque na Bomba:"
            Height          =   195
            Left            =   840
            TabIndex        =   110
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Operações:"
            Height          =   195
            Left            =   2880
            TabIndex        =   11
            Top             =   1800
            Width           =   825
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            Top             =   1800
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label Label86 
            Caption         =   "Data Borderô:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Tela 
         Height          =   4335
         Index           =   1
         Left            =   -74880
         TabIndex        =   124
         Top             =   360
         Width           =   7695
         Begin VB.TextBox txtObsAdicional 
            Height          =   285
            Left            =   3360
            TabIndex        =   29
            Top             =   1080
            Visible         =   0   'False
            Width           =   3015
         End
         Begin MSDBCtls.DBCombo cboSubGrupo 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":5FEC
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin VB.TextBox txtDespesaObs 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   3135
         End
         Begin VB.CommandButton cmdIncluirDespesa 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   6480
            TabIndex        =   30
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtDespesaValor 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Left            =   3480
            TabIndex        =   18
            Top             =   480
            Width           =   1575
         End
         Begin MSDBCtls.DBCombo cboDespesa 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":600D
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtDataIni 
            Height          =   300
            Left            =   3360
            TabIndex        =   23
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   158072833
            CurrentDate     =   39470
         End
         Begin MSComCtl2.DTPicker txtDataFim 
            Height          =   300
            Left            =   5040
            TabIndex        =   25
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   158072833
            CurrentDate     =   39470
         End
         Begin MSComCtl2.DTPicker txtMesAno 
            Height          =   300
            Left            =   3360
            TabIndex        =   27
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   158072835
            CurrentDate     =   39470
         End
         Begin VB.Label lblMesAno 
            AutoSize        =   -1  'True
            Caption         =   "Mês e Ano Referência:"
            Height          =   195
            Left            =   3360
            TabIndex        =   26
            Top             =   840
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            Caption         =   "Período:"
            Height          =   195
            Left            =   3360
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblPeriodoA 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   4800
            TabIndex        =   24
            Top             =   1080
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblObsAdicional 
            AutoSize        =   -1  'True
            Caption         =   "Obs. Adicional:"
            Height          =   195
            Left            =   3360
            TabIndex        =   28
            Top             =   840
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   3480
            TabIndex        =   17
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Despesa:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Tela 
         Height          =   4695
         Index           =   3
         Left            =   120
         TabIndex        =   126
         Top             =   360
         Width           =   6615
         Begin VB.TextBox txtTanqueEntra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   64
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtTotalEntra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtQtdEntra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   58
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCod 
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdProdutoEntra 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   3000
            TabIndex        =   65
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtValorUnitario 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   62
            Top             =   1080
            Width           =   735
         End
         Begin MSDBCtls.DBCombo cboProdutoEntra 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":6026
            Height          =   315
            Left            =   960
            TabIndex        =   56
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin VB.Label lblTanque 
            AutoSize        =   -1  'True
            Caption         =   "Tanque:"
            Height          =   195
            Left            =   2160
            TabIndex        =   63
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   3600
            TabIndex        =   57
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Cod.:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Left            =   960
            TabIndex        =   55
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "$ Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "$ Unitário:"
            Height          =   195
            Left            =   1320
            TabIndex        =   61
            Top             =   840
            Width           =   720
         End
      End
      Begin MSDBGrid.DBGrid DBGrid6 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":603F
         Height          =   2175
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":6057
         TabIndex        =   100
         Top             =   2940
         Width           =   7815
      End
      Begin MSDBGrid.DBGrid DBGrid7 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":7AF2
         Height          =   3135
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":7B07
         TabIndex        =   109
         Top             =   1920
         Width           =   7695
      End
      Begin MSDBGrid.DBGrid DBGrid8 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":89DA
         Height          =   2055
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":89F5
         TabIndex        =   156
         Top             =   480
         Width           =   7695
      End
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "frmFechamentoDeCaixaConfere.frx":9A78
         Height          =   2055
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":9A98
         TabIndex        =   161
         Top             =   3120
         Width           =   7695
      End
      Begin VB.Frame Tela 
         Height          =   5175
         Index           =   2
         Left            =   -74880
         TabIndex        =   125
         Top             =   360
         Width           =   8775
         Begin VB.CommandButton cmdImportarNotas 
            Caption         =   "Importar"
            Height          =   375
            Left            =   6720
            TabIndex        =   172
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chkNotaConferida 
            Caption         =   "Notas já conferidas"
            DataField       =   "NotaConferida"
            DataSource      =   "dbFechamento2"
            Height          =   255
            Left            =   2760
            TabIndex        =   169
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton cmdImprimirAutoriza 
            Caption         =   "Imprimir Autorização"
            Height          =   375
            Left            =   120
            TabIndex        =   167
            Top             =   4680
            Width           =   1815
         End
         Begin VB.CommandButton cmdAutorizar 
            Caption         =   "Mudar Autorização"
            Height          =   375
            Left            =   5760
            TabIndex        =   162
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtCodProduto 
            Height          =   285
            Left            =   840
            TabIndex        =   43
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdRemoveNota 
            Caption         =   "Remover"
            Height          =   375
            Left            =   1440
            TabIndex        =   51
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtLitros 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtKm 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            TabIndex        =   37
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCupom 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5880
            TabIndex        =   39
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdInclueNota 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtNotaValor 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   300
            Left            =   5400
            TabIndex        =   49
            Top             =   1080
            Width           =   1215
         End
         Begin MSDBCtls.DBCombo cboClientesNota 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":AB1B
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboPlaca 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":AB34
            Height          =   315
            Left            =   3120
            TabIndex        =   35
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Placa"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboProduto 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":AB53
            Height          =   315
            Left            =   1560
            TabIndex        =   45
            Top             =   1080
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboBico 
            Bindings        =   "frmFechamentoDeCaixaConfere.frx":AB6D
            Height          =   315
            Left            =   4560
            TabIndex        =   47
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Bico"
            Text            =   ""
         End
         Begin VB.Label lblTotalNota 
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
            Left            =   3000
            TabIndex        =   166
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   2400
            TabIndex        =   165
            Top             =   4680
            Width           =   405
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Total na Bomba:"
            Height          =   195
            Left            =   4800
            TabIndex        =   164
            Top             =   4680
            Width           =   1170
         End
         Begin VB.Label lblTotalNabomba 
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
            Left            =   6000
            TabIndex        =   163
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label lblBico 
            AutoSize        =   -1  'True
            Caption         =   "Bico:"
            Height          =   195
            Left            =   4560
            TabIndex        =   46
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Left            =   1560
            TabIndex        =   44
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label18 
            Caption         =   "Cod:"
            Height          =   255
            Left            =   840
            TabIndex        =   42
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Qtd:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Km:"
            Height          =   195
            Left            =   4680
            TabIndex        =   36
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
            Height          =   195
            Left            =   3120
            TabIndex        =   34
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Cupom:"
            Height          =   195
            Left            =   5880
            TabIndex        =   38
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   5400
            TabIndex        =   48
            Top             =   840
            Width           =   405
         End
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -70200
         TabIndex        =   160
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblTotalPagamentos 
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
         Left            =   -69000
         TabIndex        =   159
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lblTotalVale 
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
         Left            =   -69960
         TabIndex        =   146
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -71160
         TabIndex        =   145
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   -74880
         Top             =   5220
         Width           =   255
      End
      Begin VB.Label Label70 
         Caption         =   "Leitura Automática"
         Height          =   255
         Left            =   -74520
         TabIndex        =   139
         Top             =   5220
         Width           =   1815
      End
      Begin VB.Label Label54 
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
         Left            =   -69960
         TabIndex        =   121
         Top             =   5220
         Width           =   1815
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -71640
         TabIndex        =   120
         Top             =   5220
         Width           =   1605
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   118
         Top             =   5160
         Width           =   405
      End
      Begin VB.Label lblProdutoEntraTotal 
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
         Left            =   5040
         TabIndex        =   119
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label lblDespesas 
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
         Left            =   -69000
         TabIndex        =   117
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -70200
         TabIndex        =   116
         Top             =   4800
         Width           =   1005
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   114
         Top             =   5220
         Width           =   1965
      End
      Begin VB.Label lblTotalRecebido 
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
         Left            =   -70080
         TabIndex        =   115
         Top             =   5220
         Width           =   1695
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmFechamentoDeCaixaConfere.frx":AB8E
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmFechamentoDeCaixaConfere.frx":ABA9
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmFechamentoDeCaixaConfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodBar As String, Porta As Integer, Obrigatorio As String
Dim Abrindo As Boolean, Fechando As Boolean
Dim PlanoNotas As String, PlanoMicrocredito As String


Private Sub EncontraProduto(ByRef txtCodProduto As TextBox, cboProduto As ComboBox)
Dim CodigoProduto As Double, Preco As Currency
CodigoProduto = 0
With dbBicosEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from bicoencerrantes where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento & " and codigoproduto=" & CodigoProduto & " order by bico"
  .Refresh
End With
With dbProdutos2
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodProduto.Text = "" Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
    CodigoProduto = .Recordset!CodigoProduto
  End If
  If .Recordset!Combustivel = True Then
    lblBico.Visible = True
    cboBico.Visible = True
  Else
    lblBico.Visible = False
    cboBico.Visible = False
  End If
End With
With dbBicosEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from bicoencerrantes where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento & " and codigoproduto=" & CodigoProduto & " order by bico"
  .Refresh
  If .Recordset.EOF = False Then
    If cboBico.Text = "" Then
      cboBico.Text = .Recordset!Bico
    Else
      If IsNumeric(cboBico.Text) = True Then
        .Recordset.FindFirst "bico=" & cboBico.Text
        If .Recordset.NoMatch = True Then
          cboBico.Text = .Recordset!Bico
        End If
      Else
        cboBico.Text = .Recordset!Bico
      End If
    End If
  End If
End With
If dbClientes.Recordset.EOF = False Then
  Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes.Recordset!CodigoCliente)
  If IsNumeric(txtLitros.Text) = True Then
    Preco = Preco * CDbl(txtLitros.Text)
    txtNotaValor.Text = Preco
    Call txtNotaValor_LostFocus
  End If
End If
End Sub
Private Function PrecoCliente(ByVal CodigoProduto As Double, CodigoCliente As Double) As Currency
Dim Preco As Double, PrecoDif As Boolean
Dim db As New ADODB.Connection
Dim dbClientesProd As New ADODB.Recordset, Bico As Integer

db.Open CaminhoADO
dbClientesProd.CursorLocation = adUseClient
dbClientesProd.Open "select *from clientesprodutos", db, adOpenKeyset, adLockOptimistic

If dbClientesProd.RecordCount <> 0 Then
  If dbClientes.Recordset.EOF = False And dbClientes.Recordset.BOF = False Then
    dbClientesProd.Filter = "codigocliente=" & CodigoCliente & " and codigoproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
    If dbClientesProd.EOF = False Then
      If dbClientesProd!validade = dbFechamento.Recordset!DataCaixa Then
        If dbClientesProd!HoraIni >= dbFechamento.Recordset!HoraIni Then
          PrecoDif = True
        End If
      Else
        PrecoDif = True
      End If
    End If
    Preco = 0
    Bico = 0
    If cboBico.Visible = True Then
      If IsNumeric(cboBico.Text) = True Then
        Bico = CInt(cboBico.Text)
      End If
    End If
    
    Preco = PrecoAtual(CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, Bico)
    If PrecoDif = True Then
      
      If dbClientesProd!Preco <> 0 Then
        Preco = dbClientesProd!Preco
      Else
        If dbClientesProd!Porcento <> 0 Then
          Preco = Preco * dbClientesProd!Porcento
        End If
      End If
      If dbClientesProd!valorasomar <> 0 Then
        Preco = Preco + dbClientesProd!valorasomar
      End If
    Else
      If dbProdutos2.Recordset!Combustivel = True Then
        If IsNumeric(cboBico.Text) = True Then
          Preco = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, cboBico.Text)
        End If
      Else
        Preco = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno)
      End If
    End If
  End If
Else
  If dbProdutos2.Recordset!Combustivel = True Then
    If cboBico.Text = "" Then
      If dbBicosEncerrantes.Recordset.EOF = False And dbBicosEncerrantes.Recordset.BOF = False Then
        cboBico.Text = dbBicosEncerrantes.Recordset!Bico
      Else
        cboBico.Text = "0"
      End If
    End If
    Preco = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, cboBico.Text)
  Else
    Preco = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno)
  End If
End If

dbClientesProd.Close
db.Close

PrecoCliente = Preco
End Function

Private Sub ImprimeAutorizacao(ByVal YInicial As Double, ByVal Funcionario As String, ByVal Valor As Currency, ByVal NrNota As String, ByVal Cliente As String, ByVal DataCaixa As Date)
Dim StrTemp As String, Largura As String

Printer.ScaleMode = vbMillimeters

Largura = 190

Printer.CurrentY = YInicial
StrTemp = "Autorização de Débito"
Printer.FontSize = 16
Printer.FontBold = True
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 10
StrTemp = "São Paulo, " & Format(DataCaixa, "DD") & " de " & Format(DataCaixa, "MMMM") & " de " & Format(DataCaixa, "YYYY") & "."
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 10
StrTemp = "Eu, " & Funcionario & " funcionário do " & NomePosto & " autorizo o " & _
          "débito de " & Format(Valor, "Currency") & " (" & Extenso(Valor, "Reais", "Real") & _
          ") em meu pagamento de salário caso o não recebimento da nota nº " & NrNota & _
          " emitida para o cliente " & Cliente & " que conforme normas internas da firma, qualquer " & _
          "nota emitida a cliente bloqueado ou com limite excedido é de responsabilidade " & _
          "de que a tirou ao cliente."
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, 0, Printer.CurrentY, Largura

Printer.CurrentY = Printer.CurrentY + 20
Printer.DrawWidth = 2
Printer.Line (0, Printer.CurrentY)-(65, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Assinatura do Funcionário"
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = (65 / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

End Sub


Private Function CancelaPgFuncionarios() As Boolean
Dim db As New ADODB.Connection

CancelaPgFuncionarios = False

db.Open CaminhoADO
On Error GoTo 0
On Error Resume Next
db.Execute "update vendedorespagamento set confirmadonocaixa=0 where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
Do While Err.Number <> 0
  GoSub TrataErro
  On Error GoTo 0
  On Error Resume Next
  db.Execute "update vendedorespagamento set confirmadonocaixa=0 where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
Loop

CancelaPgFuncionarios = True

Exit Function

TrataErro:
  Resposta = MsgBox(Err.Number & " - " & Err.Description & Chr(vbKeyReturn) & "Deseja tentar de novo?", vbYesNo)
  If Resposta = vbYes Then Return
  
End Function

Private Function CancelaCompras() As Boolean
CancelaCompras = False
With dbProdutoEntra
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    dbProdutosNotas.Recordset.AddNew
    dbProdutosNotas.Recordset!codigofornecedor = 0
    dbProdutosNotas.Recordset!fornecedor = "Estorno Fechamento"
    dbProdutosNotas.Recordset!NrNota = dbFechamento.Recordset!CodigoFechamento
    dbProdutosNotas.Recordset!datalancada = Now
    dbProdutosNotas.Recordset!datanota = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!datanota = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!Origem = "Fechamento"
    dbProdutosNotas.Recordset!Confirmado = True
    dbProdutosNotas.Recordset!codigoPosto = dbPosto.Recordset!codigoPosto
    CodigoNota = dbProdutosNotas.Recordset!CodigoEntrada
    dbProdutosNotas.Recordset.Update
    
    Do While .Recordset.EOF = False
      If .Recordset!Tanque <> 0 Then
        dbTanques.Refresh
        dbTanques.Recordset.FindFirst "Tanque=" & .Recordset!Tanque
        If dbTanques.Recordset.NoMatch = False Then
          dbTanques.Recordset.Edit
          dbTanques.Recordset!Estoque = dbTanques.Recordset!Estoque - .Recordset!Quantidade
          dbTanques.Recordset.Update
        End If
      End If
      dbProdutos.Refresh
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = True Then
        MsgBox "Erro na tabela de produtos! Codigo produto: " & .Recordset!CodigoProduto
      End If
      'TempValor = (.Recordset!PrecoNovo - dbProdutos.Recordset!precocompra) * dbProdutos.Recordset!Estoque
      dbProdutos.Recordset.Edit
      'dbProdutos.Recordset!precocompra = .Recordset!PrecoNovo
      If IsNull(dbProdutos.Recordset!lucrominimo) = True Then
        dbProdutos.Recordset!lucrominimo = 0
      End If
      PrecoCompraNovo = .Recordset!PrecoNovo
      Venda = 0
      TempComissao = 0
      
      If IsNull(dbProdutos.Recordset!ValorEstoque) = True Then
        dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!precocompra * dbProdutos.Recordset!Estoque
      End If
      If IsNull(dbProdutos.Recordset!PrecoMedio) = True Then
        dbProdutos.Recordset!PrecoMedio = dbProdutos.Recordset!precocompra
      End If
      If IsNull(dbProdutos.Recordset!DifEstoque) = True Then
        dbProdutos.Recordset!DifEstoque = 0
      End If
      If IsNull(dbProdutos.Recordset!valordifestoque) = True Then
        dbProdutos.Recordset!valordifestoque = 0
      End If
      If IsNull(dbProdutos.Recordset!LucroMedio) = True Then
        dbProdutos.Recordset!LucroMedio = 0
      End If
      
      ValorProduto = .Recordset!valornota
      dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque - ValorProduto
      
      dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque - .Recordset!Quantidade
      dbProdutos.Recordset!qtdcomprado = dbProdutos.Recordset!qtdcomprado - .Recordset!Quantidade
      dbProdutos.Recordset!valorcomprado = dbProdutos.Recordset!valorcomprado - .Recordset!valornota
      dbProdutos.Recordset.Update
      With dbProdutosNotasCorpo
        .Recordset.AddNew
        .Recordset!codigoprodutonota = CodigoNota
        .Recordset!CodigoProduto = dbProdutoEntra.Recordset!CodigoProduto
        .Recordset!Codigo = dbProdutoEntra.Recordset!Codigo
        .Recordset!Descri = dbProdutoEntra.Recordset!Descri
        .Recordset!valorUnitario = dbProdutoEntra.Recordset!PrecoNovo
        .Recordset!Quantidade = -dbProdutoEntra.Recordset!Quantidade
        .Recordset!Total = -dbProdutoEntra.Recordset!valornota
        .Recordset!Tanque = dbProdutoEntra.Recordset!Tanque
        .Recordset.Update
      End With
      With dbDespesasLanc
        .Recordset.AddNew
        .Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
        .Recordset!Origem = "Estorno Fechamento"
        .Recordset!Data = dbFechamento.Recordset!DataCaixa
        .Recordset!Hora = Now
        .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
        .Recordset!CodigoDespesa = -1
        .Recordset!Descri = "Estorno Compra de Produto"
        .Recordset!Obs = Right("Estorno " & dbProdutos.Recordset!Descri & " " & dbProdutoEntra.Recordset!Quantidade & " - " & dbProdutoEntra.Recordset!valornota, 50)
        .Recordset!Valor = dbProdutoEntra.Recordset!valornota
        .Recordset!Produto = True
        .Recordset!compensado = True
        .Recordset.Update
        .Refresh
      End With
      .Recordset.MoveNext
    Loop
  End If
End With
CancelaCompras = True

End Function

Private Function CancelaVales() As Boolean
CancelaVales = False
Fechando = True
On Error GoTo TrataErro
'Totaliza os vales e confirma para o status
With dbVales
  .RecordSource = "Select *from vales where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!Cobrado = False Then
        .Recordset.Edit
        .Recordset!fechado = False
        .Recordset.Update
      Else
        MsgBox "Existe Vale que deverá ser estornado manualmente!"
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

CancelaVales = True
'AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
SSTab1.Tab = 5
Fechando = False
Exit Function

TrataErro:
Fechando = False
CancelaVales = False
End Function


Private Function CancelaDespesas() As Boolean
Dim db As New ADODB.Connection

Fechando = True
'Totaliza as despesas para lançar no status e no saldo das contas
db.Open CaminhoADO
db.Execute "update despesaslanc2 set fechamentodiario=0 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Execute "update fechamentodecaixa set totaldespesas=" & NumeroIngles(-CCur(lblTotalDespesas.Caption)) & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Close

'AbreFechamento .Recordset!CodigoFechamento, .Recordset!DataCaixa
Fechando = False
CancelaDespesas = True
SSTab1.Tab = 1
End Function


Private Function CancelaPagamentos() As Boolean
Dim db As New ADODB.Connection
Dim TotalRecebido As Currency, DataBordero As Date, ReceberData As Date
CancelaPagamentos = True

Fechando = True

'Totaliza os recebimentos e lança nas contas
With dbFormaDePgRecebido
  .RecordSource = "select *from qFormadePgRecebidofechamento2 where fechamentodecaixa.codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    TotalRecebido = 0
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = True Then
        TotalRecebido = TotalRecebido + .Recordset("valor")
        Dias = .Recordset("reembolso")
        Mes = .Recordset("mes")
        txtDataBordero.Value = .Recordset!Data
        If Mes = True Then
          Intervalo = "m"
        Else
          Intervalo = "d"
        End If
        If Dias > 0 Then
          DataBordero = txtDataBordero.Value
          If .Recordset!corte = True Then
            DataBordero = DataDeCorte(.Recordset!datacorte, .Recordset!diascorte, DataBordero)
          End If
          ReceberData = DateAdd(Intervalo, Dias, DataBordero)
'        Else
'          Dias = .Recordset("diadomes")
'          If Dias > 0 Then
'            If Dias >= txtData.Day Then
'              If Dias < 28 Then
'                StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
'              Else
'                StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
'                Do While IsDate(StrTemp) = False
'                  Dias = Dias - 1
'                  If Dias <= 0 Then Dias = 31
'                  StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
'                Loop
'              End If
'              ReceberData = CDate(StrTemp)
'            Else
'              If Dias < 28 Then
'                StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
'              Else
'                StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
'                Do While IsDate(StrTemp) = False
'                  Dias = Dias - 1
'                  If Dias <= 0 Then Dias = 31
'                  StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
'                Loop
'              End If
'              ReceberData = CDate(StrTemp)
'            End If
'          End If
        End If
        If Dias > 0 Then
          Select Case Weekday(ReceberData)
            Case 1 'domingo
              ReceberData = DateAdd("d", 1, ReceberData)
            Case 7 'sábado
              ReceberData = DateAdd("d", 2, ReceberData)
          End Select
          dbCartoes.Refresh
          On Error Resume Next
          NaoAcumula = dbFormaDePgRecebido.Recordset!NaoAcumula
          If Err.Number <> 0 Then
            NaoAcumula = False
          End If
          On Error GoTo 0
          If dbCartoes.Recordset.EOF = False Then
            If NaoAcumula = False Then
              dbCartoes.Refresh
              dbCartoes.Recordset.FindFirst "codigoformapg=" & .Recordset!CodigoFormadePg & " and dataprevista=#" & DataInglesa(Trim(Str(ReceberData))) & "# and confirmado=0 and fechataxa=0"
              If dbCartoes.Recordset.NoMatch = False Then
                dbCartoes.Recordset.Edit
              Else
                dbCartoes.Recordset.AddNew
                dbCartoes.Recordset!ValorBruto = 0
                dbCartoes.Recordset!valorliquido = 0
              End If
            Else
              dbCartoes.Recordset.AddNew
              dbCartoes.Recordset!ValorBruto = 0
              dbCartoes.Recordset!valorliquido = 0
            End If
          Else
            dbCartoes.Recordset.AddNew
            dbCartoes.Recordset!ValorBruto = 0
            dbCartoes.Recordset!valorliquido = 0
          End If
          dbContas.Recordset.FindFirst "codigoconta=" & .Recordset!CodigoConta
          dbCartoes.Recordset!CodigoConta = dbContas.Recordset("codigoconta")
          dbCartoes.Recordset!Conta = dbContas.Recordset("descri")
          dbCartoes.Recordset!CodigoFormaPg = .Recordset!CodigoFormadePg
          dbCartoes.Recordset!Grupo = .Recordset!Grupo
          dbCartoes.Recordset!Descri = .Recordset("formadepagamento.descri")
          dbCartoes.Recordset!DataLanc = DataBordero
          dbCartoes.Recordset!DataPrevista = ReceberData
          dbCartoes.Recordset!ValorBruto = dbCartoes.Recordset!ValorBruto - dbFormaDePgRecebido.Recordset!ValorBruto
          dbCartoes.Recordset!valorliquido = dbCartoes.Recordset!valorliquido - dbFormaDePgRecebido.Recordset!Valor
          dbCartoes.Recordset.Update
          
        Else
          DataBordero = txtDataBordero.Value
          If .Recordset!datacorte = True Then
            DataBordero = DataDeCorte(.Recordset!datacorte, .Recordset!diascorte, DataBordero)
          End If
          ReceberData = DataBordero
          Select Case Weekday(ReceberData)
            Case 1 'domingo
              ReceberData = DateAdd("d", 1, ReceberData)
            Case 7 'sábado
              ReceberData = DateAdd("d", 2, ReceberData)
          End Select
          
          With dbConciliaNova
            .Recordset.AddNew
            .Recordset!CodigoConta = dbFormaDePgRecebido.Recordset("codigoconta")
            .Recordset!DataLanc = Now
            .Recordset!compensado = True
            .Recordset!Data = Date
            .Recordset!Tipo = "Fechamento"
            .Recordset!Codigo = 999999998
            .Recordset!Descri = Left("Estorno Cx - " & Format(dbFechamento.Recordset!DataCaixa, "short date") & " - " & dbFechamento.Recordset!Turno & " - " & dbFormaDePgRecebido.Recordset("formadepagamento.descri"), 50)
            .Recordset!NrDocumento = Format(txtDataBordero.Value, "short date")
            .Recordset!Valor = -dbFormaDePgRecebido.Recordset!Valor
            .Recordset.Update
          End With
          dbContas.Refresh
          dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
          If dbContas.Recordset.NoMatch = True Then
            MsgBox "Conta para " & .Recordset("formadepagamento.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
          Else
            TempValor = -.Recordset("valor")
            dbContas.Recordset.Edit
            dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + TempValor
            dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
            dbContas.Recordset.Update
          End If
        End If
        .Recordset.Edit
        .Recordset!fechamentodiario = False
        .Recordset.Update
      End If
      
      .Recordset.MoveNext
    Loop
  End If
  .RecordSource = "select *from FormadePagamentoRecebido2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
End With

db.Open CaminhoADO
db.Execute "update fechamentodecaixa set totalrecebido=totalrecebido-" & NumeroIngles(TotalRecebido) & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Close

Fechando = False
End Function


Private Sub AlertaAtivo(ByVal Ativo As Boolean)
Dim Cor As Long
If Ativo = True Then
  lblStatus.Caption = "Ativo"
  Cor = vbButtonText
Else
  lblStatus.Caption = "Inativo"
  Cor = vbRed
End If
lblStatus.ForeColor = Cor
For i = 0 To 4
  MaskEdBox1(i).ForeColor = Cor
Next i
txtValor.ForeColor = Cor
txtCodCliente.ForeColor = Cor
cboClienteCheque.ForeColor = Cor
End Sub

Private Function GravaPgFuncionarios() As Boolean
Dim db As New ADODB.Connection

GravaPgFuncionarios = False

db.Open CaminhoADO
On Error GoTo 0
On Error Resume Next
db.Execute "update vendedorespagamento set confirmadonocaixa=0 where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
Do While Err.Number <> 0
  GoSub TrataErro
  On Error GoTo 0
  On Error Resume Next
  db.Execute "update vendedorespagamento set confirmadonocaixa=-1 where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
Loop


With dbPagamentosCaixa
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      CodigoPagamento = .Recordset!CodigoPagamento
      
      'db.Execute "update despesaslanc2 set fechamento=0 where origem='Pg Funcionários' and nrdocumento='" & CodigoPagamento & "' and datafechamento=null"
      .Recordset.Edit
      .Recordset!confirmadonocaixa = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
GravaPgFuncionarios = True
Exit Function

TrataErro:
  Resposta = MsgBox(Err.Number & " - " & Err.Description & Chr(vbKeyReturn) & "Deseja tentar de novo?", vbYesNo)
  If Resposta = vbYes Then Return

End Function

Private Function GravaCompras() As Boolean
Dim ValorProduto As Currency

GravaCompras = False
With dbProdutoEntra
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    dbProdutosNotas.Recordset.AddNew
    dbProdutosNotas.Recordset!codigofornecedor = 0
    dbProdutosNotas.Recordset!fornecedor = "Fechamento"
    dbProdutosNotas.Recordset!NrNota = dbFechamento.Recordset!CodigoFechamento
    dbProdutosNotas.Recordset!datalancada = Now
    dbProdutosNotas.Recordset!datanota = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!datanota = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
    dbProdutosNotas.Recordset!Origem = "Fechamento"
    dbProdutosNotas.Recordset!Confirmado = True
    dbProdutosNotas.Recordset!codigoPosto = dbPosto.Recordset!codigoPosto
    CodigoNota = dbProdutosNotas.Recordset!CodigoEntrada
    dbProdutosNotas.Recordset.Update
    
    Do While .Recordset.EOF = False
      If .Recordset!Tanque <> 0 Then
        dbTanques.Refresh
        dbTanques.Recordset.FindFirst "Tanque=" & .Recordset!Tanque
        If dbTanques.Recordset.NoMatch = False Then
          dbTanques.Recordset.Edit
          dbTanques.Recordset!Estoque = dbTanques.Recordset!Estoque + .Recordset!Quantidade
          dbTanques.Recordset.Update
        End If
      End If
      dbProdutos.Refresh
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = True Then
        MsgBox "Erro na tabela de produtos! Codigo produto: " & .Recordset!CodigoProduto
      End If
      TempValor = (.Recordset!PrecoNovo - dbProdutos.Recordset!precocompra) * dbProdutos.Recordset!Estoque
      dbProdutos.Recordset.Edit
      
      If IsNull(dbProdutos.Recordset!ValorEstoque) = True Then
        dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!precocompra * dbProdutos.Recordset!Estoque
      End If
      If IsNull(dbProdutos.Recordset!PrecoMedio) = True Then
        dbProdutos.Recordset!PrecoMedio = dbProdutos.Recordset!precocompra
      End If
      If IsNull(dbProdutos.Recordset!DifEstoque) = True Then
        dbProdutos.Recordset!DifEstoque = 0
      End If
      If IsNull(dbProdutos.Recordset!valordifestoque) = True Then
        dbProdutos.Recordset!valordifestoque = 0
      End If
      If IsNull(dbProdutos.Recordset!LucroMedio) = True Then
        dbProdutos.Recordset!LucroMedio = 0
      End If
      
      ValorProduto = .Recordset!valornota
      dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque + ValorProduto
      
      dbProdutos.Recordset!precocompra = .Recordset!PrecoNovo
      If IsNull(dbProdutos.Recordset!lucrominimo) = True Then
        dbProdutos.Recordset!lucrominimo = 0
      End If
      PrecoCompraNovo = .Recordset!PrecoNovo
      Venda = 0
      TempComissao = 0
'      If dbProdutos.Recordset!lucrominimo <> 0 Then
'        Venda = PrecoCompraNovo + (PrecoCompraNovo * (dbProdutos.Recordset!lucrominimo / 100))
'        If dbProdutos.Recordset!Comissao <> 0 Then
'          TempComissao = Venda / (1 - dbProdutos.Recordset!Comissao)
'          Venda = TempComissao
'        End If
'        If dbProdutos.Recordset!comissaovalor <> 0 Then
'          Venda = Venda + dbProdutos.Recordset!comissaovalor
'        End If
'      Else
'        Venda = dbProdutos.Recordset!PrecoVenda
'      End If
'      If Venda <> dbProdutos.Recordset!PrecoVenda Then
'TentaDeNovo:
'        strTemp = InputBox("O produto " & dbProdutos.Recordset!Codigo & " - " & dbProdutos.Recordset!Descri & " está sendo alterado o preço de venda de " & Format(dbProdutos.Recordset!PrecoVenda, "Currency") & " para " & Format(Venda, "Currency") & "!", "Alteração de preço!", Format(Venda, "Currency"))
'        Do While IsNumeric(strTemp) = False
'          DoEvents
'          strTemp = InputBox("O produto " & dbProdutos.Recordset!Codigo & " - " & dbProdutos.Recordset!Descri & " está sendo alterado o preço de venda de " & Format(dbProdutos.Recordset!PrecoVenda, "Currency") & " para " & Format(Venda, "Currency") & "!", "Alteração de preço!", Format(Venda, "Currency"))
'        Loop
'        If Venda < (CCur(strTemp) - 0.5) Or Venda > (CCur(strTemp) + 0.5) Then
'          Permissao = False
'          frmPermissao.Show vbModal
'          If Permissao = False Then
'            GoTo TentaDeNovo
'          Else
'            Venda = CCur(strTemp)
'          End If
'        Else
'          Venda = CCur(strTemp)
'        End If
'
'        dbProdutos.Recordset!PrecoVenda = Venda
'      End If
      dbProdutos.Recordset!Variacao = dbProdutos.Recordset!Variacao + TempValor
      dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque + .Recordset!Quantidade
      If IsNull(dbProdutos.Recordset!qtdcomprado) = True Then
        dbProdutos.Recordset!qtdcomprado = 0
      End If
      If IsNull(dbProdutos.Recordset!valorcomprado) = True Then
        dbProdutos.Recordset!valorcomprado = 0
      End If
      dbProdutos.Recordset!qtdcomprado = dbProdutos.Recordset!qtdcomprado + .Recordset!Quantidade
      dbProdutos.Recordset!valorcomprado = dbProdutos.Recordset!valorcomprado + .Recordset!valornota
      dbProdutos.Recordset.Update
      With dbProdutosNotasCorpo
        .Recordset.AddNew
        .Recordset!codigoprodutonota = CodigoNota
        .Recordset!CodigoProduto = dbProdutoEntra.Recordset!CodigoProduto
        .Recordset!Codigo = dbProdutoEntra.Recordset!Codigo
        .Recordset!Descri = dbProdutoEntra.Recordset!Descri
        .Recordset!valorUnitario = dbProdutoEntra.Recordset!PrecoNovo
        .Recordset!Quantidade = dbProdutoEntra.Recordset!Quantidade
        .Recordset!Total = dbProdutoEntra.Recordset!valornota
        .Recordset!Tanque = dbProdutoEntra.Recordset!Tanque
        .Recordset.Update
      End With
      With dbDespesasLanc
        .Recordset.AddNew
        .Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
        .Recordset!Origem = "Fechamento"
        .Recordset!Data = dbFechamento.Recordset!DataCaixa
        .Recordset!Hora = Now
        .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
        .Recordset!CodigoDespesa = -1
        .Recordset!Descri = "Compra de Produto"
        .Recordset!Obs = Right(dbProdutos.Recordset!Descri & " " & dbProdutoEntra.Recordset!Quantidade & " - " & dbProdutoEntra.Recordset!valornota, 50)
        .Recordset!Valor = -dbProdutoEntra.Recordset!valornota
        .Recordset!Produto = True
        .Recordset!compensado = True
        .Recordset.Update
        .Refresh
      End With
      .Recordset.MoveNext
    Loop
  End If
End With
GravaCompras = True
End Function

Private Function GravaVales() As Boolean
GravaVales = False
Fechando = True
If FechamentoBloqueado = True Then
  GravaVales = False
  Fechando = False
  Exit Function
End If
On Error GoTo TrataErro
'Totaliza os vales e confirma para o status
With dbVales
  .RecordSource = "Select *from vales where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechado = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

GravaVales = True
AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
SSTab1.Tab = 5
Fechando = False
Exit Function

TrataErro:
Fechando = False
GravaVales = False
End Function

Private Function GravaDespesas() As Boolean
Dim db As New ADODB.Connection

If FechamentoBloqueado = True Then
  GravaDespesas = False
  Fechando = False
  Exit Function
End If

Fechando = True
'Totaliza as despesas para lançar no status e no saldo das contas
db.Open CaminhoADO
db.Execute "update despesaslanc2 set fechamentodiario=-1 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Execute "update fechamentodecaixa set totaldespesas=" & NumeroIngles(CCur(lblTotalDespesas.Caption)) & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Close

AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
Fechando = False
GravaDespesas = True
SSTab1.Tab = 1

End Function

Private Function GravaPagamentos() As Boolean
Dim db As New ADODB.Connection
Dim TotalRecebido As Currency, DataBordero As Date, ReceberData As Date
Dim CodigoCartao As Double
GravaPagamentos = True

Fechando = True
If FechamentoBloqueado = True Then
  GravaPagamentos = False
  Fechando = False
  Exit Function
End If

With dbFormaDePgRecebido
  StrTemp = .RecordSource
  .RecordSource = "select *from QFormadePgContasRecebido2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("formadepagamento.codigoconta")
      If dbContas.Recordset.NoMatch = True Then
        MsgBox "Existe forma de pagamento sem conta destino! Verifique o cadastro de forma de pagamento!"
        Fechando = False
        GravaPagamentos = False
        Exit Function
      End If
      .Recordset.MoveNext
    Loop
  End If
  .RecordSource = StrTemp
  .Refresh
End With

'Totaliza os recebimentos e lança nas contas
With dbFormaDePgRecebido
  .RecordSource = "select *from qFormadePgRecebidofechamento2 where fechamentodecaixa.codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    TotalRecebido = 0
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = False Then
        TotalRecebido = TotalRecebido + .Recordset("valor")
        Dias = .Recordset("reembolso")
        Mes = .Recordset("mes")
        txtDataBordero.Value = .Recordset!Data
        If Mes = True Then
          Intervalo = "m"
        Else
          Intervalo = "d"
        End If
        If Dias > 0 Then
          DataBordero = txtDataBordero.Value
          If IsNull(.Recordset!corte) = False Then
            If .Recordset!corte = True Then
              DataBordero = DataDeCorte(.Recordset!datacorte, .Recordset!diascorte, DataBordero)
            End If
          End If
          ReceberData = DateAdd(Intervalo, Dias, DataBordero)
        End If
        If Dias > 0 Then
          Select Case Weekday(ReceberData)
            Case 1 'domingo
              ReceberData = DateAdd("d", 1, ReceberData)
            Case 7 'sábado
              ReceberData = DateAdd("d", 2, ReceberData)
          End Select
          dbCartoes.Refresh
          On Error Resume Next
          NaoAcumula = dbFormaDePgRecebido.Recordset!NaoAcumula
          If Err.Number <> 0 Then
            NaoAcumula = False
          End If
          On Error GoTo 0
          If dbCartoes.Recordset.EOF = False Then
            If NaoAcumula = False Then
              dbCartoes.Recordset.FindFirst "codigoformapg=" & .Recordset!CodigoFormadePg & " and dataprevista=#" & DataInglesa(Trim(Str(ReceberData))) & "# and confirmado=0"
              If dbCartoes.Recordset.NoMatch = False Then
                dbCartoes.Recordset.Edit
              Else
                dbCartoes.Recordset.AddNew
                dbCartoes.Recordset!ValorBruto = 0
                dbCartoes.Recordset!valorliquido = 0
              End If
            Else
              dbCartoes.Recordset.AddNew
              dbCartoes.Recordset!ValorBruto = 0
              dbCartoes.Recordset!valorliquido = 0
            End If
          Else
            dbCartoes.Recordset.AddNew
            dbCartoes.Recordset!ValorBruto = 0
            dbCartoes.Recordset!valorliquido = 0
          End If
          
          dbContas.Recordset.FindFirst "codigoconta=" & .Recordset!CodigoConta
          dbCartoes.Recordset!CodigoConta = dbContas.Recordset("codigoconta")
          dbCartoes.Recordset!Conta = dbContas.Recordset("descri")
          dbCartoes.Recordset!CodigoFormaPg = .Recordset!CodigoFormadePg
          dbCartoes.Recordset!Grupo = .Recordset!Grupo
          dbCartoes.Recordset!Descri = .Recordset("formadepagamento.descri")
          dbCartoes.Recordset!DataLanc = DataBordero
          dbCartoes.Recordset!DataPrevista = ReceberData
          dbCartoes.Recordset!ValorBruto = dbCartoes.Recordset!ValorBruto + dbFormaDePgRecebido.Recordset!ValorBruto
          dbCartoes.Recordset!valorliquido = dbCartoes.Recordset!valorliquido + dbFormaDePgRecebido.Recordset!Valor
          CodigoCartao = dbCartoes.Recordset!CodigoCartao
          dbCartoes.Recordset.Update
        Else
          DataBordero = txtDataBordero.Value
          If IsNull(.Recordset!corte) = False Then
            If .Recordset!corte = True Then
              DataBordero = DataDeCorte(.Recordset!datacorte, .Recordset!diascorte, DataBordero)
            End If
          End If
          ReceberData = DataBordero
          
          Select Case Weekday(ReceberData)
            Case 1 'domingo
              ReceberData = DateAdd("d", 1, ReceberData)
            Case 7 'sábado
              ReceberData = DateAdd("d", 2, ReceberData)
          End Select
          
          With dbConciliaNova
            .Recordset.AddNew
            .Recordset!CodigoConta = dbFormaDePgRecebido.Recordset("codigoconta")
            .Recordset!DataLanc = Now
            .Recordset!compensado = True
            .Recordset!Data = ReceberData
            .Recordset!Tipo = "Fechamento"
            .Recordset!Codigo = 999999998
            .Recordset!Descri = Left("Caixa - " & Format(dbFechamento.Recordset!DataCaixa, "short date") & " - " & dbFechamento.Recordset!Turno & " - " & dbFormaDePgRecebido.Recordset("formadepagamento.descri"), 50)
            .Recordset!NrDocumento = Format(txtDataBordero.Value, "short date")
            .Recordset!Valor = dbFormaDePgRecebido.Recordset!Valor
            .Recordset.Update
          End With
          dbContas.Refresh
          dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
          If dbContas.Recordset.NoMatch = True Then
            MsgBox "Conta para " & .Recordset("formadepagamento.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
          Else
            TempValor = .Recordset("valor")
            dbContas.Recordset.Edit
            dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + TempValor
            dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
            dbContas.Recordset.Update
          End If
        End If
        .Recordset.Edit
        .Recordset!fechamentodiario = True
        .Recordset("formadepagamentorecebido2.Confirma") = CodigoCartao
        .Recordset("formadepagamentorecebido2.fechames") = False
        .Recordset.Update
        CodigoCartao = -1
      End If
      
      .Recordset.MoveNext
    Loop
  End If
  .RecordSource = "select *from FormadePagamentoRecebido2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  .Refresh
End With

db.Open CaminhoADO
db.Execute "update fechamentodecaixa set totalrecebido=totalrecebido+" & NumeroIngles(TotalRecebido) & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
db.Close

AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa

Fechando = False
End Function

Private Function FechamentoBloqueado() As Boolean
FechamentoBloqueado = False
With dbBloqueiaFechamento
  If .Recordset.RecordCount <> 0 Then
    If .Recordset!bloqueia2 = True Then
      If .Recordset!Data2 <= dbFechamento.Recordset!DataCaixa Then
        If .Recordset!HoraIni <= dbFechamento.Recordset!HoraIni Then
          MsgBox "Caixa não pode ser confirmado por estar bloqueado pelo administrador!"
          FechamentoBloqueado = True
        End If
      End If
    End If
  End If
End With
End Function

Private Sub TotalizaCheque()
Dim Total As Currency, TotalBomba As Currency
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

On Error Resume Next
db.Close
dbTemp.Close
On Error GoTo 0
db.Open CaminhoADO
dbTemp.Open "select sum(valornabomba) as TotalBomba, sum(valor) as Total from cheques where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento, db

If IsNull(dbTemp!Total) = False Then
  Total = dbTemp!Total
Else
  Total = 0
End If
If IsNull(dbTemp!TotalBomba) = False Then
  TotalBomba = dbTemp!TotalBomba
Else
  TotalBomba = 0
End If
Label54.Caption = Format(Total, "Currency")
txtChequeBomba.Text = Format(TotalBomba, "Currency")
txtChequeJuros.Text = Format(Total, "Currency")
lblJuros.Caption = Format((Total - TotalBomba), "Currency")

On Error Resume Next
db.Execute "update Fechamentodecaixa set chequebomba=" & NumeroIngles(TotalBomba) & ", chequejuros=" & NumeroIngles(Total) & ", juros=" & NumeroIngles(Total - TotalBomba) & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
On Error GoTo 0

dbTemp.Close
db.Close
End Sub
Private Sub TotalRecebimento()
Dim TotalRecebido As Currency
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
TotalRecebido = 0

TotalizaCheque

If IsNumeric(txtChequeJuros.Text) = True Then
  TotalRecebido = TotalRecebido + CCur(txtChequeJuros.Text)
End If

On Error Resume Next
db.Close
dbTemp.Close
On Error GoTo 0
db.Open CaminhoADO
dbTemp.Open "select sum(valorbruto) as total from FormaDePagamentoRecebido2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento, db
If IsNumeric(dbTemp!Total) = True Then
  TotalRecebido = TotalRecebido + dbTemp!Total
End If

dbTemp.Close
db.Close
lblTotalRecebido.Caption = Format(TotalRecebido, "Currency")
End Sub

Private Sub TotalResumo()
Dim TotalResumo As Currency, TempValor As Currency, TotalDespesas As Currency
TempValor = 0
With dbFechamento
  If .Recordset.EOF = False Then
    TempValor = .Recordset!TotalCombustivel + .Recordset!TotalProdutos
  End If
End With
lblTotalVendas.Caption = Format(TempValor, "Currency")
If IsNumeric(lblJuros.Caption) = True Then
  TempValor = TempValor + CCur(lblJuros.Caption)
End If

TotalResumo = TempValor

TempValor = 0
If IsNumeric(lblTotalRecebido.Caption) = True Then
  TempValor = CCur(lblTotalRecebido.Caption)
End If
If IsNumeric(lblTotalNota.Caption) = True Then
  TempValor = TempValor + CCur(lblTotalNota.Caption)
End If
If IsNumeric(lblJurosNota.Caption) = True Then
  TempValor = TempValor - CCur(lblJurosNota.Caption)
End If
If IsNumeric(lblProdutoEntraTotal.Caption) = True Then
  TempValor = TempValor + CCur(lblProdutoEntraTotal.Caption)
End If

lblRecebimentos.Caption = Format(TempValor, "Currency")

If IsNumeric(lblTotalValeResumo.Caption) = True Then
  TempValor = TempValor + CCur(lblTotalValeResumo.Caption)
End If

TotalDespesas = 0
If IsNumeric(lblTotalPagamentos.Caption) = True Then
  TotalDespesas = -CCur(lblTotalPagamentos.Caption)
End If
If IsNumeric(lblComissoesPagas.Caption) = True Then
  TotalDespesas = TotalDespesas - CCur(lblComissoesPagas.Caption)
End If
If IsNumeric(lblComissoesCombustiveis.Caption) = True Then
  TotalDespesas = TotalDespesas - CCur(lblComissoesCombustiveis.Caption)
End If


If IsNumeric(lblDespesas.Caption) Then
  TotalDespesas = TotalDespesas + CCur(lblDespesas.Caption)
End If
lblTotalDespesas.Caption = Format(TotalDespesas, "currency")

TempValor = TempValor - TotalDespesas - TotalResumo

lblDiferenca.Caption = Format(TempValor, "Currency")

End Sub

Private Sub TotalizaCompras()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim TotalCompra As Currency
TotalCompra = 0

On Error Resume Next
db.Close
dbTemp.Close
On Error GoTo 0
db.Open CaminhoADO
dbTemp.Open "select sum(valornota) as total from produtosentrada2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento, db

If IsNull(dbTemp!Total) = False Then
  lblProdutoEntraTotal.Caption = Format(dbTemp!Total, "currency")
Else
  lblProdutoEntraTotal.Caption = Format(0, "currency")
End If

dbTemp.Close
db.Close
End Sub

Private Sub TotalizaNotas()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim TotalNota As Currency, JurosNota As Currency, TotalNaBomba As Currency
TotalNota = 0
JurosNota = 0
TotalNaBomba = 0

On Error Resume Next
dbTemp.Close
On Error GoTo 0
db.Open CaminhoADO
dbTemp.Open "select sum(lucrodif) as DifLucro, sum(valortotaldif) as TotalDif, sum(ValorPrevisto) as total from clientesnota2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento, db
If IsNull(dbTemp!diflucro) = False Then
  JurosNota = dbTemp!diflucro
End If
If IsNull(dbTemp!totaldif) = False Then
  TotalNaBomba = dbTemp!totaldif
End If
If IsNull(dbTemp!Total) = False Then
  TotalNota = dbTemp!Total
End If

dbTemp.Close
db.Close

lblTotalNota.Caption = Format(TotalNota, "Currency")
lblJurosNota.Caption = Format(JurosNota, "Currency")
lblTotalNabomba.Caption = Format(TotalNaBomba, "currency")
End Sub

Private Sub TotalizaDespesas()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim TotalDesp As Currency
TotalDesp = 0

On Error Resume Next
dbTemp.Close
db.Close
On Error GoTo 0
db.Open CaminhoADO
dbTemp.Open "select sum(valor) as Total from DespesasLanc2 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento & " and produto=0", db
If IsNull(dbTemp!Total) = False Then
  TotalDesp = dbTemp!Total
End If
lblDespesas.Caption = Format(TotalDesp, "Currency")

On Error Resume Next
dbTemp.Close
On Error GoTo 0
dbTemp.Open "select sum(saldoapagar) as total from vendedorespagamento where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento, db
If IsNull(dbTemp!Total) = False Then
  TotalDesp = TotalDesp + dbTemp!Total
  lblTotalPagamentos.Caption = Format(dbTemp!Total, "Currency")
Else
  lblTotalPagamentos.Caption = Format(0, "Currency")
End If

dbTemp.Close
db.Close

lblTotalDespesas.Caption = Format(TotalDesp, "Currency")

End Sub

Private Sub TotalizaVales()
Dim TotalVales As Currency, Codigo As Double
If IsNull(dbFechamento.Recordset!CodigoFechamento) = False Then
  Codigo = dbFechamento.Recordset!CodigoFechamento
Else
  Codigo = 0
End If
TotalDesp = 0

With qValesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select sum(valor) as total from vales where codigocaixa=" & Codigo
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalVale.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalVale.Caption = Format(0, "Currency")
  End If
  lblTotalValeResumo.Caption = lblTotalVale.Caption
End With
End Sub

Private Sub AbreFechamento(ByVal CodigoFechamento As Double, ByVal Dia As Date)
Dim db As New ADODB.Connection
Dim TempComissoes As Currency

Abrindo = True
txtDataBordero.Value = Dia
txtBomPara.Value = Dia
SSTab1.Tab = 0
With dbFormaDePgRecebido
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from FormaDePagamentoRecebido2 where codigofechamento=" & CodigoFechamento & " order by descri"
  .Refresh
End With
With dbDespesasLanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from DespesasLanc2 where produto=0 and codigoconta=-1 and codigofechamento=" & CodigoFechamento & " order by descri"
  .Refresh
End With
With dbClientesNota
  .Connect = Conectar
  .DatabaseName = Caminho
  If PlanoNotas <> "" Then
    .RecordSource = "select *from clientesnota2 where codigofechamento=" & CodigoFechamento & " and planodeconta in (" & PlanoNotas & ") order by Nome"
  Else
    .RecordSource = "select *from clientesnota2 where codigofechamento=" & CodigoFechamento & " order by Nome"
  End If
  .Refresh
End With
With dbClientesNota2
  .Connect = Conectar
  .DatabaseName = Caminho
  If PlanoMicrocredito <> "" Then
    .RecordSource = "select *from clientesnota2 where codigofechamento=" & CodigoFechamento & " and planodeconta in (" & PlanoNotas & ") order by Nome"
  Else
    .RecordSource = "select *from clientesnota2 where codigofechamento=" & CodigoFechamento & " order by Nome"
  End If
  .Refresh
End With
With dbProdutoEntra
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from produtosentrada2 where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from vales where codigocaixa=" & CodigoFechamento
  .Refresh
End With
With qVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from qvales where codigocaixa=" & CodigoFechamento
  .Refresh
End With
With dbProdutosAltera
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from qprodutosaltera where datacaixa<=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "# order by datacaixa, horaini"
  .Refresh
End With
With dbBicosEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from bicoencerrantes where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbPagamentosCaixa
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from vendedorespagamento where codigocaixa=" & CodigoFechamento & " order by funcionario"
  .Refresh
End With
With dbVendas
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select sum(valorcomissao) as total from venda2 where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbFechamento2
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select codigofechamento, notaconferida, notaconferida2 from fechamentodecaixa where codigofechamento=" & CodigoFechamento
  .Refresh
End With

lblComissoesPagas.Caption = Format(0, "Currency")
If ComissaoAcumulativa = False Then
  If IsNull(dbVendas.Recordset!Total) = False Then
    lblComissoesPagas.Caption = Format(dbVendas.Recordset!Total, "Currency")
  End If
End If

TempComissoes = 0
With dbBicosEncerrantes
    If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
            If IsNull(.Recordset!Comissao) = False Then
                TempComissoes = TempComissoes + .Recordset!Comissao
            End If
            .Recordset.MoveNext
        Loop
        .Recordset.MoveFirst
    End If
End With
lblComissoesCombustiveis.Caption = "R$ " & Format(TempComissoes, "0.0000")


With dbFechamento
  
  txtChequeBomba.Text = Format(.Recordset!chequebomba, "Currency")
  txtChequeJuros.Text = Format(.Recordset!chequejuros, "Currency")
  lblJuros.Caption = Format(.Recordset!Juros, "Currency")
  If .Recordset!distribuido = True Then
    cmdFinalizar.Visible = False
    For i = 0 To Tela.Count - 1
      Tela(i).Enabled = False
    Next i
    DBGrid2.AllowDelete = False
    DBGrid3.AllowDelete = False
    DBGrid4.AllowDelete = False
    DBGrid5.AllowDelete = False
    DBGrid6.AllowDelete = False
    DBGrid7.AllowDelete = False
    cmdSomar.Enabled = False
    cmdSubtrair.Enabled = False
  Else
    cmdFinalizar.Visible = True
    For i = 0 To Tela.Count - 1
      Tela(i).Enabled = True
    Next i
    DBGrid2.AllowDelete = True
    DBGrid3.AllowDelete = True
    DBGrid4.AllowDelete = True
    DBGrid5.AllowDelete = True
    DBGrid6.AllowDelete = True
    DBGrid7.AllowDelete = True
    cmdSomar.Enabled = True
    cmdSubtrair.Enabled = True
  End If
End With
TotalRecebimento
TotalizaDespesas
TotalizaNotas
TotalizaCompras
TotalizaCheque
TotalizaVales
TotalResumo

Abrindo = False
End Sub

Private Sub cboBico_LostFocus()
If dbClientes.Recordset.EOF = False Then
  Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes.Recordset!CodigoCliente)
  If IsNumeric(txtLitros.Text) = True Then
    Preco = Preco * CDbl(txtLitros.Text)
    txtNotaValor.Text = Preco
    Call txtNotaValor_LostFocus
  End If
End If
End Sub

Private Sub cboClienteCheque_GotFocus()
With dbChequesContas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If MaskEdBox1(0).Text = "   " Or MaskEdBox1(1).Text = "   " Or MaskEdBox1(2).Text = "    " Or MaskEdBox1(0).Text = "      - " Then Exit Sub
    .Recordset.FindFirst "comp='" & MaskEdBox1(0).Text & "' and banconumero=" & MaskEdBox1(1).Text & " and agencia=" & MaskEdBox1(2).Text & " and conta='" & MaskEdBox1(0).Text & "'"
    If .Recordset.NoMatch = False Then
      dbClientesCheques.Recordset.FindFirst "codigocliente=" & .Recordset!CodigoCliente
      txtCodCliente.Text = dbClientesCheques.Recordset!Codigo
      Call txtCodCliente_LostFocus
    End If
  End If
End With
End Sub

Private Sub cboClienteCheque_LostFocus()
If cboClienteCheque.Text = "" Then Exit Sub
With dbClientesCheques
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboClienteCheque.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodCliente.Text = .Recordset!codigochequecliente
    cboClienteCheque.Text = .Recordset!Nome
    AlertaAtivo .Recordset!Posicao
  End If
End With
End Sub

Private Sub cboClientesNota_GotFocus()
Me.KeyPreview = False
With dbClientesCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from clientescarros where codigocliente=0 order by Placa"
  .Refresh
End With
End Sub

Private Sub cboClientesNota_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboClientesNota_LostFocus()
Me.KeyPreview = True
With dbClientes
  .Refresh
  If cboClientesNota.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboClientesNota.Text & "'"
  If .Recordset.NoMatch = False Then
    cboClientesNota.Text = .Recordset!Nome
    StrTemp = "select *from clientescarros where codigocliente=" & .Recordset!CodigoCliente
  Else
    StrTemp = "select *from clientescarros where codigocliente=0"
  End If
End With
With dbClientesCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp & " order by Placa"
  .Refresh
End With
End Sub

Private Sub cboDespesa_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboDespesa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboDespesa_LostFocus()
Me.KeyPreview = True
With dbDespesas
  .Refresh
  If cboDespesa.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboDespesa.Text & "'"
  If .Recordset.NoMatch = False Then
    cboDespesa.Text = .Recordset("descri")
  End If
End With
End Sub

Private Sub cboFuncionario_LostFocus()
With dbVendedores
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  If cboFuncionario.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboFuncionario.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub cboMicrocredito_GotFocus()
Me.KeyPreview = False
With dbClientesCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from clientescarros where codigocliente=0 order by Placa"
  .Refresh
End With
End Sub

Private Sub cboMicrocredito_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboMicrocredito_LostFocus()
Me.KeyPreview = True
With dbClientes2
  .Refresh
  If cboMicrocredito.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboMicrocredito.Text & "'"
  If .Recordset.NoMatch = False Then
    cboMicrocredito.Text = .Recordset!Nome
    StrTemp = "select *from clientescarros where codigocliente=" & .Recordset!CodigoCliente
  Else
    StrTemp = "select *from clientescarros where codigocliente=0"
  End If
End With
With dbClientesCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp & " order by Placa"
  .Refresh
End With
End Sub

Private Sub cboPlaca_LostFocus()
With dbClientesCarros
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboPlaca.Text = "" Then Exit Sub
  .Recordset.FindFirst "placa='" & cboPlaca.Text & "'"
  If .Recordset.NoMatch = False Then
    cboPlaca.Text = .Recordset!Placa
  End If
End With
End Sub

Private Sub cboPlacaMicrocredito_LostFocus()
With dbClientesCarros
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboPlaca.Text = "" Then Exit Sub
  .Recordset.FindFirst "placa='" & cboPlaca.Text & "'"
  If .Recordset.NoMatch = False Then
    cboPlaca.Text = .Recordset!Placa
  End If
End With
End Sub

Private Sub cboProduto_LostFocus()
Dim CodigoProduto As Double
CodigoProduto = 0
With dbBicosEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from bicoencerrantes where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento & " and codigoproduto=" & CodigoProduto & " order by bico"
  .Refresh
End With
With dbProdutos2
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProduto.Text & "'"
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodProduto.Text = .Recordset!Codigo
    CodigoProduto = .Recordset!CodigoProduto
  End If
  If .Recordset!Combustivel = True Then
    lblBico.Visible = True
    cboBico.Visible = True
  Else
    lblBico.Visible = False
    cboBico.Visible = False
  End If
End With
With dbBicosEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from bicoencerrantes where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento & " and codigoproduto=" & CodigoProduto & " order by bico"
  .Refresh
  If .Recordset.EOF = False Then
    If cboBico.Text = "" Then
      cboBico.Text = .Recordset!Bico
    Else
      If IsNumeric(cboBico.Text) = True Then
        .Recordset.FindFirst "bico=" & cboBico.Text
        If .Recordset.NoMatch = True Then
          cboBico.Text = .Recordset!Bico
        End If
      Else
        cboBico.Text = .Recordset!Bico
      End If
    End If
  End If
End With
If dbClientes.Recordset.EOF = False Then
  Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes.Recordset!CodigoCliente)
  If IsNumeric(txtLitros.Text) = True Then
    Preco = Preco * CDbl(txtLitros.Text)
    txtNotaValor.Text = Preco
    Call txtNotaValor_LostFocus
  End If
End If
End Sub

Private Sub cboProdutoEntra_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProdutoEntra.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProdutoEntra.Text & "'"
  If .Recordset.NoMatch = False Then
    cboProdutoEntra.Text = .Recordset!Descri
    txtCod.Text = .Recordset!Codigo
  End If
End With
End Sub

Private Sub cboRecebimento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboRecebimento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboRecebimento_LostFocus()
Me.KeyPreview = True
With dbFormaDePg
  .Refresh
  If cboRecebimento.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboRecebimento.Text & "'"
  If .Recordset.NoMatch = False Then
    cboRecebimento.Text = .Recordset("descri")
  End If
End With
End Sub

Private Sub cmdAtualiza_Click()
With dbFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "distribuido=0"
    If .Recordset.NoMatch = True Then
      .Recordset.MoveLast
    End If
  End If
End With
 
If dbFechamento.Recordset.RecordCount <> 0 Then
  If dbFechamento.Recordset.EOF = False Then
    AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
  End If
End If
 
End Sub

Private Sub cmdAutorizar_Click()
Dim Resposta As Integer
With dbClientesNota
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja autorizar a nota atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Edit
  If .Recordset!Autorizado = True Then
    .Recordset!Autorizado = False
  Else
    .Recordset!Autorizado = True
  End If
  .Recordset.Update
End With
End Sub

Private Sub cmdCancelaFinaliza_Click()
Dim Resposta As Integer, TempComissao As Currency, CodigoNota As Double
Dim NaoAcumula As Boolean
Dim db As New ADODB.Connection
Dim Estatus As New frmEstatus2

Fechando = True

CodigoNota = 0
With dbFechamento
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe fechamento para cancelar!"
    Fechando = False
    Exit Sub
  End If
  If .Recordset!fechames = True Then
    MsgBox "Este caixa pertence a mês já fechado!"
    Exit Sub
  End If
'  If .Recordset!distribuido = False Then
'    MsgBox "Fechamento não foi finalizado!"
'    Fechando = False
'    Exit Sub
'  End If
  Abrindo = True
  AbreFechamento .Recordset!CodigoFechamento, .Recordset!DataCaixa
  Abrindo = False
End With


Resposta = MsgBox("Deseja cancelar a finalização o caixa atual?", vbYesNo, "Finalizar")
If Resposta = vbNo Then
  Fechando = False
  Exit Sub
End If

If IsNumeric(lblJurosResumo.Caption) = True Then
  With dbStatus
    If dbFechamento.Recordset!distribuido = True Then
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.Edit
        .Recordset!Juros = .Recordset!Juros - CCur(lblJurosResumo.Caption)
        .Recordset.Update
      End If
    End If
  End With
End If

'Registra as formas de pagamento recebidas
If CancelaPagamentos = False Then
  Fechando = False
  Exit Sub
End If

'Totaliza as despesas para lançar no status e no saldo das contas
If CancelaDespesas = False Then
  Fechando = False
  Exit Sub
End If

'Totaliza Vales de funcionários
If CancelaVales = False Then
  Fechando = False
  Exit Sub
End If

'Registra a compra de produtos
If CancelaCompras = False Then
  Fechando = False
  Exit Sub
End If

'Registra pagamentos de funcionários
If CancelaPgFuncionarios = False Then
  Fechando = False
  Exit Sub
End If

'fecha nota de clientes
With dbClientesNota
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = False
      If IsNull(.Recordset!LucroDif) = False Then
        ClienteDif = ClienteDif + .Recordset!LucroDif
      End If
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
If dbFechamento.Recordset!distribuido = True Then
  With dbStatus
    .Recordset.Edit
    If IsNull(.Recordset!clientediferenciado) = True Then .Recordset!clientediferenciado = 0
    .Recordset!clientediferenciado = .Recordset!clientediferenciado - ClienteDif
    .Recordset.Update
  End With
End If

'fecha cheques
With dbCheques
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If IsNull(.Recordset!CPF) = False Then
        dbClientesCheques.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
      Else
        If IsNull(.Recordset!CNPJ) = False Then
          dbClientesCheques.Recordset.FindFirst "cpf='" & .Recordset!CNPJ
        Else
          dbClientesCheques.Recordset.FindFirst "codigochequecliente=0"
        End If
      End If
      If dbClientesCheques.Recordset.NoMatch = False Then
        With dbChequesContas
          'On Error Resume Next
          .RecordSource = "select *from chequescontas where codigocliente=" & dbClientesCheques.Recordset!codigochequecliente & " and banconumero=" & dbCheques.Recordset!Banco & " and agencia=" & dbCheques.Recordset!Agencia & " and conta='" & dbCheques.Recordset!Agencia & "'"
          .Refresh
          If Err.Number = 0 Then
            If .Recordset.RecordCount = 0 Then
              .Recordset.AddNew
              .Recordset!CodigoCliente = dbClientesCheques.Recordset!codigochequecliente
              .Recordset!COMP = dbCheques.Recordset!COMP
              .Recordset!banconumero = dbCheques.Recordset!Banco
              .Recordset!Agencia = dbCheques.Recordset!Agencia
              .Recordset!Conta = dbCheques.Recordset!Conta
              .Recordset.Update
            End If
          End If
          On Error GoTo 0
        End With
        dbClientesCheques.Recordset.Edit
        dbClientesCheques.Recordset!numerodecheques = dbClientesCheques.Recordset!numerodecheques - 1
        dbClientesCheques.Recordset!Total = dbClientesCheques.Recordset!Total - .Recordset!Valor
        dbClientesCheques.Recordset.Update
      End If
      .Recordset.Edit
      .Recordset!fechamentodiario = False
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

db.Open CaminhoADO

StrTemp = "update fechamentodecaixa set totaldespesas=0"
StrTemp = StrTemp & ",Totalrecebido=0"
StrTemp = StrTemp & ",diferenca=0"
StrTemp = StrTemp & ",distribuido=0"
StrTemp = StrTemp & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento

db.Execute StrTemp

With dbFechamento
  A = .Recordset!CodigoFechamento
  .Refresh
  .Refresh
  .Recordset.FindFirst "codigofechamento=" & A
End With

DBGrid1.Refresh

Load Estatus
Unload Estatus


End Sub

Private Sub cmdChequeMudaAutoriza_Click()
With dbCheques
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um cheque primeiro!"
    Exit Sub
  End If
  If IsNull(.Recordset!Autorizar) = False Then
    If .Recordset!Autorizar = True Then
      If .Recordset!Autorizado = False Then
        .Recordset.Edit
        .Recordset!Autorizado = True
        .Recordset.Update
      Else
        .Recordset.Edit
        .Recordset!Autorizado = False
        .Recordset.Update
      End If
    End If
  End If
End With
End Sub

Private Sub cmdExibir_Click()
If Abrindo = False Then
  With dbFechamento
    SSTab1.Visible = True
    AbreFechamento .Recordset!CodigoFechamento, .Recordset!DataCaixa
  End With
End If
End Sub

Private Sub cmdFinalizar_Click()
Dim Resposta As Integer, TempComissao As Currency, CodigoNota As Double
Dim NaoAcumula As Boolean, ClienteDif As Currency
Dim Estatus As New frmEstatus2
Dim db As New ADODB.Connection

Fechando = True

CodigoNota = 0
With dbFechamento
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe fechamento para finalizar!"
    Fechando = False
    Exit Sub
  End If
  If .Recordset!fechado = False Then
    MsgBox "A Primeira parte ainda não foi confirmada!"
    Fechando = False
    Exit Sub
  End If
  If .Recordset!distribuido = True Then
    MsgBox "Fechamento já finalizado!"
    Fechando = False
    Exit Sub
  End If
  Abrindo = True
  AbreFechamento .Recordset!CodigoFechamento, .Recordset!DataCaixa
  Abrindo = False
End With

If FechamentoBloqueado = True Then
  Fechando = False
  Exit Sub
End If


Resposta = MsgBox("Deseja finalizar o caixa atual?", vbYesNo, "Finalizar")
If Resposta = vbNo Then
  Fechando = False
  Exit Sub
End If


If CCur(lblTotalVale.Caption) > 25 Or CCur(lblTotalVale.Caption) < -25 Then
  MsgBox "O valor de vales não pode ultrapassar R$ 10,00! Somente usuário Administrativo pode confirmar este caixa!"
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    Exit Sub
  End If
End If

With dbClientesNota
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!Autorizar = True Then
        If .Recordset!Autorizado = False Then
          MsgBox "Existe nota lançada não autorizada!"
          .Recordset.MoveFirst
          Exit Sub
        End If
      End If
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

With dbCheques
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!Autorizar = True Then
        If .Recordset!Autorizado = False Then
          MsgBox "Existe cheque não autorizado!"
          .Recordset.MoveFirst
          Exit Sub
        End If
      End If
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

If dbFormaDePgRecebido.Recordset.RecordCount = 0 Then
  Permissao = False
  MsgBox "Não foi lançado nenhums tipo de recebimento!"
  frmPermissao.Show vbModal
  If Permissao = False Then
    Fechando = False
    Exit Sub
  End If
End If
If CCur(lblDiferenca.Caption) > 0.02 Or CCur(lblDiferenca.Caption) < -0.02 Then
  Permissao = False
  MsgBox "Caixa com diferença"
  Fechando = False
  Exit Sub
End If


If IsNumeric(lblJurosResumo.Caption) = True Then
  With dbStatus
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Edit
      .Recordset!Juros = .Recordset!Juros + CCur(lblJurosResumo.Caption)
      .Recordset.Update
    End If
  End With
End If

'Registra as formas de pagamento recebidas
If GravaPagamentos = False Then
  Fechando = False
  Exit Sub
End If

'Totaliza as despesas para lançar no status e no saldo das contas
If GravaDespesas = False Then
  Fechando = False
  Exit Sub
End If

'Totaliza Vales de funcionários
If GravaVales = False Then
  Fechando = False
  Exit Sub
End If

'Registra a compra de produtos
If GravaCompras = False Then
  Fechando = False
  Exit Sub
End If

'Registra pagamentos de funcionários
If GravaPgFuncionarios = False Then
  Fechando = False
  Exit Sub
End If

ClienteDif = 0
'fecha nota de clientes
With dbClientesNota
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      If IsNull(.Recordset!LucroDif) = False Then
        ClienteDif = ClienteDif + .Recordset!LucroDif
      End If
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
With dbStatus
  .Recordset.Edit
  If IsNull(.Recordset!clientediferenciado) = True Then .Recordset!clientediferenciado = 0
  .Recordset!clientediferenciado = .Recordset!clientediferenciado + ClienteDif
  .Recordset.Update
End With

'fecha cheques
With dbCheques
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If IsNull(.Recordset!CPF) = False Then
        dbClientesCheques.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
      Else
        If IsNull(.Recordset!CNPJ) = False Then
          dbClientesCheques.Recordset.FindFirst "cpf='" & .Recordset!CNPJ
        Else
          dbClientesCheques.Recordset.FindFirst "codigochequecliente=0"
        End If
      End If
      If dbClientesCheques.Recordset.NoMatch = False Then
        With dbChequesContas
          On Error Resume Next
          .RecordSource = "select *from chequescontas where codigocliente=" & dbClientesCheques.Recordset!codigochequecliente & " and banconumero=" & dbCheques.Recordset!Banco & " and agencia=" & dbCheques.Recordset!Agencia & " and conta='" & dbCheques.Recordset!Agencia & "'"
          .Refresh
          If Err.Number = 0 Then
            If .Recordset.RecordCount = 0 Then
              .Recordset.AddNew
              .Recordset!CodigoCliente = dbClientesCheques.Recordset!codigochequecliente
              .Recordset!COMP = dbCheques.Recordset!COMP
              .Recordset!banconumero = dbCheques.Recordset!Banco
              .Recordset!Agencia = dbCheques.Recordset!Agencia
              .Recordset!Conta = dbCheques.Recordset!Conta
              .Recordset.Update
            End If
          End If
          On Error GoTo 0
        End With
        dbClientesCheques.Recordset.Edit
        dbClientesCheques.Recordset!numerodecheques = dbClientesCheques.Recordset!numerodecheques + 1
        dbClientesCheques.Recordset!Total = dbClientesCheques.Recordset!Total + .Recordset!Valor
        dbClientesCheques.Recordset.Update
      End If
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

db.Open CaminhoADO

StrTemp = "update fechamentodecaixa set totaldespesas=" & NumeroIngles(CCur(lblTotalDespesas.Caption))
If IsNumeric(lblTotalRecebido.Caption) Then
  StrTemp = StrTemp & ",Totalrecebido=" & NumeroIngles(CCur(lblTotalRecebido.Caption))
Else
  StrTemp = StrTemp & ",Totalrecebido=0"
End If
StrTemp = StrTemp & ",diferenca=" & NumeroIngles(CCur(lblDiferenca.Caption))
StrTemp = StrTemp & ",distribuido=-1"
StrTemp = StrTemp & ",distribuidopor='" & Usuarios.Nome & "'"
StrTemp = StrTemp & " where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento

db.Execute StrTemp
AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
Fechando = False

With dbFechamento
  A = .Recordset!CodigoFechamento
  .Refresh
  .Recordset.FindFirst "codigofechamento=" & A
End With
DBGrid1.Refresh

Load Estatus
Unload Estatus

End Sub


Private Sub cmdGravaDespesas_Click()

If dbFechamento.Recordset!distribuido = True Then Exit Sub

GravaDespesas
End Sub

Private Sub cmdGravarPagamentos_Click()
If dbFechamento.Recordset!distribuido = True Then Exit Sub
GravaPagamentos
End Sub

Private Sub cmdGravaVales_Click()
GravaVales
End Sub

Private Sub cmdImportarNotas_Click()
Dim Dia As Date, strEncerrantes As String, intArquivo As Integer
Dim StrTemp As String, SoPrimeira As Boolean
Dim Codigo As String, Descri As String, Tipo As String, Valor As Currency
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency
Dim DescontoPorcento As Currency
Dim Tanque As Integer, Estoque As Double
Dim Bico As Integer, Encerrante As Double, Encontrou As Boolean, Abertura As Double
Dim Preco As Currency, Qtd As Double, Funcionario As Integer
Dim CodigoConta As String, DesteCaixaQtd As Double, DesteCaixaValor As Currency

Dim CodigoCliente As Double, Cupom As String, Placa As String
Dim Km As String, Veiculo As String, ValorTotal As Currency
Dim CodigoProduto As Double, valorUnitario As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim Autorizar As Boolean, Motivo As String, Autorizado As Boolean



Dim db As New ADODB.Connection, dbSql As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset
Dim dbImportacao As New ADODB.Recordset
Dim dbClientes As New ADODB.Recordset
Dim dbClientesCarros As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbTotalNotas As New ADODB.Recordset
Dim dbTotalCobranca As New ADODB.Recordset
Dim dbClientesProdutos As New ADODB.Recordset
Dim dbEncerrantes As New ADODB.Recordset

cmdImportarNotas.Enabled = False

db.Open CaminhoADO

dbClientes.CursorLocation = adUseClient
dbClientes.Open "select *from clientes", db, adOpenKeyset, adLockOptimistic

dbClientesCarros.CursorLocation = adUseClient
dbClientesCarros.Open "select *from clientescarros", db, adOpenForwardOnly, adLockReadOnly

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "select *from produtos", db, adOpenForwardOnly, adLockReadOnly

dbTotalNotas.CursorLocation = adUseClient
dbTotalNotas.Open "select codigocliente, sum(valorprevisto) as total from clientesnota2 where confirmado=0 group by codigocliente", db, adOpenForwardOnly, adLockReadOnly

dbTotalCobranca.CursorLocation = adUseClient
dbTotalCobranca.Open "select codigocliente, sum(valor) as total from clientescobranca where pago=0 group by codigocliente", db, adOpenForwardOnly, adLockReadOnly

dbClientesProdutos.CursorLocation = adUseClient
dbClientesProdutos.Open "select *from clientesprodutos", db, adOpenForwardOnly, adLockReadOnly

dbConfig.CursorLocation = adUseClient
dbConfig.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly

dbEncerrantes.CursorLocation = adUseClient
dbEncerrantes.Open "select *from bicoEncerrantes where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic


dbSql.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig!ftp

dbImportacao.CursorLocation = adUseClient

On Error Resume Next
  
  dbImportacao.Open "select *from caixas where datacaixa='" & dbFechamento.Recordset!DataCaixa & "' and turno='" & dbFechamento.Recordset!Turno & "' and codigoposto='" & dbConfig!Porta & "' and planodeconta='" & dbFechamento.Recordset!Codigo & "' order by linhaexportada", dbSql, adOpenForwardOnly, adLockReadOnly
  
  If Err.Number <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
    cmdImportarNotas.Enabled = True
    Exit Sub
  End If
  
  On Error GoTo 0
  
  If dbImportacao.RecordCount = 0 Then
    MsgBox "O caixa atual ainda não foi exportado!"
    cmdImportarNotas.Enabled = True
    GoTo Sair
  End If
  dbImportacao.MoveLast
  dbImportacao.MoveFirst
  
  With dbClientesNota
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        If .Recordset!Confirmado = False Then
          .Recordset.Delete
        End If
        .Recordset.MoveNext
      Loop
    End If
  End With
  
  Do While dbImportacao.EOF = False
    StrTemp = dbImportacao!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
     

      Case "003"
        'notas de clientes
        'If SoPrimeira = False Then
          On Error GoTo 0
          PrecoDif = False
          If Len(StrTemp) > 120 Then
            CodigoCliente = CDbl(Mid(StrTemp, 5, 12))
            Cupom = RemoveString(Trim(Mid(StrTemp, 18, 12)))
            Placa = Mid(StrTemp, 31, 9)
            Km = Mid(StrTemp, 41, 15)
            Veiculo = Mid(StrTemp, 57, 25)
            Qtd = Mid(StrTemp, 83, 15)
            ValorTotalDif = Mid(StrTemp, 99, 15)
            If Len(StrTemp) > 179 Then
              ValorUnitarioDif = CCur(Mid(StrTemp, 179, 15))
              If ValorUnitarioDif = 0 Then
                ValorUnitarioDif = CCur(Format(ValorTotalDif / Qtd, "0.000"))
              End If
            Else
              ValorUnitarioDif = CCur(Format(ValorTotalDif / Qtd, "0.000"))
            End If
            StrTemp2 = Mid(StrTemp, 115, 15)
            If IsNumeric(StrTemp2) = False Then
              StrTemp2 = RemoveString(StrTemp2)
            End If
            CodigoProduto = StrTemp2
            If Len(StrTemp) > 130 Then
              If IsNumeric(Mid(StrTemp, 131, 15)) = True Then
                LucroDif = Mid(StrTemp, 131, 15)
                If IsNumeric(Mid(StrTemp, 147, 15)) = True Then
                  valorUnitario = Mid(StrTemp, 147, 15)
                End If
                If IsNumeric(Mid(StrTemp, 163, 15)) = True Then
                  ValorTotal = Mid(StrTemp, 163, 15)
                Else
                  ValorTotal = valorUnitario * Qtd
                End If
              Else
                ValorTotal = ValorTotalDif
                valorUnitario = ValorUnitarioDif
                LucroDif = 0
              End If
            Else
              ValorTotal = ValorTotalDif
              valorUnitario = ValorUnitarioDif
              LucroDif = 0
            End If
            LucroDif = ValorTotal - ValorTotalDif
            Autorizar = False
            Autorizado = False
            Motivo = ""
            
            If IsNumeric(Cupom) = False Then
              Cupom = 0
            End If
            dbClientes.MoveFirst
            dbClientes.Find "codigonoposto=" & CodigoCliente
            If dbClientes.EOF = True Then
              'MsgBox "Código de cliente de nota " & CodigoCliente & " não encontrado!"
              'GravaBloqueado CodigoCliente, "Não encontrado", Cupom, ValorTotal, "Cliente não localizado"
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente não cadastrado'," & CodigoCliente & ")"
              GoTo SairDoCliente
            Else
              If dbProdutos.RecordCount <> 0 Then
                If dbClientes!protestado = True Then
                  'MsgBox "Cliente bloqueado!"
                  Autorizar = True
                  Autorizado = True
                  Motivo = "Bloqueado/Protestado"
                End If
                
                dbProdutos.MoveFirst
                dbProdutos.Find "codigo=" & CodigoProduto
                If dbProdutos.EOF = True Then
                  'MsgBox "Código do produto " & CodigoProduto & " não cadastrado!"
                  'GravaBloqueado CodigoCliente, "Código de produto não encontrado", Cupom, ValorTotal, "Cliente não localizado"
                  db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigonoposto) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cupom " & Cupom & " com produto não cadastrado'," & CodigoCliente & "," & CodigoProduto & ")"
                  GoTo SairDoCliente
                Else
                    Bico = 0
                    If dbProdutos!Combustivel = True Then
                      If dbEncerrantes.RecordCount <> 0 Then
                        dbEncerrantes.MoveFirst
                        dbEncerrantes.Find "codigoproduto=" & dbProdutos!CodigoProduto
                        If dbEncerrantes.EOF = False Then
                          Bico = dbEncerrantes!Bico
                        End If
                      End If
                    End If
                    Preco = PrecoAtual(dbProdutos!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, Bico)
                End If
                If dbClientes!mensalista = False Then
                  If dbClientes!desativado < dbFechamento.Recordset!DataCaixa Then
                    If Usuarios.Grupo.admDatas < 2 Then
                      'MsgBox "O cliente " & DbClientes!Nome & " está desativado!"
                      If Configura.NotaBloqueia = 0 Then
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
                        Autorizar = True
                        Motivo = "Desativado"
                      End If
                    Else
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está desativado! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
                      If Configura.NotaBloqueia = 0 Then
                        Autorizar = True
                        Autorizado = False
                        Motivo = "Desativado"
                      End If
                    End If
                  End If
                End If
                If dbClientes!limitar = True Then
                  If IsNull(dbClientes!Limite) = False Then
                    Limite = CCur(ValorTotal)
                    dbTotalNotas.Requery
                    If dbTotalNotas.RecordCount <> 0 Then
                      dbTotalNotas.MoveFirst
                      dbTotalNotas.Find "codigocliente=" & CodigoCliente
                      If dbTotalNotas.EOF = False Then
                        If IsNull(dbTotalNotas!Total) = False Then
                          Limite = Limite + dbTotalNotas!Total
                        End If
                      End If
                    End If
                    
                    dbTotalCobranca.Requery
                    If dbTotalCobranca.RecordCount <> 0 Then
                      dbTotalCobranca.MoveFirst
                      dbTotalCobranca.Find "codigocliente=" & CodigoCliente
                      If dbTotalCobranca.EOF = False Then
                        If IsNull(dbTotalCobranca!Total) = False Then
                          Limite = Limite + dbTotalCobranca!Total
                        End If
                      End If
                    End If
                    If Limite > dbClientes!Limite Then
                      If Usuarios.Grupo.admDatas < 2 Then
                        'MsgBox "O cliente " & DbClientes!Nome & " ultrapassará o limite dele! Somente o administrador pode lançar."
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(ValorTotal) & ")"
                        Autorizar = True
                        Motivo = "Limite"
                      Else
                        'Resposta = MsgBox("O cliente " & DbClientes!Nome & " ultrapassará o limite dele! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
                        'If Resposta = vbNo Then GoTo SairDoCliente
                        Autorizar = True
                        Autorizado = False
                        Motivo = "Ultrapassou Limite"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(ValorTotal) & ")"
                      End If
                    End If
                  Else
                    'MsgBox "O cliente " & DbClientes!Nome & " esta marcado para ser limitado mas não possue valor definido!"
                    'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Marcado para limitar mas não possue valor a ser limitado"
                    Autorizar = True
                    Motivo = "Sem Limite"
                    db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente marcado para limitar mas sem limite cadastrado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
                  End If
                End If
                If dbClientes!diapagamento <> 0 Then
                  If dbClientes!diapagamento >= 28 Then
                    DataPrevista = CDate(Format(UltimoDiaDoMes(Month(dbFechamento.Recordset!DataCaixa), Year(dbFechamento.Recordset!DataCaixa)), "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
                  Else
                    DataPrevista = CDate(Format(dbClientes!diapagamento, "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
                  End If
                Else
                  DataPrevista = DateAdd("m", 1, dbFechamento.Recordset!DataCaixa)
                End If
                If DataPrevista < dbFechamento.Recordset!DataCaixa Then
                  DataPrevista = DateAdd("m", 1, DataPrevista)
                End If
                dbClientesProdutos.Filter = ""
                If dbClientesProdutos.RecordCount <> 0 Then
                  dbClientesProdutos.MoveFirst
                  dbClientesProdutos.Filter = "codigocliente=" & dbClientes!CodigoCliente & " and codproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
                  If dbClientesProdutos.EOF = False Then
                    If dbClientesProdutos!validade = dbFechamento.Recordset!DataCaixa Then
                      If dbClientesProdutos!HoraIni >= dbFechamento.Recordset!HoraIni Then
                        PrecoDif = True
                      End If
                    Else
                      PrecoDif = True
                    End If
                  End If
                  If PrecoDif = True Then
                    If dbClientesProdutos!Preco <> 0 Then
                      TempValorPagar = Qtd * dbClientesProdutos!Preco
                    Else
                      TempValorPagar = Qtd * Preco
                      If dbClientesProdutos!Porcento <> 0 Then
                        TempValorPagar = TempValorPagar * dbClientesProdutos!Porcento
                      End If
                    End If
                    If dbClientesProdutos!valorasomar <> 0 Then
                      TempValorPagar = TempValorPagar + (Qtd * dbClientesProdutos!valorasomar)
                    End If
                    TempDif = TempValorPagar - ValorTotal
                    If TempDif > 0.2 Or TempDif < -0.2 Then
                      If Usuarios.Grupo.admDatas < 2 Then
                        'MsgBox "O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
                        Autorizar = True
                        Motivo = "Preço Diferenciado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal) & "," & NumeroIngles(TempValorPagar) & ")"
                      Else
                        'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
                        'If Resposta = vbNo Then GoTo SairDoCliente
                        Autorizar = True
                        Autorizado = False
                        Motivo = "Preço Diferenciado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal) & "," & NumeroIngles(TempValorPagar) & ")"
                      End If
                    End If
                  Else
                    'ValorUnitarioDif = Qtd * valorUnitario
                    TempDif = (ValorUnitarioDif * Qtd) - ValorTotal
                    If TempDif > 0.01 Or TempDif < -0.01 Then
                      'MsgBox "Preço unitário incorreto!"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(ValorUnitarioDif * Qtd) & ")"
                      GoTo SairDoCliente
                    End If
                  End If
                Else
                  TempDif = Preco - (ValorTotal / Qtd)
                  If TempDif > 0.2 Or TempDif < -0.02 Then
                    If Usuarios.Grupo.admDatas < 2 Then
                      'MsgBox "O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
                      Autorizar = True
                      Motivo = "Preço Diferenciado"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal / Qtd) & "," & NumeroIngles(Preco) & ")"
                    Else
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
                      'If Resposta = vbNo Then GoTo SairDoCliente
                      Autorizar = True
                      Autorizado = False
                      Motivo = "Preço incorreto!"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamento.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal / Qtd) & "," & NumeroIngles(Preco) & ")"
                    End If
                  End If
                End If
                A = Fix(valorUnitario)
                If Qtd = 0 Then
                  Qtd = ValorTotal / ValorUnitarioDif
                End If
              End If
            End If
          End If
          
          dbClientesCarros.Filter = "placa='" & Trim(Placa) & "'"
          
          StrTemp = "insert into clientesnota2 (codigofechamento,codigocliente,nome,datalanc,dataprevista,valorprevisto,Data,"
          If Trim(Cupom) <> "" Then
            StrTemp = StrTemp & "Cupom,"
          End If
          StrTemp = StrTemp & "Km,Placa,"
          On Error Resume Next
          If dbClientesCarros.EOF = False And dbClientesCarros.BOF = False Then
            StrTemp = StrTemp & "codigocarro,"
          End If
          On Error GoTo 0
          StrTemp = StrTemp & "Litros,Consumo,CodigoProduto,valorUnitario,Qtd,ValorUnitarioDif,ValorTotalDif,LucroDif,Autorizar,Autorizado,Motivo) values ("
          
          StrTemp = StrTemp & dbFechamento.Recordset!CodigoFechamento & "," & dbClientes!CodigoCliente & ",'" & dbClientes!Nome & "',#" & DataInglesa(Date) & " " & Time & "#,#" & DataInglesa(DataPrevista) & "#," & NumeroIngles(ValorTotal) & ",#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#,"
          If Trim(Cupom) <> "" Then
            StrTemp = StrTemp & Trim(Cupom) & ","
          End If
          If Trim(Km) = "" Then Km = 0
          StrTemp = StrTemp & NumeroIngles(Trim(Km)) & ",'" & Trim(Placa) & "',"
          On Error Resume Next
          If dbClientesCarros.EOF = False And dbClientesCarros.BOF = False Then
            StrTemp = StrTemp & dbClientesCarros!codigocarro & ","
          End If
          On Error GoTo 0
          If Consumo = "" Then
            Consumo = 0
          End If
          StrTemp = StrTemp & NumeroIngles(Qtd) & "," & NumeroIngles(Consumo) & "," & CodigoProduto & "," & NumeroIngles(valorUnitario) & "," & NumeroIngles(Qtd) & "," & NumeroIngles(ValorUnitarioDif) & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(LucroDif) & "," & Autorizar & "," & Autorizado & ",'" & Motivo & "')"
          
          On Error GoTo 0
          On Error Resume Next
          db.Execute StrTemp
          If Err.Number <> 0 Then
            MsgBox Err.Number & "-" & Err.Description
          End If
        
          If IsNull(dbClientes!UltimoAbastecimento) = True Then
            dbClientes!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
          End If
          If dbClientes!UltimoAbastecimento < dbFechamento.Recordset!DataCaixa Then
            dbClientes!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
          End If
          db.Execute "update clientes set TotalNotas=TotalNotas+" & NumeroIngles(ValorTotal) & " where codigocliente=" & CodigoCliente
          db.Execute "update clientes set saldo=limite-totalnotas-totalboleto where codigocliente=" & CodigoCliente
        'End If
SairDoCliente:
        
        
    End Select
    dbImportacao.MoveNext
  Loop

Sair:


dbConfig.Close
dbClientesCarros.Close
dbProdutos.Close
dbTotalNotas.Close
dbTotalCobranca.Close
dbClientesProdutos.Close
dbImportacao.Close
dbSql.Close
db.Close

cmdImportarNotas.Enabled = True
Call cmdExibir_Click
dbClientesNota.Refresh

dbClientesNota.Refresh
DBGrid4.Refresh

SSTab1.Tab = 2

End Sub

Private Sub cmdImprimir_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Y1 As Double, Y2 As Double, X1 As Double, X2 As Double
Dim Coluna As Double, Total As Currency, Salto As Double

If dbFechamento.Recordset.EOF = True Then
  MsgBox "Escolha um caixa a ser impresso!"
  DBGrid1.SetFocus
  Exit Sub
End If

AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 16
Largura = 190
Dia = Now

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 14
StrTemp = "Conferência de Valores do Caixa"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Impresso em: " & Format(Dia, "Short date") & " - " & Format(Dia, "Short Time")
Printer.Print StrTemp

StrTemp = "Responsável: " & dbFechamento.Recordset!responsavel
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno: " & dbFechamento.Recordset!Turno
Printer.CurrentX = 100
Printer.Print StrTemp;

StrTemp = "Data: " & Format(dbFechamento.Recordset!DataCaixa, "Short date")
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

With dbFormaDePgRecebido
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    StrTemp = "Recebimentos"
    Printer.FontBold = True
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    Printer.FontBold = False
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Descrição"
    Printer.CurrentX = 1
    Printer.Print StrTemp;
    
    StrTemp = "Dt. Bord."
    Printer.CurrentX = 74 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = 94 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Descrição"
    Printer.CurrentX = 1 + 95
    Printer.Print StrTemp;
    
    StrTemp = "Dt. Bord."
    Printer.CurrentX = 74 + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = 94 + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Coluna = 0
    Do While .Recordset.EOF = False
      
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 1 + Coluna
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Data, "Short Date")
      Printer.CurrentX = 74 + Coluna - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Total = Total + .Recordset!ValorBruto
      StrTemp = Format(.Recordset!ValorBruto, "Currency")
      Printer.CurrentX = 94 + Coluna - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      If Coluna = 0 Then
        Coluna = 95
      Else
        Coluna = 0
        Printer.Print ""
      End If
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst

    If Coluna = 95 Then
      Printer.Print ""
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Y2 = Printer.CurrentY
    Printer.Line (0, Y2)-(Largura, Y2)
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (55, Y1)-(55, Y2)
    Printer.Line (75, Y1)-(75, Y2)
    Printer.Line (95, Y1)-(95, Y2)
    Printer.Line (150, Y1)-(150, Y2)
    Printer.Line (170, Y1)-(170, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = Format(Total, "Currency")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
  Else
    
  End If
End With

With dbDespesasLanc
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    StrTemp = "Despesas"
    Printer.FontBold = True
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    Printer.FontBold = False
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Descrição"
    Printer.CurrentX = 1
    Printer.Print StrTemp;
    
    StrTemp = "Observação"
    Printer.CurrentX = 81
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Total = 0
    Do While .Recordset.EOF = False
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 1
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Obs
      Printer.CurrentX = 81
      Printer.Print StrTemp;
      
      Total = Total + .Recordset!Valor
      StrTemp = Format(.Recordset!Valor, "Currency")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Y2 = Printer.CurrentY
    Printer.Line (0, Y2)-(Largura, Y2)
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (80, Y1)-(80, Y2)
    Printer.Line (160, Y1)-(160, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = Format(Total, "Currency")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
  End If
End With

With dbCheques
  If .Recordset.RecordCount <> 0 Then
    If IsNumeric(txtChequeJuros.Text) = True Then
      If CCur(txtChequeJuros.Text) <> 0 Then
        Total = Total + CCur(txtChequeJuros.Text)
        ImprimeGrid DBGrid6, Printer, dbCheques, 7, True, , , , , , "Cheques"
      End If
    End If
  End If
End With

Printer.ScaleMode = vbMillimeters
With dbClientesNota
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    StrTemp = "Notas de Clientes"
    Printer.FontBold = True
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    Printer.FontBold = False
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Cliente"
    Printer.CurrentX = 1
    Printer.Print StrTemp;
    
    StrTemp = "Cupom"
    Printer.CurrentX = 74 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = 94 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Cliente"
    Printer.CurrentX = 1 + 95
    Printer.Print StrTemp;
    
    StrTemp = "Cupom"
    Printer.CurrentX = 74 + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = 94 + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Coluna = 0
    Total = 0
    Do While .Recordset.EOF = False
      
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 1 + Coluna
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Cupom
      Printer.CurrentX = 74 + Coluna - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Total = Total + .Recordset!ValorPrevisto
      StrTemp = Format(.Recordset!ValorPrevisto, "Currency")
      Printer.CurrentX = 94 + Coluna - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      If Coluna = 0 Then
        Coluna = 95
      Else
        Coluna = 0
        Printer.Print ""
      End If
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
    If Coluna = 95 Then
      Printer.Print ""
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Y2 = Printer.CurrentY
    Printer.Line (0, Y2)-(Largura, Y2)
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (55, Y1)-(55, Y2)
    Printer.Line (75, Y1)-(75, Y2)
    Printer.Line (95, Y1)-(95, Y2)
    Printer.Line (150, Y1)-(150, Y2)
    Printer.Line (170, Y1)-(170, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = Format(Total, "Currency")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
  End If
End With


With dbProdutoEntra
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    StrTemp = "Entrada de Produtos"
    Printer.FontBold = True
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    Printer.FontBold = False
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    StrTemp = "Código"
    Printer.CurrentX = 19 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Descrição"
    Printer.CurrentX = 21
    Printer.Print StrTemp;
    
    StrTemp = "Qtd."
    Printer.CurrentX = 139 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "V. Unit."
    Printer.CurrentX = 164 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "V. Total"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Total = 0
    Do While .Recordset.EOF = False
      StrTemp = .Recordset!Codigo
      Printer.CurrentX = 19 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 21
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Quantidade
      Printer.CurrentX = 139 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!PrecoNovo, "Currency")
      Printer.CurrentX = 164 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Total = Total + .Recordset!valornota
      StrTemp = Format(.Recordset!valornota, "Currency")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Y2 = Printer.CurrentY
    Printer.Line (0, Y2)-(Largura, Y2)
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (20, Y1)-(20, Y2)
    Printer.Line (120, Y1)-(120, Y2)
    Printer.Line (140, Y1)-(140, Y2)
    Printer.Line (165, Y1)-(165, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = Format(Total, "Currency")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
  End If
End With

Printer.CurrentY = Printer.CurrentY + 0.5
Y1 = Printer.CurrentY
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Tot. Vendas"
Printer.CurrentX = 1
Printer.Print StrTemp;

StrTemp = "Juros"
Printer.CurrentX = 31
Printer.Print StrTemp;

StrTemp = "Tot. Recebido"
Printer.CurrentX = 61
Printer.Print StrTemp;

StrTemp = "Despesas"
Printer.CurrentX = 91
Printer.Print StrTemp;

StrTemp = "Vales"
Printer.CurrentX = 121
Printer.Print StrTemp;

StrTemp = "Diferença"
Printer.CurrentX = 161
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = lblTotalVendas.Caption
Printer.CurrentX = 1
Printer.Print StrTemp;

StrTemp = lblJuros.Caption
Printer.CurrentX = 31
Printer.Print StrTemp;

StrTemp = lblRecebimentos.Caption
Printer.CurrentX = 61
Printer.Print StrTemp;

StrTemp = lblTotalDespesas.Caption
Printer.CurrentX = 91
Printer.Print StrTemp;

StrTemp = lblTotalVale.Caption
Printer.CurrentX = 121
Printer.Print StrTemp;

StrTemp = lblDiferenca.Caption
Printer.CurrentX = 161
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Y2 = Printer.CurrentY
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.Line (0, Y1)-(0, Y2)
Printer.Line (30, Y1)-(30, Y2)
Printer.Line (60, Y1)-(60, Y2)
Printer.Line (90, Y1)-(90, Y2)
Printer.Line (120, Y1)-(120, Y2)
Printer.Line (160, Y1)-(160, Y2)
Printer.Line (Largura, Y1)-(Largura, Y2)

'Printer.Print ""
'
'StrTemp = "Conferência de Valores"
'Printer.Print StrTemp
'
'Salto = 2
'
'Printer.CurrentX = 0
'StrTemp = "Descrição"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "Valor Conferido"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "- Valor Informado"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "= Valor Diferença"
'Printer.Print StrTemp
'
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Dineiro"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Moeda"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Cartão"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Despesas"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Cheques"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'Printer.CurrentX = 0
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Notas"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 20
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 80
'StrTemp = "___________________________"
'Printer.Print StrTemp;
'
'Printer.CurrentX = 140
'StrTemp = "___________________________"
'Printer.Print StrTemp
'
'
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Em caso de divergência nos valores conferidos e valores informados, é obrigatória a assinatura do Gerente, do Caixa e do Conferente responsáveis pelo caixa."
'ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, Printer.CurrentX, Printer.CurrentY, Largura
'
'
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Nome do Gerente:____________________________________  Assinatura:__________________________________"
'Printer.Print StrTemp
'
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Nome do Caixa:____________________________________  Assinatura:__________________________________"
'Printer.Print StrTemp
'
'Printer.CurrentY = Printer.CurrentY + Salto
'StrTemp = "Nome do Conferente:____________________________________  Assinatura:__________________________________"
'Printer.Print StrTemp

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdImprimirAutoriza_Click()
Dim Resposta As Integer, YInicial As Double

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

With dbClientesNota
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  
  Resposta = MsgBox("Deseja imprimir somente a selecionada?", vbYesNoCancel)
  
  
  Select Case Resposta
    Case vbYes
      If .Recordset.EOF = True Then
        MsgBox "Selecione um registro primeiro!"
        Exit Sub
      End If
      If .Recordset!Autorizar = False Then
        MsgBox "Esta nota não precisa de autorização!"
        Exit Sub
      End If
      ImprimeAutorizacao 0, dbFechamento.Recordset!responsavel, .Recordset!ValorPrevisto, .Recordset!Cupom, .Recordset!Nome, dbFechamento.Recordset!DataCaixa
    Case vbNo
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        If .Recordset!Autorizar = True Then
          ImprimeAutorizacao YInicial, dbFechamento.Recordset!responsavel, .Recordset!ValorPrevisto, .Recordset!Cupom, .Recordset!Nome, dbFechamento.Recordset!DataCaixa
          If YInicial > 0 Then
            If .Recordset.AbsolutePosition = .Recordset.RecordCount + 1 Then
              Printer.EndDoc
            Else
              Printer.NewPage
            End If
            YInicial = 0
          Else
            YInicial = 135
          End If
        End If
        .Recordset.MoveNext
      Loop
  End Select
End With

Printer.EndDoc

NaoImprime:
End Sub

Private Sub cmdInclueNota_Click()
Dim DataPrevista As Date, Consumo As Double
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim Limite As Currency, valorUnitario As Currency, Qtd As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim ValorNocaixa As Currency, TempDif As Currency
Dim Autorizar As Boolean, Autorizado As Boolean, Motivo As String

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If cboClientesNota.Text <> dbClientes.Recordset("nome") Then
  MsgBox "Selecione um cliente válido!", vbCritical, "Erro!"
  cboClientesNota.SetFocus
  Exit Sub
End If
If IsNumeric(txtNotaValor.Text) = False Then
  MsgBox "Informe um valor válido!", vbCritical, "Erro!"
  txtNotaValor.SetFocus
  Exit Sub
End If
If CCur(txtNotaValor.Text) <= 0 Then
  If CDbl(txtLitros.Text) >= 0 Then
    MsgBox "Para extornar uma nota a quantidade e o valor devem ser negativos!"
    Exit Sub
  End If
End If
If IsNumeric(txtCupom.Text) = False Then
  MsgBox "Informe o número do cupom válido!"
  txtCupom.SetFocus
  Exit Sub
End If

If Configura.NotaNoCaixa = 0 Then
  If IsNumeric(txtLitros.Text) = False Then
    Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes.Recordset!CodigoCliente)
    txtLitros.Text = CCur(txtNotaValor.Text) / Preco
  End If
  
'  If CDbl(txtLitros.Text) <= 0 Then
'    If CCur(txtNotaValor.Text) >= 0 Then
'      MsgBox "Para extornar uma nota a quantidade e o valor devem ser negativos!"
'    End If
'  End If
  If dbClientesCarros.Recordset.RecordCount <> 0 Then
    If cboPlaca.Text <> dbClientesCarros.Recordset!Placa Then
      MsgBox "Informe uma placa correta!"
      cboPlaca.SetFocus
      Exit Sub
    End If
  End If
  If IsNumeric(txtKm.Text) = False Then
    MsgBox "Informe um Km correto!"
    txtKm.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtLitros.Text) = False Then
    MsgBox "Informe a quantidade de litros!"
    txtLitros.SetFocus
    Exit Sub
  End If
  If dbProdutos2.Recordset.EOF = True Then
    MsgBox "Produto inválido!"
    Exit Sub
  End If
  If txtCodProduto.Text <> dbProdutos2.Recordset!Codigo Then
    MsgBox "Produto inválido!"
    txtCodProduto.SetFocus
    Exit Sub
  End If
  If dbClientes.Recordset!mensalista = False Then
    If dbClientes.Recordset!desativado < dbFechamento.Recordset!DataCaixa Then
      If Usuarios.Grupo.admDatas < 2 Then
        MsgBox "O cliente está desativado! Somente o administrador pode lançar."
        If Configura.NotaBloqueia = 0 Then
          Autorizar = True
          Motivo = "Desativado"
        End If
      Else
        MsgBox "O cliente está desativado!"
        If Configura.NotaBloqueia = 0 Then
          Autorizar = True
          Autorizado = True
          Motivo = "Desativado"
        End If
      End If
    End If
  End If
End If
If dbClientes.Recordset!protestado = True Then
  MsgBox "Cliente bloqueado!"
  Autorizar = True
  Autorizado = True
  Motivo = "Bloqueado/Protestado"
End If
If dbClientes.Recordset!limitar = True Then
  If IsNull(dbClientes.Recordset!Limite) = False Then
    Limite = CCur(txtNotaValor.Text)
    db.Open CaminhoADO
    dbTemp.Open "select sum(valorprevisto) as total from clientesnota2 where codigocliente=" & dbClientes.Recordset!CodigoCliente & " and confirmado=0", db
    If IsNull(dbTemp!Total) = False Then
      Limite = Limite + dbTemp!Total
    End If
    dbTemp.Close
    dbTemp.Open "select sum(valor) as total from clientescobranca where codigocliente=" & dbClientes.Recordset!CodigoCliente & " and pago=0", db
    If IsNull(dbTemp!Total) = False Then
      Limite = Limite + dbTemp!Total
    End If
    If Limite > dbClientes.Recordset!Limite Then
      If Usuarios.Grupo.admDatas < 2 Then
        MsgBox "O cliente ultrapassará o limite de notas dele! Somente o administrador pode lançar."
        Autorizar = True
        Motivo = "Ultrapassou Limite"
      Else
        MsgBox "O cliente ultrapassará o limite de notas dele!"
        Autorizar = True
        Autorizado = True
        Motivo = "Ultrapassou Limite"
      End If
    End If
  Else
    MsgBox "Este cliente esta marcado para ser limitado mas não possue valor definido!"
    Autorizar = True
    Motivo = "Sem limite"
  End If
End If

If dbClientes.Recordset("diapagamento") <> 0 Then
  If dbClientes.Recordset!diapagamento >= 28 Then
    DataPrevista = CDate(Format(UltimoDiaDoMes(Month(dbFechamento.Recordset!DataCaixa), Year(dbFechamento.Recordset!DataCaixa)), "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
  Else
    DataPrevista = CDate(Format(dbClientes.Recordset("diapagamento"), "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
  End If
Else
  DataPrevista = DateAdd("m", 1, dbFechamento.Recordset!DataCaixa)
End If
If DataPrevista < dbFechamento.Recordset!DataCaixa Then
  DataPrevista = DateAdd("m", 1, DataPrevista)
End If
'Calcula a kilometragem
If cboPlaca.Text <> "" Then
  With dbClientesNota2Temp
    Consumo = 0
    .Refresh
    If dbClientesCarros.Recordset.RecordCount <> 0 Then
      .Recordset.FindLast "codigocarro=" & dbClientesCarros.Recordset!codigocarro
      If .Recordset.NoMatch = False Then
        If IsNull(.Recordset!Km) = False Then
          If txtKm.Text <> "0" And txtKm.Text <> "" And txtLitros.Text <> "0" And txtLitros.Text <> "" Then
            Consumo = (CDbl(txtKm.Text) - .Recordset!Km) / CDbl(txtLitros.Text)
          Else
            Consumo = 0
          End If
        End If
      End If
    End If
  End With
End If
If IsNumeric(txtLitros.Text) = False Then txtLitros.Text = "0"
Qtd = CDbl(txtLitros.Text)
If cboProduto.Text = dbProdutos2.Recordset!Descri Then
  If dbProdutos2.Recordset!Combustivel = False Then
    With dbProdutosAltera
      .DatabaseName = Caminho
      .Connect = Conectar
      .RecordSource = "select *from qprodutosaltera where datacaixa<=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "# and codigoproduto=" & dbProdutos2.Recordset!CodigoProduto & " order by datacaixa desc, horaini desc"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "datacaixa<=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
        If .Recordset.NoMatch = True Then
          ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
          valorUnitario = dbProdutos2.Recordset!PrecoVenda
        Else
          If .Recordset!DataCaixa = dbFechamento.Recordset!DataCaixa Then
            If .Recordset!HoraIni <= dbFechamento.Recordset("fechamentodecaixa.HoraIni") Then
              ValorNocaixa = Qtd * .Recordset!PrecoVenda
              valorUnitario = .Recordset!PrecoVenda
            Else
              ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
              valorUnitario = dbProdutos2.Recordset!PrecoVenda
            End If
          Else
            ValorNocaixa = Qtd * .Recordset!PrecoVenda
            valorUnitario = .Recordset!PrecoVenda
          End If
        End If
      Else
        ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
        valorUnitario = dbProdutos2.Recordset!PrecoVenda
      End If
    End With
  Else
    If IsNumeric(cboBico.Text) = False Then
      MsgBox "Informe um bico correto!"
      cboBico.SetFocus
      Exit Sub
    End If
    With dbBicosEncerrantes
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "bico=" & cboBico.Text
        If .Recordset.NoMatch = True Then
          MsgBox "Informe um bico correto!"
          cboBico.SetFocus
          Exit Sub
        Else
          If .Recordset!Preco = 0 Then
            MsgBox "Erro no preço do bico!"
            Exit Sub
          End If
          ValorNocaixa = Qtd * .Recordset!Preco
          valorUnitario = .Recordset!Preco
        End If
      Else
        MsgBox "Não foi encontrado bico com este produto!"
        Exit Sub
      End If
    End With
  End If
End If

With dbClientesProdutos
  If .Recordset.RecordCount <> 0 Then
    If cboProduto.Text <> dbProdutos2.Recordset!Descri Then
      MsgBox "Este cliente tem preço diferenciado e não pode ser lançado sem indicar o produto!"
      Exit Sub
    End If
    .Recordset.FindFirst "codigocliente=" & dbClientes.Recordset!CodigoCliente & " and codigoproduto=" & dbProdutos2.Recordset!CodigoProduto & " and validade>=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
    If .Recordset.NoMatch = False Then
      If .Recordset!validade = dbFechamento.Recordset!DataCaixa Then
        If .Recordset!HoraIni >= dbFechamento.Recordset!HoraIni Then
          PrecoDif = True
        End If
      Else
        PrecoDif = True
      End If
    End If
    ValorUnitarioDif = 0
    If PrecoDif = True Then
      If .Recordset!Preco <> 0 Then
        TempValorPagar = Qtd * .Recordset!Preco
        ValorUnitarioDif = .Recordset!Preco
      Else
        TempValorPagar = Qtd * valorUnitario
        ValorUnitarioDif = valorUnitario
        If .Recordset!Porcento <> 0 Then
          TempValorPagar = TempValorPagar * .Recordset!Porcento
          ValorUnitarioDif = ValorUnitarioDif * .Recordset!Porcento
        End If
      End If
      If .Recordset!valorasomar <> 0 Then
        TempValorPagar = TempValorPagar + (Qtd * .Recordset!valorasomar)
        ValorUnitarioDif = ValorUnitarioDif + .Recordset!valorasomar
      End If
      TempDif = TempValorPagar - CCur(txtNotaValor.Text)
      If TempDif > 0.2 Or TempDif < -0.2 Then
        If Usuarios.Grupo.admDatas < 2 Then
          MsgBox "O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
          Autorizar = True
          Motivo = "Preço Diferenciado"
        Else
          Resposta = MsgBox("O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then Exit Sub
          Autorizar = True
          Autorizado = True
          Motivo = "Preço Diferenciado"
        End If
      End If
    Else
      ValorUnitarioDif = Qtd * valorUnitario
      TempDif = ValorUnitarioDif - CCur(txtNotaValor.Text)
      If TempDif > 0.01 Or TempDif < -0.01 Then
        MsgBox "Preço unitário incorreto!"
        Exit Sub
      End If
    End If
  Else
    If dbProdutos2.Recordset!Combustivel = True Then
      TempDif = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, cboBico.Text)
    Else
      TempDif = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno)
    End If
    TempDif = TempDif - (CCur(txtNotaValor.Text) / Qtd)
    If TempDif > 0.01 Or TempDif < -0.01 Then
      MsgBox "Preço unitário incorreto!"
      Exit Sub
    End If
  End If
End With
If ValorUnitarioDif = 0 Then ValorUnitarioDif = valorUnitario
If CCur(txtNotaValor.Text) < 0 Then
  ValorUnitarioDif = ValorUnitarioDif * -1
End If
txtLitros.Text = CCur(txtNotaValor.Text) / ValorUnitarioDif
With dbClientesNota
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset("codigocliente") = dbClientes.Recordset("codigoCliente")
  .Recordset("nome") = dbClientes.Recordset("nome")
  .Recordset("datalanc") = Now
  .Recordset("dataprevista") = DataPrevista
  .Recordset("valorprevisto") = CCur(txtNotaValor.Text)
  .Recordset!Data = dbFechamento.Recordset!DataCaixa
  .Recordset!Cupom = txtCupom.Text
  If IsNumeric(txtKm.Text) = True Then
    .Recordset!Km = CDbl(txtKm.Text)
  End If
  If cboPlaca.Text <> "" Then
    If dbClientesCarros.Recordset.RecordCount <> 0 Then
      .Recordset!Placa = dbClientesCarros.Recordset!Placa
      .Recordset!codigocarro = dbClientesCarros.Recordset!codigocarro
    Else
      .Recordset!Placa = cboPlaca.Text
    End If
  End If
  .Recordset!Litros = Qtd
  .Recordset!Consumo = Consumo
  .Recordset!Qtd = Qtd
  If cboProduto.Text = dbProdutos2.Recordset!Descri Then
    .Recordset!CodigoProduto = dbProdutos2.Recordset!Codigo
  End If
  If PrecoDif = True Then
    .Recordset!LucroDif = CCur(txtNotaValor.Text) - ValorNocaixa
    .Recordset!ValorUnitarioDif = CCur(txtNotaValor.Text) / Qtd
    .Recordset!valorUnitario = ValorNocaixa / Qtd
    .Recordset!ValorTotalDif = ValorNocaixa
  Else
    .Recordset!LucroDif = 0
    If Qtd <> 0 Then
      .Recordset!valorUnitario = CCur(txtNotaValor.Text) / Qtd
    End If
    .Recordset!ValorUnitarioDif = CCur(txtNotaValor.Text) / Qtd
    .Recordset!ValorTotalDif = CCur(txtNotaValor.Text)
  End If
  .Recordset!Autorizar = Autorizar
  .Recordset!Autorizado = Autorizado
  .Recordset!Motivo = Motivo
  .Recordset.Update
  .Refresh
End With
With dbClientes
  .Recordset.Edit
  If IsNull(.Recordset!UltimoAbastecimento) = True Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  If .Recordset!UltimoAbastecimento < dbFechamento.Recordset!DataCaixa Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  .Recordset!TotalNotas = .Recordset!TotalNotas + CCur(txtNotaValor.Text)
  .Recordset.Update
End With
cboClientesNota.Text = ""
txtNotaValor.Text = ""
txtCupom.Text = ""
txtLitros.Text = ""


TotalizaNotas
TotalResumo
On Error Resume Next
dbTemp.Close
db.Close

cboClientesNota.SetFocus
End Sub

Private Sub cmdIncluirDespesa_Click()
Dim StrObs As String

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If cboDespesa.Text <> dbDespesas.Recordset("descri") Then
  MsgBox "Selecione uma despesa válida!", vbCritical, "Erro!"
  cboDespesa.SetFocus
  Exit Sub
End If
If IsNumeric(txtDespesaValor.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtDespesaValor.SetFocus
  Exit Sub
End If

If txtDespesaObs.Visible = True Then
  StrObs = txtDespesaObs.Text
Else
  StrObs = cboSubGrupo.Text
End If

Select Case Obrigatorio
  Case "Mes e Ano Referência"
    StrObs = StrObs & " - Ref. " & Format(txtMesAno.Value, "MM/yyyy")
  Case "Período"
    StrObs = StrObs & " - De: " & Format(txtDataIni.Value, "short date") & " até " & Format(txtDataFim.Value, "short date")
  Case "Obs. Adicional"
    If Trim(txtObsAdicional.Text) = "" Then
      MsgBox "É preciso incluir uma observação adicional!"
      txtObsAdicional.SetFocus
      Exit Sub
    End If
    If Len(txtObsAdicional.Text) < 5 Then
      MsgBox "Observação adicional muito curta!"
      txtObsAdicional.SetFocus
      Exit Sub
    End If
    StrObs = StrObs & " - " & txtObsAdicional.Text
End Select

With dbDespesasLanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset!Origem = "Fechamento"
  .Recordset("data") = dbFechamento.Recordset!DataCaixa
  .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
  .Recordset("hora") = Now
  .Recordset("codigoconta") = -1
  .Recordset("conta") = "Fechamento de Caixa"
  .Recordset("codigodespesa") = dbDespesas.Recordset("codigodespesa")
  .Recordset("descri") = dbDespesas.Recordset("descri")
  .Recordset("obs") = StrObs
  .Recordset!compensado = True
  .Recordset("valor") = -CCur(txtDespesaValor.Text)
  .Recordset!valorpago = -CCur(txtDespesaValor.Text)
  .Recordset!codigoenviar = "1"
  .Recordset.Update
  .Refresh
End With

cboDespesa.Text = ""
txtDespesaValor.Text = ""

TotalizaDespesas
TotalResumo

cboDespesa.SetFocus

End Sub

Private Sub cmdIncluirMicrocredito_Click()
Dim DataPrevista As Date, Consumo As Double
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim Limite As Currency, valorUnitario As Currency, Qtd As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim ValorNocaixa As Currency, TempDif As Currency
Dim Autorizar As Boolean, Autorizado As Boolean, Motivo As String

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If cboMicrocredito.Text <> dbClientes2.Recordset("nome") Then
  MsgBox "Selecione um cliente válido!", vbCritical, "Erro!"
  cboClientesNota.SetFocus
  Exit Sub
End If
If IsNumeric(txtValorMicrocredito.Text) = False Then
  MsgBox "Informe um valor válido!", vbCritical, "Erro!"
  txtValorMicrocredito.SetFocus
  Exit Sub
End If
If CCur(txtValorMicrocredito.Text) <= 0 Then
  If CDbl(txtQtdMicrocredito.Text) >= 0 Then
    MsgBox "Para extornar uma nota a quantidade e o valor devem ser negativos!"
    Exit Sub
  End If
End If
If IsNumeric(txtCupomMicrocredito.Text) = False Then
  MsgBox "Informe o número do cupom válido!"
  txtCupom.SetFocus
  Exit Sub
End If

If Configura.NotaNoCaixa = 0 Then
  If IsNumeric(txtQtdMicrocredito.Text) = False Then
    Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes2.Recordset!CodigoCliente)
    txtQtdMicrocredito.Text = CCur(txtValorMicrocredito.Text) / Preco
  End If
  
'  If CDbl(txtLitros.Text) <= 0 Then
'    If CCur(txtNotaValor.Text) >= 0 Then
'      MsgBox "Para extornar uma nota a quantidade e o valor devem ser negativos!"
'    End If
'  End If
  If dbClientesCarros.Recordset.RecordCount <> 0 Then
    If cboPlacaMicrocredito.Text <> dbClientesCarros.Recordset!Placa Then
      MsgBox "Informe uma placa correta!"
      cboPlacaMicrocredito.SetFocus
      Exit Sub
    End If
  End If
  If IsNumeric(txtKmMicrocredito.Text) = False Then
    MsgBox "Informe um Km correto!"
    txtKm.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtQtdMicrocredito.Text) = False Then
    MsgBox "Informe a quantidade de litros!"
    txtLitros.SetFocus
    Exit Sub
  End If
  If dbProdutos2.Recordset.EOF = True Then
    MsgBox "Produto inválido!"
    Exit Sub
  End If
  If txtCodMicrocredito.Text <> dbProdutos2.Recordset!Codigo Then
    MsgBox "Produto inválido!"
    txtCodMicrocredito.SetFocus
    Exit Sub
  End If
  If dbClientes2.Recordset!mensalista = False Then
    If dbClientes2.Recordset!desativado < dbFechamento.Recordset!DataCaixa Then
      If Usuarios.Grupo.admDatas < 2 Then
        MsgBox "O cliente está desativado! Somente o administrador pode lançar."
        If Configura.NotaBloqueia = 0 Then
          Autorizar = True
          Motivo = "Desativado"
        End If
      Else
        MsgBox "O cliente está desativado!"
        If Configura.NotaBloqueia = 0 Then
          Autorizar = True
          Autorizado = True
          Motivo = "Desativado"
        End If
      End If
    End If
  End If
End If
If dbClientes2.Recordset!protestado = True Then
  MsgBox "Cliente bloqueado!"
  Autorizar = True
  Autorizado = True
  Motivo = "Bloqueado/Protestado"
End If
If dbClientes2.Recordset!limitar = True Then
  If IsNull(dbClientes2.Recordset!Limite) = False Then
    Limite = CCur(txtValorMicrocredito.Text)
    db.Open CaminhoADO
    dbTemp.Open "select sum(valorprevisto) as total from clientesnota2 where codigocliente=" & dbClientes2.Recordset!CodigoCliente & " and confirmado=0", db
    If IsNull(dbTemp!Total) = False Then
      Limite = Limite + dbTemp!Total
    End If
    dbTemp.Close
    dbTemp.Open "select sum(valor) as total from clientescobranca where codigocliente=" & dbClientes2.Recordset!CodigoCliente & " and pago=0", db
    If IsNull(dbTemp!Total) = False Then
      Limite = Limite + dbTemp!Total
    End If
    If Limite > dbClientes2.Recordset!Limite Then
      If Usuarios.Grupo.admDatas < 2 Then
        MsgBox "O cliente ultrapassará o limite de notas dele! Somente o administrador pode lançar."
        Autorizar = True
        Motivo = "Ultrapassou Limite"
      Else
        MsgBox "O cliente ultrapassará o limite de notas dele!"
        Autorizar = True
        Autorizado = True
        Motivo = "Ultrapassou Limite"
      End If
    End If
  Else
    MsgBox "Este cliente esta marcado para ser limitado mas não possue valor definido!"
    Autorizar = True
    Motivo = "Sem limite"
  End If
End If

If dbClientes2.Recordset("diapagamento") <> 0 Then
  If dbClientes2.Recordset!diapagamento >= 28 Then
    DataPrevista = CDate(Format(UltimoDiaDoMes(Month(dbFechamento.Recordset!DataCaixa), Year(dbFechamento.Recordset!DataCaixa)), "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
  Else
    DataPrevista = CDate(Format(dbClientes.Recordset("diapagamento"), "00") & "/" & Month(dbFechamento.Recordset!DataCaixa) & "/" & Year(dbFechamento.Recordset!DataCaixa))
  End If
Else
  DataPrevista = DateAdd("m", 1, dbFechamento.Recordset!DataCaixa)
End If
If DataPrevista < dbFechamento.Recordset!DataCaixa Then
  DataPrevista = DateAdd("m", 1, DataPrevista)
End If
'Calcula a kilometragem
If cboPlaca.Text <> "" Then
  With dbClientesNota2Temp
    Consumo = 0
    .Refresh
    If dbClientesCarros.Recordset.RecordCount <> 0 Then
      .Recordset.FindLast "codigocarro=" & dbClientesCarros.Recordset!codigocarro
      If .Recordset.NoMatch = False Then
        If IsNull(.Recordset!Km) = False Then
          If txtKm.Text <> "0" And txtKm.Text <> "" And txtLitros.Text <> "0" And txtLitros.Text <> "" Then
            Consumo = (CDbl(txtKm.Text) - .Recordset!Km) / CDbl(txtLitros.Text)
          Else
            Consumo = 0
          End If
        End If
      End If
    End If
  End With
End If
If IsNumeric(txtQtdMicrocredito.Text) = False Then txtQtdMicrocredito.Text = "0"
Qtd = CDbl(txtQtdMicrocredito.Text)
If cboProduto.Text = dbProdutos2.Recordset!Descri Then
  If dbProdutos2.Recordset!Combustivel = False Then
    With dbProdutosAltera
      .DatabaseName = Caminho
      .Connect = Conectar
      .RecordSource = "select *from qprodutosaltera where datacaixa<=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "# and codigoproduto=" & dbProdutos2.Recordset!CodigoProduto & " order by datacaixa desc, horaini desc"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "datacaixa<=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
        If .Recordset.NoMatch = True Then
          ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
          valorUnitario = dbProdutos2.Recordset!PrecoVenda
        Else
          If .Recordset!DataCaixa = dbFechamento.Recordset!DataCaixa Then
            If .Recordset!HoraIni <= dbFechamento.Recordset("fechamentodecaixa.HoraIni") Then
              ValorNocaixa = Qtd * .Recordset!PrecoVenda
              valorUnitario = .Recordset!PrecoVenda
            Else
              ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
              valorUnitario = dbProdutos2.Recordset!PrecoVenda
            End If
          Else
            ValorNocaixa = Qtd * .Recordset!PrecoVenda
            valorUnitario = .Recordset!PrecoVenda
          End If
        End If
      Else
        ValorNocaixa = Qtd * dbProdutos2.Recordset!PrecoVenda
        valorUnitario = dbProdutos2.Recordset!PrecoVenda
      End If
    End With
  Else
    If IsNumeric(cboBicoMicrocredito.Text) = False Then
      MsgBox "Informe um bico correto!"
      cboBico.SetFocus
      Exit Sub
    End If
    With dbBicosEncerrantes
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "bico=" & cboBicoMicrocredito.Text
        If .Recordset.NoMatch = True Then
          MsgBox "Informe um bico correto!"
          cboBicoMicrocredito.SetFocus
          Exit Sub
        Else
          ValorNocaixa = Qtd * .Recordset!Preco
          valorUnitario = .Recordset!Preco
        End If
      Else
        MsgBox "Não foi encontrado bico com este produto!"
        Exit Sub
      End If
    End With
  End If
End If

With dbClientesProdutos
  If .Recordset.RecordCount <> 0 Then
    If cboProdutoMicrocredito.Text <> dbProdutos2.Recordset!Descri Then
      MsgBox "Este cliente tem preço diferenciado e não pode ser lançado sem indicar o produto!"
      Exit Sub
    End If
    .Recordset.FindFirst "codigocliente=" & dbClientes2.Recordset!CodigoCliente & " and codigoproduto=" & dbProdutos2.Recordset!CodigoProduto & " and validade>=#" & DataInglesa(dbFechamento.Recordset!DataCaixa) & "#"
    If .Recordset.NoMatch = False Then
      If .Recordset!validade = dbFechamento.Recordset!DataCaixa Then
        If .Recordset!HoraIni >= dbFechamento.Recordset!HoraIni Then
          PrecoDif = True
        End If
      Else
        PrecoDif = True
      End If
    End If
    ValorUnitarioDif = 0
    If PrecoDif = True Then
      If .Recordset!Preco <> 0 Then
        TempValorPagar = Qtd * .Recordset!Preco
        ValorUnitarioDif = .Recordset!Preco
      Else
        TempValorPagar = Qtd * valorUnitario
        ValorUnitarioDif = valorUnitario
        If .Recordset!Porcento <> 0 Then
          TempValorPagar = TempValorPagar * .Recordset!Porcento
          ValorUnitarioDif = ValorUnitarioDif * .Recordset!Porcento
        End If
      End If
      If .Recordset!valorasomar <> 0 Then
        TempValorPagar = TempValorPagar + (Qtd * .Recordset!valorasomar)
        ValorUnitarioDif = ValorUnitarioDif + .Recordset!valorasomar
      End If
      TempDif = TempValorPagar - CCur(txtNotaValor.Text)
      If TempDif > 0.2 Or TempDif < -0.2 Then
        If Usuarios.Grupo.admDatas < 2 Then
          MsgBox "O cliente " & dbClientes2.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
          Autorizar = True
          Motivo = "Preço Diferenciado"
        Else
          Resposta = MsgBox("O cliente " & dbClientes2.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then Exit Sub
          Autorizar = True
          Autorizado = True
          Motivo = "Preço Diferenciado"
        End If
      End If
    Else
      ValorUnitarioDif = Qtd * valorUnitario
      TempDif = ValorUnitarioDif - CCur(txtValorMicrocredito.Text)
      If TempDif > 0.01 Or TempDif < -0.01 Then
        MsgBox "Preço unitário incorreto!"
        Exit Sub
      End If
    End If
  Else
    If dbProdutos2.Recordset!Combustivel = True Then
      TempDif = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, cboBico.Text)
    Else
      TempDif = PrecoAtual(dbProdutos2.Recordset!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno)
    End If
    TempDif = TempDif - (CCur(txtValorMicrocredito.Text) / Qtd)
    If TempDif > 0.01 Or TempDif < -0.01 Then
      MsgBox "Preço unitário incorreto!"
      Exit Sub
    End If
  End If
End With
If ValorUnitarioDif = 0 Then ValorUnitarioDif = valorUnitario
If CCur(txtValorMicrocredito.Text) < 0 Then
  ValorUnitarioDif = ValorUnitarioDif * -1
End If
txtLitros.Text = CCur(txtNotaValor.Text) / ValorUnitarioDif
With dbClientesNota
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset("codigocliente") = dbClientes2.Recordset("codigoCliente")
  .Recordset("nome") = dbClientes2.Recordset("nome")
  .Recordset("datalanc") = Now
  .Recordset("dataprevista") = DataPrevista
  .Recordset("valorprevisto") = CCur(txtValorMicrocredito.Text)
  .Recordset!Data = dbFechamento.Recordset!DataCaixa
  .Recordset!Cupom = txtCupomMicrocredito.Text
  If IsNumeric(txtKmMicrocredito.Text) = True Then
    .Recordset!Km = CDbl(txtKmMicrocredito.Text)
  End If
  If cboPlacaMicrocredito.Text <> "" Then
    If dbClientesCarros.Recordset.RecordCount <> 0 Then
      .Recordset!Placa = dbClientesCarros.Recordset!Placa
      .Recordset!codigocarro = dbClientesCarros.Recordset!codigocarro
    Else
      .Recordset!Placa = cboPlacaMicrocredito.Text
    End If
  End If
  .Recordset!Litros = Qtd
  .Recordset!Consumo = Consumo
  .Recordset!Qtd = Qtd
  If cboProduto.Text = dbProdutos2.Recordset!Descri Then
    .Recordset!CodigoProduto = dbProdutos2.Recordset!Codigo
  End If
  If PrecoDif = True Then
    .Recordset!LucroDif = CCur(txtValorMicrocredito.Text) - ValorNocaixa
    .Recordset!ValorUnitarioDif = CCur(txtValorMicrocredito.Text) / Qtd
    .Recordset!valorUnitario = ValorNocaixa / Qtd
    .Recordset!ValorTotalDif = ValorNocaixa
  Else
    .Recordset!LucroDif = 0
    If Qtd <> 0 Then
      .Recordset!valorUnitario = CCur(txtValorMicrocredito.Text) / Qtd
    End If
    .Recordset!ValorUnitarioDif = CCur(txtValorMicrocredito.Text) / Qtd
    .Recordset!ValorTotalDif = CCur(txtValorMicrocredito.Text)
  End If
  .Recordset!Autorizar = Autorizar
  .Recordset!Autorizado = Autorizado
  .Recordset!Motivo = Motivo
  .Recordset.Update
  .Refresh
End With
With dbClientes2
  .Recordset.Edit
  If IsNull(.Recordset!UltimoAbastecimento) = True Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  If .Recordset!UltimoAbastecimento < dbFechamento.Recordset!DataCaixa Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  .Recordset!TotalNotas = .Recordset!TotalNotas + CCur(txtValorMicrocredito.Text)
  .Recordset.Update
End With
cboMicrocredito.Text = ""
txtValorMicrocredito.Text = ""
txtCupomMicrocredito.Text = ""
txtQtdMicrocredito.Text = ""


TotalizaNotas
TotalResumo
On Error Resume Next
dbTemp.Close
db.Close

cboMicrocredito.SetFocus
End Sub

Private Sub cmdIncluirRecebimento_Click()
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency, DescontoPorcento As Currency

If dbFechamento.Recordset!distribuido = True Then Exit Sub

'If DateDiff("d", dbFechamento.Recordset!DataCaixa, txtDataBordero.Value) >= 30 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar borderô com data futura acima de 30 dias do caixa!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", dbFechamento.Recordset!DataCaixa, txtDataBordero.Value) <= -30 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar borderô com data anterior a 30 dias do caixa!"
'    Exit Sub
'  End If
'End If

If cboRecebimento.Text <> dbFormaDePg.Recordset("descri") Then
  MsgBox "Escolha uma forma de Pagamento válida!", vbCritical, "Erro!"
  cboRecebimento.SetFocus
  Exit Sub
End If
If IsNumeric(txtValorRecebe.Text) = False Then
  MsgBox "Informe um valor válido!", vbCritical, "Erro!"
  txtValorRecebe.SetFocus
  Exit Sub
End If
Tarifa = dbFormaDePg.Recordset("descontovalor")
Operacao = dbFormaDePg.Recordset("descontoporOperacao")
Porcento = dbFormaDePg.Recordset("descontoPorcento") / 100

TotalOper = 0
If Operacao <> 0 Then
  If IsNumeric(txtOperacoes.Text) = True Then
    TotalOper = CDbl(txtOperacoes.Text)
    If TotalOper = 0 Then
      MsgBox "Informe um valor correto para desconto por operação!"
      txtOperacoes.SetFocus
      Exit Sub
    Else
      Operacao = Operacao * TotalOper
    End If
  Else
    MsgBox "Informe um valor correto para desconto por operação!"
    txtOperacoes.SetFocus
    Exit Sub
  End If
End If
ValorBruto = CCur(txtValorRecebe.Text)

If Porcento <> 0 Then
  DescontoPorcento = ValorBruto * Porcento
End If

Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao
With dbFormaDePg
  If .Recordset!CodigoConta = 0 Then
    MsgBox "Esta forma de pagamento está sem conta destino!"
    Exit Sub
  End If
End With
With dbFormaDePgRecebido
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset("codigoformadepg") = dbFormaDePg.Recordset("codigoPagamento")
  .Recordset("descri") = dbFormaDePg.Recordset("descri")
  .Recordset("valorbruto") = ValorBruto
  .Recordset("valordescoper") = Operacao
  .Recordset("valordesctarifa") = Tarifa
  .Recordset("valordesconto") = DescontoPorcento
  .Recordset("valor") = Liquido
  .Recordset("operacoes") = TotalOper
  .Recordset("data") = txtDataBordero.Value
  .Recordset("hora") = Now
  .Recordset.Update
  .Refresh
End With

TotalRecebimento
TotalResumo

cboRecebimento.Text = ""
txtValorRecebe.Text = ""
txtOperacoes.Text = ""
cboRecebimento.SetFocus

End Sub

Private Sub cmdIncluirVale_Click()

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If cboFuncionario.Text <> dbVendedores.Recordset!Nome Then
  Call cboFuncionario_LostFocus
  If cboFuncionario.Text <> dbVendedores.Recordset!Nome Then
    MsgBox "Funcionário inválido!"
    cboFuncionario.SetFocus
    Exit Sub
  End If
End If
If cboMotivo.Text = "" Then
  MsgBox "Escolha um motivo válido!"
  cboMotivo.SetFocus
  Exit Sub
End If
If IsNumeric(txtValeValor.Text) = False Then
  MsgBox "Informe um valor válido!"
  txtValeValor.SetFocus
  Exit Sub
End If
With dbVales
  .Recordset.AddNew
  .Recordset!CodigoCaixa = dbFechamento.Recordset!CodigoFechamento
  .Recordset!codfun = dbVendedores.Recordset!codigovendedor
  .Recordset!Descri = cboMotivo.Text
  .Recordset!Valor = txtValeValor.Text
  .Recordset.Update
End With
qVales.Refresh
TotalizaVales
TotalResumo
End Sub

Private Sub cmdProdutoEntra_Click()
Dim PrecoAntigo As Currency, PrecoNovo As Currency, TempValor As Currency
Dim Variacao As Currency

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
If dbProdutos.Recordset.EOF = True Then
  MsgBox "Produto inválido!"
  txtCod.SetFocus
  Exit Sub
End If
If cboProdutoEntra.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Produto inválido!"
  txtCod.SetFocus
  Exit Sub
End If
If IsNumeric(txtQtdEntra.Text) = False Then
  MsgBox "Quantidade inválida!"
  txtQtdEntra.SetFocus
  Exit Sub
End If
If IsNumeric(txtTotalEntra.Text) = False Then
  MsgBox "Valor inválido!"
  txtTotalEntra.SetFocus
  Exit Sub
End If
If dbProdutos.Recordset!Combustivel = True Then
  If IsNumeric(txtTanqueEntra.Text) = False Then
    MsgBox "Tanque inválido!"
    txtTanqueEntra.SetFocus
    Exit Sub
  End If
End If
PrecoAntigo = dbProdutos.Recordset!precocompra
PrecoNovo = CCur(txtTotalEntra.Text) / CDbl(txtQtdEntra.Text)
Variacao = PrecoAntigo - PrecoNovo
'TempValor = PrecoNovo + dbProdutos.Recordset!comissaovalor + (dbProdutos.Recordset!precovenda * (dbProdutos.Recordset!Comissao / 100))
'
'If TempValor >= (dbProdutos.Recordset!precovenda / 2) Then
'  MsgBox "Margem de lucro abaixo de 50%! Custo= " & Format(TempValor, "Currency") & " / Venda= " & Format(dbProdutos.Recordset!precovenda, "Currency"), vbCritical
'End If

With dbProdutoEntra
  .Recordset.AddNew
  .Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
  .Recordset!Data = dbFechamento.Recordset!DataCaixa
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!Codigo = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!PrecoAntigo = PrecoAntigo
  .Recordset!PrecoNovo = PrecoNovo
  .Recordset!VariaEstoque = Variacao
  .Recordset!Quantidade = CDbl(txtQtdEntra.Text)
  .Recordset!valornota = CCur(txtTotalEntra.Text)
  If IsNumeric(txtTanqueEntra.Text) = True Then
    .Recordset!Tanque = CDbl(txtTanqueEntra.Text)
  Else
    .Recordset!Tanque = 0
  End If
  .Recordset.Update
  .Refresh
End With
txtCod.Text = ""
cboProdutoEntra.Text = ""
txtQtdEntra.Text = ""
txtTotalEntra.Text = ""
txtTanqueEntra.Text = ""


TotalizaCompras
TotalResumo

txtCod.SetFocus
End Sub

Private Sub cmdRelaciona_Click()
Dim Autorizar As Boolean, Autorizado As Boolean

If dbFechamento.Recordset!distribuido = True Then Exit Sub

If IsNumeric(MaskEdBox1(0).Text) = False Or IsNumeric(MaskEdBox1(1).Text) = False Or IsNumeric(MaskEdBox1(2).Text) = False Or IsNumeric(MaskEdBox1(4).Text) = False Or MaskEdBox1(3).Text = "      - " Then
  MsgBox "Dados incompletos!"
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If
If lblJurosTabelado.Caption = "ERR" Then
  MsgBox "Prazo não cadastrado!"
  Exit Sub
End If
If dbClientesCheques.Recordset.EOF = True Then
  MsgBox "Não foi cadastrado o cliente!"
  txtCodCliente.SetFocus
  Exit Sub
End If
If cboClienteCheque.Text <> dbClientesCheques.Recordset!Nome Then
  MsgBox "Não foi cadastrado o cliente!"
  txtCodCliente.SetFocus
  Exit Sub
End If
Autorizar = False
Autorizado = False
If Configura.ChequesNoCaixa = 0 Then
  If dbClientesCheques.Recordset!Posicao = False Then
    MsgBox "Este cliente está inativo!"
    Exit Sub
  End If
  If CCur(txtValor.Text) > CCur(lblSaldo.Caption) Then
    MsgBox "Este cheque irá ultrapassar o limite do cliente!"
    Autorizar = True
    If Usuarios.Grupo.admLiberaNotas = 2 Then
      Resposta = MsgBox("Deseja continuar o lançamento do cheque atual?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      Autorizado = True
    End If
  End If
End If
With dbCheques
  .Refresh
  .Recordset.FindFirst "comp='" & MaskEdBox1(0).Text & "' and banco='" & MaskEdBox1(1).Text & "' and agencia='" & MaskEdBox1(2).Text & "' and conta='" & MaskEdBox1(3).Text & "' and chequeNr='" & MaskEdBox1(4).Text & "'"
  If .Recordset.NoMatch = False Then
    MsgBox "Cheque já cadastrado!"
    Exit Sub
  End If
  .Recordset.AddNew
  .Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
  .Recordset!CMC7 = CodBar
  .Recordset!COMP = MaskEdBox1(0).Text
  .Recordset!Banco = MaskEdBox1(1).Text
  .Recordset!Agencia = MaskEdBox1(2).Text
  .Recordset!Conta = MaskEdBox1(3).Text
  .Recordset!chequenr = MaskEdBox1(4).Text
  .Recordset!DataLanc = Now
  .Recordset!datacheque = txtBomPara.Value
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset!codigoSoma = "1"
  .Recordset!valornabomba = CCur(lblValorNaBomba.Caption)
  .Recordset!diaspre = lblJurosTabelado.Caption
  If IsNull(dbClientesCheques.Recordset!CIC) = False Then
    If dbClientesCheques.Recordset!CIC = "" Then
      .Recordset!CPF = dbClientesCheques.Recordset!CNPJ
    Else
      .Recordset!CPF = dbClientesCheques.Recordset!CIC
    End If
  Else
    .Recordset!CPF = dbClientesCheques.Recordset!CNPJ
  End If
  .Recordset!CodigoCliente = dbClientesCheques.Recordset!codigochequecliente
  .Recordset!Usuario = Usuarios.Nome
  .Recordset!usuariolanc = Usuarios.Nome
  .Recordset!Autorizar = Autorizar
  .Recordset!Autorizado = Autorizado
  .Recordset.Update
  dbClientesCheques.Recordset.Edit
  If IsNull(dbClientesCheques.Recordset!saldopendente) = True Then
    dbClientesCheques.Recordset!saldopendente = 0
  End If
  dbClientesCheques.Recordset!saldopendente = dbClientesCheques.Recordset!saldopendente + CCur(txtValor.Text)
  dbClientesCheques.Recordset!numerodecheques = dbClientesCheques.Recordset!numerodecheques + 1
  dbClientesCheques.Recordset!Total = dbClientesCheques.Recordset!Total + CCur(txtValor.Text)
  dbClientesCheques.Recordset.Update
  .Refresh
End With

TotalizaCheque
TotalRecebimento
TotalResumo

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
txtValor.Text = ""
txtCMC7.Text = ""

lblStatus = ""
txtCodCliente.Text = ""
cboClienteCheque.Text = ""
txtCodCliente.SetFocus


End Sub

Private Sub cmdRemoveNota_Click()
Dim Resposta As Integer

If dbFechamento.Recordset!distribuido = True Then Exit Sub

With dbClientesNota
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione primeiro uma nota!"
    Exit Sub
  End If
  If .Recordset!Confirmado = True Then
    MsgBox "A nota atual já foi confirmada para cobrança! Para estornála deverá ser lançada com o valor negativo!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover a nota atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  With dbClientes
    .Recordset.FindFirst "codigocliente=" & dbClientesNota.Recordset!CodigoCliente
    .Recordset.Edit
    .Recordset!TotalNotas = .Recordset!TotalNotas - dbClientesNota.Recordset!ValorPrevisto
    .Recordset.Update
  End With
  .Recordset.Delete
  .Refresh
  
  TotalizaNotas
  TotalResumo
End With
End Sub

Private Sub cmdRemoverMicrocredito_Click()
Dim Resposta As Integer

If dbFechamento.Recordset!distribuido = True Then Exit Sub

With dbClientesNota2
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione primeiro uma nota!"
    Exit Sub
  End If
  If .Recordset!Confirmado = True Then
    MsgBox "A nota atual já foi confirmada para cobrança! Para estornála deverá ser lançada com o valor negativo!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover a nota atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  With dbClientes
    .Recordset.FindFirst "codigocliente=" & dbClientesNota.Recordset!CodigoCliente
    .Recordset.Edit
    .Recordset!TotalNotas = .Recordset!TotalNotas - dbClientesNota.Recordset!ValorPrevisto
    .Recordset.Update
  End With
  .Recordset.Delete
  .Refresh
  
  TotalizaNotas
  TotalResumo
End With
End Sub

Private Sub cmdRemoverVale_Click()
Dim Resposta As Integer

If dbFechamento.Recordset!distribuido = True Then Exit Sub

Resposta = MsgBox("Deseja remover o vale atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

If qVales.Recordset.EOF = True Then
  MsgBox "Selecione um registro a ser removido!"
  Exit Sub
End If
If qVales.Recordset!fechado = True Or qVales.Recordset!Cobrado = True Then
  MsgBox "Este vale já foi gravado!"
  Exit Sub
End If
With dbVales
  .RecordSource = "Select *from vales"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Tabela de Vales vazia!"
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  .Recordset.FindFirst "codigovale=" & qVales.Recordset!codigovale
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de vales!"
    Exit Sub
  End If
  .Recordset.Delete
  .Refresh
End With
qVales.Refresh
TotalizaVales
TotalResumo
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSomar_Click()

If dbFechamento.Recordset!distribuido = True Then Exit Sub

With dbPagamentos
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  A = .Recordset.AbsolutePosition
  .Recordset.Edit
  .Recordset!CodigoCaixa = dbFechamento.Recordset!CodigoFechamento
  .Recordset!DataCaixa = dbFechamento.Recordset!DataCaixa
  .Recordset!Turno = dbFechamento.Recordset!Turno
  .Recordset!usuarioconfirmou = Usuarios.Nome
  .Recordset.Update
  .Refresh
  
  dbPagamentosCaixa.Refresh
  
  TotalizaDespesas
  TotalResumo
  
  On Error Resume Next
  DBGrid8.SetFocus
  .Recordset.AbsolutePosition = A
End With
End Sub

Private Sub cmdSubtrair_Click()

If dbFechamento.Recordset!distribuido = True Then Exit Sub

With dbPagamentosCaixa
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  A = .Recordset.AbsolutePosition
  .Recordset.Edit
  .Recordset!CodigoCaixa = 0
  .Recordset!DataCaixa = Null
  .Recordset!Turno = Null
  .Recordset!usuarioconfirmou = Null
  .Recordset.Update
  .Refresh
  dbPagamentos.Refresh
  
  TotalizaDespesas
  TotalResumo
  
  On Error Resume Next
  DBGrid8.SetFocus
  .Recordset.AbsolutePosition = A
End With
End Sub

Private Sub dbClientesCheques_Reposition()
Dim Saldo As Currency
If Fechando = True Then Exit Sub
If dbClientesCheques.Recordset.EOF = True Then Exit Sub
lblLimite.Caption = Format(dbClientesCheques.Recordset!Limitevalor2, "Currency")
Saldo = dbClientesCheques.Recordset!Limitevalor2
If IsNull(dbClientesCheques.Recordset!saldopendente) = False Then
  lblPre.Caption = Format(dbClientesCheques.Recordset!saldopendente, "Currency")
Else
  lblPre.Caption = Format(0, "Currency")
End If
Saldo = Saldo - CCur(lblPre.Caption)

lblSaldo.Caption = Format(Saldo, "Currency")
End Sub

Private Sub dbDespesas_Reposition()
Dim CodigoDespesa As Double

If dbDespesas.Recordset.EOF = False Then
  CodigoDespesa = dbDespesas.Recordset!CodigoDespesa
  If IsNull(dbDespesas.Recordset!Obrigatorio) = False Then
    Obrigatorio = dbDespesas.Recordset!Obrigatorio
  Else
    Obrigatorio = "Nenhum"
  End If
Else
  CodigoDespesa = 0
  Obrigatorio = "Nenhum"
End If
With dbDespesaTipoGrupo
  .RecordSource = "Select *from despesatiposubgrupo where codigodespesatipo=" & CodigoDespesa & " order by descri"
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount = 0 Then
    cboSubGrupo.Visible = False
    txtDespesaObs.Visible = True
  Else
    cboSubGrupo.Visible = True
    txtDespesaObs.Visible = False
  End If
End With

Select Case Obrigatorio
  Case "Mes e Ano Referência"
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = True
    txtMesAno.Visible = True
  Case "Período"
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = True
    lblPeriodoA.Visible = True
    txtDataIni.Visible = True
    txtDataFim.Visible = True
    lblMesAno.Visible = False
    txtMesAno.Visible = False
  Case "Obs. Adicional"
    lblObsAdicional.Visible = True
    txtObsAdicional.Visible = True
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = False
    txtMesAno.Visible = False
  Case Else
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = False
    txtMesAno.Visible = False
End Select
End Sub

Private Sub dbFechamento_Reposition()
Dim db As New ADODB.Connection
If IsNull(dbFechamento.Recordset("fechamentodecaixa.codigopdv")) = True Then
  db.Open CaminhoADO
  db.Execute "update fechamentodecaixa set codigopdv=1 where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
  db.Close
End If
If Abrindo = False Then
  SSTab1.Visible = False
End If
End Sub

Private Sub DBGrid2_BeforeDelete(Cancel As Integer)

If dbFechamento.Recordset!distribuido = True Then Cancel = True

With dbFormaDePgRecebido
  If .Recordset.EOF = False Then
    If .Recordset!fechamentodiario = True Then
      MsgBox "Registro já gravado!"
      Cancel = True
    End If
  End If
End With
End Sub

Private Sub DBGrid3_BeforeDelete(Cancel As Integer)

If dbFechamento.Recordset!distribuido = True Then Cancel = True

If dbDespesasLanc.Recordset!fechamentodiario = True Then
  MsgBox "Esta despesa já foi gravada!"
  Cancel = True
End If
End Sub

Private Sub DBGrid4_DblClick()
Call cmdAutorizar_Click
End Sub

Private Sub DBGrid5_BeforeDelete(Cancel As Integer)
If dbFechamento.Recordset!distribuido = True Then Cancel = True
End Sub

Private Sub DBGrid5_LostFocus()
TotalRecebimento
TotalResumo
End Sub

Private Sub DBGrid6_BeforeDelete(Cancel As Integer)

If dbFechamento.Recordset!distribuido = True Then Cancel = True

If dbCheques.Recordset!codigoSoma <> "1" Or dbCheques.Recordset!compensado = True Or dbCheques.Recordset!devolvido = True Or dbCheques.Recordset!cobrando = True Or dbCheques.Recordset!protesto = True Or dbCheques.Recordset!Custodia = True Then
  MsgBox "Este cheque não poderá ser removido porque já deu sequencia nos procedimentos!"
  Cancel = True
End If
With dbClientesCheques
  .Recordset.FindFirst "codigochequecliente=" & dbCheques.Recordset!CodigoCliente
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de clientes!"
    Cancel = True
  Else
    .Recordset.Edit
    .Recordset!saldopendente = .Recordset!saldopendente - dbCheques.Recordset!Valor
    If .Recordset!saldopendente < 0 Then .Recordset!saldopendente = 0
    .Recordset!numerodecheques = .Recordset!numerodecheques - 1
    .Recordset!Total = .Recordset!Total - dbCheques.Recordset!Valor
    .Recordset.Update
  End If
End With
End Sub

Private Sub DBGrid6_DblClick()
Call cmdChequeMudaAutoriza_Click
End Sub

Private Sub DBGrid6_LostFocus()
TotalRecebimento
TotalResumo
End Sub

Private Sub DBGrid7_BeforeDelete(Cancel As Integer)

If dbFechamento.Recordset!distribuido = True Then Cancel = True

If dbVales.Recordset!fechado = True Then Cancel = True

End Sub

Private Sub DBGrid7_LostFocus()
TotalRecebimento
TotalResumo
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
Dim db As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset
Dim A As Double

Abrindo = True

StrTemp = GetSetting(App.EXEName, "Base", "COM")

StrTemp2 = GetSetting(App.EXEName, "Base", "Baud", "9600")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "Paridade", "n")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "DataBit", "8")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "StopBit", "1")

MSComm1.Settings = StrTemp2

If Usuarios.Grupo.AdmEstatus = 2 Then
  cmdCancelaFinaliza.Visible = True
Else
  cmdCancelaFinaliza.Visible = False
End If
If Usuarios.Grupo.admLiberaNotas = 2 Then
  cmdAutorizar.Enabled = True
  cmdChequeMudaAutoriza.Enabled = True
  cmdMudarAutorizaMicrocredito.Enabled = True
Else
  cmdAutorizar.Enabled = False
  cmdChequeMudaAutoriza.Enabled = False
  cmdMudarAutorizaMicrocredito.Enabled = False
End If
If Usuarios.Nome = "Usuário Master" Then
  cmdAutorizar.Enabled = True
  cmdChequeMudaAutoriza.Enabled = True
  cmdMudarAutorizaMicrocredito.Enabled = True
End If


If StrTemp <> "" Then
  If StrTemp <> "Sem" Then
    Porta = CInt(Right(StrTemp, 1))
  Else
    Porta = -1
  End If
End If
If Porta > 0 Then
  Timer1.Enabled = True
  MSComm1.CommPort = Porta
  Call Image1_DblClick
  On Error GoTo 0
End If

db.Open CaminhoADO
dbConfig.CursorLocation = adUseClient
dbConfig.Open "Select *from config", db, adOpenKeyset, adLockOptimistic


txtDataBordero.Value = Date
With dbFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select fechamentodecaixa.*, pdvs.* from fechamentodecaixa left join pdvs on fechamentodecaixa.codigopdv=pdvs.codigopdv where datacaixa>=#" & DataInglesa(DateAdd("m", -12, Date)) & "# order by datacaixa, fechamentodecaixa.horaini"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    .Recordset.FindFirst "distribuido=0"
    If .Recordset.NoMatch = True Then
      .Recordset.MoveLast
    End If
  End If
End With
With dbFormaDePg
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
If dbConfig.RecordCount <> 0 Then
  On Error Resume Next
  If IsNull(dbConfig!clientesnotaplano) = False Then
    PlanoNotas = "'" & dbConfig!clientesnotaplano & "'"
    PlanoNotas = Replace(PlanoNotas, ",", "','")
  End If
  If IsNull(dbConfig!microcreditoplano) = False Then
    PlanoMicrocredito = "'" & dbConfig!microcreditoplano & "'"
    PlanoMicrocredito = Replace(PlanoMicrocredito, ",", "','")
  End If
End If

With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  If PlanoNotas = "" Then
    .RecordSource = "select *from Clientes order by Nome"
  Else
    .RecordSource = "select *from Clientes where PlanoDeConta in (" & PlanoNotas & ") order by Nome"
  End If
  .Refresh
End With
With dbClientes2
  .Connect = Conectar
  .DatabaseName = Caminho
  If PlanoNotas = "" Then
    .RecordSource = "select *from Clientes order by Nome"
  Else
    .RecordSource = "select *from Clientes where planodeconta in (" & PlanoMicrocredito & ") order by Nome"
  End If
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientesCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbChequesContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbJuros
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Data1 = CDate("31/12/2060")
    .Recordset!Data2 = CDate("31/12/2060")
    .Recordset!bloqueia1 = 0
    .Recordset!bloqueia2 = 0
    .Recordset.Update
    .Refresh
  End If
  .RecordSource = "select bloqueiafechamento.*, Turnos.* from bloqueiafechamento, turnos where bloqueiafechamento.coditoturno2=turnos.codigoturno"
  .Refresh
End With

With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from concilianova where datalanc>=#" & DataInglesa(DateAdd("m", -1, Date)) & "#"
  .Refresh
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTanques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutosNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from produtosnotas where codigofornecedor=-1"
  .Refresh
End With
With dbProdutosNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from produtosnotascorpo where codigoprodutonota=0"
  .Refresh
End With
With dbPosto
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVendedores
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientesNota2Temp
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaTipoGrupo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientesCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from clientescarros where codigocliente=0 order by Placa"
  .Refresh
End With
With dbClientesProdutos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbPagamentos
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from vendedorespagamento where codigocaixa=0 and pago=-1 order by funcionario"
  .Refresh
End With
With dbPagamentosCaixa
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from vendedorespagamento where codigocaixa=0 order by funcionario"
  .Refresh
End With
With dbProdutos2
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With

Abrindo = False
If dbFechamento.Recordset.RecordCount <> 0 Then
  If dbFechamento.Recordset.EOF = False Then
    AbreFechamento dbFechamento.Recordset!CodigoFechamento, dbFechamento.Recordset!DataCaixa
  End If
End If

If Usuarios.Nome = "Usuário Master" Then
  cmdImportarNotas.Visible = True
Else
  cmdImportarNotas.Visible = False
End If

With Usuarios.Grupo
  If .ClientesPlanos <> "" Then
    A = InStr(1, Usuarios.Grupo.ClientesPlanos, ",")
  Else
    
  End If
End With

Select Case Usuarios.Grupo.ControleConferencia
  Case 0 'Somente leitura
    cmdIncluirRecebimento.Enabled = False
    DBGrid2.AllowDelete = False
    cmdIncluirDespesa.Enabled = False
    DBGrid3.AllowDelete = False
    cmdInclueNota.Enabled = False
    DBGrid4.AllowDelete = False
    cmdProdutoEntra.Enabled = False
    DBGrid5.AllowDelete = False
    cmdRelaciona.Enabled = False
    DBGrid6.AllowDelete = False
    cmdIncluirVale.Enabled = False
    cmdRemoverVale.Enabled = False
    DBGrid7.AllowDelete = False
    cmdFinalizar.Enabled = False
    txtChequeBomba.Enabled = False
    txtChequeJuros.Enabled = False
    cmdIncluirMicrocredito.Enabled = False
    cmdRemoverMicrocredito.Enabled = False
    cmdImportarMicrocredito.Enabled = False
    
  Case 2 'Liberado
    
End Select

dbConfig.Close
db.Close

End Sub

Private Sub Image1_DblClick()
On Error Resume Next

With MSComm1
  If .PortOpen = True Then
    .PortOpen = False
  Else
    StrTemp = GetSetting(App.EXEName, "Base", "COM")

    StrTemp2 = GetSetting(App.EXEName, "Base", "Baud", "9600")
    StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "Paridade", "n")
    StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "DataBit", "8")
    StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "StopBit", "1")
    
    MSComm1.Settings = StrTemp2
    
    If StrTemp <> "" Then
      If StrTemp <> "Sem" Then
        Porta = CInt(Right(StrTemp, 1))
      Else
        Porta = -1
      End If
    End If
    If Porta > 0 Then
      Timer1.Enabled = True
      MSComm1.CommPort = Porta
    End If
    .PortOpen = True
  End If
  If .PortOpen = False Then
    Image1.Picture = LoadResPicture(102, vbResBitmap)
  Else
    Image1.Picture = LoadResPicture(101, vbResBitmap)
  End If
End With
Select Case Usuarios.Grupo.ControleConferencia
  Case 0, 1
    cmdFinalizar.Enabled = False
  Case 2
    cmdFinalizar.Enabled = True
End Select

End Sub

Private Sub lblJuros_Change()
lblJurosResumo.Caption = lblJuros.Caption
End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)
With MaskEdBox1(Index)
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case SSTab1.Tab
  Case 0
    If SSTab1.Enabled = True Then txtChequeBomba.SetFocus
  Case 1
    If SSTab1.Enabled = True Then cboDespesa.SetFocus
  Case 2
    If SSTab1.Enabled = True Then cboClientesNota.SetFocus
  Case 3
    If SSTab1.Enabled = True Then txtCod.SetFocus
  Case 4
    If SSTab1.Enabled = True Then txtCodCliente.SetFocus
  Case 5
    If SSTab1.Enabled = True Then cboFuncionario.SetFocus
  Case 6
    If SSTab1.Enabled = True Then cboFuncionario.SetFocus
End Select
End Sub

Private Sub Timer1_Timer()

If MSComm1.InBufferCount > 0 Then
  Timer1.Enabled = False

  'recebeu o codigo de barras armazena na variavel o codigo de barras
  CodBar = ""
  CodBar = MSComm1.Input
  If Len(CodBar) > 1 Then
    Do While Asc(Mid(CodBar, Len(CodBar) - 1, 1)) <> 3
      DoEvents
      CodBar = CodBar & MSComm1.Input
      If MSComm1.InBufferCount = 0 Then Exit Do
    Loop
    CodBar = Mid(CodBar, 1, Len(CodBar) - 1)
    CodBar = Converte(Trim(CodBar))
    If Len(CodBar) >= 33 Then
      'txtCodigo.Text = CodBar
      On Error Resume Next
      MaskEdBox1(0).Text = Mid(CodBar, 11, 3)
      MaskEdBox1(1).Text = Mid(CodBar, 2, 3)
      MaskEdBox1(2).Text = Mid(CodBar, 5, 4)
      MaskEdBox1(3).Text = Mid(CodBar, 26, 6) & "-" & Mid(CodBar, 32, 1)
      MaskEdBox1(4).Text = Mid(CodBar, 14, 6)
      
      txtBomPara.SetFocus
      
      With dbCheques
        .Refresh
        If .Recordset.RecordCount <> 0 Then
          .Recordset.FindFirst "comp='" & MaskEdBox1(0) & "' and banco='" & MaskEdBox1(1) & "' and agencia='" & MaskEdBox1(2) & "' and conta='" & MaskEdBox1(3) & "' and chequenr='" & MaskEdBox1(4) & "'"
          If .Recordset.NoMatch = False Then
            txtBomPara.Value = .Recordset("datacheque")
            txtValor.Text = Format(.Recordset("valor"), "Currency")
          End If
        End If
      End With
    End If
  End If
  Timer1.Enabled = True
End If


End Sub

Private Sub txtBomPara_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtBomPara_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtBomPara_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtChequeBomba_GotFocus()
With txtChequeBomba
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtChequeBomba_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtChequeBomba_LostFocus()
With txtChequeBomba
  If .Text = "" Then .Text = "0"
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "Currency")
  End If
End With
TotalRecebimento
TotalResumo
End Sub

Private Sub txtChequeJuros_GotFocus()
With txtChequeJuros
  .SelStart = 0
  .SelLength = Len(.Text)
End With

End Sub

Private Sub txtChequeJuros_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtChequeJuros_LostFocus()
With txtChequeJuros
  If .Text = "" Then .Text = "0"
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "Currency")
  End If
End With
TotalRecebimento
TotalResumo
End Sub

Private Sub txtCMC7_Change()
Dim Cheque As DadosCheque
Dim Cheque2 As CMC7

Cheque = ConverteCMC7(txtCMC7.Text)

If Cheque.COMP = "" Then Exit Sub

Cheque2.CMC7 = txtCMC7.Text
Cheque2 = CMC7Define(Cheque2)
If Cheque2.Validado = False Then
  MsgBox "Não foi possível validar a leitura do cheque! Código incorreto!"
  Exit Sub
End If
With Cheque
  MaskEdBox1(0).Text = .COMP
  MaskEdBox1(1).Text = .Banco
  MaskEdBox1(2).Text = .Agencia
  MaskEdBox1(3).Text = .Conta
  MaskEdBox1(4).Text = .Cheque
  CodBar = txtCMC7.Text
  txtBomPara.SetFocus
End With

End Sub

Private Sub txtCod_GotFocus()
With txtCod
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCod_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCod.Text = "" Then Exit Sub
  If IsNumeric(txtCod.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCod.Text
  If .Recordset.NoMatch = False Then
    cboProdutoEntra.Text = .Recordset!Descri
    txtCod.Text = .Recordset!Codigo
  End If
End With
End Sub

Private Sub txtCodCliente_GotFocus()
With dbChequesContas
  .RecordSource = "Select *from chequescontas"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    If MaskEdBox1(0).Text = "   " Or MaskEdBox1(1).Text = "   " Or MaskEdBox1(2).Text = "    " Or MaskEdBox1(0).Text = "      - " Then Exit Sub
    .Recordset.FindFirst "comp='" & MaskEdBox1(0).Text & "' and banconumero=" & CDbl(MaskEdBox1(1).Text) & " and agencia=" & CDbl(MaskEdBox1(2).Text) & " and conta='" & MaskEdBox1(3).Text & "'"
    If .Recordset.NoMatch = False Then
      dbClientesCheques.Recordset.FindFirst "codigochequecliente=" & .Recordset!CodigoCliente
      txtCodCliente.Text = dbClientesCheques.Recordset!codigochequecliente
      Call txtCodCliente_LostFocus
    End If
  End If
End With
End Sub

Private Sub txtCodCliente_LostFocus()
If txtCodCliente.Text = "" Then Exit Sub
With dbClientesCheques
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigochequecliente=" & txtCodCliente.Text
  If .Recordset.NoMatch = False Then
    txtCodCliente.Text = .Recordset!codigochequecliente
    cboClienteCheque.Text = .Recordset!Nome
    AlertaAtivo .Recordset!Posicao
  End If
End With
End Sub

Private Sub txtCodMicrocredito_GotFocus()
With txtCodMicrocredito
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodProduto_GotFocus()
With txtCodProduto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCupom_GotFocus()
With txtCupom
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDataBordero_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataBordero_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataBordero_LostFocus()
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

Private Sub txtDespesaObs_GotFocus()
With txtDespesaObs
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDespesaValor_GotFocus()
With txtDespesaValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDespesaValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtDespesaValor_LostFocus()
With txtDespesaValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
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

Private Sub txtNotaValor_GotFocus()
With txtNotaValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNotaValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtNotaValor_LostFocus()
With txtNotaValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "0.000")
  If IsNumeric(txtLitros.Text) = False Then
    Preco = PrecoCliente(dbProdutos2.Recordset!CodigoProduto, dbClientes.Recordset!CodigoCliente)
    txtLitros.Text = Format(CCur(txtNotaValor.Text) / Preco, "#,###.###")
  End If
End With
End Sub

Private Sub txtOperacoes_GotFocus()
With txtOperacoes
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtQtdEntra_GotFocus()
With txtQtdEntra
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTanqueEntra_GotFocus()
With txtChequeBomba
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotalEntra_GotFocus()
With txtTotalEntra
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotalEntra_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtTotalEntra_LostFocus()
Dim Qtd As Double, ICMS As Double, IPI As Double, Unitario As Currency
Dim Total As Currency

With txtTotalEntra
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.000")
  Total = CCur(.Text)
  If IsNumeric(txtQtdEntra.Text) = True Then
    Qtd = CDbl(txtQtdEntra.Text)
  End If
  Unitario = Total / Qtd
  txtValorUnitario.Text = Format(Unitario, "#,##0.000")
End With
End Sub

Private Sub txtValeValor_GotFocus()
With txtValeValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValeValor_LostFocus()
With txtValeValor
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtValor_LostFocus()
Dim Dias As Double, Taxa As Double
Dim Valor As Currency
With txtValor
  lblValorNaBomba.Caption = ""
  lblJurosTabelado.Caption = "ERR"
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
  Dias = DateDiff("d", dbFechamento.Recordset!DataCaixa, CDate(txtBomPara.Value))
  If Dias < 0 Then
    Exit Sub
  End If
  With dbJuros
    .Refresh
    If .Recordset.RecordCount = 0 Then
      Exit Sub
    Else
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        If .Recordset!Inicio <= Dias And .Recordset!final >= Dias Then
          lblJurosTabelado.Caption = Format((.Recordset!Taxa * 100), "0.00")
          Valor = CCur(txtValor.Text)
          Valor = Valor / (.Recordset!Taxa + 1)
          lblValorNaBomba.Caption = Format(Valor, "currency")
          Exit Sub
        End If
        .Recordset.MoveNext
      Loop
    End If
    
  End With
End With
End Sub

Private Sub txtValorRecebe_GotFocus()
With txtValorRecebe
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValorRecebe_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValorRecebe_LostFocus()
With txtValorRecebe
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtValorUnitario_GotFocus()
With txtValorUnitario
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValorUnitario_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValorUnitario_LostFocus()
Dim Qtd As Double, ICMS As Double, IPI As Double, Unitario As Currency
Dim Total As Currency

'lblUnitarioCalc.Caption = ""

With txtValorUnitario
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.000")
  If IsNumeric(txtQtdEntra.Text) = False Then Exit Sub
  Qtd = CDbl(txtQtdEntra.Text)
  Total = CDbl(.Text) * Qtd
  txtTotalEntra.Text = Format(Total, "#,##0.000")
End With
End Sub

