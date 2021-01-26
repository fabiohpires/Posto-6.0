VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFechamentoDeCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   6930
   ClientLeft      =   -180
   ClientTop       =   450
   ClientWidth     =   11295
   Icon            =   "frmFechamentoDeCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDesconfirmar 
      Caption         =   "Cancela Confirmação"
      Height          =   255
      Left            =   6600
      TabIndex        =   57
      ToolTipText     =   "Vulgo Papel Higiênico"
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   375
      Left            =   9960
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Im&primir"
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin MSDataListLib.DataCombo cboPdvs 
      Bindings        =   "frmFechamentoDeCaixa.frx":0442
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Frame Frame5 
      Caption         =   "DBFs"
      Height          =   6495
      Left            =   4560
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Data dbDespesasLanc2 
         Caption         =   "dbDespesasLanc2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesasLanc2"
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbTurno 
         Caption         =   "dbTurno"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Turnos order by HoraIni"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Data dbClientesProdutos 
         Caption         =   "dbClientesProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesProdutos order by codigocliente, codigoproduto, validade, horaini"
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSAdodcLib.Adodc dbImportacao 
         Height          =   375
         Left            =   5640
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbImportacao"
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
      Begin VB.Data dbDespesasTipo 
         Caption         =   "dbDespesasTipo"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesaTipo"
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbDespesasLanc 
         Caption         =   "dbDespesasLanc"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesasLanc2"
         Top             =   600
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
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   5280
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbClientesNotas 
         Caption         =   "dbClientesNotas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ClientesNota2"
         Top             =   4920
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data qProdutosAltera 
         Caption         =   "qProdutosAltera"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QProdutosAltera"
         Top             =   4560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbConfig 
         Caption         =   "dbConfig"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "config"
         Top             =   4200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbEstatus 
         Caption         =   "dbEstatus"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Status"
         Top             =   3840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbNotasCorpo 
         Caption         =   "dbNotasCorpo"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ProdutosNotasCorpo where aguardando=-1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbAlteracao 
         Caption         =   "dbAlteracao"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Alteracoes"
         Top             =   2760
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbAlteraBico 
         Caption         =   "dbAlteraBico"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AlteraBico"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbUltimoEncerrante 
         Caption         =   "dbUltimoEncerrante"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from BicoEncerrantes where codigofechamento=0 order by bico"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Data dbUltimoFechamento 
         Caption         =   "dbUltimoFechamento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FechamentoDeCaixa where fechado=-1 order by datacaixa, HoraIni desc"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Data dbBicoEncerrantes2 
         Caption         =   "dbBicoEncerrantes2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from BicoEncerrantes where codigofechamento=0 order by bico"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Data dbFechamento2 
         Caption         =   "dbFechamento2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FechamentoDeCaixa order by datacaixa, HoraIni"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Data dbPostos 
         Caption         =   "dbPostos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Postos"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Data dbResponsavel2 
         Caption         =   "dbResponsavel2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Vendedores order by nome"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Data dbDifComb 
         Caption         =   "dbDifComb"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DiferencaCombustivel where codigofechamento=0"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Data dbFormaDePg 
         Caption         =   "dbFormaDePg"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   2  'ServerSideCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FormaDePagamento"
         Top             =   5280
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbFormaDePgRecebido 
         Caption         =   "dbFormaDePgRecebido"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FormaDePagamentoRecebido2"
         Top             =   5640
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbClientesCarros 
         Caption         =   "dbClientesCarros"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ClientesNota2"
         Top             =   4920
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbProdutosHistorico 
         Caption         =   "dbProdutosHistorico"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ProdutosHistorico"
         Top             =   4560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbBloqueiaFechamento 
         Caption         =   "dbBloqueiaFechamento"
         Connect         =   "Access 2000;"
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
         Top             =   4200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbEstacionamento 
         Caption         =   "dbEstacionamento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Estacionamento"
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbEstacionamentoCaixa 
         Caption         =   "dbEstacionamentoCaixa"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "EstacionamentoCaixa"
         Top             =   3840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Produtos where combustivel=0 order by codigo"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Data dbVendas 
         Caption         =   "dbVendas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Venda2 where codigofechamento=0 order by codproduto"
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Data dbTanquesEstoque 
         Caption         =   "dbTanquesEstoque"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from TanqueEstoque where codigofechamento=0 order by tanque"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Data dbTanques 
         Caption         =   "dbTanques"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Tanques order by tanque"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Data dbBicoEncerrantes 
         Caption         =   "dbBicoEncerrantes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from BicoEncerrantes where codigofechamento=0 order by bico"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Data dbBicos 
         Caption         =   "dbBicos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Bicos order by bico"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Data dbFechamento 
         Caption         =   "dbFechamento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FechamentoDeCaixa order by datacaixa, HoraIni"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Data dbResponsavel 
         Caption         =   "dbResponsavel"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Vendedores order by nome"
         Top             =   240
         Width           =   3015
      End
      Begin MSAdodcLib.Adodc dbCupons 
         Height          =   375
         Left            =   5640
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbCupons"
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
      Begin MSAdodcLib.Adodc dbVendasLeituraX 
         Height          =   375
         Left            =   5640
         Top             =   2640
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbVendasLeituraX"
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
      Begin MSAdodcLib.Adodc dbResultado 
         Height          =   375
         Left            =   5640
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbResultado"
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
      Begin MSAdodcLib.Adodc dbEncerrantesNovos 
         Height          =   375
         Left            =   5640
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "bicosencerrantesnovos"
         Caption         =   "dbEncerrantesNovos"
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
      Begin MSAdodcLib.Adodc dbPDVs 
         Height          =   375
         Left            =   5640
         Top             =   3720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         RecordSource    =   "select *from pdvs order by descri"
         Caption         =   "dbPDVs"
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
   Begin VB.CommandButton cmdUltimo 
      Caption         =   ">>|"
      Height          =   255
      Left            =   1200
      TabIndex        =   61
      ToolTipText     =   "Último (F4)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimeiro 
      Caption         =   "|<<"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      ToolTipText     =   "Primeiro (F1)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdPosterior 
      Caption         =   ">>"
      Height          =   255
      Left            =   840
      TabIndex        =   59
      ToolTipText     =   "Prócimo (F3)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<<"
      Height          =   255
      Left            =   480
      TabIndex        =   58
      ToolTipText     =   "Anterior (F2)"
      Top             =   480
      Width           =   375
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   3720
      TabIndex        =   56
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   161
      FullHeight      =   33
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmFechamentoDeCaixa.frx":0457
      Height          =   3015
      Left            =   240
      OleObjectBlob   =   "frmFechamentoDeCaixa.frx":0477
      TabIndex        =   14
      Top             =   1560
      Width           =   3495
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmFechamentoDeCaixa.frx":1376
      Height          =   2655
      Left            =   3720
      OleObjectBlob   =   "frmFechamentoDeCaixa.frx":1395
      TabIndex        =   15
      Top             =   1560
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmFechamentoDeCaixa.frx":1F4C
      Height          =   2295
      Left            =   6120
      OleObjectBlob   =   "frmFechamentoDeCaixa.frx":1F63
      TabIndex        =   23
      Top             =   1560
      Width           =   4935
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Bindings        =   "frmFechamentoDeCaixa.frx":3032
      Height          =   1815
      Left            =   4440
      OleObjectBlob   =   "frmFechamentoDeCaixa.frx":304A
      TabIndex        =   29
      Top             =   4800
      Width           =   6495
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Ca&ncelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin MSDBCtls.DBCombo cboTurno 
      Bindings        =   "frmFechamentoDeCaixa.frx":4105
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37600
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Con&firmar"
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   5520
         Width           =   1215
      End
      Begin MSDBCtls.DBCombo cboResponsavel 
         Bindings        =   "frmFechamentoDeCaixa.frx":411B
         DataField       =   "CodigoResponsavel"
         DataSource      =   "dbFechamento"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   "CodigoVendedor"
         Text            =   ""
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar dados"
         Height          =   255
         Left            =   3720
         TabIndex        =   55
         ToolTipText     =   "Importa os dados do caixa (F5)"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtEstacionaCanc 
         Alignment       =   1  'Right Justify
         DataField       =   "Cancelados"
         DataSource      =   "dbEstacionamentoCaixa"
         Height          =   285
         Left            =   9000
         TabIndex        =   26
         Text            =   "0"
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEstacionaFim 
         Alignment       =   1  'Right Justify
         DataField       =   "Final"
         DataSource      =   "dbEstacionamentoCaixa"
         Height          =   285
         Left            =   8160
         TabIndex        =   25
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtEstacionaIni 
         Alignment       =   1  'Right Justify
         DataField       =   "Inicial"
         DataSource      =   "dbEstacionamentoCaixa"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         TabIndex        =   24
         Top             =   3480
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtInformado 
         DataField       =   "Informado"
         DataSource      =   "dbFechamento"
         Height          =   300
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   10200
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtCodFunc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtCodProduto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6120
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   1680
         TabIndex        =   27
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   " Resumo "
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   10815
         Begin VB.CommandButton cmdEntraCombustivel 
            Caption         =   "Entra Tanque"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblDiferenca 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2400
            TabIndex        =   46
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Diferença calculada:"
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblComissoes 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2400
            TabIndex        =   40
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Comissões a pagar:"
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Venda de Combustível:"
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Venda de Produtos+Estac.:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Faturamento calculado:"
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblTotalCombustivel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2400
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblTotalProdutos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblFaturamento 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2400
            TabIndex        =   33
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Label lblEstacionaTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9600
         TabIndex        =   54
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Total:"
         Height          =   255
         Left            =   9600
         TabIndex        =   53
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Canc.:"
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Final:"
         Height          =   255
         Left            =   8160
         TabIndex        =   51
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Inicial:"
         Height          =   255
         Left            =   7320
         TabIndex        =   50
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Estacionamento:"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblEstoque 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8400
         TabIndex        =   48
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Faturamento Informado:"
         Height          =   255
         Left            =   4080
         TabIndex        =   44
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Responsável:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   6000
         X2              =   6000
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Label lblTotalVenda 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9000
         TabIndex        =   42
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Total:"
         Height          =   195
         Left            =   9720
         TabIndex        =   41
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Func.:"
         Height          =   195
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtd:"
         Height          =   195
         Left            =   6600
         TabIndex        =   18
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod.:"
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Estoque:"
         Height          =   255
         Left            =   8400
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label10 
      Caption         =   "PDV:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Data:"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      Caption         =   "&Turno:"
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmFechamentoDeCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechamentoAnterior As Double, ErroNaSoma As Boolean
Dim Abrindo As Boolean, AlteraAnterior As Double, FechandoLote As Boolean

Private Sub GravaBloqueado(ByVal CodigoCliente As Double, ByVal Cliente As String, ByVal Cupom As Double, ByVal ValorTotal As Currency, ByVal Motivo As String)
Dim intArquivo As Integer, StrTemp As String

intArquivo = FreeFile()

Open App.Path & "\NotasBloqueadas.txt" For Append As intArquivo

Print #intArquivo, String(80, "*")
StrTemp = Motivo
Print #intArquivo, StrTemp
StrTemp = "Data - Turno: " & txtData.Value & " - " & cboTurno.Text & " - Codigo: " & CodigoCliente & " - Nome: " & Cliente
Print #intArquivo, StrTemp
StrTemp = "Nota nr: " & Cupom & " - Valor: " & Format(ValorTotal, "Currency")
Print #intArquivo, StrTemp

Print #intArquivo, String(80, "*")

Close intArquivo


End Sub



Private Function FecharCaixa() As Boolean
Dim Resposta As Integer, LucroVenda As Currency, StrTemp As String
Dim Estacionamento As Currency, ValorEstoque As Currency
Dim Vendas As Double, LucroMedio As Currency, PrecoMedio As Currency


FecharCaixa = False
Call cmdAbrir_Click

AtualizaSequenciaCaixa

With dbBloqueiaFechamento
  If .Recordset.RecordCount <> 0 Then
    If .Recordset!bloqueia1 = True Then
      If .Recordset!Data1 <= dbFechamento.Recordset!DataCaixa And .Recordset!bloqueia1 = True Then
        If .Recordset!HoraIni <= dbFechamento.Recordset!HoraIni Then
          MsgBox "Caixa não pode ser confirmado por estar bloqueado pelo administrador!"
          Exit Function
        End If
      End If
    End If
  End If
End With

With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!Tanque <> 0 Then
        If .Recordset!Diferenca > 1000 Or .Recordset!Diferenca < -1000 Then
          If Usuarios.Grupo.AdmEstatus = 2 Then
            Resposta = MsgBox("Diferenca de combustivel muito alta! Deseja continuar?", vbYesNo + vbDefaultButton2)
            If Resposta = vbNo Then Exit Function
          Else
            MsgBox "Diferenca de combustivel muito alta!"
            Exit Function
          End If
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

StrTemp = ReadINI("Fechamento", "NaoImportado", "", App.Path & "\Posto.ini")
If StrTemp = "" Then
  A = FreeFile()
  Open App.Path & "\Posto.ini" For Append As #A
  Print #A, ""
  Print #A, ""
  Print #A, "[Fechamento]"
  Print #A, ";0 - bloqueia o fechamento não importado"
  Print #A, ";1 - não bloqueia"
  Print #A, "NaoImportado=0"
  Print #A, "VendasCombustivel = 1100000000"
  Print #A, "VendasProdutos = 1200000000"
  Print #A, "Diferenca = 2100000000"
  Print #A, ""
  Close #A
End If

If StrTemp = "0" Then
  If PodeFechar = False Then
    Exit Function
  End If
End If

If IsNumeric(lblFaturamento.Caption) = False Then
  MsgBox "Faturamento inválido!"
  Exit Function
End If
If cboResponsavel.Text = "" Or cboResponsavel.Text = "0" Then
  MsgBox "Responsável inválido!"
  cboResponsavel.SetFocus
  Exit Function
End If
CodigoFechamento = dbFechamento.Recordset!CodigoFechamento

With dbTanquesEstoque
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!Entrada <> 0 Then
        MsgBox "Existe lançamento de nota pendente para confirmar! Confirme o lançamento e zere o campo de entrada!"
        Exit Function
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

If FechandoLote = False Then
  Resposta = MsgBox("Deseja finalizar o fechamento atual?", vbYesNo, "Fechamento de Caixa!")
  If Resposta = vbNo Then Exit Function
End If

Frame1.Enabled = False

With Animation1
  .Top = 4080
  .Left = 3720
  .Width = 2415
  .Height = 495
  .Visible = True
  .Open App.Path & "\Loading.avi"
  .Play
End With

With dbProdutos
  .RecordSource = "Select *from produtos where combustivel=-1"
  .Refresh
End With

With dbBicoEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If .Recordset!apurado = False Then
        Vendas = .Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno
        dbTanques.Refresh
        dbTanques.Recordset.FindFirst "tanque=" & .Recordset!Tanque
        If dbTanques.Recordset.NoMatch = False Then
          dbTanques.Recordset.Edit
          dbTanques.Recordset!Estoque = dbTanques.Recordset!Estoque - Vendas
          dbTanques.Recordset.Update
        Else
          MsgBox "Tanque '" & .Recordset!Tanque & "' não encontrado!"
        End If
        
        dbBicos.Recordset.FindFirst "bico=" & .Recordset("bico")
        If .Recordset.EOF = True Then
          MsgBox "O bico " & .Recordset("bico") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
        End If
        dbBicos.Recordset.Edit
        dbBicos.Recordset("ultimonumero") = .Recordset("encerrante")
        If dbBicos.Recordset!PrecoVenda <> .Recordset!Preco Then
          dbBicos.Recordset!PrecoVenda = .Recordset!Preco
        End If
        dbBicos.Recordset.Update
        
        dbProdutos.Refresh
        If dbProdutos.Recordset.RecordCount <> 0 Then
          dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
          If dbProdutos.Recordset.NoMatch = False Then
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
            PrecoMedio = dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque
            LucroMedio = (.Recordset!Preco - PrecoMedio) * Vendas
            dbProdutos.Recordset!LucroMedio = dbProdutos.Recordset!LucroMedio + LucroMedio
            ValorEstoque = PrecoMedio * Vendas
            dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque - ValorEstoque
            LucroVenda = (.Recordset!Preco - dbProdutos.Recordset!precocompra) * Vendas
            dbProdutos.Recordset!PrecoVenda = .Recordset!Preco
            dbProdutos.Recordset!TotalVendido = dbProdutos.Recordset!TotalVendido + .Recordset!ValorTotal
            dbProdutos.Recordset("estoque") = dbProdutos.Recordset("estoque") - Vendas
            dbProdutos.Recordset("acumulativo") = dbProdutos.Recordset("acumulativo") + Vendas
            dbProdutos.Recordset!LucroVenda = dbProdutos.Recordset!LucroVenda + LucroVenda
            dbProdutos.Recordset.Update
            .Recordset.Edit
            .Recordset!LucroMedio = LucroMedio
            .Recordset!PrecoMedio = PrecoMedio
            .Recordset.Update
          End If
        End If
        
        .Recordset.Edit
        .Recordset!apurado = True
        .Recordset.Update
        'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, .Recordset!CodigoProduto, .Recordset!Tanque, , Vendas
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With dbProdutos
  .RecordSource = "Select *from produtos where combustivel=0"
  .Refresh
End With
With dbVendas
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = False Then
        dbProdutos.Refresh
        dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
        If dbProdutos.Recordset.NoMatch = True Then
          MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
        Else
          LucroVenda = (.Recordset("valorunitario") * .Recordset("quantidade")) - (dbProdutos.Recordset("precocompra") * .Recordset("quantidade")) - .Recordset("ValorDesconto") - .Recordset!ValorComissao
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
          
          If dbProdutos.Recordset!ValorEstoque <> 0 And dbProdutos.Recordset!Estoque <> 0 Then
            PrecoMedio = dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque
          Else
            PrecoMedio = 0
          End If
          'LucroMedio = (.Recordset!valorUnitario - PrecoMedio) * .Recordset!Quantidade
          LucroMedio = (.Recordset("valorunitario") * .Recordset("quantidade")) - (PrecoMedio * .Recordset("quantidade")) - .Recordset("ValorDesconto") - .Recordset!ValorComissao
          dbProdutos.Recordset!LucroMedio = dbProdutos.Recordset!LucroMedio + LucroMedio
          ValorEstoque = PrecoMedio * .Recordset!Quantidade
          dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque - ValorEstoque
          
          
          EstoqueAnterior = dbProdutos.Recordset!Estoque
          dbProdutos.Recordset("estoque") = dbProdutos.Recordset("estoque") - .Recordset("quantidade")
          dbProdutos.Recordset!LucroVenda = dbProdutos.Recordset!LucroVenda + LucroVenda
          dbProdutos.Recordset!acumulativo = dbProdutos.Recordset!acumulativo + .Recordset("quantidade")
          If IsNull(dbProdutos.Recordset!TotalVendido) = True Then dbProdutos.Recordset!TotalVendido = 0
          dbProdutos.Recordset!TotalVendido = dbProdutos.Recordset!TotalVendido + .Recordset!ValorTotal
          dbProdutos.Recordset.Update
        End If
        With dbProdutosHistorico
          .Recordset.AddNew
          .Recordset!lancadoem = Now
          .Recordset!dataalteracao = Date
          .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
          .Recordset!Codigo = dbProdutos.Recordset!Codigo
          .Recordset!descriproduto = dbProdutos.Recordset!Descri
          .Recordset!descrioperacao = "Venda no Caixa: " & dbFechamento.Recordset!DataCaixa & " turno: " & dbFechamento.Recordset!Turno
          .Recordset!precocompra = dbProdutos.Recordset!precocompra
          .Recordset!PrecoVenda = dbProdutos.Recordset!PrecoVenda
          .Recordset!EstoqueAnterior = EstoqueAnterior
          .Recordset!Quantidade = dbVendas.Recordset!Quantidade
          .Recordset!estoquefinal = EstoqueAnterior - dbVendas.Recordset!Quantidade
          .Recordset.Update
        End With
  
        On Error Resume Next
        .Recordset.Edit
        If ComissaoAcumulativa = False Then
          With dbDespesasLanc2
            .Connect = Conectar
            .DatabaseName = Caminho
            .RecordSource = "select *from despesaslanc2 where descri='Comissões paga no caixa'"
            .Refresh
            .Refresh
            If .Recordset.RecordCount <> 0 Then
              .Recordset.MoveLast
              .Recordset.MoveFirst
              .Recordset.FindFirst "Descri='Comissões paga no caixa' and fechamento=0"
              If .Recordset.NoMatch = False Then
                .Recordset.Edit
              Else
                .Recordset.AddNew
                .Recordset!Valor = 0
              End If
            Else
              .Recordset.AddNew
              .Recordset!Valor = 0
            End If
            
            .Recordset!CodigoFechamento = -1
            .Recordset!Origem = "Despesa"
            .Recordset!Data = dbFechamento.Recordset!DataCaixa
            .Recordset!Hora = Now
            .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
            .Recordset!CodigoConta = 0
            .Recordset!CodigoDespesa = 0
            .Recordset!Descri = "Comissões paga no caixa"
            .Recordset!Obs = dbFechamento.Recordset!DataCaixa & " Turno " & dbFechamento.Recordset!Turno
            .Recordset!Valor = .Recordset!Valor - dbVendas.Recordset!ValorComissao
            .Recordset!valorpago = .Recordset!valorpago - dbVendas.Recordset!ValorComissao
            .Recordset!Fechamento = False
            .Recordset!compensado = True
            .Recordset!distribuido = True
            .Recordset!codigoenviar = "1"
            .Recordset!fechamentodiario = True
            .Recordset.Update
          End With
          .Recordset!Pago = True
        End If
        .Recordset!fechamentodiario = True
        .Recordset.Update
        On Error GoTo 0
      End If
      
      RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, , , dbVendas.Recordset!Quantidade
      
      .Recordset.MoveNext
    Loop
  End If
End With
If IsNumeric(lblEstacionaTotal.Caption) = True Then
  Estacionamento = CCur(lblEstacionaTotal.Caption)
Else
  Estacionamento = 0
End If


'*******************************************************************************************
'Registra diferença de estoque no estatus
'*******************************************************************************************
With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    StrTemp = dbProdutos.RecordSource
    dbProdutos.RecordSource = "Select *from produtos where combustivel=-1"
    dbProdutos.Refresh
    Do While .Recordset.EOF = False
      If .Recordset!apurado = False Then
        If .Recordset!Tanque <> 0 Then
          If .Recordset!Diferenca <> 0 Then
            dbTanques.Refresh
            dbTanques.Recordset.FindFirst "tanque=" & .Recordset!tanquenr
            dbProdutos.Refresh
            dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
            ValorEstoque = (dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque)
            ValorEstoque = ValorEstoque * .Recordset!Diferenca
            dbProdutos.Recordset.Edit
            dbProdutos.Recordset!DifEstoque = dbProdutos.Recordset!DifEstoque + .Recordset!Diferenca
            dbProdutos.Recordset!valordifestoque = dbProdutos.Recordset!valordifestoque + ValorEstoque
            dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque + .Recordset!Diferenca
            dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque + ValorEstoque
            dbProdutos.Recordset.Update
            dbTanques.Recordset.Edit
            dbTanques.Recordset!Estoque = .Recordset!Estoque + .Recordset!Diferenca
            dbTanques.Recordset.Update
            .Recordset.Edit
            .Recordset!ValorDiferenca = ValorEstoque
            .Recordset.Update
          Else
            .Recordset.Edit
            .Recordset!ValorDiferenca = 0
            .Recordset.Update
          End If
          'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, .Recordset!tanquenr, , , .Recordset!Diferenca
        End If
        .Recordset.Edit
        .Recordset!apurado = True
        .Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
    dbProdutos.RecordSource = StrTemp
    dbProdutos.Refresh
  End If
End With

With dbEstatus
  .Recordset.Edit
  If IsNull(.Recordset!Estacionamento) = True Then
    .Recordset!Estacionamento = 0
  End If
  .Recordset!Estacionamento = .Recordset!Estacionamento + Estacionamento
  .Recordset.Update
End With
With dbEstacionamento
  If IsNumeric(txtEstacionaFim.Text) = True Then
    .Recordset.Edit
    .Recordset!ultimonumero = txtEstacionaFim.Text
    .Recordset.Update
  End If
End With


With qProdutosAltera
  .RecordSource = "select *from qprodutosaltera where produtosaltera.codigoprodutoaltera=" & AlteraAnterior & " order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbProdutos.Refresh
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = False Then
        If .Recordset!PrecoVenda <> dbProdutos.Recordset!PrecoVenda Then
          dbProdutos.Recordset.Edit
          dbProdutos.Recordset!PrecoVenda = .Recordset!PrecoVenda
          dbProdutos.Recordset.Update
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With dbFechamento
  .Recordset.Edit
  If IsNumeric(lblTotalCombustivel.Caption) = True Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
  End If
  If IsNumeric(lblTotalProdutos.Caption) = True Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
  End If
  .Recordset!responsavel = cboResponsavel.Text
  .Recordset!fechado = True
  .Recordset!finalizadopor = Usuarios.Nome
  .Recordset!ComissaoAcumulativa = ComissaoAcumulativa
  .Recordset.Update
End With

'Load frmEstatus2
'Unload frmEstatus2
Dim Estatus As New frmEstatus2
Load Estatus
Unload Estatus

Animation1.Visible = False


Call cmdAbrir_Click
FecharCaixa = True
End Function

Private Function Desconfirmar() As Boolean
Dim Resposta As Integer, LucroVenda As Currency, StrTemp As String
Dim Estacionamento As Currency
Dim Vendas As Double
Dim db As New ADODB.Connection, dbTemp As New ADODB.Recordset

Desconfirmar = False
Call cmdAbrir_Click

Frame1.Enabled = False

With Animation1
  .Top = 4080
  .Left = 3720
  .Width = 2415
  .Height = 495
  .Visible = True
  .Open App.Path & "\Loading.avi"
  .Play
End With

With dbFechamento
  .Recordset.Edit
  If IsNumeric(lblTotalCombustivel.Caption) = True Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
  End If
  If IsNumeric(lblTotalProdutos.Caption) = True Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
  End If
  .Recordset!responsavel = cboResponsavel.Text
  .Recordset!fechado = False
  .Recordset.Update
End With
With dbProdutos
  .RecordSource = "Select *from produtos where combustivel=-1"
  .Refresh
End With


'*******************************************************************************************
'Registra diferença de estoque no estatus
'*******************************************************************************************
With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    StrTemp = dbProdutos.RecordSource
    dbProdutos.RecordSource = "Select *from produtos where combustivel=-1"
    dbProdutos.Refresh
    Do While .Recordset.EOF = False
      If .Recordset!Tanque <> 0 Then
        If .Recordset!Diferenca <> 0 Then
          dbTanques.Refresh
          dbTanques.Recordset.FindFirst "tanque=" & .Recordset!tanquenr
          dbProdutos.Refresh
          dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
          If IsNull(.Recordset!ValorDiferenca) = True Then
            ValorEstoque = (dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque) * .Recordset!Diferenca
          Else
            ValorEstoque = .Recordset!ValorDiferenca
          End If
          dbProdutos.Recordset.Edit
          dbProdutos.Recordset!DifEstoque = dbProdutos.Recordset!DifEstoque - .Recordset!Diferenca
          dbProdutos.Recordset!valordifestoque = dbProdutos.Recordset!valordifestoque - .Recordset!ValorDiferenca
          dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque - .Recordset!Diferenca
          dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque - ValorEstoque
          dbProdutos.Recordset.Update
          dbTanques.Recordset.Edit
          dbTanques.Recordset!Estoque = dbTanques.Recordset!Estoque - .Recordset!Diferenca
          dbTanques.Recordset.Update
          .Recordset.Edit
          .Recordset!apurado = False
          .Recordset.Update
        Else
          .Recordset.Edit
          .Recordset!apurado = False
          .Recordset.Update
        End If
        'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, .Recordset!tanquenr, , , -.Recordset!Diferenca
      End If
      .Recordset.MoveNext
    Loop
    dbProdutos.RecordSource = StrTemp
    dbProdutos.Refresh
  End If
End With


With dbBicoEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If .Recordset!apurado = True Then
        .Recordset.Edit
        .Recordset!apurado = False
        .Recordset.Update
        Vendas = .Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno
        dbTanques.Refresh
        dbTanques.Recordset.FindFirst "tanque=" & .Recordset!Tanque
        If dbTanques.Recordset.NoMatch = False Then
          dbTanques.Recordset.Edit
          dbTanques.Recordset!Estoque = dbTanques.Recordset!Estoque + Vendas
          dbTanques.Recordset.Update
        Else
          MsgBox "Tanque '" & .Recordset!Tanque & "' não encontrado!"
        End If
        
        dbBicos.Recordset.FindFirst "bico=" & .Recordset("bico")
        If .Recordset.EOF = True Then
          MsgBox "O bico " & .Recordset("bico") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
        End If
        dbBicos.Recordset.Edit
        dbBicos.Recordset("ultimonumero") = .Recordset("abertura")
        If dbBicos.Recordset!PrecoVenda <> .Recordset!Preco Then
          dbBicos.Recordset!PrecoVenda = .Recordset!Preco
        End If
        dbBicos.Recordset.Update
        
        dbProdutos.Refresh
        If dbProdutos.Recordset.RecordCount <> 0 Then
          dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
          If dbProdutos.Recordset.NoMatch = False Then
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
            If IsNull(.Recordset!LucroMedio) = True Then
              PrecoMedio = dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque
              LucroMedio = (.Recordset("preco") * .Recordset("vendas")) - (PrecoMedio * .Recordset("vendas"))
            Else
              LucroMedio = .Recordset!LucroMedio
              PrecoMedio = .Recordset!PrecoMedio
            End If
            ValorEstoque = PrecoMedio * Vendas
            dbProdutos.Recordset!LucroMedio = dbProdutos.Recordset!LucroMedio - LucroMedio
            dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque + ValorEstoque
            LucroVenda = (.Recordset!Preco - dbProdutos.Recordset!precocompra) * Vendas
            dbProdutos.Recordset!PrecoVenda = .Recordset!Preco
            dbProdutos.Recordset("estoque") = dbProdutos.Recordset("estoque") + Vendas
            dbProdutos.Recordset("acumulativo") = dbProdutos.Recordset("acumulativo") - Vendas
            dbProdutos.Recordset!TotalVendido = dbProdutos.Recordset!TotalVendido - .Recordset!ValorTotal
            dbProdutos.Recordset!LucroVenda = dbProdutos.Recordset!LucroVenda - LucroVenda
            dbProdutos.Recordset.Update
          End If
        End If
      End If
      'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, .Recordset!CodigoProduto, .Recordset!Tanque, , -Vendas
      
      .Recordset.MoveNext
    Loop
  End If
End With

With dbProdutos
  .RecordSource = "Select *from produtos where combustivel=0"
  .Refresh
End With
With dbVendas
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = True Then
        .Recordset.Edit
        .Recordset!fechamentodiario = False
        dbProdutos.Refresh
        dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
        If dbProdutos.Recordset.NoMatch = True Then
          MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
        Else
          LucroVenda = (.Recordset("valorunitario") * .Recordset("quantidade")) - (dbProdutos.Recordset("precocompra") * .Recordset("quantidade")) - .Recordset("ValorDesconto") - .Recordset!ValorComissao
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
          If dbProdutos.Recordset!ValorEstoque <> 0 And dbProdutos.Recordset!Estoque <> 0 Then
            PrecoMedio = dbProdutos.Recordset!ValorEstoque / dbProdutos.Recordset!Estoque
          Else
            PrecoMedio = 0
          End If
          LucroMedio = (.Recordset("valorunitario") * .Recordset("quantidade")) - (PrecoMedio * .Recordset("quantidade")) - .Recordset("ValorDesconto") - .Recordset!ValorComissao
          dbProdutos.Recordset!LucroMedio = dbProdutos.Recordset!LucroMedio - LucroMedio
          ValorEstoque = PrecoMedio * .Recordset!Quantidade
          dbProdutos.Recordset!ValorEstoque = dbProdutos.Recordset!ValorEstoque + ValorEstoque
          
          EstoqueAnterior = dbProdutos.Recordset!Estoque
          dbProdutos.Recordset("estoque") = dbProdutos.Recordset("estoque") + .Recordset("quantidade")
          dbProdutos.Recordset!LucroVenda = dbProdutos.Recordset!LucroVenda - LucroVenda
          dbProdutos.Recordset!acumulativo = dbProdutos.Recordset!acumulativo - .Recordset("quantidade")
          If IsNull(dbProdutos.Recordset!TotalVendido) = True Then dbProdutos.Recordset!TotalVendido = 0
          dbProdutos.Recordset!TotalVendido = dbProdutos.Recordset!TotalVendido - .Recordset!ValorTotal
          dbProdutos.Recordset.Update
        End If
        If ComissaoAcumulativa = False Then
          With dbDespesasLanc2
            .Connect = Conectar
            .DatabaseName = Caminho
            .RecordSource = "select *from despesaslanc2 where descri='Comissões paga no caixa'"
            .Refresh
            .Refresh
            If .Recordset.RecordCount <> 0 Then
              .Recordset.MoveLast
              .Recordset.MoveFirst
              .Recordset.FindFirst "Descri='Comissões paga no caixa' and fechamento=0"
              If .Recordset.NoMatch = False Then
                .Recordset.Edit
              Else
                .Recordset.AddNew
                .Recordset!Valor = 0
              End If
            Else
              .Recordset.AddNew
              .Recordset!Valor = 0
            End If
            .Recordset!CodigoFechamento = -1
            .Recordset!Origem = "Despesa"
            .Recordset!Data = dbFechamento.Recordset!DataCaixa
            .Recordset!Hora = Now
            .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
            .Recordset!CodigoConta = 0
            .Recordset!CodigoDespesa = 0
            .Recordset!Descri = "Comissões paga no caixa"
            .Recordset!Obs = dbFechamento.Recordset!DataCaixa & " Turno " & dbFechamento.Recordset!Turno
            .Recordset!Valor = .Recordset!Valor + dbVendas.Recordset!ValorComissao
            .Recordset!valorpago = .Recordset!valorpago + dbVendas.Recordset!ValorComissao
            .Recordset!Fechamento = False
            .Recordset!compensado = True
            .Recordset!fechamentodiario = True
            .Recordset!codigoenviar = "1"
            .Recordset.Update
          End With
          .Recordset!Pago = False
          .Recordset!fechamentodiario = False
          .Recordset.Update
        End If
        With dbProdutosHistorico
          .Recordset.AddNew
          .Recordset!lancadoem = Now
          .Recordset!dataalteracao = Date
          .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
          .Recordset!Codigo = dbProdutos.Recordset!Codigo
          .Recordset!descriproduto = dbProdutos.Recordset!Descri
          .Recordset!descrioperacao = "Estorno de Venda: " & dbFechamento.Recordset!DataCaixa & " turno: " & dbFechamento.Recordset!Turno
          .Recordset!precocompra = dbProdutos.Recordset!precocompra
          .Recordset!PrecoVenda = dbProdutos.Recordset!PrecoVenda
          .Recordset!EstoqueAnterior = EstoqueAnterior
          .Recordset!Quantidade = -dbVendas.Recordset!Quantidade
          .Recordset!estoquefinal = EstoqueAnterior + dbVendas.Recordset!Quantidade
          .Recordset.Update
        End With
  
        On Error Resume Next
        .Recordset.Update
        On Error GoTo 0
        
        RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, , , -dbVendas.Recordset!Quantidade
      End If
      .Recordset.MoveNext
    Loop
  End If
End With
If IsNumeric(lblEstacionaTotal.Caption) = True Then
  Estacionamento = CCur(lblEstacionaTotal.Caption)
Else
  Estacionamento = 0
End If


With dbEstatus
  .Recordset.Edit
  If IsNull(.Recordset!Estacionamento) = True Then
    .Recordset!Estacionamento = 0
  End If
  .Recordset!Estacionamento = .Recordset!Estacionamento - Estacionamento
  .Recordset.Update
End With
With dbEstacionamento
  If IsNumeric(txtEstacionaIni.Text) = True Then
    .Recordset.Edit
    .Recordset!ultimonumero = txtEstacionaIni.Text
    .Recordset.Update
  End If
End With


With qProdutosAltera
  .RecordSource = "select *from qprodutosaltera where produtosaltera.codigoprodutoaltera=" & AlteraAnterior & " order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = False Then
        If .Recordset!PrecoVenda <> dbProdutos.Recordset!PrecoVenda Then
          dbProdutos.Recordset.Edit
          dbProdutos.Recordset!PrecoVenda = .Recordset!PrecoVenda
          dbProdutos.Recordset.Update
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

db.Open CaminhoADO
dbTemp.Open "select *from produtosnotascorpo where codigocaixa=" & dbFechamento.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic

If dbTemp.RecordCount <> 0 Then
  Do While dbTemp.EOF = False
    With dbTanques
      .Recordset.FindFirst "tanque=" & dbTemp!Tanque
      .Recordset.Edit
      .Recordset!Estoque = .Recordset!Estoque - dbTemp!Quantidade
      .Recordset.Update
      dbTemp!Aguardando = True
      dbTemp.Update
    End With
    dbTemp.MoveNext
  Loop
End If
dbTemp.Close
db.Close

Animation1.Close
Animation1.Visible = False

Call cmdAbrir_Click
Desconfirmar = True
End Function

Public Function PodeFechar() As Boolean
Dim TempValor As Currency, Tolerancia As Double

PodeFechar = False

Tolerancia = 0.1


With dbResultado
  .Refresh
  If .Recordset.RecordCount = 0 Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  End If
  .Recordset.MoveFirst
  'encontra venda de combustiveis
  StrTemp = ReadINI("Fechamento", "VendasCombustivel", "1100000000", App.Path & "\Posto.ini")
  .Recordset.Find "codigoconta='" & StrTemp & "'"
  If .Recordset.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  Else
    If IsNumeric(lblTotalCombustivel.Caption) = True Then
      TempValor = CCur(lblTotalCombustivel.Caption) - .Recordset!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como venda de combustiveis " & Format(.Recordset!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          Else
            PodeFechar = True
            Exit Function
          End If
        Else
          MsgBox "Este caixa deveria ter como venda de combustiveis " & Format(.Recordset!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
  
  
  .Recordset.MoveFirst
  'encontra venda de Produtos
  StrTemp = ReadINI("Fechamento", "VendasProdutos", "1200000000", App.Path & "\Posto.ini")
  .Recordset.Find "codigoconta='" & StrTemp & "'"
  If .Recordset.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  Else
    If IsNumeric(lblTotalProdutos.Caption) = True Then
      TempValor = CCur(lblTotalProdutos.Caption) - .Recordset!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como venda de produtos " & Format(.Recordset!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          Else
            PodeFechar = True
            Exit Function
          End If
        Else
          MsgBox "Este caixa deveria ter como venda de produtos " & Format(.Recordset!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
  
  .Recordset.MoveFirst
  'encontra Diferença de Caixa
  StrTemp = ReadINI("Fechamento", "Diferenca", "2100000000", App.Path & "\Posto.ini")
  .Recordset.Find "codigoconta='" & StrTemp & "'"
  If .Recordset.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  Else
    If IsNumeric(lblDiferenca.Caption) = True Then
      TempValor = CCur(lblDiferenca.Caption) - .Recordset!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como diferença de caixa " & Format(.Recordset!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          Else
            PodeFechar = True
            Exit Function
          End If
        Else
          MsgBox "Este caixa deveria ter como diferença de caixa " & Format(.Recordset!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
End With
PodeFechar = True
End Function

Private Sub GravaResultado(ByVal StrTemp As String)
Dim db As New ADODB.Connection, CodigoConta As String, Valor As Currency

db.Open CaminhoADO
dbResultado.CursorLocation = adUseClient

'998|     2100000000|1,54

CodigoConta = Trim(Mid(StrTemp, 5, 15))
Valor = CCur(Trim(Mid(StrTemp, 21)))

dbResultado.Recordset.AddNew
dbResultado.Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
dbResultado.Recordset!CodigoConta = CodigoConta
dbResultado.Recordset!Valor = Valor
dbResultado.Recordset.Update

db.Close

End Sub



Private Sub PegaCupons()
  Dim strEncerrantes As String
  
  cmdImportar.Enabled = False
  Animation1.Visible = True
  Animation1.Open App.Path & "\engrenagem.avi"
  Animation1.Play
  
  strEncerrantes = "CuponsClientes.txt"
  
  If Inet1.StillExecuting = False Then
    With Inet1
      On Error Resume Next
      .Execute , "cd /download"
      Do While .StillExecuting = True
        DoEvents
      Loop
      If Dir(strEncerrantes) <> "" Then
        Kill strEncerrantes
      End If
      
      .Execute , "get " & strEncerrantes & " " & strEncerrantes
      Do While .StillExecuting = True
        DoEvents
      Loop
      If Dir(strEncerrantes) = "" Then
        MsgBox "Não foi possível fazer download do arquivo."
        cmdImportar.Enabled = True
        Animation1.Close
        Animation1.Visible = False
        Exit Sub
      End If
    End With
    
    intArquivo = FreeFile()
    
    Open strEncerrantes For Input As #intArquivo
    Do While EOF(intArquivo) = False
      Line Input #intArquivo, StrTemp
      GravaCupons StrTemp, dbVendasLeituraX
    Loop
    
    Close #intArquivo
    cmdImportar.Enabled = True
  Else
    MsgBox "Ainda está sendo executada a última importação, agurarde...!"
  End If

cmdImportar.Enabled = True
Animation1.Close
Animation1.Visible = False
End Sub

Private Function ApagaRegistros() As Boolean
Dim SoPrimeira As Boolean
Dim db As New ADODB.Connection

db.Open CaminhoADO
db.Execute "delete *from fechamentodecaixapista where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento


SoPrimeira = False
ApagaRegistros = False
If dbFechamento.Recordset!notaconferida Then
  SoPrimeira = True
End If
With dbClientesNotas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "confirmado=-1"
    If .Recordset.NoMatch = False Then
      SoPrimeira = True
    End If
  End If
End With
With dbFormaDePgRecebido
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "fechamentodiario=-1"
    If .Recordset.NoMatch = False Then
      SoPrimeira = True
    End If
  End If
End With
With dbDespesasLanc
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "fechamentodiario=-1"
    If .Recordset.NoMatch = False Then
      SoPrimeira = True
    End If
  End If
End With

If SoPrimeira = False Then
  ApagaRegistros = True
End If

With dbVendas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.RecordCount <> 0
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
If SoPrimeira = False Then
  With dbClientesNotas
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      Do While .Recordset.RecordCount <> 0
        With dbClientes
          .Recordset.FindFirst "codigocliente=" & dbClientesNotas.Recordset!CodigoCliente
          .Recordset.Edit
          .Recordset!TotalNotas = .Recordset!TotalNotas - dbClientesNotas.Recordset!ValorPrevisto
          .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
          .Recordset.Update
        End With
        .Recordset.Delete
        .Refresh
      Loop
    End If
  End With
  With dbFormaDePgRecebido
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      Do While .Recordset.RecordCount <> 0
        .Recordset.Delete
        .Refresh
      Loop
    End If
  End With
  With dbDespesasLanc
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      Do While .Recordset.RecordCount <> 0
        .Recordset.Delete
        .Refresh
      Loop
    End If
  End With
End If

End Function

Private Sub GravaDespesas(ByVal StrTemp As String)
Dim Codigo As String, Descri As String, Tipo As String, Valor As Currency
Codigo = Trim(Mid(StrTemp, 5, 15))
Descri = Trim(Mid(StrTemp, 21, 50))
Tipo = Trim(Mid(StrTemp, 72, 5))
Valor = CCur(Mid(StrTemp, 78))

If Tipo = "PAG" Then
  Valor = Valor * -1
End If
With dbDespesasTipo
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigonoposto='" & Codigo & "'"
  If .Recordset.NoMatch = True Then Exit Sub
End With
With dbDespesasLanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset!Origem = "Fechamento"
  .Recordset("data") = dbFechamento.Recordset!DataCaixa
  .Recordset!Vencimento = dbFechamento.Recordset!DataCaixa
  .Recordset("hora") = Now
  .Recordset("codigoconta") = -1
  .Recordset("conta") = "Fechamento de Caixa"
  .Recordset("codigodespesa") = dbDespesasTipo.Recordset("codigodespesa")
  .Recordset("descri") = dbDespesasTipo.Recordset("descri")
  .Recordset("obs") = Descri
  .Recordset!compensado = True
  .Recordset("valor") = Valor
  .Recordset!valorpago = Valor
  .Recordset.Update
  .Refresh
End With

End Sub

Private Sub GravaNumerarios(ByVal StrTemp As String)
Dim Codigo As String, Valor As Currency


Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency
Dim DescontoPorcento As Currency

Codigo = CDbl(Trim(Mid(StrTemp, 5, 15)))
Valor = CCur(Mid(StrTemp, 37))

With dbFormaDePg
  .RecordSource = "select *from formadepagamento"
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigonoposto='" & Trim(Codigo) & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  Tarifa = .Recordset("descontovalor")
  Operacao = .Recordset("descontoporOperacao")
  Porcento = .Recordset("descontoPorcento") / 100
End With


ValorBruto = Valor

If Porcento <> 0 Then
  DescontoPorcento = ValorBruto * Porcento
End If

Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao
With dbFormaDePg
  If .Recordset!CodigoConta = 0 Then
    MsgBox "A forma de pagamento " & .Recordset!Descri & " está sem conta destino!"
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
  .Recordset("data") = dbFechamento.Recordset!DataCaixa
  .Recordset("hora") = Now
  .Recordset.Update
  .Refresh
End With

End Sub

Private Sub GravaNotas(ByVal StrTemp As String)
Dim CodigoCliente As Double, Cupom As String, Placa As String
Dim Km As String, Veiculo As String, Qtd As Double, ValorTotal As Currency
Dim CodigoProduto As Double, valorUnitario As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim Autorizar As Boolean, Motivo As String, Autorizado As Boolean
Dim Ws As Workspace, db As Database, dbTemp As Recordset
Dim Preco As Currency

On Error GoTo 0
PrecoDif = False
With dbClientesNotas
  
  If Len(StrTemp) < 120 Then Exit Sub
  
  CodigoCliente = CDbl(Mid(StrTemp, 5, 12))
  Cupom = Mid(StrTemp, 18, 12)
  Placa = Mid(StrTemp, 31, 9)
  Km = Mid(StrTemp, 41, 15)
  Veiculo = Mid(StrTemp, 57, 25)
  Qtd = Mid(StrTemp, 83, 15)
  ValorTotalDif = Mid(StrTemp, 99, 15)
  If Len(StrTemp) > 179 Then
    ValorUnitarioDif = CCur(Mid(StrTemp, 179, 15))
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
      valorUnitario = Mid(StrTemp, 147, 15)
      ValorTotal = Mid(StrTemp, 163, 15)
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
  Autorizar = False
  Autorizado = False
  Motivo = ""
  
  If IsNumeric(Cupom) = False Then
    Cupom = 0
  End If
  dbClientes.Refresh
  dbClientes.Recordset.FindFirst "codigonoposto=" & CodigoCliente
  If dbClientes.Recordset.NoMatch = True Then
    MsgBox "Código de cliente de nota " & CodigoCliente & " não encontrado!"
    GravaBloqueado CodigoCliente, "Não encontrado", Cupom, ValorTotal, "Cliente não localizado"
    Exit Sub
  End If
  Set Ws = DBEngine.Workspaces(0)
  Set db = Ws.OpenDatabase(Caminho, , , Conectar)
  Set dbTemp = db.OpenRecordset("select *from produtos")
  dbTemp.FindFirst "codigo=" & CodigoProduto
  
  If dbClientes.Recordset!protestado = True Then
    MsgBox "Cliente bloqueado!"
    Autorizar = True
    Autorizado = True
    Motivo = "Bloqueado/Protestado"
  End If
  
  If dbTemp.NoMatch = True Then
    MsgBox "Código do produto " & CodigoProduto & " não cadastrado!"
    GravaBloqueado CodigoCliente, "Código de produto não encontrado", Cupom, ValorTotal, "Cliente não localizado"
    Exit Sub
  Else
    If dbTemp!Combustivel = True Then
      dbBicoEncerrantes.Recordset.FindFirst "codigoproduto=" & dbTemp!CodigoProduto
      Preco = PrecoAtual(dbTemp!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbBicoEncerrantes.Recordset!Bico)
    Else
      Preco = PrecoAtual(dbTemp!CodigoProduto, dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno)
    End If
  End If
  If dbClientes.Recordset!mensalista = False Then
    If dbClientes.Recordset!desativado < dbFechamento.Recordset!DataCaixa Then
      If Usuarios.Grupo.admDatas < 2 Then
        MsgBox "O cliente " & dbClientes.Recordset!Nome & " está desativado!"
        If Configura.NotaBloqueia = 0 Then
          GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Cliente Desativado"
          Autorizar = True
          Motivo = "Desativado"
        End If
      Else
        Resposta = MsgBox("O cliente " & dbClientes.Recordset!Nome & " está desativado! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
        GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Cliente Desativado"
        If Resposta = vbNo Then Exit Sub
        If Configura.NotaBloqueia = 0 Then
          Autorizar = True
          Autorizado = True
          Motivo = "Desativado"
        End If
      End If
    End If
  End If
  If dbClientes.Recordset!limitar = True Then
    If IsNull(dbClientes.Recordset!Limite) = False Then
      Limite = CCur(ValorTotal)
      Set dbTemp = db.OpenRecordset("select sum(valorprevisto) as total from clientesnota2 where codigocliente=" & dbClientes.Recordset!CodigoCliente & " and confirmado=0")
      If IsNull(dbTemp!Total) = False Then
        Limite = Limite + dbTemp!Total
      End If
      Set dbTemp = db.OpenRecordset("select sum(valor) as total from clientescobranca where codigocliente=" & dbClientes.Recordset!CodigoCliente & " and pago=0")
      If IsNull(dbTemp!Total) = False Then
        Limite = Limite + dbTemp!Total
      End If
      If Limite > dbClientes.Recordset!Limite Then
        If Usuarios.Grupo.admDatas < 2 Then
          MsgBox "O cliente " & dbClientes.Recordset!Nome & " ultrapassará o limite dele! Somente o administrador pode lançar."
          GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
          Autorizar = True
          Motivo = "Limite"
        Else
          Resposta = MsgBox("O cliente " & dbClientes.Recordset!Nome & " ultrapassará o limite dele! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
          GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
          If Resposta = vbNo Then Exit Sub
          Autorizar = True
          Autorizado = True
          Motivo = "Ultrapassou Limite"
        End If
      End If
    Else
      MsgBox "O cliente " & dbClientes.Recordset!Nome & " esta marcado para ser limitado mas não possue valor definido!"
      GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Marcado para limitar mas não possue valor a ser limitado"
      Autorizar = True
      Motivo = "Sem Limite"
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
  With dbClientesProdutos
    If .Recordset.RecordCount <> 0 Then
      .Recordset.FindFirst "codigocliente=" & dbClientes.Recordset!CodigoCliente & " and codproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(txtData.Value) & "#"
      If .Recordset.NoMatch = False Then
        If .Recordset!validade = txtData.Value Then
          If .Recordset!HoraIni >= dbTurno.Recordset!HoraIni Then
            PrecoDif = True
          End If
        Else
          PrecoDif = True
        End If
      End If
      If PrecoDif = True Then
        If .Recordset!Preco <> 0 Then
          TempValorPagar = Qtd * .Recordset!Preco
        Else
          TempValorPagar = Qtd * valorUnitario
          If .Recordset!Porcento <> 0 Then
            TempValorPagar = TempValorPagar * .Recordset!Porcento
          End If
        End If
        If .Recordset!valorasomar <> 0 Then
          TempValorPagar = TempValorPagar + (Qtd * .Recordset!valorasomar)
        End If
        TempDif = TempValorPagar - ValorTotal
        If TempDif > 0.2 Or TempDif < -0.2 Then
          If Usuarios.Grupo.admDatas < 2 Then
            MsgBox "O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
            GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
            Autorizar = True
            Motivo = "Preço Diferenciado"
          Else
            Resposta = MsgBox("O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
            GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
            If Resposta = vbNo Then Exit Sub
            Autorizar = True
            Autorizado = True
            Motivo = "Preço Diferenciado"
          End If
        End If
      Else
        'ValorUnitarioDif = Qtd * valorUnitario
        TempDif = (ValorUnitarioDif * Qtd) - ValorTotalDif
        If TempDif > 0.01 Or TempDif < -0.01 Then
          MsgBox "Preço unitário incorreto!"
          Exit Sub
        End If
      End If
    Else
      TempDif = Preco - (ValorTotal / Qtd)
      If TempDif > 0.2 Or TempDif < -0.02 Then
        If Usuarios.Grupo.admDatas < 2 Then
          MsgBox "O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Somente o administrador pode lançar."
          GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
          Autorizar = True
          Motivo = "Preço Diferenciado"
        Else
          Resposta = MsgBox("O cliente " & dbClientes.Recordset!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
          GravaBloqueado dbClientes.Recordset!CodigoCliente, dbClientes.Recordset!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
          If Resposta = vbNo Then Exit Sub
          Autorizar = True
          Autorizado = True
          Motivo = "Preço incorreto!"
        End If
      End If
    End If
  End With
  A = Fix(valorUnitario)
  If Qtd = 0 Then
    Qtd = ValorTotal / ValorUnitarioDif
  End If
  .Recordset.AddNew
  .Recordset("codigofechamento") = dbFechamento.Recordset!CodigoFechamento
  .Recordset("codigocliente") = dbClientes.Recordset!CodigoCliente
  .Recordset("nome") = dbClientes.Recordset!Nome
  .Recordset("datalanc") = Now
  .Recordset("dataprevista") = DataPrevista
  .Recordset("valorprevisto") = ValorTotal
  .Recordset!Data = dbFechamento.Recordset!DataCaixa
  If Trim(Cupom) <> "" Then
    .Recordset!Cupom = Cupom
  End If
  If Trim(Km) = "" Then Km = 0
  .Recordset!Km = Km
  .Recordset!Placa = Placa
  On Error Resume Next
  If dbClientesCarros.Recordset.EOF = False And dbClientesCarros.Recordset.BOF = False Then
    .Recordset!codigocarro = dbClientesCarros.Recordset!codigocarro
  End If
  On Error GoTo 0
  .Recordset!Litros = Qtd
  .Recordset!Consumo = Consumo
  .Recordset!CodigoProduto = CodigoProduto
  .Recordset!valorUnitario = valorUnitario
  .Recordset!Qtd = Qtd
  .Recordset!ValorUnitarioDif = ValorUnitarioDif
  .Recordset!ValorTotalDif = ValorTotalDif
  .Recordset!LucroDif = LucroDif
  .Recordset!Autorizar = Autorizar
  .Recordset!Autorizado = Autorizado
  .Recordset!Motivo = Motivo
  .Recordset.Update
End With

With dbClientes
  .Recordset.Edit
  If IsNull(.Recordset!UltimoAbastecimento) = True Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  If .Recordset!UltimoAbastecimento < dbFechamento.Recordset!DataCaixa Then
    .Recordset!UltimoAbastecimento = dbFechamento.Recordset!DataCaixa
  End If
  .Recordset!TotalNotas = .Recordset!TotalNotas + ValorTotal
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With

End Sub

Private Sub GravaTanque(ByVal StrTemp As String)
Dim Tanque As Integer, Estoque As Double
Tanque = Mid(StrTemp, 5, 5)
StrTemp2 = Mid(StrTemp, 11)
For i = 1 To Len(StrTemp2)
  If Mid(StrTemp2, i, 1) <> 0 Then
    StrTemp2 = Mid(StrTemp2, i)
    Exit For
  End If
Next i
Estoque = CDbl(StrTemp2)

With dbTanquesEstoque
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "tanque=" & Tanque
    If .Recordset.NoMatch = False Then
      .Recordset.Edit
      .Recordset!Estoque = Estoque
      .Recordset.Update
    End If
  End If
End With
End Sub

Private Sub GravaBico(ByVal StrTemp As String)
Dim Bico As Integer, Encerrante As Double, Encontrou As Boolean, Abertura As Double
Bico = CInt(Mid(StrTemp, 5, 6))
Encerrante = CDbl(Mid(StrTemp, 29, 16))
Abertura = CDbl(Mid(StrTemp, 12, 16))
If Encerrante > 1000000 Then
  Do While Encerrante > 1000000
    Encerrante = Encerrante - 1000000
  Loop
End If
If Abertura > 1000000 Then
  Do While Abertura > 1000000
    Abertura = Abertura - 1000000
  Loop
End If
With dbBicoEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "bico=" & Bico
    If .Recordset.NoMatch = True Then
      MsgBox "Bico " & Bico & " cadastrado no posto mas não localizado no sistema."
      Encontrou = False
    Else
      Encontrou = True
      .Recordset.Edit
      '.Recordset!Abertura = Abertura
      If .Recordset!Abertura > 1000000 Then .Recordset!Abertura = Abertura
      .Recordset!Encerrante = Encerrante
      .Recordset.Update
    End If
  End If
End With

End Sub

Private Sub GravaVenda(ByVal StrTemp As String)
Dim Bico As Integer, Preco As Currency, Codigo As Double, Qtd As Double, Funcionario As Integer

If Trim(Mid(StrTemp, 18, 6)) <> "" Then
  Bico = CInt(Mid(StrTemp, 18, 6))
Else
  Bico = 0
End If
Preco = CCur(Mid(StrTemp, 38, 12))
StrTemp2 = Mid(StrTemp, 5, 12)
If IsNumeric(StrTemp2) = False Then
  StrTemp2 = RemoveString(StrTemp2)
End If
Codigo = CDbl(StrTemp2)
Qtd = CDbl(Mid(StrTemp, 25, 12))
If Trim(Mid(StrTemp, 64)) <> "" Then
  Funcionario = CInt(Mid(StrTemp, 64))
Else
  Funcionario = 0
End If

If Bico <> 0 Then
  With dbBicoEncerrantes
    .Recordset.FindFirst "bico=" & Bico
    If .Recordset!Preco <> Preco Then
      MsgBox "O preço da bomba " & Bico & " está cadastrado " & Format(.Recordset!Preco, "#,##0.000") & " mas no posto está " & Format(Preco, "#,##0.000")
      '.Recordset.Edit
      '.Recordset!Preco = Preco
      '.Recordset.Update
    End If
  End With
Else
  txtCodProduto.Text = Codigo
  Call txtCodProduto_LostFocus
  txtQtd.Text = Qtd
  Call txtQtd_LostFocus
  If Funcionario <> 0 Then
    txtCodFunc.Text = Funcionario
  Else
    txtCodFunc.Text = ""
  End If
  If dbProdutos.Recordset!Codigo <> Codigo Then
    MsgBox "O produto " & Codigo & " não está cadastrado!"
  Else
    If qProdutosAltera.Recordset.RecordCount <> 0 Then
      qProdutosAltera.Recordset.FindFirst "codigo=" & txtCodProduto.Text
      If qProdutosAltera.Recordset.NoMatch = False Then
        If qProdutosAltera.Recordset!PrecoVenda <> Preco Then
          MsgBox "O produto " & Codigo & " está cadastrado " & Format(qProdutosAltera.Recordset!PrecoVenda, "#,##0.000") & " mas no posto está " & Format(Preco, "#,##0.000")
        End If
      Else
        MsgBox "O produto " & Codigo & " não foi encontrado na tabela de preços atual!"
      End If
    End If
    Call cmdIncluir_Click
  End If
  
End If

End Sub

Private Sub TotalVenda()
Dim Unitario As Currency, Qtd As Double, Desconto As Currency, Total As Currency
lblTotalVenda.Caption = ""
With dbProdutos
  .Refresh
  If IsNumeric(txtQtd.Text) = False Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = True Then Exit Sub
  If .Recordset!Codigo <> txtCodProduto.Text Then Exit Sub
  Unitario = .Recordset!PrecoVenda
  If qProdutosAltera.Recordset.RecordCount <> 0 Then
    qProdutosAltera.Recordset.FindFirst "codigo=" & txtCodProduto.Text
    If qProdutosAltera.Recordset.NoMatch = False Then
      Unitario = qProdutosAltera.Recordset!PrecoVenda
    End If
  End If
End With
Qtd = CDbl(txtQtd.Text)
Total = (Qtd * Unitario) - Desconto
lblTotalVenda.Caption = Format(Total, "Currency")

End Sub

Private Sub cboResponsavel_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboResponsavel_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case vbKeyF5
    KeyCode = 0
    If Shift = 1 Then
      PegaCupons
    Else
      If cmdImportar.Visible = True Then
        Call cmdImportar_Click
      End If
    End If
  Case vbKeyF1
    KeyCode = 0
    Call cmdPrimeiro_Click
  Case vbKeyF2
    KeyCode = 0
    Call cmdAnterior_Click
  Case vbKeyF3
    KeyCode = 0
    Call cmdPosterior_Click
  Case vbKeyF4
    KeyCode = 0
    Call cmdUltimo_Click
End Select
End Sub

Private Sub cboResponsavel_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cboTurno_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboTurno_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
End Select
End Sub

Private Sub cboTurno_LostFocus()
Me.KeyPreview = True
With dbTurno
  .Refresh
  If cboTurno.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboTurno.Text & "'"
  If .Recordset.NoMatch = False Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdAnterior_Click()
With dbFechamento
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.EOF = False And .Recordset.BOF = False Then
      .Recordset.MovePrevious
      If .Recordset.BOF = False Then
        txtData.Value = .Recordset!DataCaixa
        cboTurno = .Recordset!Turno
        Call cmdAbrir_Click
      End If
    End If
  End If
End With
End Sub

Private Sub cmdCalcular_Click()
Dim MedeAntes As Boolean, TanqueNegativo As Boolean
Dim Combustivel As Currency, Produtos As Currency, Comissoes As Currency
Dim TempValor As Double, TempValor2 As Double, VendaTanque As Double
Dim TempValor3 As Double, B As Double

'Verifica se o último número do bico está correto
TanqueNegativo = False

With dbUltimoFechamento
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    dbUltimoEncerrante.RecordSource = "select *from BicoEncerrantes where codigofechamento=" & .Recordset!CodigoFechamento & " order by bico"
    If dbUltimoEncerrante.Recordset.RecordCount <> 0 Then
      dbBicos.Refresh
      If dbBicos.Recordset.RecordCount <> 0 Then
        dbBicos.Recordset.MoveLast
        dbBicos.Recordset.MoveFirst
        Do While dbBicos.Recordset.EOF = False
          If dbBicos.Recordset!ultimonumero <> dbUltimoEncerrante.Recordset!Encerrante Then
            dbBicos.Recordset.Edit
            dbBicos.Recordset!ultimonumero = dbUltimoEncerrante.Recordset!Encerrante
            dbBicos.Recordset.Update
          End If
          dbBicos.Recordset.MoveNext
        Loop
      End If
    End If
  End If
End With

With dbBicoEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If Frame1.Enabled = True Then
        With dbEncerrantesNovos
          .Refresh
          .Recordset.Filter = "datacaixa=#" & dbFechamento.Recordset!DataCaixa & "# and horaini=#" & dbFechamento.Recordset!HoraIni & "# and bico=" & dbBicoEncerrantes.Recordset!Bico
          If .Recordset.RecordCount <> 0 Then
            dbBicoEncerrantes.Recordset.Edit
            dbBicoEncerrantes.Recordset!Abertura = .Recordset!inicial
            dbBicoEncerrantes.Recordset.Update
          End If
        End With
        If .Recordset!Encerrante < .Recordset!Abertura Then
          If .Recordset!Abertura < 1000000 Then
            If .Recordset!Encerrante <> 0 Then
              MsgBox "Encerrante inválido no bico " & .Recordset!Bico
            End If
            .Recordset.Edit
            .Recordset!Encerrante = .Recordset!Abertura
            .Recordset.Update
          End If
        End If
      End If
      If dbFechamento.Recordset!fechado = False Then
        If .Recordset!Encerrante <> 0 Then
          .Recordset.Edit
          If .Recordset!Abertura > 1000000 Then
              Do While .Recordset!Abertura > 1000000
                .Recordset!Abertura = .Recordset!Abertura - 1000000
              Loop
          End If
          'If .Recordset!Encerrante > 1000000 Then .Recordset!Encerrante = .Recordset!Encerrante - 1000000
'          If .Recordset!Encerrante > 1000000 Then
'            Do While .Recordset!Encerrante > 1000000
'              .Recordset!Encerrante = .Recordset!Encerrante - 1000000
'            Loop
'          End If
          .Recordset!Vendas = .Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno
          .Recordset!ValorTotal = (.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno) * .Recordset!Preco
          .Recordset.Update
        Else
          .Recordset.Edit
          If .Recordset!Abertura > 1000000 Then .Recordset!Abertura = .Recordset!Abertura - 1000000
          .Recordset!Encerrante = .Recordset!Abertura
          .Recordset!Vendas = 0
          .Recordset!ValorTotal = 0
          .Recordset.Update
        End If
      End If
      On Error Resume Next
      Combustivel = Combustivel + .Recordset!ValorTotal
      On Error GoTo 0
      .Recordset.MoveNext
    Loop
  End If
End With
With dbVendas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Produtos = Produtos + .Recordset!ValorTotal
      Comissoes = Comissoes + .Recordset!ValorComissao
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With
If Frame1.Enabled = True Then
  With dbProdutos
    .RecordSource = "Select *from produtos where combustivel=-1"
    .Refresh
  End With
  If Abrindo = True Then
    With dbTanques
      StrTemp = .RecordSource
      .RecordSource = "Select codigoproduto, sum(estoque) as total from tanques group by codigoproduto"
      .Refresh
      If IsNull(.Recordset!Total) = False Then
        If dbProdutos.Recordset.EOF = False Then
          dbProdutos.Recordset.MoveLast
          dbProdutos.Recordset.MoveFirst
          Do While dbProdutos.Recordset.EOF = False
            .Recordset.FindFirst "codigoproduto=" & dbProdutos.Recordset!CodigoProduto
            If .Recordset.NoMatch = False Then
              B = dbProdutos.Recordset!Estoque - .Recordset!Total
              If B > 1 Or B < -1 Then
                dbNotasCorpo.RecordSource = "select sum(quantidade) as total2 from produtosnotascorpo where aguardando=-1 and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
                dbNotasCorpo.Refresh
                If IsNull(dbNotasCorpo.Recordset!Total2) = False Then
                  B = CLng(dbProdutos.Recordset!Estoque - dbNotasCorpo.Recordset!Total2) - CLng(.Recordset!Total)
                  If B > 1 Or B < -1 Then
                    'MsgBox "Erro na soma dos tanques de " & dbProdutos.Recordset!Descri & "!"
                    CorrigeTanque dbProdutos.Recordset!CodigoProduto
                  End If
                Else
                  'MsgBox "Erro na soma dos tanques de " & dbProdutos.Recordset!Descri & "!"
                  CorrigeTanque dbProdutos.Recordset!CodigoProduto
                End If
              End If
            End If
            dbProdutos.Recordset.MoveNext
          Loop
          dbProdutos.Recordset.MoveFirst
        End If
      End If
      .RecordSource = StrTemp
      .Refresh
    End With
  End If
  With dbTanquesEstoque
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbTanques.Refresh
        dbTanques.Recordset.FindFirst "tanque=" & .Recordset!Tanque
        If dbTanques.Recordset.NoMatch = False Then
          TempValor = 0
          TempValor2 = 0
          TempValor3 = 0
          VendaTanque = 0
          MedeAntes = dbPostos.Recordset!MedetanqueAntes
          If MedeAntes = True Then
            If IsNull(dbFechamento.Recordset!Sequencia) = False Then
              dbPostos.RecordSource = "select sum(vendas) as total from qbicoencerrantes where tanque=" & .Recordset!Tanque & " and sequencia<" & dbFechamento.Recordset!Sequencia & " and fechado=0"
            Else
              dbPostos.RecordSource = "select sum(vendas) as total from qbicoencerrantes where tanque=" & .Recordset!Tanque & " and sequencia<0 and fechado=0"
            End If
            dbPostos.Refresh
            If IsNull(dbPostos.Recordset!Total) = False Then
              TempValor2 = dbPostos.Recordset!Total
            Else
              TempValor2 = 0
            End If
          Else
            TempValor2 = 0
            If IsNull(dbFechamento.Recordset!Sequencia) = False Then
              dbPostos.RecordSource = "select sum(vendas) as total from qbicoencerrantes where tanque=" & .Recordset!Tanque & " and sequencia<=" & dbFechamento.Recordset!Sequencia & " and fechado=0"
              dbPostos.Refresh
              If IsNull(dbPostos.Recordset!Total) = False Then
                TempValor2 = dbPostos.Recordset!Total
              Else
                TempValor2 = 0
              End If
            Else
              AtualizaSequenciaCaixa
            End If
          End If
          TempValor3 = 0
          If IsNull(dbFechamento.Recordset!Sequencia) = False Then
            dbPostos.RecordSource = "select sum(vendas) as total from qbicoencerrantes where tanque=" & .Recordset!Tanque & " and sequencia=" & dbFechamento.Recordset!Sequencia
            dbPostos.Refresh
            If IsNull(dbPostos.Recordset!Total) = False Then
              TempValor3 = dbPostos.Recordset!Total
            Else
              TempValor3 = 0
            End If
          End If
          dbPostos.RecordSource = "select *from postos"
          dbPostos.Refresh
        End If
        dbProdutos.Refresh
        dbProdutos.Recordset.FindFirst "codigoproduto=" & dbTanques.Recordset!CodigoProduto
        TempValor = .Recordset!Estoque - (dbTanques.Recordset!Estoque + .Recordset!Entrada - TempValor2)
        dbDifComb.Refresh
        If dbDifComb.Recordset.RecordCount > dbTanques.Recordset.RecordCount Then
          Do While dbDifComb.Recordset.RecordCount <> 0
            dbDifComb.Recordset.Delete
            dbDifComb.Refresh
          Loop
        End If
        If dbDifComb.Recordset.RecordCount = 0 Then
          dbDifComb.Recordset.AddNew
        Else
          dbDifComb.Recordset.FindFirst "tanquenr=" & dbTanques.Recordset!Tanque
          If dbDifComb.Recordset.NoMatch = False Then
            dbDifComb.Recordset.Edit
          Else
            dbDifComb.Recordset.AddNew
          End If
        End If
        If MedeAntes = True Then
          If dbDifComb.Recordset!Tanque - VendaTanque < 0 Then
            MsgBox "O tanque " & dbDifComb.Recordset!tanquenr & " vai ficar com o estoque negativo!"
            TanqueNegativo = True
          End If
        End If
        
        dbDifComb.Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
        dbDifComb.Recordset!CodigoProduto = dbTanques.Recordset!CodigoProduto
        dbProdutos.Recordset.FindFirst "codigoproduto=" & dbTanques.Recordset!CodigoProduto
        dbDifComb.Recordset!Descri = dbProdutos.Recordset!Descri
        dbDifComb.Recordset!Estoque = dbTanques.Recordset!Estoque - TempValor2
        dbDifComb.Recordset!Tanque = dbTanquesEstoque.Recordset!Estoque
        dbDifComb.Recordset!Diferenca = TempValor
        dbDifComb.Recordset!Vendido = TempValor3
        dbDifComb.Recordset!tanquenr = dbTanques.Recordset!Tanque
        dbDifComb.Recordset.Update
        
        .Recordset.MoveNext
      Loop
    End If
  End With
  With dbProdutos
    .RecordSource = "Select *from produtos where combustivel=0"
    .Refresh
  End With
End If
With dbEstacionamentoCaixa
  .Refresh
  'If dbFechamento.Recordset!fechado = False Then
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Edit
      .Recordset!totalun = .Recordset!final - .Recordset!inicial - .Recordset!cancelados
      If .Recordset!totalun < 0 Then
        MsgBox "Erro no controle de estacionamento!"
        .Recordset!final = .Recordset!inicial
        On Error Resume Next
        txtEstacionaFim.SetFocus
        Exit Sub
      End If
      .Recordset!Total = .Recordset!Preco * .Recordset!totalun
      .Recordset.Update
      lblEstacionaTotal.Caption = Format(.Recordset!Total, "Currency")
      Produtos = Produtos + .Recordset!Total
    End If
  'End If
End With

dbDifComb.RecordSource = "select *from DiferencaCombustivel where codigofechamento=" & dbFechamento.Recordset!CodigoFechamento
dbDifComb.Refresh

lblComissoes.Caption = Format(Comissoes, "Currency")
lblTotalCombustivel.Caption = Format(Combustivel, "Currency")
lblTotalProdutos.Caption = Format(Produtos, "Currency")
If ComissaoAcumulativa = True Then
  lblFaturamento.Caption = Format(Combustivel + Produtos, "Currency")
Else
  lblFaturamento.Caption = Format(Combustivel + Produtos - Comissoes, "Currency")
End If
If txtInformado.Text <> "" Then
  If IsNumeric(txtInformado.Text) = False Then
    txtInformado.Text = Format(0, "Currency")
  End If
Else
  txtInformado.Text = Format(0, "Currency")
End If
lblDiferenca.Caption = Format(CCur(txtInformado.Text) - CCur(lblFaturamento.Caption), "Currency")
If TanqueNegativo = False Then
  'ErroNaSoma = False
End If

With dbFechamento
  If .Recordset!fechado = False Then
    On Error Resume Next
    .Recordset.Edit
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
    .Recordset.Update
  End If
End With
End Sub

Private Sub cmdCancelar_Click()

With dbFechamento
  If .Recordset.BOF = False And .Recordset.EOF = False Then
    On Error Resume Next
    .Recordset.Edit
    .Recordset!responsavel = cboResponsavel.Text
    .Recordset.Update
    On Error GoTo 0
  End If
End With

Frame1.Visible = False
cboResponsavel.Visible = False
DBGrid1.Visible = False
DBGrid2.Visible = False
DBGrid3.Visible = False
DBGrid4.Visible = False

txtData.Enabled = True
cboTurno.Enabled = True
cboPdvs.Enabled = True
cmdAbrir.Enabled = True
cmdSair.Enabled = True
cmdCancelar.Visible = True
cmdSair.Cancel = True
cmdRemover.Enabled = False
cboPdvs.SetFocus
End Sub

Private Sub cmdConfirmar_Click()
Dim CodigoFechamento As Double
Dim db As New ADODB.Connection
Dim dbFechados As New ADODB.Recordset
Dim Estatus As New frmEstatus2

With dbFechamento
  CodigoFechamento = .Recordset!CodigoFechamento
  .Recordset.MovePrevious
  If .Recordset.BOF = False Then
    If .Recordset!fechado = False Then
      Resposta = MsgBox("Existe fechamento anterior para ser confirmado! Deseja confirmar em lote?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      FechandoLote = True
      db.Open CaminhoADO
      dbFechados.CursorLocation = adUseClient
      dbFechados.Open "select codigofechamento, datacaixa, turno, horaini, fechado from fechamentodecaixa order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
      If dbFechados.RecordCount = 0 Then Exit Sub
      dbFechados.Find "fechado=0"
      TempFechamento = CodigoFechamento
      If dbFechados.EOF = True Then Exit Sub
      Do While TempFechamento <> dbFechados!CodigoFechamento
        txtData.Value = dbFechados!DataCaixa
        cboTurno.Text = dbFechados!Turno
        Call cmdAbrir_Click
        Me.Refresh
        If FecharCaixa = False Then
          FechandoLote = False
          Exit Sub
        End If
        Call cmdCancelar_Click
        Load Estatus
        Unload Estatus
        dbFechados.MoveNext
      Loop
      FechandoLote = False
      Exit Sub
    End If
  End If
  .Refresh
  .Recordset.FindFirst "codigofechamento=" & CodigoFechamento
  FecharCaixa
End With


End Sub

Private Sub cmdDesconfirmar_Click()
Dim Resposta As Integer, TempFechamento As Double
Dim db As New ADODB.Connection
Dim dbFechados As New ADODB.Recordset
Dim Estatus As New frmEstatus2


CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
TempFechamento = dbFechamento.Recordset!CodigoFechamento

With dbFechamento
  .Recordset.MoveNext
  If .Recordset.EOF = False Then
    If .Recordset!fechames = True Then
      MsgBox "Este caixa pertence a mês já fechado!"
      Exit Sub
    End If
    If .Recordset!fechado = True Then
      MsgBox "Existe fechamento posterior confirmado!"
      Resposta = MsgBox("Deseja cancelar o fechamento em lote?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      Call cmdCancelar_Click
      db.Open CaminhoADO
      dbFechados.CursorLocation = adUseClient
      dbFechados.Open "select codigofechamento, datacaixa, turno, horaini, fechado from fechamentodecaixa order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic
      If dbFechados.RecordCount = 0 Then Exit Sub
      dbFechados.Find "fechado=-1"
      If dbFechados.EOF = True Then Exit Sub
      Do While TempFechamento <> dbFechados!CodigoFechamento
        txtData.Value = dbFechados!DataCaixa
        cboTurno.Text = dbFechados!Turno
        Call cmdAbrir_Click
        Me.Refresh
        If Desconfirmar() = False Then Exit Sub
        Call cmdCancelar_Click
        Load Estatus
        Unload Estatus
        dbFechados.MoveNext
      Loop
      'If Desconfirmar() = False Then Exit Sub
      Exit Sub
    End If
  End If
  .Refresh
  .Recordset.FindFirst "codigofechamento=" & CodigoFechamento
  If .Recordset!fechado = False Then
    MsgBox "Este caixa não está finalizado!"
    Exit Sub
  End If
End With


Resposta = MsgBox("Deseja retornar o fechamento atual?", vbYesNo, "Fechamento de Caixa!")
If Resposta = vbNo Then Exit Sub

Desconfirmar


Load Estatus
Unload Estatus
End Sub

Private Sub cmdEntraCombustivel_Click()
Load frmFechamentoConfirmaEntrada
frmFechamentoConfirmaEntrada.CodigoCaixa = dbFechamento.Recordset!CodigoFechamento
frmFechamentoConfirmaEntrada.Show vbModal
Call cmdCalcular_Click
End Sub

Private Sub cmdImportar_Click()
Dim Dia As Date, strEncerrantes As String, intArquivo As Integer
Dim StrTemp As String, SoPrimeira As Boolean

cmdImportar.Enabled = False
With Animation1
  .Top = 4080
  .Left = 3720
  .Width = 2415
  .Height = 495
  .Visible = True
  .Open App.Path & "\engrenagem.avi"
  .Play
End With

With dbConfig
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from cuponsfiscais where datacupom=#" & DateAdd("d", -1, Date) & "#"
  .Refresh
End With
With dbImportacao
  .ConnectionString = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig.Recordset!ftp
  .RecordSource = "select *from caixas where datacaixa='" & txtData.Value & "' and turno='" & cboTurno.Text & "' and codigoposto='" & dbConfig.Recordset!Porta & "' order by linhaexportada"
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
  End If
  If .Recordset.RecordCount = 0 Then
    MsgBox "O caixa atual ainda não foi exportado!"
    cmdImportar.Enabled = True
    Animation1.Visible = False
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  SoPrimeira = False
  If ApagaRegistros = False Then
    MsgBox "Este caixa não pode ser importado a segunda parte porque existe registro já gravado!"
    SoPrimeira = True
  End If
  
  Do While .Recordset.EOF = False
    StrTemp = .Recordset!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
      Case "001"
        GravaBico StrTemp
      Case "002"
        GravaVenda StrTemp
      Case "003"
        If SoPrimeira = False Then GravaNotas StrTemp
      Case "004"
        GravaTanque StrTemp
      Case "005"
        If SoPrimeira = False Then GravaNumerarios StrTemp
      Case "006"
        If SoPrimeira = False Then GravaDespesas StrTemp
      Case "007"
        GravaCupons StrTemp, dbVendasLeituraX
      Case "998"
        GravaResultado StrTemp
    End Select
    .Recordset.MoveNext
  Loop
End With

Call cmdCalcular_Click

Animation1.Close
Animation1.Visible = False
cmdImportar.Enabled = True

Shell "notepad " & App.Path & "\NotasBloqueadas.txt"

End Sub

Private Sub cmdImprimir_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Combustivel As Currency, Produtos As Currency
Dim Vendido As Double, Estoque As Double, Posto As Double, Diferenca As Double
Dim Y1 As Double, Y2 As Double, X1 As Double, X2 As Double
Dim Colunas As Double, Linhas As Double
Dim ColunaAtual As Double, LinhaAtual As Double
Dim Inicio As Double, Fim As Double
Dim VendasCombustivel As Double, Entrada As Double

If cmdAbrir.Enabled = True Then
  MsgBox "Escolha um caixa a ser impresso!"
  cboPdvs.SetFocus
  Exit Sub
End If

Call cmdCalcular_Click

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
StrTemp = "Demonstrativo de Movimento"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8
StrTemp = "Impresso em: " & Format(Dia, "Short date") & " - " & Format(Dia, "Short Time")
Printer.Print StrTemp

StrTemp = "Responsável: " & cboResponsavel.Text
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno: " & cboTurno.Text
Printer.CurrentX = 100
Printer.Print StrTemp;

StrTemp = "Data: " & Format(txtData.Value, "Short date")
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
With dbBicoEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Leitura de Bomba"
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y2 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Bico"
    Printer.CurrentX = 9 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Combustivel"
    Printer.CurrentX = 11
    Printer.Print StrTemp;
    
    StrTemp = "Inicial"
    Printer.CurrentX = 69 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Final"
    Printer.CurrentX = 99 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Preço"
    Printer.CurrentX = 119 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Retorno"
    Printer.CurrentX = 134 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Vendas"
    Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "R$ Total"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    With dbProdutos
      .RecordSource = "Select *From produtos where combustivel=-1"
      .Refresh
    End With
    
    .Recordset.MoveFirst
    VendasCombustivel = 0
    Do While .Recordset.EOF = False
      
      StrTemp = .Recordset!Bico
      Printer.CurrentX = 9 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      dbProdutos.Refresh
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      StrTemp = dbProdutos.Recordset!Descri
      Printer.CurrentX = 11
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Abertura, "#,##0.00")
      Printer.CurrentX = 69 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Encerrante, "#,##0.00")
      Printer.CurrentX = 99 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Preco, "#,##0.000")
      Printer.CurrentX = 119 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Retorno, "#,##0.00")
      Printer.CurrentX = 134 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno, "#,##0.00")
      Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      VendasCombustivel = VendasCombustivel + (.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno)
      
      If IsNull(.Recordset!ValorTotal) = False Then
        Combustivel = Combustivel + .Recordset!ValorTotal
      End If
      StrTemp = Format(.Recordset!ValorTotal, "Currency")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
  
      
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
    Y2 = Printer.CurrentY - 0.5
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (10, Y1)-(10, Y2)
    Printer.Line (40, Y1)-(40, Y2)
    Printer.Line (70, Y1)-(70, Y2)
    Printer.Line (100, Y1)-(100, Y2)
    Printer.Line (120, Y1)-(120, Y2)
    Printer.Line (135, Y1)-(135, Y2)
    Printer.Line (155, Y1)-(155, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = Format(VendasCombustivel, "#,##0.0")
    Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Combustivel, "Currency")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
  End If
End With


With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Diferença de Combustível"
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y2 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    StrTemp = "Tq."
    Printer.CurrentX = 9 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Combustivel"
    Printer.CurrentX = 11
    Printer.Print StrTemp;
    
    StrTemp = "Vendido"
    Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "No Sistema"
    Printer.CurrentX = 124 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "No Posto"
    Printer.CurrentX = 159 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Diferença"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5

    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      
      StrTemp = .Recordset!tanquenr
      Printer.CurrentX = 9 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 11
      Printer.Print StrTemp;
      
      Vendido = Vendido + .Recordset!Vendido
      StrTemp = Format(.Recordset!Vendido, "#,##0.0")
      Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      With dbTanquesEstoque
        .Refresh
        .Recordset.FindFirst "tanque=" & dbDifComb.Recordset!tanquenr
        If .Recordset.NoMatch = False Then
          Entrada = .Recordset!Entrada
        Else
          Entrada = 0
        End If
      End With
      Estoque = Estoque + .Recordset!Estoque + Entrada
      StrTemp = Format(.Recordset!Estoque + Entrada, "#,##0.0")
      Printer.CurrentX = 124 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Posto = Posto + .Recordset!Tanque
      StrTemp = Format(.Recordset!Tanque, "#,##0.0")
      Printer.CurrentX = 159 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Diferenca = Diferenca + .Recordset!Diferenca
      StrTemp = Format(.Recordset!Diferenca, "#,##0.0")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
    Y2 = Printer.CurrentY - 0.5
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (10, Y1)-(10, Y2)
    Printer.Line (55, Y1)-(55, Y2)
    Printer.Line (90, Y1)-(90, Y2)
    Printer.Line (125, Y1)-(125, Y2)
    Printer.Line (160, Y1)-(160, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    StrTemp = Format(Vendido, "#,##0.0")
    Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Estoque, "#,##0.0")
    Printer.CurrentX = 124 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Posto, "#,##0.0")
    Printer.CurrentX = 159 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Diferenca, "#,##0.0")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
  End If
  
End With

With dbVendas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Venda de Produtos"
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y2 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.Line (0, Y1)-(0, Y2)
    Printer.Line (Largura, Y1)-(Largura, Y2)
    
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    For i = 1 To 3
      StrTemp = "Cod."
      Printer.CurrentX = 14 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = "Qtd."
      Printer.CurrentX = 29 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = "Func."
      Printer.CurrentX = 44 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = "Total"
      Printer.CurrentX = 62 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      ColunaAtual = ColunaAtual + 63
    Next i
    ColunaAtual = 0
    Printer.Print ""
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5

    
    
    If .Recordset.RecordCount >= 3 Then
      Colunas = 3
      If .Recordset.RecordCount Mod 3 = 0 Then
        Linhas = .Recordset.RecordCount / 3
      Else
        If (.Recordset.RecordCount + 1) Mod 3 = 0 Then
          Linhas = (.Recordset.RecordCount + 1) / 3
        Else
          Linhas = (.Recordset.RecordCount + 2) / 3
        End If
      End If
    Else
      Colunas = .Recordset.RecordCount
      Linhas = 1
    End If
    LinhaAtual = 1
    ColunaAtual = 0
    Inicio = Printer.CurrentY
    
    Do While .Recordset.EOF = False
      If LinhaAtual > Linhas Then
        LinhaAtual = 1
        ColunaAtual = ColunaAtual + 63
        Y2 = Printer.CurrentY
        Printer.CurrentY = Inicio
      End If
      StrTemp = .Recordset!CodProduto
      Printer.CurrentX = 14 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Quantidade
      Printer.CurrentX = 29 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!codigovendedor
      Printer.CurrentX = 44 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Produtos = Produtos + .Recordset!ValorTotal
      StrTemp = Format(.Recordset!ValorTotal, "Currency")
      Printer.CurrentX = 62 + ColunaAtual - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      LinhaAtual = LinhaAtual + 1
      .Recordset.MoveNext
    Loop
    If Colunas < 3 Then
      Y2 = Printer.CurrentY
    End If
    .Recordset.MoveFirst
    
    ColunaAtual = 0
    Y2 = Y2 + 0.5
    For i = 1 To 3
      Printer.Line (0 + ColunaAtual, Y1)-(0 + ColunaAtual, Y2)
      Printer.Line (15 + ColunaAtual, Y1)-(15 + ColunaAtual, Y2)
      Printer.Line (30 + ColunaAtual, Y1)-(30 + ColunaAtual, Y2)
      Printer.Line (45 + ColunaAtual, Y1)-(45 + ColunaAtual, Y2)
      If i < 3 Then
        Printer.Line (63 + ColunaAtual, Y1)-(63 + ColunaAtual, Y2)
      Else
        Printer.Line (Largura, Y1)-(Largura, Y2)
      End If
      ColunaAtual = ColunaAtual + 63
    Next i
    
    'printer.CurrentY = printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5

    Printer.CurrentY = Y2
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    If IsNumeric(lblComissoes.Caption) = True Then
      Comissoes = CCur(lblComissoes.Caption)
    Else
      Comissoes = 0
    End If
    StrTemp = "Comissões a pagar: " & Format(Comissoes, "Currency")
    Printer.CurrentX = 0
    Printer.Print StrTemp;

    StrTemp = Format(Produtos, "Currency")
    Printer.CurrentX = 62 + ColunaAtual - 63 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
  End If
End With
StrTemp = "Estacionamento Inicial: " & txtEstacionaIni.Text & " / Final: " & txtEstacionaFim.Text & " / Cancelados: " & txtEstacionaCanc.Text & " / Total: " & lblEstacionaTotal.Caption
Produtos = Produtos + CCur(lblEstacionaTotal.Caption)
Printer.CurrentX = 0
Printer.Print StrTemp
Printer.CurrentY = Printer.CurrentY + 3

With dbProdutos
  .RecordSource = "Select *From produtos where combustivel=0"
  .Refresh
End With

StrTemp = "Total de Combustível:"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = Format(Combustivel, "Currency")
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Total de Produtos + Estacionamento:"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = Format(Produtos, "Currency")
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(120, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Faturamento calculado pelo sistema:"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = Format((Produtos + Combustivel), "Currency")
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Faturamento informado pelo caixa:"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = Format(txtInformado.Text, "Currency")
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(120, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

Printer.FontBold = True
StrTemp = "Diferença de Caixa:"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = Format(lblDiferenca.Caption, "Currency")
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.CurrentY = Printer.CurrentY + 5
If CCur(lblDiferenca.Caption) < 0 Then
  StrTemp = "Assinatura do caixa " & cboResponsavel.Text & ":"
  Printer.CurrentX = 0
  Printer.Print StrTemp;
    
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (Printer.TextWidth(StrTemp) + 1, Printer.CurrentY + Printer.TextHeight(StrTemp))-(Largura, Printer.CurrentY + Printer.TextHeight(StrTemp))
  Printer.Print ""
  Printer.CurrentY = Printer.CurrentY + 0.5
  
End If

StrTemp = "Observações:"
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Esta diferença de caixa poderá variar caso haja diferença na conferência de valores."
Printer.CurrentX = 0
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, 0, Printer.CurrentY, Largura

For i = 0 To 2
  Printer.CurrentY = Printer.CurrentY + 5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Next i
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Obs: A diferença de caixa será descontada do responsável."
Printer.CurrentX = 0
Printer.Print StrTemp



StrTemp = "Conferência de Valores"
Printer.FontSize = 12
Printer.CurrentX = 0
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, 0, Printer.CurrentY, Largura

Printer.FontSize = 8
StrTemp = "Caso haja diferença de valores na conferência, é obrigatório descrever no campo abaixo."
Printer.CurrentX = 0
Printer.Print StrTemp

For i = 0 To 3
  Printer.CurrentY = Printer.CurrentY + 5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Next i
Printer.CurrentY = Printer.CurrentY + 0.5


Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdIncluir_Click()
Dim Comissao As Currency, Unitario As Currency
TotalVenda
If IsNumeric(lblTotalVenda.Caption) = False Then
  MsgBox "Erro no total da venda!"
  Exit Sub
End If
If IsNumeric(txtCodProduto.Text) = False Then
  MsgBox "Informe um código válido!"
  txtCodProduto.SetFocus
  Exit Sub
End If
With dbProdutos
  .Refresh
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = True Then
    MsgBox "Produto não encontrado!"
    txtCodProduto.SetFocus
    Exit Sub
  End If
  If .Recordset!Estoque < CDbl(txtQtd.Text) Then
    Resposta = MsgBox("A venda atual tornará o estoque negativo! Deseja continuar?", vbYesNo + vbDefaultButton2)
    If Resposta = vbNo Then Exit Sub
  End If
  If .Recordset!Comissao <> 0 Or .Recordset!ComissaoValor <> 0 Then
    If IsNumeric(txtCodFunc.Text) = False Or txtCodFunc.Text = "0" Then
      MsgBox "Informe um codigo de funcionário válido!"
      txtCodFunc.SetFocus
      Exit Sub
    End If
    dbResponsavel2.Refresh
    dbResponsavel2.Recordset.FindFirst "codigo=" & txtCodFunc.Text
    If dbResponsavel2.Recordset.NoMatch = True Then
      MsgBox "Funcionário " & txtCodFunc.Text & " não encontrado!"
      txtCodFunc.SetFocus
    End If
  Else
    txtCodFunc.Text = ""
  End If
End With
If IsNumeric(txtQtd.Text) = False Then
  MsgBox "Informe uma quantidade válida!"
  txtQtd.SetFocus
  Exit Sub
End If
If qProdutosAltera.Recordset.RecordCount <> 0 Then
  qProdutosAltera.Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If qProdutosAltera.Recordset.NoMatch = False Then
    Unitario = qProdutosAltera.Recordset!PrecoVenda
  Else
    Unitario = CCur(lblTotalVenda.Caption) / CDbl(txtQtd.Text)
  End If
Else
  Unitario = dbProdutos.Recordset!PrecoVenda
End If

With dbProdutos
  If .Recordset!Comissao <> 0 Then
    Comissao = (Unitario * txtQtd.Text) * (.Recordset!Comissao)
  End If
  If .Recordset!ComissaoValor <> 0 Then
    Comissao = Comissao + (.Recordset!ComissaoValor * txtQtd.Text)
  End If
End With

With dbVendas
  .Recordset.AddNew
  .Recordset!CodigoFechamento = dbFechamento.Recordset!CodigoFechamento
  .Recordset!Hora = Now
  .Recordset!Data = txtData.Value
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!CodProduto = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!Quantidade = txtQtd.Text
  .Recordset!valorUnitario = Unitario
  .Recordset!ValorTotal = CCur(lblTotalVenda.Caption)
  If IsNumeric(txtCodFunc.Text) = True Then
    .Recordset!codigovendedor = txtCodFunc.Text
    .Recordset!CodigoPagamento = dbResponsavel2.Recordset!codigovendedor
  End If
  .Recordset!ValorComissao = Comissao
  .Recordset.Update
End With

Call cmdCalcular_Click

txtCodFunc.Text = ""
txtCodProduto.Text = ""
txtQtd.Text = ""
txtCodProduto.SetFocus
lblEstoque.Caption = ""
End Sub

Private Sub cmdAbrir_Click()
Dim CodigoFechamento As Double, Abertura As Double
Dim CaixaAnterior As Double, AlteraPreco As Boolean
Dim UltimoEstacionamento As Double, AnteriorFechado As Boolean
Dim CaixaFechado As Boolean

Abrindo = True
ErroNaSoma = False
If DateDiff("d", Date, txtData.Value) >= 1 Then
  Resposta = MsgBox("Deseja criar um caixa futuro?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
End If


With dbFechamento
  .RecordSource = "select *from fechamentodecaixa order by datacaixa, horaini"
  .Refresh
  If cboTurno.Text <> dbTurno.Recordset!Descri Then
    Call cboTurno_LostFocus
    If cboTurno.Text <> dbTurno.Recordset!Descri Then
      MsgBox "Turno não encontrado!"
      On Error Resume Next
      cboTurno.SetFocus
      Exit Sub
    End If
  End If
  .Refresh
  .Recordset.FindFirst "datacaixa=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and codigoturno=" & dbTurno.Recordset!CodigoTurno
  If .Recordset.NoMatch = True Then
    'verifica se não existe caixa finalizado posterior
    .Recordset.FindLast "datacaixa>=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and fechado=-1"
    If .Recordset.NoMatch = False Then
      If .Recordset!DataCaixa = txtData.Value Then
        If .Recordset!HoraIni > dbTurno.Recordset!HoraIni Then
          MsgBox "Já existe caixa posterior finalizado!"
          Exit Sub
        End If
      Else
        MsgBox "Já existe caixa posterior finalizado!"
        Exit Sub
      End If
    End If
    .Recordset.AddNew
    .Recordset!DataCaixa = txtData.Value
    .Recordset!CodigoTurno = dbTurno.Recordset!CodigoTurno
    .Recordset!Turno = dbTurno.Recordset!Descri
    .Recordset!HoraIni = dbTurno.Recordset!HoraIni
    .Recordset!horafim = dbTurno.Recordset!horafim
    .Recordset.Update
  End If
  .Refresh
  .Recordset.FindFirst "datacaixa=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and codigoturno=" & dbTurno.Recordset!CodigoTurno
  If .Recordset.NoMatch = True Then
    MsgBox "Erro ao criar folha de caixa!"
    Exit Sub
  End If
  txtData.Enabled = False
  cboTurno.Enabled = False
  cboPdvs.Enabled = False
  cmdAbrir.Enabled = False
  cmdSair.Enabled = False
  cmdCancelar.Visible = True
  cmdCancelar.Cancel = True
  CodigoFechamento = .Recordset!CodigoFechamento
  If .Recordset!fechado = True Then
    CaixaFechado = True
    Frame1.Enabled = False
    DBGrid1.AllowDelete = False
    DBGrid1.AllowUpdate = False
    DBGrid2.AllowDelete = False
    DBGrid2.AllowUpdate = False
    DBGrid3.AllowDelete = False
    DBGrid3.AllowUpdate = False
    DBGrid4.AllowDelete = False
    DBGrid4.AllowUpdate = False
    cmdImportar.Visible = False
    cmdEntraCombustivel.Visible = False
  Else
    CaixaFechado = False
    Frame1.Enabled = True
    DBGrid1.AllowUpdate = True
    DBGrid2.AllowUpdate = True
    DBGrid3.AllowDelete = True
    cmdImportar.Visible = True
    cmdEntraCombustivel.Visible = True
  End If
  If IsNumeric(.Recordset!TotalCombustivel) = True Then
    lblTotalCombustivel.Caption = Format(.Recordset!TotalCombustivel, "Currency")
  End If
  If IsNumeric(.Recordset!TotalProdutos) = True Then
    lblTotalProdutos.Caption = Format(.Recordset!TotalProdutos, "Currency")
  End If
End With
dbFechamento2.Refresh
dbFechamento2.Recordset.FindFirst "codigofechamento=" & CodigoFechamento
If dbFechamento2.Recordset.NoMatch = False Then
  dbFechamento2.Recordset.MovePrevious
  If dbFechamento2.Recordset.BOF = False Then
    FechamentoAnterior = dbFechamento2.Recordset!CodigoFechamento
    AnteriorFechado = dbFechamento2.Recordset!fechado
    dbBicoEncerrantes2.RecordSource = "select *from BicoEncerrantes where codigofechamento=" & FechamentoAnterior & " order by bico"
    dbBicoEncerrantes2.Refresh
    dbFechamento2.Recordset.MoveNext
  Else
    FechamentoAnterior = -1
  End If
End If

If CaixaFechado = False Then
  With dbAlteracao
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      Do While .Recordset.BOF = False
        If .Recordset!dataalteracao <= dbFechamento.Recordset!DataCaixa Then
          If .Recordset!dataalteracao < dbFechamento.Recordset!DataCaixa Then
            AlteraPreco = True
            Exit Do
          Else
            If .Recordset!HoraIni <= dbFechamento.Recordset!HoraIni Then
              AlteraPreco = True
              Exit Do
            Else
              AlteraPreco = False
            End If
          End If
        End If
        .Recordset.MovePrevious
      Loop
    End If
  End With
End If

If dbFechamento.Recordset!fechado = True Then
  AlteraPreco = False
End If
With dbBicoEncerrantes
  .RecordSource = "select *from BicoEncerrantes where codigofechamento=" & CodigoFechamento & " order by bico"
  .Refresh
  If CaixaFechado = False Then
    If .Recordset.RecordCount = 0 Then
      dbBicos.Refresh
      If dbBicos.Recordset.RecordCount <> 0 Then
        dbBicos.Recordset.MoveLast
        dbBicos.Recordset.MoveFirst
        Do While dbBicos.Recordset.EOF = False
          .Recordset.AddNew
          .Recordset!CodigoFechamento = CodigoFechamento
          .Recordset!Bico = dbBicos.Recordset!Bico
          If FechamentoAnterior >= 0 And AnteriorFechado = False Then
            dbBicoEncerrantes2.Refresh
            dbBicoEncerrantes2.Recordset.FindFirst "bico=" & dbBicos.Recordset!Bico
            If dbBicoEncerrantes2.Recordset.NoMatch = False Then
              Abertura = dbBicoEncerrantes2.Recordset!Abertura
            End If
          Else
            Abertura = dbBicos.Recordset!ultimonumero
          End If
          
          .Recordset!Abertura = Abertura
          If AlteraPreco = True Then
            dbAlteraBico.Recordset.FindFirst "bico=" & .Recordset!Bico
            If .Recordset.NoMatch = False Then
              .Recordset!Preco = dbAlteraBico.Recordset!Preco
            Else
              .Recordset!Preco = dbBicos.Recordset!PrecoVenda
            End If
          Else
            .Recordset!Preco = dbBicos.Recordset!PrecoVenda
          End If
          .Recordset!CodigoProduto = dbBicos.Recordset!CodigoProduto
          .Recordset!Tanque = dbBicos.Recordset!Tanque
          .Recordset.Update
          dbBicos.Recordset.MoveNext
        Loop
        .Recordset.MoveFirst
      End If
    Else
      If AlteraPreco = True Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          .Recordset.Edit
          dbAlteraBico.Recordset.FindFirst "bico=" & .Recordset!Bico
          If .Recordset.NoMatch = False Then
            .Recordset!Preco = dbAlteraBico.Recordset!Preco
          End If
          .Recordset.Update
          .Recordset.MoveNext
        Loop
        .Recordset.MoveFirst
      End If
    End If
    .Refresh
  End If
End With
With dbTanquesEstoque
  .RecordSource = "select *from TanqueEstoque where codigofechamento=" & CodigoFechamento & " order by tanque"
  .Refresh
  If CaixaFechado = False Then
    If .Recordset.RecordCount = 0 Then
      dbTanques.RecordSource = "select *from tanques order by tanque"
      dbTanques.Refresh
      If dbTanques.Recordset.RecordCount <> 0 Then
        dbTanques.Recordset.MoveLast
        dbTanques.Recordset.MoveFirst
        Do While dbTanques.Recordset.EOF = False
          .Recordset.AddNew
          .Recordset!CodigoFechamento = CodigoFechamento
          .Recordset!Tanque = dbTanques.Recordset!Tanque
          .Recordset.Update
          dbTanques.Recordset.MoveNext
        Loop
      End If
    End If
  End If
End With
With dbVendas
  .Connect = Conectar
  .RecordSource = "select *from venda2 where codigofechamento=" & CodigoFechamento & " order by codproduto"
  .Refresh
End With
With dbClientesNotas
  .Connect = Conectar
  .RecordSource = "select *from ClientesNota2 where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbFormaDePgRecebido
  .Connect = Conectar
  .RecordSource = "select *from formadepagamentorecebido2 where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbDespesasLanc
  .Connect = Conectar
  .RecordSource = "select *from despesaslanc2 where codigofechamento=" & CodigoFechamento
  .Refresh
End With
With dbDespesasLanc2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from despesaslanc2 where descri='Comissões paga no caixa'"
  .Refresh
End With
With dbDifComb
  .Connect = Conectar
  .RecordSource = "select *from diferencacombustivel where codigofechamento=" & CodigoFechamento
  .Refresh
End With

If dbFechamento.Recordset!fechado = False Then
  cmdConfirmar.Visible = True
  cmdCalcular.Visible = True
Else
  cmdConfirmar.Visible = False
  cmdCalcular.Visible = False
End If

If CaixaFechado = False Then
  If FechamentoAnterior > 0 Then
    With dbBicoEncerrantes2
      .RecordSource = "select *from BicoEncerrantes where codigofechamento=" & FechamentoAnterior & " order by bico"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          dbBicoEncerrantes.Recordset.FindFirst "bico=" & .Recordset!Bico
          If dbBicoEncerrantes.Recordset.NoMatch = False Then
            dbBicoEncerrantes.Recordset.Edit
            dbBicoEncerrantes.Recordset!Abertura = .Recordset!Encerrante
            dbBicoEncerrantes.Recordset.Update
          End If
          .Recordset.MoveNext
        Loop
      End If
    End With
  End If
End If
With dbEstacionamentoCaixa
  If FechamentoAnterior > 0 Then
    .RecordSource = "Select *from estacionamentocaixa where codigocaixa=" & FechamentoAnterior
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      If IsNull(.Recordset!final) = False Then
        UltimoEstacionamento = .Recordset!final
      End If
    End If
    If UltimoEstacionamento = 0 Then
      UltimoEstacionamento = dbEstacionamento.Recordset!ultimonumero
    End If
    .RecordSource = "Select *from estacionamentocaixa where codigocaixa=" & CodigoFechamento
    .Refresh
    If .Recordset.RecordCount = 0 Then
      .Recordset.AddNew
      .Recordset!CodigoCaixa = CodigoFechamento
      .Recordset!inicial = UltimoEstacionamento
      .Recordset!final = UltimoEstacionamento
      .Recordset!Preco = dbEstacionamento.Recordset!Preco
      .Recordset!totalun = 0
      .Recordset!cancelados = 0
      .Recordset!Total = 0
      .Recordset.Update
      .Refresh
    Else
      If dbFechamento.Recordset!fechado = False Then
        .Recordset.Edit
        .Recordset!inicial = UltimoEstacionamento
        If IsNull(.Recordset!final) = False Then
          If .Recordset!final = 0 Then
            .Recordset.final = UltimoEstacionamento
          End If
        Else
          .Recordset.final = UltimoEstacionamento
        End If
        .Recordset!Preco = dbEstacionamento.Recordset!Preco
        .Recordset.Update
      End If
      .Refresh
    End If
  End If
End With

With qProdutosAltera
  .RecordSource = "select *from produtosaltera order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "datacaixa<=#" & DataInglesa(txtData.Value) & "#"
    If .Recordset.NoMatch = True Then
      AlteraAnterior = 0
    Else
      If .Recordset!DataCaixa < dbFechamento.Recordset!DataCaixa Then
        AlteraAnterior = .Recordset!codigoprodutoaltera
      Else
        If .Recordset!DataCaixa = dbFechamento.Recordset!DataCaixa And .Recordset!HoraIni = dbFechamento.Recordset!HoraIni Then
          AlteraAnterior = .Recordset!codigoprodutoaltera
        ElseIf .Recordset!HoraIni > dbFechamento.Recordset!HoraIni Then
          .Recordset.MoveNext
          Do While .Recordset.EOF = False
            If .Recordset!DataCaixa = dbFechamento.Recordset!DataCaixa Then
              If .Recordset!HoraIni <= dbFechamento.Recordset!HoraIni Then
                AlteraAnterior = .Recordset!codigoprodutoaltera
                Exit Do
              End If
            Else
              If .Recordset!DataCaixa < dbFechamento.Recordset!DataCaixa Then
                AlteraAnterior = .Recordset!codigoprodutoaltera
                Exit Do
              End If
            End If
            .Recordset.MoveNext
          Loop
        Else
          AlteraAnterior = .Recordset!codigoprodutoaltera
        End If
      End If
    End If
  Else
    AlteraAnterior = 0
  End If
  .RecordSource = "select *from qprodutosaltera where produtosaltera.codigoprodutoaltera=" & AlteraAnterior & " order by datacaixa desc, horaini desc, codigo"
  .Refresh
End With
With dbResultado
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from fechamentodecaixapista where codigofechamento=" & CodigoFechamento
  .Refresh
End With

Call cmdCalcular_Click
dbBicoEncerrantes.Refresh
dbTanques.Refresh
dbVendas.Refresh
dbDifComb.Refresh

cmdRemover.Enabled = True
Frame1.Visible = True
cboResponsavel.Visible = True
cboPdvs.Enabled = False
DBGrid1.Visible = True
DBGrid2.Visible = True
DBGrid3.Visible = True
DBGrid4.Visible = True

If Frame1.Enabled = True Then
  cboResponsavel.SetFocus
End If

With dbFechamento
  On Error Resume Next
  .Recordset.Edit
  .Recordset!responsavel = cboResponsavel.Text
  .Recordset.Update
End With

Abrindo = False
End Sub

Private Sub cmdPosterior_Click()
With dbFechamento
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.EOF = False Then
      .Recordset.MoveNext
      If .Recordset.EOF = False Then
        txtData.Value = .Recordset!DataCaixa
        cboTurno = .Recordset!Turno
        Call cmdAbrir_Click
      End If
    End If
  End If
End With
End Sub

Private Sub cmdPrimeiro_Click()
With dbFechamento
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    If .Recordset.BOF = False Then
      txtData.Value = .Recordset!DataCaixa
      cboTurno = .Recordset!Turno
      Call cmdAbrir_Click
    End If
  End If
End With
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer, CodigoFechamento As Double
Dim Ws As Workspace, db As Database
If Usuarios.Nome <> "Usuário Master" Then
  MsgBox "Você não tem permissão para remover um caixa!"
  Exit Sub
End If
With dbFechamento
  If Frame1.Visible = False Then Exit Sub
  If .Recordset!fechado = True Then
    MsgBox "Não é possível remover um caixa já finalizado!"
    Exit Sub
  End If
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  CodigoFechamento = .Recordset!CodigoFechamento
  Resposta = MsgBox("Deseja remover o caixa atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  Permissao = False
  frmPermissao.Show vbModal
  If Permissao = False Then
    Exit Sub
  End If
  .Recordset.Delete
  Set Ws = DBEngine.Workspaces(0)
  Set db = Ws.OpenDatabase(Caminho, , , Conectar)
  db.Execute "delete *from BicoEncerrantes where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from TanqueEstoque where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from venda2 where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from diferencacombustivel where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from estacionamentocaixa where codigocaixa=" & CodigoFechamento
  Call cmdCancelar_Click
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdUltimo_Click()
With dbFechamento
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    If .Recordset.EOF = False Then
      txtData.Value = .Recordset!DataCaixa
      cboTurno = .Recordset!Turno
      Call cmdAbrir_Click
    End If
  End If
End With
End Sub

Private Sub dbAlteracao_Reposition()
If dbAlteracao.Recordset.EOF = True Then Exit Sub
If dbAlteracao.Recordset.BOF = True Then Exit Sub
With dbAlteraBico
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from alterabico where codalteracao=" & dbAlteracao.Recordset!codalteracao & " order by bico"
  .Refresh
End With
End Sub

Private Sub DBGrid1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DBGrid1_LostFocus()
Me.KeyPreview = True
Call cmdCalcular_Click
End Sub

Private Sub DBGrid2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DBGrid2_LostFocus()
Me.KeyPreview = True
Call cmdCalcular_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF5
    If Shift = 1 Then
      PegaCupons
    Else
      If cmdImportar.Visible = True Then
        Call cmdImportar_Click
      End If
    End If
  Case vbKeyF1
    Call cmdPrimeiro_Click
  Case vbKeyF2
    Call cmdAnterior_Click
  Case vbKeyF3
    Call cmdPosterior_Click
  Case vbKeyF4
    Call cmdUltimo_Click
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyF5
    Call cmdImportar_Click
End Select
End Sub

Private Sub Form_Load()
If Usuarios.Grupo.AdmEstatus = 2 Then
  cmdDesconfirmar.Visible = True
Else
  cmdDesconfirmar.Visible = False
End If
With dbConfig
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

txtData.Value = Date
FechamentoAnterior = -1

With dbPDVs
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from pdvs order by descri"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If IsNull(.Recordset!Descri) = False Then
      cboPdvs.Text = .Recordset!Descri
    End If
  End If
End With

With dbEstacionamento
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from Estacionamento"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Preco = 0
    .Recordset!ultimonumero = 0
    .Recordset.Update
    .Refresh
  End If
End With
With dbEstacionamentoCaixa
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from estacionamentocaixa where codigocaixa=0"
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
  .RecordSource = "select bloqueiafechamento.*, Turnos.* from bloqueiafechamento, turnos where bloqueiafechamento.codigoturno1=turnos.codigoturno"
  .Refresh
End With
With dbEstatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbAlteracao
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select alteracoes.*, turnos.* from alteracoes, turnos where turnos.codigoturno=alteracoes.codigoturno order by dataalteracao, horaini"
  .Refresh
End With
With dbAlteraBico
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbResponsavel
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbTurno
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbFechamento
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from fechamentodecaixa where codigofechamento=0"
  .Refresh
End With
With dbBicos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbBicoEncerrantes
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from bicoencerrantes where codigofechamento=0"
  .Refresh
End With
With dbTanques
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbTanquesEstoque
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from tanqueestoque where codigofechamento=0"
  .Refresh
End With
With dbProdutos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbVendas
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from venda2"
  .Refresh
End With
With dbDifComb
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from diferencacombustivel where codigofechamento=0"
  .Refresh
End With
With dbResponsavel2
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbPostos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbFechamento2
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "Select *from fechamentodecaixa where datacaixa>=#" & DataInglesa(DateAdd("m", -2, Date)) & "# order by datacaixa, horaini"
  .Refresh
End With
With dbUltimoFechamento
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from FechamentoDeCaixa where datacaixa>=#" & DataInglesa(DateAdd("m", -2, Date)) & "# and fechado=-1 order by datacaixa, HoraIni desc"
  .Refresh
End With
With dbUltimoEncerrante
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from BicoEncerrantes where codigofechamento=0"
  .Refresh
End With
With dbBicoEncerrantes2
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbNotasCorpo
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select sum(quantidade) as total2 from produtosnotascorpo where aguardando=-1 and codigoproduto=0"
  .Refresh
End With
With dbDespesasTipo
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbDespesasLanc
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from despesaslanc2 where codigofechamento=-1"
  .Refresh
End With
With dbFechamento
  .RecordSource = "select *from fechamentodecaixa where datacaixa>=#" & DataInglesa(DateAdd("m", -2, Date)) & "# and fechado=-1 order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    txtData.Value = .Recordset!DataCaixa
    cboTurno.Text = .Recordset!Turno
  End If
End With
With qProdutosAltera
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbProdutosHistorico
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from produtoshistorico where lancadoem>=#" & DataInglesa(DateAdd("m", -2, Date)) & "#"
  .Refresh
End With
With dbClientesNotas
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from clientesnota2 where codigofechamento=0"
  .Refresh
End With
With dbClientes
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbClientesCarros
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbFormaDePg
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbFormaDePgRecebido
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from formadepagamentorecebido2 where codigofechamento=0"
  .Refresh
End With
With dbVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from cuponsfiscais where datacupom=#" & Date & "#"
  .Refresh
End With
With dbClientesProdutos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbEncerrantesNovos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from bicosencerrantesnovo order by datacaixa, horaini, bico"
  .Refresh
End With

Frame1.Visible = False
cboResponsavel.Visible = False
DBGrid1.Visible = False
DBGrid2.Visible = False
DBGrid3.Visible = False
DBGrid4.Visible = False

txtData.Enabled = True
cboTurno.Enabled = True
cboPdvs.Enabled = True
cmdAbrir.Enabled = True
cmdSair.Enabled = True
cmdCancelar.Visible = True
cmdSair.Cancel = True
cmdRemover.Enabled = False


Select Case Usuarios.Grupo.ControleFechamentoDiario
  Case 0, 1
    cmdConfirmar.Enabled = False
  Case 2
    cmdConfirmar.Enabled = True
End Select

Select Case Usuarios.Grupo.ControleFechamentoDiario
  Case 0 'Somente leitura
    cmdRemover.Enabled = False
    cmdEntraCombustivel.Enabled = False
    cmdConfirmar.Enabled = False
    DBGrid1.AllowUpdate = False
    DBGrid2.AllowUpdate = False
    cmdIncluir.Enabled = False
    DBGrid3.AllowDelete = False
    cboResponsavel.Enabled = False
    txtInformado.Enabled = False
  Case 2 'Liberado
    
End Select
On Error Resume Next
cboPdvs.SetFocus
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
  Case 0
    'No state to report.
  Case 1
    'The control is looking up the IP address
    'of the specified host computer.
  Case 2
    'The control successfully found the IP address
    'of the specified host computer.
  Case 3
    'The control is connecting to the host computer.
  Case 4
    'The control successfully connected to the host computer.
  Case 5
    'The control is sending a request to the host computer.
  Case 6
    'The control successfully sent the request.
  Case 7
    'The control is receiving a response from the host computer.
  Case 8
    'The control successfully received a response from the host computer.
  Case 9
    'The control is disconnecting from the host computer.
  Case 10
    'The control successfully disconnected from the host computer.
  Case 11
    'An error occurred in communicating with the host computer.
  Case 12
    'The request has completed and all data has been received.
End Select
End Sub

Private Sub txtCodFunc_GotFocus()
With txtCodFunc
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

Private Sub txtCodProduto_LostFocus()
lblEstoque.Caption = ""
With dbProdutos
  .Refresh
  If txtCodProduto.Text = "" Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    lblEstoque.Caption = .Recordset!Estoque
  End If
End With
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtEstacionaCanc_GotFocus()
With txtEstacionaCanc
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEstacionaCanc_LostFocus()
Call cmdCalcular_Click
End Sub

Private Sub txtEstacionaFim_GotFocus()
With txtEstacionaFim
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEstacionaFim_LostFocus()
Call cmdCalcular_Click
End Sub

Private Sub txtInformado_LostFocus()
Call cmdCalcular_Click
End Sub

Private Sub txtQtd_GotFocus()
With txtQtd
  .SelStart = 0
  .SelLength = Len(.Text)
End With

End Sub

Private Sub txtQtd_LostFocus()
TotalVenda
End Sub

