VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmUsuariosGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupos de Usuários"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmUsuariosGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   1320
      TabIndex        =   133
      Top             =   720
      Width           =   1215
   End
   Begin VB.Data dbUsuarios 
      Caption         =   "dbUsuarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Usuarios"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data dbGrupos 
      Caption         =   "dbGrupos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from UsuariosGrupos order by descri"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDescri 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin MSDBCtls.DBList lstGrupos 
      Bindings        =   "frmUsuariosGrupo.frx":0442
      Height          =   5130
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9049
      _Version        =   393216
      ListField       =   "Descri"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "frmUsuariosGrupo.frx":0459
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label36"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboCadBombasDeCombustivel"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboCadClientes"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboCadClientesCheques"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboCadContas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cboCadDespesasTipo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboCadDespesasBancarias"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboCadFormaDePagamento"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cboCadFornecedores"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cboCadFuncionarios"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboCadJuros"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboCadPostos"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboCadProdutos"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cboCadProdutosFornecedores"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cboCadTanques"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboCadTurnos"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cboCadConfiguracao"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtClientesPlanos"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Controle"
      TabPicture(1)   =   "frmUsuariosGrupo.frx":0475
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboCtrlVales"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboCtrlLavagem"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboCtrlLuz"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboCtrlAgua"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboCtrlCobrancaDeClientes"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboCtrlContasAPagar"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboCtrlLancamentoDeContasAPagar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboCtrlLancamentoDeNotas"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cboCtrlPagamentosAntecipados"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboCtrlCartoesPendentes"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cboCtrlConferencia"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cboCtrlFechamento"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label31"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label30"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label29"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label28"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label27"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label26"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label25"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label24"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label23"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label22"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label21"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label20"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Cheques"
      TabPicture(2)   =   "frmUsuariosGrupo.frx":0491
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboChkPorData"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboChkEmpresaCobranca"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboChkProtesto"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboChkCobranca"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cboChkDevolucao"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboChkDeposito"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label42"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label41"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label40"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label39"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label38"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label37"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Banco"
      TabPicture(3)   =   "frmUsuariosGrupo.frx":04AD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboBcoTransfere"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cboBcoConcilia"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label55"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label54"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Relatórios"
      TabPicture(4)   =   "frmUsuariosGrupo.frx":04C9
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboRelatKilometragemDeClientes"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cboRelatFatuamentoDeCheques"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cboRelatRetornoDeCombustivel"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cboRelatCadastroIncompleto"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cboRelatProtestoDeCheques"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cboRelatVendaDiariaCombustivel"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cboRelatVendaMediaProdutos"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cboRelatVendaLucro"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cboRelatVendasDetalhada"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "cboRelatVendasDeProdutos"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "cboRelatGalonagemTotal"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cboRelatGalonagem"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "cboRelatFormaDePagamento"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "cboRelatDifCombustivel"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "cboRelatDifRecebimentos"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "cboRelatDifCaixa"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "cboRelatComprasVendas"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "cboRelatProdutosComprados"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "cboRelatChequeCliente"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "cboRelatAcertoEstoque"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Label34"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Label33"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "Label32"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "Label72"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Label71"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "Label70"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "Label69"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "Label68"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "Label67"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "Label66"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "Label65"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "Label64"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).Control(32)=   "Label63"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "Label62"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "Label61"
      Tab(4).Control(34).Enabled=   0   'False
      Tab(4).Control(35)=   "Label60"
      Tab(4).Control(35).Enabled=   0   'False
      Tab(4).Control(36)=   "Label59"
      Tab(4).Control(36).Enabled=   0   'False
      Tab(4).Control(37)=   "Label58"
      Tab(4).Control(37).Enabled=   0   'False
      Tab(4).Control(38)=   "Label57"
      Tab(4).Control(38).Enabled=   0   'False
      Tab(4).Control(39)=   "Label56"
      Tab(4).Control(39).Enabled=   0   'False
      Tab(4).ControlCount=   40
      TabCaption(5)   =   "Administração"
      TabPicture(5)   =   "frmUsuariosGrupo.frx":04E5
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cboLiberaNotas"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cboAdmDatas"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cboAdmGruposDeUsuarios"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cboAdmUsuarios"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cboAdmLMC"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cboAdmTotalDeVendas"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "cboAdmEstatus"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "cboAdmConfirmaDespesas"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label35"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label19"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label78"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label77"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label76"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "Label75"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Label74"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Label73"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).ControlCount=   16
      Begin VB.TextBox txtClientesPlanos 
         Height          =   285
         Left            =   2880
         TabIndex        =   139
         Top             =   4560
         Width           =   3015
      End
      Begin VB.ComboBox cboLiberaNotas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0501
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":050E
         TabIndex        =   136
         Text            =   "Bloqueado"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmDatas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0538
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0545
         TabIndex        =   134
         Text            =   "Bloqueado"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatKilometragemDeClientes 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":056F
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":057C
         TabIndex        =   132
         Text            =   "Bloqueado"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatFatuamentoDeCheques 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":05A6
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":05B3
         TabIndex        =   130
         Text            =   "Bloqueado"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatRetornoDeCombustivel 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":05DD
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":05EA
         TabIndex        =   128
         Text            =   "Bloqueado"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmGruposDeUsuarios 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0614
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0621
         TabIndex        =   126
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmUsuarios 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":064B
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0658
         TabIndex        =   124
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmLMC 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0682
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":068F
         TabIndex        =   122
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmTotalDeVendas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":06B9
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":06C6
         TabIndex        =   120
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmEstatus 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":06F0
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":06FD
         TabIndex        =   118
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboAdmConfirmaDespesas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0727
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0734
         TabIndex        =   116
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatCadastroIncompleto 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":075E
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":076B
         TabIndex        =   114
         Text            =   "Bloqueado"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatProtestoDeCheques 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0795
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":07A2
         TabIndex        =   112
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatVendaDiariaCombustivel 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":07CC
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":07D9
         TabIndex        =   110
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatVendaMediaProdutos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0803
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":0810
         TabIndex        =   108
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatVendaLucro 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":083A
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":0847
         TabIndex        =   106
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatVendasDetalhada 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0871
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":087E
         TabIndex        =   104
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatVendasDeProdutos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":08A8
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":08B5
         TabIndex        =   102
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatGalonagemTotal 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":08DF
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":08EC
         TabIndex        =   100
         Text            =   "Bloqueado"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatGalonagem 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0916
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0923
         TabIndex        =   98
         Text            =   "Bloqueado"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatFormaDePagamento 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":094D
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":095A
         TabIndex        =   96
         Text            =   "Bloqueado"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatDifCombustivel 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0984
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0991
         TabIndex        =   94
         Text            =   "Bloqueado"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatDifRecebimentos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":09BB
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":09C8
         TabIndex        =   92
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatDifCaixa 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":09F2
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":09FF
         TabIndex        =   90
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatComprasVendas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0A29
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0A36
         TabIndex        =   88
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatProdutosComprados 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0A60
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0A6D
         TabIndex        =   86
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatChequeCliente 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0A97
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0AA4
         TabIndex        =   84
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboRelatAcertoEstoque 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0ACE
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0ADB
         TabIndex        =   82
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboBcoTransfere 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0B05
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0B12
         TabIndex        =   80
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboBcoConcilia 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0B3C
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0B49
         TabIndex        =   78
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboChkPorData 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0B73
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0B80
         TabIndex        =   76
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboChkEmpresaCobranca 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0BAA
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0BB7
         TabIndex        =   74
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboChkProtesto 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0BE1
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0BEE
         TabIndex        =   72
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboChkCobranca 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0C18
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0C25
         TabIndex        =   70
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboChkDevolucao 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0C4F
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0C5C
         TabIndex        =   68
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboChkDeposito 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0C86
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0C93
         TabIndex        =   66
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlVales 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0CBD
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":0CCA
         TabIndex        =   64
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlLavagem 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0CF4
         Left            =   -72120
         List            =   "frmUsuariosGrupo.frx":0D01
         TabIndex        =   62
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlLuz 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0D2B
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0D38
         TabIndex        =   60
         Text            =   "Bloqueado"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlAgua 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0D62
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0D6F
         TabIndex        =   58
         Text            =   "Bloqueado"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlCobrancaDeClientes 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0D99
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0DA6
         TabIndex        =   56
         Text            =   "Bloqueado"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlContasAPagar 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0DD0
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0DDD
         TabIndex        =   54
         Text            =   "Bloqueado"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlLancamentoDeContasAPagar 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0E07
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0E14
         TabIndex        =   52
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlLancamentoDeNotas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0E3E
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0E4B
         TabIndex        =   50
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlPagamentosAntecipados 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0E75
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0E82
         TabIndex        =   48
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlCartoesPendentes 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0EAC
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0EB9
         TabIndex        =   46
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlConferencia 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0EE3
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0EF0
         TabIndex        =   44
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboCtrlFechamento 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0F1A
         Left            =   -74760
         List            =   "frmUsuariosGrupo.frx":0F27
         TabIndex        =   42
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCadConfiguracao 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0F51
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":0F5E
         TabIndex        =   40
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboCadTurnos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0F88
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":0F95
         TabIndex        =   38
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboCadTanques 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0FBF
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":0FCC
         TabIndex        =   36
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboCadProdutosFornecedores 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":0FF6
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":1003
         TabIndex        =   34
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboCadProdutos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":102D
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":103A
         TabIndex        =   32
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboCadPostos 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":1064
         Left            =   2880
         List            =   "frmUsuariosGrupo.frx":1071
         TabIndex        =   30
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCadJuros 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":109B
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":10A8
         TabIndex        =   28
         Text            =   "Bloqueado"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.ComboBox cboCadFuncionarios 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":10D2
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":10DF
         TabIndex        =   26
         Text            =   "Bloqueado"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ComboBox cboCadFornecedores 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":1109
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":1116
         TabIndex        =   24
         Text            =   "Bloqueado"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox cboCadFormaDePagamento 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":1140
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":114D
         TabIndex        =   22
         Text            =   "Bloqueado"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox cboCadDespesasBancarias 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":1177
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":1184
         TabIndex        =   20
         Text            =   "Bloqueado"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ComboBox cboCadDespesasTipo 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":11AE
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":11BB
         TabIndex        =   18
         Text            =   "Bloqueado"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cboCadContas 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":11E5
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":11F2
         TabIndex        =   16
         Text            =   "Bloqueado"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboCadClientesCheques 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":121C
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":1229
         TabIndex        =   14
         Text            =   "Bloqueado"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboCadClientes 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":1253
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":1260
         TabIndex        =   12
         Text            =   "Bloqueado"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboCadBombasDeCombustivel 
         Height          =   315
         ItemData        =   "frmUsuariosGrupo.frx":128A
         Left            =   240
         List            =   "frmUsuariosGrupo.frx":1297
         TabIndex        =   10
         Text            =   "Bloqueado"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Clientes Planos (separados por vírgula):"
         Height          =   195
         Left            =   2880
         TabIndex        =   138
         Top             =   4320
         Width           =   2805
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Liberação de Cheques e Notas bloqueados:"
         Height          =   315
         Left            =   -74760
         TabIndex        =   137
         Top             =   4920
         Width           =   3225
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Liberação de datas:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   135
         Top             =   4320
         Width           =   1410
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Kilometragem de Clientes/Estacionamento:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   131
         Top             =   6120
         Width           =   3030
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Faturamento de Cheques:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   129
         Top             =   5520
         Width           =   1830
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Retorno de Combustível:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   127
         Top             =   4920
         Width           =   1770
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "Grupos de Usuários:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   125
         Top             =   3720
         Width           =   1440
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "Usuários:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   123
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "LMC:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   121
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "Total de Vendas:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   119
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   117
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar Despesas:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   115
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Cadastro Incompleto:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   113
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Protesto de Cheques:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   111
         Top             =   3720
         Width           =   1530
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Venda Diária de Combustível:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   109
         Top             =   3120
         Width           =   2115
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "Venda Média de Produtos:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   107
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "Venda/Lucro de Produtos Comissionados:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   105
         Top             =   1920
         Width           =   2985
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "Venda Detalhada:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   103
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Vendas de Produtos:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   101
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Galonagem Total:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   99
         Top             =   6120
         Width           =   1260
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Galonagem:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   97
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   95
         Top             =   4920
         Width           =   1560
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "Diferença de Combustível:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   93
         Top             =   4320
         Width           =   1890
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Diferença de Recebimentos:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   91
         Top             =   3720
         Width           =   2025
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Diferença de Caixa:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   89
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Compras / Vendas:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   87
         Top             =   2520
         Width           =   1365
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Produtos Comprados/Extrato:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   85
         Top             =   1920
         Width           =   2085
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cheques p/ Cliente:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   83
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Acerto de Estoque:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   81
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Transferência:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   79
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Conciliação:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   77
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cheques p/ Data:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   75
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Enviar p/ Empresa de Cobrança:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   3120
         Width           =   2325
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Protesto de Cheques:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   71
         Top             =   2520
         Width           =   1530
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Cobrança de Cheques:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   69
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Devolução de Cheques:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   67
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Depósito de Cheques:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   65
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Vales de Funcionários:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   63
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Controle de Lavagem:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   61
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Controle de Luz:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   59
         Top             =   6120
         Width           =   1155
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Controle de água:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   57
         Top             =   5520
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Cobrança de Clientes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   55
         Top             =   4920
         Width           =   1560
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Contas a pagar:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   53
         Top             =   4320
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Lançamento de Contas a pagar:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   51
         Top             =   3720
         Width           =   2280
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Lançamento de Notas:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   49
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Pagamentos Antecipados:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   47
         Top             =   2520
         Width           =   1860
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cartões Pendentes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   45
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Conferência de Caixa"
         Height          =   195
         Left            =   -74760
         TabIndex        =   43
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Fechamento Diário:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Configuração:"
         Height          =   195
         Left            =   2880
         TabIndex        =   39
         Top             =   3720
         Width           =   990
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Turnos:"
         Height          =   195
         Left            =   2880
         TabIndex        =   37
         Top             =   3120
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tanques:"
         Height          =   195
         Left            =   2880
         TabIndex        =   35
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Produtos / Fornecedores:"
         Height          =   195
         Left            =   2880
         TabIndex        =   33
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Produtos:"
         Height          =   195
         Left            =   2880
         TabIndex        =   31
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Postos:"
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Juros:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   6120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Funcionários:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   5520
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedores:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Despesas Bancárias / Tipo:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1980
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Despesas / Tipo"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Contas:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Clientes de Cheque:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clientes:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bombas de Combustível:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1770
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Novo Grupo:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Grupos:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   555
   End
End
Attribute VB_Name = "frmUsuariosGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Grupos As Permite, Alterado As Boolean

Private Sub Limpar()

'Cadastro
cboCadBombasDeCombustivel.ListIndex = 0
cboCadClientes.ListIndex = 0
cboCadClientesCheques.ListIndex = 0
cboCadContas.ListIndex = 0
cboCadDespesasTipo.ListIndex = 0
cboCadDespesasBancarias.ListIndex = 0
cboCadFormaDePagamento.ListIndex = 0
cboCadFornecedores.ListIndex = 0
cboCadFuncionarios.ListIndex = 0
cboCadJuros.ListIndex = 0
cboCadPostos.ListIndex = 0
cboCadProdutos.ListIndex = 0
cboCadProdutosFornecedores.ListIndex = 0
cboCadTanques.ListIndex = 0
cboCadTurnos.ListIndex = 0
cboCadConfiguracao.ListIndex = 0
txtClientesPlanos.Text = ""
'Controle
cboCtrlFechamento.ListIndex = 0
cboCtrlConferencia.ListIndex = 0
cboCtrlCartoesPendentes.ListIndex = 0
cboCtrlPagamentosAntecipados.ListIndex = 0
cboCtrlLancamentoDeNotas.ListIndex = 0
cboCtrlLancamentoDeContasAPagar.ListIndex = 0
cboCtrlContasAPagar.ListIndex = 0
cboCtrlCobrancaDeClientes.ListIndex = 0
cboCtrlAgua.ListIndex = 0
cboCtrlLuz.ListIndex = 0
cboCtrlLavagem.ListIndex = 0
cboCtrlVales.ListIndex = 0
'Cheques
cboChkDeposito.ListIndex = 0
cboChkDevolucao.ListIndex = 0
cboChkCobranca.ListIndex = 0
cboChkProtesto.ListIndex = 0
cboChkEmpresaCobranca.ListIndex = 0
cboChkPorData.ListIndex = 0
'Banco
cboBcoConcilia.ListIndex = 0
cboBcoTransfere.ListIndex = 0
'Relatórios
cboRelatAcertoEstoque.ListIndex = 0
cboRelatChequeCliente.ListIndex = 0
cboRelatProdutosComprados.ListIndex = 0
cboRelatComprasVendas.ListIndex = 0
cboRelatDifCaixa.ListIndex = 0
cboRelatDifRecebimentos.ListIndex = 0
cboRelatDifCombustivel.ListIndex = 0
cboRelatFormaDePagamento.ListIndex = 0
cboRelatGalonagem.ListIndex = 0
cboRelatGalonagemTotal.ListIndex = 0
cboRelatVendasDeProdutos.ListIndex = 0
cboRelatVendasDetalhada.ListIndex = 0
cboRelatVendaLucro.ListIndex = 0
cboRelatVendaMediaProdutos.ListIndex = 0
cboRelatVendaDiariaCombustivel.ListIndex = 0
cboRelatProtestoDeCheques.ListIndex = 0
cboRelatCadastroIncompleto.ListIndex = 0
cboRelatRetornoDeCombustivel.ListIndex = 0
cboRelatFatuamentoDeCheques.ListIndex = 0
cboRelatKilometragemDeClientes.ListIndex = 0
'Administração
cboAdmConfirmaDespesas.ListIndex = 0
cboAdmEstatus.ListIndex = 0
cboAdmTotalDeVendas.ListIndex = 0
cboAdmLMC.ListIndex = 0
cboAdmUsuarios.ListIndex = 0
cboAdmGruposDeUsuarios.ListIndex = 0
cboAdmDatas.ListIndex = 0
cboLiberaNotas.ListIndex = 0

With Grupos
  .Descri = ""
  'Cadastro
  .CadBomba = 0
  .CadCliente = 0
  .CadClienteCheque = 0
  .CadConta = 0
  .CadDespesaTipo = 0
  .CadDespesaBancaria = 0
  .CadFormaDePg = 0
  .CadFornecedores = 0
  .CadFuncionarios = 0
  .CadJuros = 0
  .CadPostos = 0
  .CadProdutos = 0
  .CadProdutosFornecedores = 0
  .CadTanques = 0
  .CadTurnos = 0
  .CadConfiguracao = 0
  'Controle
  .ControleFechamentoDiario = 0
  .ControleConferencia = 0
  .ControleCartoes = 0
  .ControlePgAntecipado = 0
  .ControleNotas = 0
  .ControleLancContas = 0
  .ControleContasPg = 0
  .ControleCobranca = 0
  .ControleAgua = 0
  .ControleLuz = 0
  .ControleLavagem = 0
  .ControleVales = 0
  'Cheques
  .ChequeDeposito = 0
  .ChequeDevolucao = 0
  .ChequeCobranca = 0
  .ChequeProtesto = 0
  .ChequeEnviarPEmpresaCobranca = 0
  .ChequePorData = 0
  'Banco
  .BancoConcilia = 0
  .BancoTransfere = 0
  'Relatórios
  .RelatAcertoEstoque = 0
  .RelatChequeCliente = 0
  .RelatProdutosComprados = 0
  .RelatCompraVenda = 0
  .RelatDifCaixa = 0
  .RelatDifRecebe = 0
  .RelatDifCombustivel = 0
  .RelatFormaDePg = 0
  .RelatGalonagem = 0
  .RelatGalonagemTotal = 0
  .RelatVendaProdutos = 0
  .RelatVendaDetalhada = 0
  .RelatVendaLucro = 0
  .RelatVendaMedia = 0
  .RelatDiariaCombustivel = 0
  .RelatProtestoDeCheques = 0
  .RelatCadastroIncompleto = 0
  .RelatRetornoCombustivel = 0
  .RelatFaturamentoCheques = 0
  .RelatKilometragem = 0
  'Administração
  .AdmConfirma = 0
  .AdmEstatus = 0
  .AdmTotalVenda = 0
  .AdmLMC = 0
  .AdmUsuarios = 0
  .AdmUsuariosGrupos = 0
  .admDatas = 0
End With
End Sub

Private Sub Gravar()
Grupos.Descri = txtDescri.Text

With dbGrupos
  .Recordset.Edit
  .Recordset!Descri = Grupos.Descri
  'Cadastro
  .Recordset!CadBomba = Criptografa(cboCadBombasDeCombustivel.ListIndex, 1)
  .Recordset!CadCliente = Criptografa(cboCadClientes.ListIndex, 2)
  .Recordset!CadClienteCheque = Criptografa(cboCadClientesCheques.ListIndex, 3)
  .Recordset!CadConta = Criptografa(cboCadContas.ListIndex, 4)
  .Recordset!CadDespesaTipo = Criptografa(cboCadDespesasTipo.ListIndex, 5)
  .Recordset!CadDespesaBancaria = Criptografa(cboCadDespesasBancarias.ListIndex, 6)
  .Recordset!CadFormaDePg = Criptografa(cboCadFormaDePagamento.ListIndex, 7)
  .Recordset!CadFornecedores = Criptografa(cboCadFornecedores.ListIndex, 8)
  .Recordset!CadFuncionarios = Criptografa(cboCadFuncionarios.ListIndex, 9)
  .Recordset!CadJuros = Criptografa(cboCadJuros.ListIndex, 10)
  .Recordset!CadPostos = Criptografa(cboCadPostos.ListIndex, 11)
  .Recordset!CadProdutos = Criptografa(cboCadProdutos.ListIndex, 12)
  .Recordset!CadProdutosFornecedores = Criptografa(cboCadProdutosFornecedores.ListIndex, 13)
  .Recordset!CadTanques = Criptografa(cboCadTanques.ListIndex, 14)
  .Recordset!CadTurnos = Criptografa(cboCadTurnos.ListIndex, 15)
  .Recordset!CadConfiguracao = Criptografa(cboCadConfiguracao.ListIndex, 16)
  If txtClientesPlanos.Text = "" Then
    .Recordset!ClientesPlanos = ""
  Else
    .Recordset!ClientesPlanos = Criptografa(txtClientesPlanos.Text, 16)
  End If
  'Controle
  .Recordset!ControleFechamentoDiario = Criptografa(cboCtrlFechamento.ListIndex, 17)
  .Recordset!ControleConferencia = Criptografa(cboCtrlConferencia.ListIndex, 18)
  .Recordset!ControleCartoes = Criptografa(cboCtrlCartoesPendentes.ListIndex, 19)
  .Recordset!ControlePgAntecipado = Criptografa(cboCtrlPagamentosAntecipados.ListIndex, 20)
  .Recordset!ControleNotas = Criptografa(cboCtrlLancamentoDeNotas.ListIndex, 21)
  .Recordset!ControleLancContas = Criptografa(cboCtrlLancamentoDeContasAPagar.ListIndex, 22)
  .Recordset!ControleContasPg = Criptografa(cboCtrlContasAPagar.ListIndex, 23)
  .Recordset!ControleCobranca = Criptografa(cboCtrlCobrancaDeClientes.ListIndex, 24)
  .Recordset!ControleAgua = Criptografa(cboCtrlAgua.ListIndex, 25)
  .Recordset!ControleLuz = Criptografa(cboCtrlLuz.ListIndex, 26)
  .Recordset!ControleLavagem = Criptografa(cboCtrlLavagem.ListIndex, 27)
  .Recordset!ControleVales = Criptografa(cboCtrlVales.ListIndex, 28)
  'Cheques
  .Recordset!ChequeDeposito = Criptografa(cboChkDeposito.ListIndex, 29)
  .Recordset!ChequeDevolucao = Criptografa(cboChkDevolucao.ListIndex, 30)
  .Recordset!ChequeCobranca = Criptografa(cboChkCobranca.ListIndex, 31)
  .Recordset!ChequeProtesto = Criptografa(cboChkProtesto.ListIndex, 32)
  .Recordset!ChequeEnviarPEmpresaCobranca = Criptografa(cboChkEmpresaCobranca.ListIndex, 33)
  .Recordset!ChequePorData = Criptografa(cboChkPorData.ListIndex, 34)
  'Banco
  .Recordset!BancoConcilia = Criptografa(cboBcoConcilia.ListIndex, 35)
  .Recordset!BancoTransfere = Criptografa(cboBcoTransfere.ListIndex, 36)
  'Relatórios
  .Recordset!RelatAcertoEstoque = Criptografa(cboRelatAcertoEstoque.ListIndex, 37)
  .Recordset!RelatChequeCliente = Criptografa(cboRelatChequeCliente.ListIndex, 38)
  .Recordset!RelatProdutosComprados = Criptografa(cboRelatProdutosComprados.ListIndex, 40)
  .Recordset!RelatCompraVenda = Criptografa(cboRelatComprasVendas.ListIndex, 41)
  .Recordset!RelatDifCaixa = Criptografa(cboRelatDifCaixa.ListIndex, 42)
  .Recordset!RelatDifRecebe = Criptografa(cboRelatDifRecebimentos.ListIndex, 43)
  .Recordset!RelatDifCombustivel = Criptografa(cboRelatDifCombustivel.ListIndex, 44)
  .Recordset!RelatFormaDePg = Criptografa(cboRelatFormaDePagamento.ListIndex, 45)
  .Recordset!RelatGalonagem = Criptografa(cboRelatGalonagem.ListIndex, 46)
  .Recordset!RelatGalonagemTotal = Criptografa(cboRelatGalonagemTotal.ListIndex, 47)
  .Recordset!RelatVendaProdutos = Criptografa(cboRelatVendasDeProdutos.ListIndex, 48)
  .Recordset!RelatVendaDetalhada = Criptografa(cboRelatVendasDetalhada.ListIndex, 49)
  .Recordset!RelatVendaLucro = Criptografa(cboRelatVendaLucro.ListIndex, 50)
  .Recordset!RelatVendaMedia = Criptografa(cboRelatVendaMediaProdutos.ListIndex, 51)
  .Recordset!RelatDiariaCombustivel = Criptografa(cboRelatVendaDiariaCombustivel.ListIndex, 52)
  .Recordset!RelatProtestoDeCheques = Criptografa(cboRelatProtestoDeCheques.ListIndex, 53)
  .Recordset!RelatCadastroIncompleto = Criptografa(cboRelatCadastroIncompleto.ListIndex, 54)
  .Recordset!RelatRetornoCombustivel = Criptografa(cboRelatRetornoDeCombustivel.ListIndex, 55)
  .Recordset!RelatFaturamentoCheques = Criptografa(cboRelatFatuamentoDeCheques.ListIndex, 56)
  .Recordset!RelatKilometragem = Criptografa(cboRelatKilometragemDeClientes.ListIndex, 57)
  'Administração
  .Recordset!AdmConfirma = Criptografa(cboAdmConfirmaDespesas.ListIndex, 58)
  .Recordset!AdmEstatus = Criptografa(cboAdmEstatus.ListIndex, 59)
  .Recordset!AdmTotalVenda = Criptografa(cboAdmTotalDeVendas.ListIndex, 60)
  .Recordset!AdmLMC = Criptografa(cboAdmLMC.ListIndex, 61)
  .Recordset!AdmUsuarios = Criptografa(cboAdmUsuarios.ListIndex, 62)
  .Recordset!AdmUsuariosGrupos = Criptografa(cboAdmGruposDeUsuarios.ListIndex, 63)
  .Recordset!admDatas = Criptografa(cboAdmDatas.ListIndex, 64)
  .Recordset!liberanotas = Criptografa(cboLiberaNotas.ListIndex, 65)
  .Recordset.Update
End With
End Sub

Private Sub Ler()
With dbGrupos
  If .Recordset.EOF = True Then
    With Grupos
      .Descri = ""
      'Cadastro
      .CadBomba = 0
      .CadCliente = 0
      .CadClienteCheque = 0
      .CadConta = 0
      .CadDespesaTipo = 0
      .CadDespesaBancaria = 0
      .CadFormaDePg = 0
      .CadFornecedores = 0
      .CadFuncionarios = 0
      .CadJuros = 0
      .CadPostos = 0
      .CadProdutos = 0
      .CadProdutosFornecedores = 0
      .CadTanques = 0
      .CadTurnos = 0
      .CadConfiguracao = 0
      .ClientesPlanos = ""
      'Controle
      .ControleFechamentoDiario = 0
      .ControleConferencia = 0
      .ControleCartoes = 0
      .ControlePgAntecipado = 0
      .ControleNotas = 0
      .ControleLancContas = 0
      .ControleContasPg = 0
      .ControleCobranca = 0
      .ControleAgua = 0
      .ControleLuz = 0
      .ControleLavagem = 0
      .ControleVales = 0
      'Cheques
      .ChequeDeposito = 0
      .ChequeDevolucao = 0
      .ChequeCobranca = 0
      .ChequeProtesto = 0
      .ChequeEnviarPEmpresaCobranca = 0
      .ChequePorData = 0
      'Banco
      .BancoConcilia = 0
      .BancoTransfere = 0
      'Relatórios
      .RelatAcertoEstoque = 0
      .RelatChequeCliente = 0
      .RelatProdutosComprados = 0
      .RelatCompraVenda = 0
      .RelatDifCaixa = 0
      .RelatDifRecebe = 0
      .RelatDifCombustivel = 0
      .RelatFormaDePg = 0
      .RelatGalonagem = 0
      .RelatGalonagemTotal = 0
      .RelatVendaProdutos = 0
      .RelatVendaDetalhada = 0
      .RelatVendaLucro = 0
      .RelatVendaMedia = 0
      .RelatDiariaCombustivel = 0
      .RelatProtestoDeCheques = 0
      .RelatCadastroIncompleto = 0
      .RelatRetornoCombustivel = 0
      .RelatFaturamentoCheques = 0
      .RelatKilometragem = 0
      'Administração
      .AdmConfirma = 0
      .AdmEstatus = 0
      .AdmTotalVenda = 0
      .AdmLMC = 0
      .AdmUsuarios = 0
      .AdmUsuariosGrupos = 0
      .admDatas = 0
      .admLiberaNotas = 0
    End With
  Else
    If IsNull(.Recordset!Descri) = True Then Exit Sub
    Grupos.Descri = .Recordset!Descri
    txtDescri = .Recordset!Descri
    'Cadastro
    Grupos.CadBomba = Criptografa(.Recordset!CadBomba, 1)
    Grupos.CadCliente = Criptografa(.Recordset!CadCliente, 2)
    Grupos.CadClienteCheque = Criptografa(.Recordset!CadClienteCheque, 3)
    Grupos.CadConta = Criptografa(.Recordset!CadConta, 4)
    Grupos.CadDespesaTipo = Criptografa(.Recordset!CadDespesaTipo, 5)
    Grupos.CadDespesaBancaria = Criptografa(.Recordset!CadDespesaBancaria, 6)
    Grupos.CadFormaDePg = Criptografa(.Recordset!CadFormaDePg, 7)
    Grupos.CadFornecedores = Criptografa(.Recordset!CadFornecedores, 8)
    Grupos.CadFuncionarios = Criptografa(.Recordset!CadFuncionarios, 9)
    Grupos.CadJuros = Criptografa(.Recordset!CadJuros, 10)
    Grupos.CadPostos = Criptografa(.Recordset!CadPostos, 11)
    Grupos.CadProdutos = Criptografa(.Recordset!CadProdutos, 12)
    Grupos.CadProdutosFornecedores = Criptografa(.Recordset!CadProdutosFornecedores, 13)
    Grupos.CadTanques = Criptografa(.Recordset!CadTanques, 14)
    Grupos.CadTurnos = Criptografa(.Recordset!CadTurnos, 15)
    Grupos.CadConfiguracao = Criptografa(.Recordset!CadConfiguracao, 16)
    If IsNull(.Recordset!ClientesPlanos) = False Then
      Grupos.ClientesPlanos = Criptografa(.Recordset!ClientesPlanos, 16)
    End If
    'Controle
    Grupos.ControleFechamentoDiario = Criptografa(.Recordset!ControleFechamentoDiario, 17)
    Grupos.ControleConferencia = Criptografa(.Recordset!ControleConferencia, 18)
    Grupos.ControleCartoes = Criptografa(.Recordset!ControleCartoes, 19)
    Grupos.ControlePgAntecipado = Criptografa(.Recordset!ControlePgAntecipado, 20)
    Grupos.ControleNotas = Criptografa(.Recordset!ControleNotas, 21)
    Grupos.ControleLancContas = Criptografa(.Recordset!ControleLancContas, 22)
    Grupos.ControleContasPg = Criptografa(.Recordset!ControleContasPg, 23)
    Grupos.ControleCobranca = Criptografa(.Recordset!ControleCobranca, 24)
    Grupos.ControleAgua = Criptografa(.Recordset!ControleAgua, 25)
    Grupos.ControleLuz = Criptografa(.Recordset!ControleLuz, 26)
    Grupos.ControleLavagem = Criptografa(.Recordset!ControleLavagem, 27)
    Grupos.ControleVales = Criptografa(.Recordset!ControleVales, 28)
    'Cheques
    Grupos.ChequeDeposito = Criptografa(.Recordset!ChequeDeposito, 29)
    Grupos.ChequeDevolucao = Criptografa(.Recordset!ChequeDevolucao, 30)
    Grupos.ChequeCobranca = Criptografa(.Recordset!ChequeCobranca, 31)
    Grupos.ChequeProtesto = Criptografa(.Recordset!ChequeProtesto, 32)
    Grupos.ChequeEnviarPEmpresaCobranca = Criptografa(.Recordset!ChequeEnviarPEmpresaCobranca, 33)
    Grupos.ChequePorData = Criptografa(.Recordset!ChequePorData, 34)
    'Banco
    Grupos.BancoConcilia = Criptografa(.Recordset!BancoConcilia, 35)
    Grupos.BancoTransfere = Criptografa(.Recordset!BancoTransfere, 36)
    'Relatórios
    Grupos.RelatAcertoEstoque = Criptografa(.Recordset!RelatAcertoEstoque, 37)
    Grupos.RelatChequeCliente = Criptografa(.Recordset!RelatChequeCliente, 38)
    Grupos.RelatProdutosComprados = Criptografa(.Recordset!RelatProdutosComprados, 40)
    Grupos.RelatCompraVenda = Criptografa(.Recordset!RelatCompraVenda, 41)
    Grupos.RelatDifCaixa = Criptografa(.Recordset!RelatDifCaixa, 42)
    Grupos.RelatDifRecebe = Criptografa(.Recordset!RelatDifRecebe, 43)
    Grupos.RelatDifCombustivel = Criptografa(.Recordset!RelatDifCombustivel, 44)
    Grupos.RelatFormaDePg = Criptografa(.Recordset!RelatFormaDePg, 45)
    Grupos.RelatGalonagem = Criptografa(.Recordset!RelatGalonagem, 46)
    Grupos.RelatGalonagemTotal = Criptografa(.Recordset!RelatGalonagemTotal, 47)
    Grupos.RelatVendaProdutos = Criptografa(.Recordset!RelatVendaProdutos, 48)
    Grupos.RelatVendaDetalhada = Criptografa(.Recordset!RelatVendaDetalhada, 49)
    Grupos.RelatVendaLucro = Criptografa(.Recordset!RelatVendaLucro, 50)
    Grupos.RelatVendaMedia = Criptografa(.Recordset!RelatVendaMedia, 51)
    Grupos.RelatDiariaCombustivel = Criptografa(.Recordset!RelatDiariaCombustivel, 52)
    Grupos.RelatProtestoDeCheques = Criptografa(.Recordset!RelatProtestoDeCheques, 53)
    Grupos.RelatCadastroIncompleto = Criptografa(.Recordset!RelatCadastroIncompleto, 54)
    Grupos.RelatRetornoCombustivel = Criptografa(.Recordset!RelatRetornoCombustivel, 55)
    Grupos.RelatFaturamentoCheques = Criptografa(.Recordset!RelatFaturamentoCheques, 56)
    Grupos.RelatKilometragem = Criptografa(.Recordset!RelatKilometragem, 57)
    'Administração
    Grupos.AdmConfirma = Criptografa(.Recordset!AdmConfirma, 58)
    Grupos.AdmEstatus = Criptografa(.Recordset!AdmEstatus, 59)
    Grupos.AdmTotalVenda = Criptografa(.Recordset!AdmTotalVenda, 60)
    Grupos.AdmLMC = Criptografa(.Recordset!AdmLMC, 61)
    Grupos.AdmUsuarios = Criptografa(.Recordset!AdmUsuarios, 62)
    Grupos.AdmUsuariosGrupos = Criptografa(.Recordset!AdmUsuariosGrupos, 63)
    If IsNull(.Recordset!admDatas) = False Then
      Grupos.admDatas = Criptografa(.Recordset!admDatas, 64)
    Else
      Grupos.admDatas = Grupos.AdmEstatus
    End If
    If IsNull(.Recordset!liberanotas) = False Then
      Grupos.admLiberaNotas = Criptografa(.Recordset!liberanotas, 65)
    Else
      Grupos.admLiberaNotas = Grupos.AdmEstatus
    End If
  End If
End With
With Grupos
  txtDescri.Text = .Descri
  'Cadastro
  cboCadBombasDeCombustivel.ListIndex = .CadBomba
  cboCadClientes.ListIndex = .CadCliente
  cboCadClientesCheques.ListIndex = .CadClienteCheque
  cboCadContas.ListIndex = .CadConta
  cboCadDespesasTipo.ListIndex = .CadDespesaTipo
  cboCadDespesasBancarias.ListIndex = .CadDespesaBancaria
  cboCadFormaDePagamento.ListIndex = .CadFormaDePg
  cboCadFornecedores.ListIndex = .CadFornecedores
  cboCadFuncionarios.ListIndex = .CadFuncionarios
  cboCadJuros.ListIndex = .CadJuros
  cboCadPostos.ListIndex = .CadPostos
  cboCadProdutos.ListIndex = .CadProdutos
  cboCadProdutosFornecedores.ListIndex = .CadProdutosFornecedores
  cboCadTanques.ListIndex = .CadTanques
  cboCadTurnos.ListIndex = .CadTurnos
  cboCadConfiguracao.ListIndex = .CadConfiguracao
  txtClientesPlanos.Text = Grupos.ClientesPlanos
  'Controle
  cboCtrlFechamento.ListIndex = .ControleFechamentoDiario
  cboCtrlConferencia.ListIndex = .ControleConferencia
  cboCtrlCartoesPendentes.ListIndex = .ControleCartoes
  cboCtrlPagamentosAntecipados.ListIndex = .ControlePgAntecipado
  cboCtrlLancamentoDeNotas.ListIndex = .ControleNotas
  cboCtrlLancamentoDeContasAPagar.ListIndex = .ControleLancContas
  cboCtrlContasAPagar.ListIndex = .ControleContasPg
  cboCtrlCobrancaDeClientes.ListIndex = .ControleCobranca
  cboCtrlAgua.ListIndex = .ControleAgua
  cboCtrlLuz.ListIndex = .ControleLuz
  cboCtrlLavagem.ListIndex = .ControleLavagem
  cboCtrlVales.ListIndex = .ControleVales
  'Cheques
  cboChkDeposito.ListIndex = .ChequeDeposito
  cboChkDevolucao.ListIndex = .ChequeDevolucao
  cboChkCobranca.ListIndex = .ChequeCobranca
  cboChkProtesto.ListIndex = .ChequeProtesto
  cboChkEmpresaCobranca.ListIndex = .ChequeEnviarPEmpresaCobranca
  cboChkPorData.ListIndex = .ChequePorData
  'Banco
  cboBcoConcilia.ListIndex = .BancoConcilia
  cboBcoTransfere.ListIndex = .BancoTransfere
  'Relatórios
  cboRelatAcertoEstoque.ListIndex = .RelatAcertoEstoque
  cboRelatChequeCliente.ListIndex = .RelatChequeCliente
  cboRelatProdutosComprados.ListIndex = .RelatProdutosComprados
  cboRelatComprasVendas.ListIndex = .RelatCompraVenda
  cboRelatDifCaixa.ListIndex = .RelatDifCaixa
  cboRelatDifRecebimentos.ListIndex = .RelatDifRecebe
  cboRelatDifCombustivel.ListIndex = .RelatDifCombustivel
  cboRelatFormaDePagamento.ListIndex = .RelatFormaDePg
  cboRelatGalonagem.ListIndex = .RelatGalonagem
  cboRelatGalonagemTotal.ListIndex = .RelatGalonagemTotal
  cboRelatVendasDeProdutos.ListIndex = .RelatVendaProdutos
  cboRelatVendasDetalhada.ListIndex = .RelatVendaDetalhada
  cboRelatVendaLucro.ListIndex = .RelatVendaLucro
  cboRelatVendaMediaProdutos.ListIndex = .RelatVendaMedia
  cboRelatVendaDiariaCombustivel.ListIndex = .RelatDiariaCombustivel
  cboRelatProtestoDeCheques.ListIndex = .RelatProtestoDeCheques
  cboRelatCadastroIncompleto.ListIndex = .RelatCadastroIncompleto
  cboRelatRetornoDeCombustivel.ListIndex = .RelatRetornoCombustivel
  cboRelatFatuamentoDeCheques.ListIndex = .RelatFaturamentoCheques
  cboRelatKilometragemDeClientes.ListIndex = .RelatKilometragem
  'Administração
  cboAdmConfirmaDespesas.ListIndex = .AdmConfirma
  cboAdmEstatus.ListIndex = .AdmEstatus
  cboAdmTotalDeVendas.ListIndex = .AdmTotalVenda
  cboAdmLMC.ListIndex = .AdmLMC
  cboAdmUsuarios.ListIndex = .AdmUsuarios
  cboAdmGruposDeUsuarios.ListIndex = .AdmUsuariosGrupos
  cboAdmDatas.ListIndex = .admDatas
  cboLiberaNotas.ListIndex = .admLiberaNotas
End With
End Sub

Private Sub chkCadBomba_Click()

End Sub

Private Sub cmdAlterar_Click()
Dim StrTemp As String

If txtDescri.Text = "" Then
  MsgBox "Indique uma descrição para o grupo atual!"
  txtDescri.SetFocus
  Exit Sub
End If
StrTemp = txtDescri.Text
With dbGrupos
  .Recordset.Edit
  .Recordset!Descri = StrTemp
  .Recordset.Update
  .Refresh
  lstGrupos.Refresh
  .Recordset.FindFirst "descri='" & StrTemp & "'"
End With
Ler
End Sub

Private Sub cmdGravar_Click()
Gravar
End Sub

Private Sub cmdIncluir_Click()
Dim StrTemp As String

If txtDescri.Text = "" Then
  MsgBox "Indique uma descrição para o grupo atual!"
  txtDescri.SetFocus
  Exit Sub
End If
StrTemp = txtDescri.Text
Limpar
With dbGrupos
  .Refresh
  .Recordset.AddNew
  .Recordset!Descri = StrTemp
  .Recordset!CadBomba = Criptografa("0", 1)
  .Recordset!CadCliente = Criptografa("0", 2)
  .Recordset!CadClienteCheque = Criptografa("0", 3)
  .Recordset!CadConta = Criptografa("0", 4)
  .Recordset!CadDespesaTipo = Criptografa("0", 5)
  .Recordset!CadDespesaBancaria = Criptografa("0", 6)
  .Recordset!CadFormaDePg = Criptografa("0", 7)
  .Recordset!CadFornecedores = Criptografa("0", 8)
  .Recordset!CadFuncionarios = Criptografa("0", 9)
  .Recordset!CadJuros = Criptografa("0", 10)
  .Recordset!CadPostos = Criptografa("0", 11)
  .Recordset!CadProdutos = Criptografa("0", 12)
  .Recordset!CadProdutosFornecedores = Criptografa("0", 13)
  .Recordset!CadTanques = Criptografa("0", 14)
  .Recordset!CadTurnos = Criptografa("0", 15)
  .Recordset!CadConfiguracao = Criptografa("0", 16)
  'Controle
  .Recordset!ControleFechamentoDiario = Criptografa("0", 17)
  .Recordset!ControleConferencia = Criptografa("0", 18)
  .Recordset!ControleCartoes = Criptografa("0", 19)
  .Recordset!ControlePgAntecipado = Criptografa("0", 20)
  .Recordset!ControleNotas = Criptografa("0", 21)
  .Recordset!ControleLancContas = Criptografa("0", 22)
  .Recordset!ControleContasPg = Criptografa("0", 23)
  .Recordset!ControleCobranca = Criptografa("0", 24)
  .Recordset!ControleAgua = Criptografa("0", 25)
  .Recordset!ControleLuz = Criptografa("0", 26)
  .Recordset!ControleLavagem = Criptografa("0", 27)
  .Recordset!ControleVales = Criptografa("0", 28)
  'Cheques
  .Recordset!ChequeDeposito = Criptografa("0", 29)
  .Recordset!ChequeDevolucao = Criptografa("0", 30)
  .Recordset!ChequeCobranca = Criptografa("0", 31)
  .Recordset!ChequeProtesto = Criptografa("0", 32)
  .Recordset!ChequeEnviarPEmpresaCobranca = Criptografa("0", 33)
  .Recordset!ChequePorData = Criptografa("0", 34)
  'Banco
  .Recordset!BancoConcilia = Criptografa("0", 35)
  .Recordset!BancoTransfere = Criptografa("0", 36)
  'Relatórios
  .Recordset!RelatAcertoEstoque = Criptografa("0", 37)
  .Recordset!RelatChequeCliente = Criptografa("0", 38)
  .Recordset!RelatProdutosComprados = Criptografa("0", 40)
  .Recordset!RelatCompraVenda = Criptografa("0", 41)
  .Recordset!RelatDifCaixa = Criptografa("0", 42)
  .Recordset!RelatDifRecebe = Criptografa("0", 43)
  .Recordset!RelatDifCombustivel = Criptografa("0", 44)
  .Recordset!RelatFormaDePg = Criptografa("0", 45)
  .Recordset!RelatGalonagem = Criptografa("0", 46)
  .Recordset!RelatGalonagemTotal = Criptografa("0", 47)
  .Recordset!RelatVendaProdutos = Criptografa("0", 48)
  .Recordset!RelatVendaDetalhada = Criptografa("0", 49)
  .Recordset!RelatVendaLucro = Criptografa("0", 50)
  .Recordset!RelatVendaMedia = Criptografa("0", 51)
  .Recordset!RelatDiariaCombustivel = Criptografa("0", 52)
  .Recordset!RelatProtestoDeCheques = Criptografa("0", 53)
  .Recordset!RelatCadastroIncompleto = Criptografa("0", 54)
  .Recordset!RelatRetornoCombustivel = Criptografa("0", 55)
  .Recordset!RelatFaturamentoCheques = Criptografa("0", 56)
  .Recordset!RelatKilometragem = Criptografa("0", 57)
  'Administração
  .Recordset!AdmConfirma = Criptografa("0", 58)
  .Recordset!AdmEstatus = Criptografa("0", 59)
  .Recordset!AdmTotalVenda = Criptografa("0", 60)
  .Recordset!AdmLMC = Criptografa("0", 61)
  .Recordset!AdmUsuarios = Criptografa("0", 62)
  .Recordset!AdmUsuariosGrupos = Criptografa("0", 63)
  .Recordset.Update
  .Refresh
  lstGrupos.Refresh
  .Recordset.FindFirst "descri='" & StrTemp & "'"
End With
Ler

End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With dbUsuarios
  .Refresh
  If .Recordset.EOF = False Then
    .Recordset.FindFirst "codigogrupo=" & dbGrupos.Recordset!CodigoGrupo
    If .Recordset.NoMatch = False Then
      MsgBox "Não é possível excluir o registro atual. Existe usuario cadastrado neste grupo."
      Exit Sub
    End If
  End If
  dbGrupos.Recordset.Delete
  dbGrupos.Refresh
  Ler
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Ws As Workspace, db As Database, dbTemp As Recordset, dbClientesPlanos As Recordset
Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from UsuariosGrupos order by ClientesPlanos")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table UsuariosGrupos add column ClientesPlanos text(200)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela UsuariosGrupos->ClientesPlanos!"
  End If
End If
On Error GoTo 0

Set dbTemp = db.OpenRecordset("select *from UsuariosGrupos order by LiberaNotas")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table UsuariosGrupos add column LiberaNotas text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela UsuariosGrupos->LiberaNotas!"
  End If
End If
On Error GoTo 0

On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from UsuariosGrupos order by admDatas")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table UsuariosGrupos add column admDatas text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela UsuariosGrupos->admDatas!"
  End If
End If
On Error GoTo 0
With dbGrupos
  .Connect = Conectar
  .DatabaseName = CaminhoUsuarios
  .RecordSource = "select *from usuariosgrupos order by AdmLMC"
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
    On Error GoTo 0
    On Error Resume Next
    db.Execute "drop table usuariosgrupos"
    If Err.Number <> 0 And Err.Number <> 3010 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela de grupos de usuarios!"
      End
    End If
    On Error GoTo 0
    On Error Resume Next
    db.Execute "create table UsuariosGrupos (CodigoGrupo counter, Descri Text(50), CadBomba Text(50), CadCliente Text(50), CadClienteCheque Text(50), CadConta Text(50), CadDespesaTipo Text(50), CadDespesaBancaria Text(50)," & _
               "CadFormaDePg Text(50), CadFornecedores Text(50), CadFuncionarios Text(50), CadJuros Text(50), CadPostos Text(50), CadProdutos Text(50), CadProdutosFornecedores Text(50), CadTanques Text(50), CadTurnos Text(50)," & _
               "CadConfiguracao Text(50), ControleFechamentoDiario Text(50), ControleConferencia Text(50), ControleCartoes Text(50), ControlePgAntecipado Text(50), ControleNotas Text(50), ControleLancContas Text(50)," & _
               "ControleContasPg Text(50), ControleCobranca Text(50), ControleAgua Text(50), ControleLuz Text(50), ControleLavagem Text(50), ControleVales Text(50), ChequeDeposito Text(50), ChequeDevolucao Text(50), ChequeCobranca Text(50)," & _
               "ChequeProtesto Text(50), ChequeEnviarPEmpresaCobranca Text(50), ChequePorData Text(50), BancoConcilia Text(50), BancoTransfere Text(50), RelatAcertoEstoque Text(50), RelatChequeCliente Text(50)," & _
               "RelatProdutosComprados Text(50), RelatCompraVenda Text(50), RelatDifCaixa Text(50), RelatDifRecebe Text(50), RelatDifCombustivel Text(50), RelatFormaDePg Text(50), RelatGalonagem Text(50), RelatGalonagemTotal Text(50)," & _
               "RelatVendaProdutos Text(50), RelatVendaDetalhada Text(50), RelatVendaLucro Text(50), RelatVendaMedia Text(50), RelatDiariaCombustivel Text(50), RelatProtestoDeCheques Text(50), RelatCadastroIncompleto Text(50)," & _
               "RelatRetornoCombustivel Text(50), RelatFaturamentoCheques Text(50), RelatKilometragem Text(50), AdmConfirma Text(50), AdmEstatus Text(50), AdmTotalVenda Text(50), AdmLMC Text(50), AdmUsuarios Text(50), AdmUsuariosGrupos Text(50))"
    If Err.Number <> 0 And Err.Number <> 3010 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela UsuariosGrupos!"
      End
    End If
  End If
  On Error GoTo 0
  On Error Resume Next
  .RecordSource = "select *from usuariosgrupos"
  .Refresh
  If Err.Number <> 0 Then
    For i = 0 To 15
      On Error GoTo 0
      On Error Resume Next
      .Refresh
      If Err.Number = 0 Then Exit For
    Next i
  End If
End With
With dbUsuarios
  .Connect = Conectar
  .DatabaseName = CaminhoUsuarios
  .Refresh
End With
Select Case Usuarios.Grupo.ControleCartoes
  Case 1 'Somente leitura
    cmdGravar.Enabled = False
    cmdIncluir.Enabled = False
    cmdRemover.Enabled = False
    For Each ComboBox In Me
      If TypeName(ComboBox) = "ComboBox" Then
        ComboBox.Enabled = False
      End If
    Next
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Alterado = True
End Sub

Private Sub Form_Terminate()
Dim Resposta As Integer
If Alterado = True Then
  Resposta = MsgBox("Deseja gravar as alterações?", vbYesNo)
  If Resposta = vbYes Then Gravar
End If
End Sub

Private Sub lstGrupos_Click()
With dbGrupos
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "Descri='" & lstGrupos.Text & "'"
    If .Recordset.NoMatch = False Then
      Ler
    Else
      Limpar
    End If
  End If
End With
End Sub

Private Sub txtDescri_GotFocus()
With txtDescri
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub
