VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFechamentoDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assistente para Fechamento Diário de Caixa"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   9390
   Icon            =   "frmFechamentoDiario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "DBFs DAO"
      Height          =   5535
      Left            =   2760
      TabIndex        =   178
      Top             =   6360
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Data QTemp 
         Caption         =   "QTemp"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QBicoMovimentoTotalTanque"
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Data QGalonagem 
         Caption         =   "QGalonagem"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QGalonagemProduto"
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Data dbJuros 
         Caption         =   "dbJuros"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from juros order by inicio, final"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Data dbContas 
         Caption         =   "dbContas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from contas order by descri"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Data dbStatus 
         Caption         =   "dbStatus"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from status"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Data QComissoes 
         Caption         =   "QComissoes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesCarros where codigocliente=0"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Data QProdutoEntra 
         Caption         =   "QProdutoEntra"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(valornota) as total from produtosentrada where codigofechamento=-1"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Data dbProdutoEntra 
         Caption         =   "dbProdutoEntra"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtosentrada where codigofechamento=-1"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Data dbChequesContas 
         Caption         =   "dbChequesContas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from chequescontas"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Data dbClientesCheques 
         Caption         =   "dbClientesCheques"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from chequesclientes order by nome"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Data QCheques 
         Caption         =   "QCheques"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(valor) as Total from cheques where codigofechamento=0"
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Data dbCheques 
         Caption         =   "dbCheques"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from cheques "
         Top             =   4560
         Width           =   3015
      End
      Begin VB.Data QClientesNota 
         Caption         =   "QClientesNota"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QClientesNota"
         Top             =   4200
         Width           =   3015
      End
      Begin VB.Data dbClientesNota 
         Caption         =   "dbClientesNota"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientesnota where confirmado=0"
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Data dbCarros 
         Caption         =   "dbCarros"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesCarros where codigocliente=0"
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientes where mensalista=-1 order by nome"
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Data QFormaDePgRecTotaliza 
         Caption         =   "QFormaDePgRecTotaliza"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QFormaDePagamentoRecebidoTotaliza"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Data dbFormaDePgRecebido 
         Caption         =   "dbFormaDePgRecebido"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FormaDePagamentoRecebido order by descri"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Data dbFormaDePg 
         Caption         =   "dbFormaDePg"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FormaDePagamento order by descri"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Data QDespesaLancTotaliza 
         Caption         =   "QDespesaLancTotaliza"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QDespesaLancTotaliza"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Data dbDespesasLanc 
         Caption         =   "dbDespesasLanc"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DespesasLanc order by descri"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Data dbDespesas 
         Caption         =   "dbDespesas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DespesaTipo order by descri"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Data QVendaTotaliza 
         Caption         =   "QVendaTotaliza"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QVendaTotaliza"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Data QBicoMovimentaTotal 
         Caption         =   "QBicoMovimentaTotal"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QBicoMovimentoTotaliza"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Data dbVendas 
         Caption         =   "dbVendas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from venda order by descri"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Data dbProdutos2 
         Caption         =   "dbProdutos2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtos order by descri"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtos order by descri"
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Data dbDifComb 
         Caption         =   "dbDifComb"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DiferencaCombustivel"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Data dbTanquesMovimento 
         Caption         =   "dbTanquesMovimento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from tanquesMovimento order by tanque"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Data dbTanques 
         Caption         =   "dbTanques"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from tanques order by tanque"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Data dbFechamento 
         Caption         =   "dbFechamento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FechamentoDiario order by Data, Hora"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Data dbTurno 
         Caption         =   "dbTurno"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from turnos order by descri"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Data dbResponsavel 
         Caption         =   "dbResponsavel"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from vendedores where gerente=-1 order by nome"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Data dbPosto 
         Caption         =   "dbPosto"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from postos order by nome"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Data dbBicoMovimento 
         Caption         =   "dbBicoMovimento"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from bicomovimento order by bico"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Data dbBico 
         Caption         =   "dbBico"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from bicos order by bico"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Data dbConciliaNova 
         Caption         =   "dbConciliaNova"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from conciliaNova"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Data dbCartoes 
         Caption         =   "dbCartoes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Cartoes"
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSDBCtls.DBCombo cboTurno 
      Bindings        =   "frmFechamentoDiario.frx":0442
      Height          =   315
      Left            =   4680
      TabIndex        =   55
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo cboResponsavel 
      Bindings        =   "frmFechamentoDiario.frx":0458
      Height          =   315
      Left            =   120
      TabIndex        =   51
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin TabDlg.SSTab Paginas 
      Height          =   5295
      Left            =   120
      TabIndex        =   104
      Top             =   840
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   5
      TabHeight       =   485
      TabCaption(0)   =   "Controle de Bomba"
      TabPicture(0)   =   "frmFechamentoDiario.frx":0474
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Tela(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DBGrid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Controle de Tanque"
      TabPicture(1)   =   "frmFechamentoDiario.frx":0490
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Tela(1)"
      Tab(1).Control(1)=   "DBGrid2"
      Tab(1).Control(2)=   "DBGrid3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Vendas"
      TabPicture(2)   =   "frmFechamentoDiario.frx":04AC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label33"
      Tab(2).Control(1)=   "Label35"
      Tab(2).Control(2)=   "Tela(2)"
      Tab(2).Control(3)=   "DBGrid4"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Despesas"
      TabPicture(3)   =   "frmFechamentoDiario.frx":04C8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label39"
      Tab(3).Control(1)=   "Label40"
      Tab(3).Control(2)=   "Tela(3)"
      Tab(3).Control(3)=   "DBGrid5"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Recebimentos"
      TabPicture(4)   =   "frmFechamentoDiario.frx":04E4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label41"
      Tab(4).Control(1)=   "Label42"
      Tab(4).Control(2)=   "Tela(4)"
      Tab(4).Control(3)=   "Frame4"
      Tab(4).Control(4)=   "DBGrid6"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Notas"
      TabPicture(5)   =   "frmFechamentoDiario.frx":0500
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label43"
      Tab(5).Control(1)=   "Label44"
      Tab(5).Control(2)=   "Tela(5)"
      Tab(5).Control(3)=   "DBGrid7"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Cheques"
      TabPicture(6)   =   "frmFechamentoDiario.frx":051C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image1"
      Tab(6).Control(1)=   "Label70"
      Tab(6).Control(2)=   "Label54"
      Tab(6).Control(3)=   "Label53"
      Tab(6).Control(4)=   "Tela(6)"
      Tab(6).Control(5)=   "DBGrid8"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Compra de Produtos"
      TabPicture(7)   =   "frmFechamentoDiario.frx":0538
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "DBGrid9"
      Tab(7).Control(1)=   "Tela(7)"
      Tab(7).Control(2)=   "lblProdutoEntraTotal"
      Tab(7).Control(3)=   "Label55"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Fechamento"
      TabPicture(8)   =   "frmFechamentoDiario.frx":0554
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Tela(8)"
      Tab(8).ControlCount=   1
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "frmFechamentoDiario.frx":0570
         Height          =   3135
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":058D
         TabIndex        =   98
         Top             =   1680
         Width           =   7215
      End
      Begin MSDBGrid.DBGrid DBGrid8 
         Bindings        =   "frmFechamentoDiario.frx":1480
         Height          =   2535
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":1498
         TabIndex        =   194
         Top             =   2280
         Width           =   8655
      End
      Begin MSDBGrid.DBGrid DBGrid7 
         Bindings        =   "frmFechamentoDiario.frx":2873
         Height          =   3135
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":2890
         TabIndex        =   193
         Top             =   1680
         Width           =   7095
      End
      Begin MSDBGrid.DBGrid DBGrid6 
         Bindings        =   "frmFechamentoDiario.frx":35CB
         Height          =   2295
         Left            =   -70920
         OleObjectBlob   =   "frmFechamentoDiario.frx":35ED
         TabIndex        =   192
         Top             =   2400
         Width           =   4695
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "frmFechamentoDiario.frx":4194
         Height          =   2535
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":41B1
         TabIndex        =   191
         Top             =   2280
         Width           =   7215
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "frmFechamentoDiario.frx":4D30
         Height          =   3135
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":4D47
         TabIndex        =   190
         Top             =   1680
         Width           =   8655
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "frmFechamentoDiario.frx":5FA2
         Height          =   3375
         Left            =   -72000
         OleObjectBlob   =   "frmFechamentoDiario.frx":5FBA
         TabIndex        =   189
         Top             =   1680
         Width           =   5775
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmFechamentoDiario.frx":6EA5
         Height          =   3375
         Left            =   -74880
         OleObjectBlob   =   "frmFechamentoDiario.frx":6EC6
         TabIndex        =   188
         Top             =   1680
         Width           =   2775
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmFechamentoDiario.frx":78B9
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "frmFechamentoDiario.frx":78D7
         TabIndex        =   179
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Frame Frame4 
         Caption         =   " Valores Declarados "
         Height          =   3975
         Left            =   -74880
         TabIndex        =   164
         Top             =   720
         Width           =   3735
         Begin MSMask.MaskEdBox txtDinheiro 
            DataField       =   "Dinheiro"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   180
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChequeRecebido 
            DataField       =   "ChequeAVista"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   181
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChequesPre 
            DataField       =   "ChequePre"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   182
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtCartoes 
            DataField       =   "Cartoes"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   183
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtNotas 
            DataField       =   "Notas"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   184
            Top             =   1920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtVT 
            DataField       =   "VT"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   185
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtTickets 
            DataField       =   "Tickets"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   186
            Top             =   2640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtDespesas 
            DataField       =   "Despesas"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   1800
            TabIndex        =   187
            Top             =   3000
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            Caption         =   "Tickets"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   1635
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   166
            Top             =   3360
            Width           =   1635
         End
         Begin VB.Label lblTotalDeclarado 
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
            Left            =   1800
            TabIndex        =   165
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            Caption         =   "Dinheiro:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Caption         =   "Despesas:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   3000
            Width           =   1635
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "Vale Transporte:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   1635
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            Caption         =   "Notas:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            Caption         =   "Cartões:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            Caption         =   "Cheque à Vista:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "Cheque Pré-Datado:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   1635
         End
      End
      Begin VB.Frame Tela 
         Height          =   975
         Index           =   5
         Left            =   -74880
         TabIndex        =   148
         Top             =   600
         Width           =   8775
         Begin MSDBCtls.DBCombo cboPlaca 
            Bindings        =   "frmFechamentoDiario.frx":8B52
            Height          =   315
            Left            =   3240
            TabIndex        =   60
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Placa"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cboClientesNota 
            Bindings        =   "frmFechamentoDiario.frx":8B69
            Height          =   315
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
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
            Left            =   6000
            TabIndex        =   64
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdInclueNota 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   7440
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCupom 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            TabIndex        =   62
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   6000
            TabIndex        =   63
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
            Height          =   195
            Left            =   3240
            TabIndex        =   59
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Cupom:"
            Height          =   195
            Left            =   4680
            TabIndex        =   61
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Tela 
         Height          =   1575
         Index           =   4
         Left            =   -70920
         TabIndex        =   147
         Top             =   720
         Width           =   4695
         Begin MSDBCtls.DBCombo cboRecebimento 
            Bindings        =   "frmFechamentoDiario.frx":8B82
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin VB.CommandButton cmdIncluirRecebimento 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   3840
            TabIndex        =   49
            Top             =   960
            Width           =   735
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
            TabIndex        =   46
            Top             =   1080
            Width           =   1215
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
            TabIndex        =   48
            Top             =   1080
            Width           =   855
         End
         Begin MSComCtl2.DTPicker txtDataBordero 
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52887553
            CurrentDate     =   37600
         End
         Begin VB.Label Label86 
            Caption         =   "Data Borderô:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   1560
            TabIndex        =   45
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Operações:"
            Height          =   195
            Left            =   2880
            TabIndex        =   47
            Top             =   840
            Width           =   825
         End
      End
      Begin VB.Frame Tela 
         Height          =   1575
         Index           =   3
         Left            =   -74880
         TabIndex        =   146
         Top             =   600
         Width           =   8775
         Begin MSDBCtls.DBCombo cboDespesa 
            Bindings        =   "frmFechamentoDiario.frx":8B9C
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
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
            TabIndex        =   29
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdIncluirDespesa 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   5280
            TabIndex        =   32
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtDespesaObs 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   4935
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Despesa:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   3480
            TabIndex        =   28
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   915
         End
      End
      Begin VB.Frame Tela 
         Height          =   975
         Index           =   2
         Left            =   -74880
         TabIndex        =   143
         Top             =   600
         Width           =   8775
         Begin MSDBCtls.DBCombo cboProduto 
            Bindings        =   "frmFechamentoDiario.frx":8BB5
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin VB.TextBox txtDesconto 
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
            Left            =   4920
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtProdutoQuantidade 
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
            Left            =   4200
            TabIndex        =   20
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtCodProduto 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdIncluirVendas 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   7680
            TabIndex        =   25
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCodVendedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5760
            TabIndex        =   24
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label84 
            Caption         =   "Desconto:"
            Height          =   255
            Left            =   4920
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Qtd.:"
            Height          =   195
            Left            =   4200
            TabIndex        =   19
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cod.:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Left            =   1200
            TabIndex        =   17
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   6360
            TabIndex        =   145
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblProdutoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6360
            TabIndex        =   144
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Fuc.:"
            Height          =   195
            Left            =   5760
            TabIndex        =   23
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Tela 
         Height          =   975
         Index           =   1
         Left            =   -74880
         TabIndex        =   140
         Top             =   600
         Width           =   8655
         Begin MSDBCtls.DBCombo cboTanque 
            Bindings        =   "frmFechamentoDiario.frx":8BCF
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Tanque"
            Text            =   ""
         End
         Begin VB.TextBox txtReposicao 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   4320
            TabIndex        =   141
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdIncluirTanque 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtRegua 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Reposição:"
            Height          =   195
            Left            =   4320
            TabIndex        =   142
            Top             =   240
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Régua:"
            Height          =   195
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tanque:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Tela 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   139
         Top             =   600
         Width           =   8775
         Begin MSDBCtls.DBCombo cboBico 
            Bindings        =   "frmFechamentoDiario.frx":8BE7
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Bico"
            Text            =   ""
         End
         Begin VB.TextBox txtMecanico 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmdInclueBico 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   5040
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtBicoEncerra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtRetorno 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3960
            TabIndex        =   7
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdRemover 
            Caption         =   "Remover"
            Height          =   375
            Left            =   6360
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Mecânico:"
            Height          =   195
            Left            =   1080
            TabIndex        =   2
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Eletônico:"
            Height          =   195
            Left            =   2520
            TabIndex        =   4
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bico:"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Retorno:"
            Height          =   195
            Left            =   3960
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Tela 
         Height          =   4575
         Index           =   8
         Left            =   -74880
         TabIndex        =   107
         Top             =   600
         Width           =   8775
         Begin VB.TextBox txtPontos 
            Alignment       =   1  'Right Justify
            DataField       =   "Pontos"
            DataSource      =   "dbFechamento"
            Height          =   285
            Left            =   4680
            TabIndex        =   198
            Top             =   3600
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtJuros 
            DataField       =   "Juros"
            DataSource      =   "dbFechamento"
            Height          =   300
            Left            =   1320
            TabIndex        =   195
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "#,##0.00;-#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Frame Frame2 
            Caption         =   " Resumo "
            Height          =   3255
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   3735
            Begin VB.Label lblClientesNota 
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
               Left            =   1800
               TabIndex        =   127
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Notas:"
               Height          =   195
               Left            =   1170
               TabIndex        =   126
               Top             =   1320
               Width           =   465
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Vendas Combustível:"
               Height          =   195
               Left            =   120
               TabIndex        =   125
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lblVendasCombustivel 
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
               Left            =   1800
               TabIndex        =   124
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Vendas Produtos:"
               Height          =   195
               Left            =   375
               TabIndex        =   123
               Top             =   600
               Width           =   1260
            End
            Begin VB.Label lblVendasProdutos 
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
               Left            =   1800
               TabIndex        =   122
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Despesas:"
               Height          =   195
               Left            =   885
               TabIndex        =   121
               Top             =   960
               Width           =   750
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
               Left            =   1800
               TabIndex        =   120
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Comissão:"
               Height          =   195
               Left            =   915
               TabIndex        =   119
               Top             =   2760
               Width           =   720
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
               Left            =   1800
               TabIndex        =   118
               Top             =   2760
               Width           =   1695
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
               Left            =   1800
               TabIndex        =   117
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Total Recebido:"
               Height          =   195
               Left            =   495
               TabIndex        =   116
               Top             =   2040
               Width           =   1140
            End
            Begin VB.Label lblTotalChequeResumo 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Total"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "QCheques"
               Height          =   255
               Left            =   1800
               TabIndex        =   115
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Cheques:"
               Height          =   195
               Left            =   960
               TabIndex        =   114
               Top             =   1680
               Width           =   675
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Compra de Produtos:"
               Height          =   195
               Left            =   150
               TabIndex        =   113
               Top             =   2400
               Width           =   1485
            End
            Begin VB.Label Label63 
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
               Left            =   1800
               TabIndex        =   112
               Top             =   2400
               Width           =   1695
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Galonagem "
            Height          =   3255
            Left            =   4200
            TabIndex        =   108
            Top             =   240
            Width           =   4215
            Begin MSDBGrid.DBGrid DBGrid10 
               Bindings        =   "frmFechamentoDiario.frx":8BFC
               Height          =   2295
               Left            =   120
               OleObjectBlob   =   "frmFechamentoDiario.frx":8C15
               TabIndex        =   196
               Top             =   360
               Width           =   3975
            End
            Begin VB.Label lblGalonagem 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
               Height          =   255
               Left            =   2325
               TabIndex        =   110
               Top             =   2880
               Width           =   1695
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   1800
               TabIndex        =   109
               Top             =   2880
               Width           =   405
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Encerrante de Pontuação:"
            Height          =   195
            Left            =   2640
            TabIndex        =   197
            Top             =   3600
            Width           =   1875
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Vendas:"
            Height          =   195
            Left            =   240
            TabIndex        =   138
            Top             =   3960
            Width           =   585
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
            Left            =   240
            TabIndex        =   137
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Despesas:"
            Height          =   195
            Left            =   2040
            TabIndex        =   136
            Top             =   3960
            Width           =   750
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
            Left            =   2040
            TabIndex        =   135
            Top             =   4200
            Width           =   1695
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
            Left            =   3840
            TabIndex        =   134
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Recebimentos:"
            Height          =   195
            Left            =   3840
            TabIndex        =   133
            Top             =   3960
            Width           =   1065
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
            Left            =   5760
            TabIndex        =   132
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   5760
            TabIndex        =   131
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Juros Cobrado:"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   3600
            Width           =   1065
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Valor Pré-Datado:"
            Height          =   195
            Left            =   6960
            TabIndex        =   129
            Top             =   3960
            Width           =   1260
         End
         Begin VB.Label lblPreDatado 
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
            Left            =   6960
            TabIndex        =   128
            Top             =   4200
            Width           =   1575
         End
      End
      Begin VB.Frame Tela 
         Height          =   975
         Index           =   7
         Left            =   -74880
         TabIndex        =   106
         Top             =   600
         Width           =   8655
         Begin MSDBCtls.DBCombo cboProdutoEntra 
            Bindings        =   "frmFechamentoDiario.frx":9601
            Height          =   315
            Left            =   960
            TabIndex        =   88
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
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
            Left            =   5760
            TabIndex        =   94
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdProdutoEntra 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   7320
            TabIndex        =   97
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCod 
            Height          =   285
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   735
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
            TabIndex        =   90
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTotalEntra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   92
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtTanqueEntra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6600
            TabIndex        =   96
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "$ Unitário:"
            Height          =   195
            Left            =   5760
            TabIndex        =   93
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "$ Total:"
            Height          =   195
            Left            =   4560
            TabIndex        =   91
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Left            =   960
            TabIndex        =   87
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Cod.:"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   3600
            TabIndex        =   89
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lblTanque 
            AutoSize        =   -1  'True
            Caption         =   "Tanque:"
            Height          =   195
            Left            =   6600
            TabIndex        =   95
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Tela 
         Height          =   1575
         Index           =   6
         Left            =   -74880
         TabIndex        =   105
         Top             =   600
         Width           =   8775
         Begin MSDBCtls.DBCombo cboClienteCheque 
            Bindings        =   "frmFechamentoDiario.frx":961A
            Height          =   315
            Left            =   4560
            TabIndex        =   79
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            Text            =   ""
         End
         Begin VB.TextBox txtCodCliente 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   77
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   83
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdRelaciona 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   5040
            TabIndex        =   84
            Top             =   960
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   480
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
            Left            =   720
            TabIndex        =   69
            Top             =   480
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
            Left            =   1320
            TabIndex        =   71
            Top             =   480
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
            Left            =   2040
            TabIndex        =   73
            Top             =   480
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
            Left            =   3000
            TabIndex        =   75
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   6
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Index           =   5
            Left            =   1200
            TabIndex        =   81
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99/99/99"
            PromptChar      =   " "
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   177
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblStatus 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   176
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   4560
            TabIndex        =   78
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lblNome 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4560
            TabIndex        =   175
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label83 
            Caption         =   "Código:"
            Height          =   255
            Left            =   3840
            TabIndex        =   76
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblJurosTabelado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4200
            TabIndex        =   170
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
            Height          =   195
            Left            =   4200
            TabIndex        =   169
            Top             =   840
            Width           =   420
         End
         Begin VB.Label lblValorNaBomba 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3120
            TabIndex        =   168
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "Na Bomba:"
            Height          =   195
            Left            =   3120
            TabIndex        =   167
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Comp:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   720
            TabIndex        =   68
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
            Height          =   195
            Left            =   1320
            TabIndex        =   70
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   195
            Left            =   2040
            TabIndex        =   72
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            Height          =   195
            Left            =   3000
            TabIndex        =   74
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label47 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   2160
            TabIndex        =   82
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   1200
            TabIndex        =   80
            Top             =   840
            Width           =   390
         End
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
         Left            =   -69360
         TabIndex        =   163
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   -69840
         TabIndex        =   162
         Top             =   4920
         Width           =   405
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -69840
         TabIndex        =   161
         Top             =   4920
         Width           =   1605
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
         Left            =   -68160
         TabIndex        =   160
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label70 
         Caption         =   "Leitura Automática"
         Height          =   255
         Left            =   -74520
         TabIndex        =   159
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   -74880
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label44 
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
         Left            =   -69600
         TabIndex        =   158
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   -70080
         TabIndex        =   157
         Top             =   4920
         Width           =   405
      End
      Begin VB.Label Label42 
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
         Left            =   -67920
         TabIndex        =   156
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -69960
         TabIndex        =   155
         Top             =   4800
         Width           =   1965
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   195
         Left            =   -70440
         TabIndex        =   154
         Top             =   4920
         Width           =   1005
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
         Left            =   -69360
         TabIndex        =   153
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   -68400
         TabIndex        =   152
         Top             =   4920
         Width           =   405
      End
      Begin VB.Label Label33 
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
         Left            =   -67920
         TabIndex        =   151
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   5280
         TabIndex        =   150
         Top             =   4920
         Width           =   405
      End
      Begin VB.Label Label25 
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
         Left            =   5760
         TabIndex        =   149
         Top             =   4920
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2040
      Top             =   6120
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   102
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "&Finalizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   101
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "&Próximo >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   99
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<< &Anterior"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   100
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdInlueBomba 
      Caption         =   "&Novo"
      Height          =   375
      Left            =   6120
      TabIndex        =   56
      Top             =   240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   3240
      TabIndex        =   53
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52690945
      CurrentDate     =   37600
   End
   Begin VB.CommandButton cmdResumo 
      Caption         =   "&Resumo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   103
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblUltimoConfirmado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   174
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label82 
      AutoSize        =   -1  'True
      Caption         =   "Último Caixa Confirmado:"
      Height          =   195
      Left            =   120
      TabIndex        =   173
      Top             =   1200
      Width           =   1755
   End
   Begin VB.Label lblUltimoLancado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   172
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label81 
      AutoSize        =   -1  'True
      Caption         =   "Último Caixa Lançado:"
      Height          =   195
      Left            =   120
      TabIndex        =   171
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      Caption         =   "&Turno:"
      Height          =   195
      Left            =   4680
      TabIndex        =   54
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Responsável:"
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Data:"
      Height          =   195
      Left            =   3240
      TabIndex        =   52
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmFechamentoDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public intTela As Integer, CodigoFechamento As Double
Dim CodBar As String, Porta As Integer

Private Sub TotalizaVenda()
Dim Valor As Currency, Desconto As Currency
lblProdutoTotal.Caption = ""
With txtProdutoQuantidade
  If IsNumeric(.Text) = False Then Exit Sub
  If dbProdutos2.Recordset.EOF = True Then Exit Sub
  If dbProdutos2.Recordset("descri") <> cboProduto.Text Then Exit Sub
  Valor = dbProdutos2.Recordset("precovenda")
  Valor = Valor * CDbl(.Text)
  Desconto = 0
  If IsNumeric(txtDesconto.Text) = True Then
    Desconto = CCur(txtDesconto.Text)
  End If
  Valor = Valor - Desconto
  lblProdutoTotal.Caption = Format(Valor, "currency")
End With
End Sub
Private Sub SomaDeclarados()
Dim TotalDeclarado As Currency
With dbFechamento
  'Soma os valores declarados
  TotalDeclarado = 0
  If IsNumeric(txtDinheiro.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtDinheiro.Text)
  End If
  If IsNumeric(txtCartoes.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtCartoes.Text)
  End If
  If IsNumeric(txtVT.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtVT.Text)
  End If
  If IsNumeric(txtDespesas.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtDespesas.Text)
  End If
  If IsNumeric(txtNotas.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtNotas.Text)
  End If
  If IsNumeric(txtChequeRecebido.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtChequeRecebido.Text)
  End If
  If IsNumeric(txtChequesPre.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtChequesPre.Text)
  End If
  If IsNumeric(txtTickets.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtTickets.Text)
  End If
End With
lblTotalDeclarado.Caption = Format(TotalDeclarado, "Currency")
End Sub

Private Sub Confirmado()
cmdFinalizar.Visible = False
For i = 0 To Tela.Count - 1
  Tela(i).Enabled = False
Next i

txtChequeRecebido.Enabled = False
txtChequesPre.Enabled = False
txtJuros.Enabled = False
DBGrid1.AllowDelete = False
DBGrid2.AllowAddNew = False
DBGrid3.AllowDelete = False
DBGrid4.AllowAddNew = False
DBGrid5.AllowDelete = False
DBGrid6.AllowDelete = False
DBGrid7.AllowDelete = False
DBGrid8.AllowDelete = False
DBGrid9.AllowDelete = False
DBGrid10.AllowDelete = False

a = dbFechamento.Recordset!Codigoresponsavel
With dbResponsavel
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codigovendedor=" & a
    If .Recordset.NoMatch = False Then
      cboResponsavel.Text = .Recordset!Nome
    End If
  End If
End With
End Sub

Private Sub NaoConfirmado()
cmdFinalizar.Visible = True
For i = 0 To Tela.Count - 1
  Tela(i).Enabled = True
Next i
txtChequeRecebido.Enabled = True
txtChequesPre.Enabled = True
DBGrid1.AllowDelete = True
DBGrid2.AllowDelete = True
DBGrid4.AllowDelete = True
DBGrid5.AllowDelete = True
DBGrid6.AllowDelete = True
DBGrid7.AllowDelete = True
DBGrid8.AllowDelete = True
DBGrid9.AllowDelete = True


End Sub

Private Sub Totaliza()
Dim TempValor As Currency, totalGeral As Currency, totalPrazo As Currency
Dim TotalDeclarado As Currency

With QBicoMovimentaTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QBicoMovimentoTotaliza where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!totalvendido) = False Then
    If IsNumeric(.Recordset!totalvendido) = True Then
      Label25.Caption = Format(.Recordset!totalvendido, "Currency")
      lblVendasCombustivel.Caption = Format(.Recordset!totalvendido, "Currency")
    End If
  End If
End With
With QDespesaLancTotaliza
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QDespesaLancTotaliza where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!TotalDespesa) = False Then
    If IsNumeric(.Recordset!TotalDespesa) = True Then
      Label39.Caption = Format(.Recordset!TotalDespesa, "Currency")
      lblDespesas.Caption = Format(.Recordset!TotalDespesa, "Currency")
      lblTotalDespesas.Caption = Format(.Recordset!TotalDespesa, "Currency")
    Else
      Label39.Caption = Format(0, "Currency")
      lblDespesas.Caption = Format(0, "Currency")
      lblTotalDespesas.Caption = Format(0, "Currency")
    End If
  Else
    Label39.Caption = Format(0, "Currency")
    lblDespesas.Caption = Format(0, "Currency")
    lblTotalDespesas.Caption = Format(0, "Currency")
  End If
End With

With QFormaDePgRecTotaliza
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QFormaDePagamentoRecebidoTotaliza where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      Label42.Caption = Format(.Recordset!Total, "Currency")
      lblTotalRecebido.Caption = Format(.Recordset!Total, "Currency")
    Else
      Label42.Caption = Format(0, "Currency")
      lblTotalRecebido.Caption = Format(0, "Currency")
    End If
  Else
    Label42.Caption = Format(0, "Currency")
    lblTotalRecebido.Caption = Format(0, "Currency")
  End If
End With
With QVendaTotaliza
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QVendaTotaliza where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblVendasProdutos.Caption = Format(.Recordset!Total, "Currency")
      Label33.Caption = Format(.Recordset!Total, "Currency")
      lblComissao.Caption = Format(.Recordset!Comissao, "Currency")
    Else
      lblVendasProdutos.Caption = Format(0, "Currency")
      Label33.Caption = Format(0, "Currency")
      lblComissao.Caption = Format(0, "Currency")
    End If
  Else
    lblVendasProdutos.Caption = Format(0, "Currency")
    Label33.Caption = Format(0, "Currency")
    lblComissao.Caption = Format(0, "Currency")
  End If
End With
With QClientesNota
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from qclientesnota where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!Valor) = False Then
    If IsNumeric(.Recordset!Valor) = True Then
      lblClientesNota.Caption = Format(.Recordset!Valor, "Currency")
      Label44.Caption = Format(.Recordset!Valor, "Currency")
    Else
      lblClientesNota.Caption = Format(0, "Currency")
      Label44.Caption = Format(0, "Currency")
    End If
  Else
    lblClientesNota.Caption = Format(0, "Currency")
    Label44.Caption = Format(0, "Currency")
  End If
End With
With QProdutoEntra
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valornota) as total from produtosentrada where codigofechamento=" & CodigoFechamento
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      Label63.Caption = Format(.Recordset!Total, "Currency")
      lblProdutoEntraTotal.Caption = Format(.Recordset!Total, "Currency")
    Else
      lblProdutoEntraTotal.Caption = Format(0, "Currency")
      Label63.Caption = Format(0, "Currency")
    End If
  Else
    lblProdutoEntraTotal.Caption = Format(0, "Currency")
    Label63.Caption = Format(0, "Currency")
  End If
End With
With qGalonagem
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from qgalonagemproduto where codigofechamento=" & CodigoFechamento
  .Refresh
  TempValor = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      TempValor = TempValor + .Recordset("vendido")
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

totalPrazo = 0
With dbBicoMovimento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!prazo > 0 Then
        totalPrazo = totalPrazo + .Recordset!valorvendido
      End If
      .Recordset.MoveNext
    Loop
  Else
    lblPreDatado.Caption = Format(0, "Currency")
  End If
End With

Juros = 0
'If Tela(0).Enabled = True Then
  With dbCheques
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      Do While .Recordset.EOF = False
        Juros = Juros + (.Recordset!Valor - .Recordset!valornabomba)
        .Recordset.MoveNext
      Loop
      If IsNull(Juros) = False Then
        txtJuros.Text = Juros
      Else
        txtJuros.Text = 0
      End If
    End If
  End With
'End If
If IsNumeric(txtJuros.Text) = True Then
  totalPrazo = totalPrazo + (CCur(txtJuros.Text) / 0.1)
End If
lblPreDatado.Caption = Format(totalPrazo, "Currency")

With qCheques
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      Label54.Caption = Format(.Recordset!Total, "Currency")
    Else
      Label54.Caption = Format(0, "Currency")
    End If
  Else
    Label54.Caption = Format(0, "Currency")
  End If
End With

lblGalonagem.Caption = TempValor

TempValor = 0
If IsNumeric(lblVendasCombustivel.Caption) = True Then
  TempValor = TempValor + CCur(lblVendasCombustivel.Caption)
End If
If IsNumeric(lblVendasProdutos.Caption) = True Then
  TempValor = TempValor + CCur(lblVendasProdutos.Caption)
End If

If IsNumeric(txtJuros.Text) = True Then
  TempValor = TempValor + CCur(txtJuros.Text)
End If

totalGeral = -TempValor
lblTotalVendas.Caption = Format(TempValor, "Currency")

TempValor = 0
If IsNumeric(lblDespesas.Caption) = True Then
  TempValor = TempValor - CCur(lblDespesas.Caption)
End If
If IsNumeric(lblProdutoEntraTotal.Caption) = True Then
  TempValor = TempValor + CCur(lblProdutoEntraTotal.Caption)
End If
If IsNumeric(lblComissao.Caption) = True Then
  If ComissaoAcumulativa = False Then
    If cmdFinalizar.Visible = True Then
      TempValor = TempValor + CCur(lblComissao.Caption)
    End If
  End If
End If

totalGeral = totalGeral + TempValor
lblTotalDespesas.Caption = Format(TempValor, "Currency")
TempValor = 0

If IsNumeric(lblClientesNota.Caption) = True Then
  TempValor = TempValor + CCur(lblClientesNota.Caption)
End If

If IsNumeric(lblTotalChequeResumo.Caption) = True Then
  TempValor = TempValor + CCur(lblTotalChequeResumo.Caption)
Else
  If IsNumeric(txtChequesPre.Text) = True Then
    TempValor = TempValor + CCur(txtChequesPre.Text)
  End If
  If IsNumeric(txtChequeRecebido.Text) = True Then
    TempValor = TempValor + txtChequeRecebido.Text
  End If
End If

If IsNumeric(lblTotalRecebido.Caption) = True Then
  TempValor = TempValor + CCur(lblTotalRecebido.Caption)
End If

'Soma os valores declarados
TotalDeclarado = 0
If IsNumeric(Label42.Caption) = False Then
  If IsNumeric(txtDinheiro.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtDinheiro.Text)
  End If
  If IsNumeric(txtCartoes.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtCartoes.Text)
  End If
  If IsNumeric(txtVT.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtVT.Text)
  End If
Else
  If CCur(Label42.Caption) = 0 Then
    If IsNumeric(txtDinheiro.Text) = True Then
      TotalDeclarado = TotalDeclarado + CCur(txtDinheiro.Text)
    End If
    If IsNumeric(txtCartoes.Text) = True Then
      TotalDeclarado = TotalDeclarado + CCur(txtCartoes.Text)
    End If
    If IsNumeric(txtVT.Text) = True Then
      TotalDeclarado = TotalDeclarado + CCur(txtVT.Text)
    End If
  End If
End If
If IsNumeric(Label39.Caption) = False Then
  If IsNumeric(txtDespesas.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtDespesas.Text)
  End If
End If

If IsNumeric(Label44.Caption) = False Then
  If IsNumeric(txtNotas.Text) = True Then
    TotalDeclarado = TotalDeclarado + CCur(txtNotas.Text)
  End If
End If


TempValor = TempValor + TotalDeclarado

totalGeral = totalGeral + TempValor
lblRecebimentos.Caption = Format(TempValor, "currency")
lblDiferenca.Caption = Format(totalGeral, "Currency")


End Sub

Private Sub AbreFechamento(ByVal Fechamento As Double, ByVal Posto As Double)
  With dbPosto
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from postos order by nome"
    .Refresh
  End With
  With dbResponsavel
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from vendedores order by nome"
    .Refresh
  End With
  With dbFechamento
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from FechamentoDiario where codigofechamento=" & CodigoFechamento & " order by Data, Hora "
    .Refresh
  End With
  With dbBicoMovimento
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from bicomovimento where codigofechamento=" & Fechamento & " order by bico"
    .Refresh
  End With
  With dbBico
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from bicos where codigoposto=" & Posto & " order by bico"
    .Refresh
  End With
  With dbTanques
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from tanques where codigoposto=" & Posto & " order by tanque"
    .Refresh
  End With
  With dbTanquesMovimento
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from tanquesMovimento where codigofechamento=" & Fechamento & " and codigoposto=" & Posto & " order by tanque"
    .Refresh
  End With
  With dbProdutos
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from produtos order by descri"
    .Refresh
  End With
  With dbProdutos2
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from produtos where combustivel=0 order by descri"
    .Refresh
  End With
  With dbVendas
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from venda where codigofechamento=" & Fechamento & " order by descri"
    .Refresh
  End With
  With dbDespesas
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from DespesaTipo order by descri"
    .Refresh
  End With
  With dbDespesasLanc
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from DespesasLanc where codigofechamento=" & Fechamento & " and codigoconta=-1 order by descri"
    .Refresh
  End With
  With dbFormaDePg
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from FormaDePagamento order by descri"
    .Refresh
  End With
  With dbFormaDePgRecebido
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from FormaDePagamentoRecebido where codigofechamento=" & Fechamento & " order by descri"
    .Refresh
  End With
  With dbContas
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from contas order by descri"
    .Refresh
  End With
  With dbStatus
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from status"
    .Refresh
  End With
  With QTemp
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from QBicoMovimentoTotalTanque"
    .Refresh
  End With
  With dbClientes
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from clientes where mensalista=-1 order by nome"
    .Refresh
  End With
  With dbClientesNota
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from clientesNota where codigofechamento=" & CodigoFechamento & " order by nome"
    .Refresh
  End With
  With QClientesNota
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from qclientesnota where codigofechamento=" & CodigoFechamento
    .Refresh
  End With
  With dbTurno
    .Connect = Conectar
    .DatabaseName = Caminho
    .Refresh
  End With
  With dbCheques
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from cheques where codigofechamento=" & CodigoFechamento
    .Refresh
  End With
  With qCheques
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select sum(valor) as Total from cheques where codigofechamento=" & CodigoFechamento
    .Refresh
  End With
  With dbProdutoEntra
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from produtosentrada where codigofechamento=" & CodigoFechamento
    .Refresh
  End With
  With dbCarros
    .Connect = Conectar
    .DatabaseName = Caminho
    .Refresh
  End With
  With QComissoes
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "Select sum(valorcomissao) as comissao, codigovendedor from venda where codigofechamento=" & CodigoFechamento & " group by codigovendedor"
    .Refresh
  End With
  With dbJuros
    .Connect = Conectar
    .DatabaseName = Caminho
    .Refresh
  End With
  With dbDifComb
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from DiferencaCombustivel where codigofechamento=" & CodigoFechamento
    .Refresh
  End With
  
End Sub

Private Sub cboBico_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboBico_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboBico_LostFocus()
Me.KeyPreview = True
With dbBico
  .Refresh
  If cboBico.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "bico=" & CLng(cboBico.Text)
  If .Recordset.NoMatch = False Then
    cboBico.Text = .Recordset("bico")
  End If
End With
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
    If .Recordset!posicao = True Then
      lblStatus.ForeColor = vbBlack
      lblStatus.Caption = "Ativo"
    Else
      lblStatus.ForeColor = vbRed
      lblStatus.Caption = "Inativo"
    End If
  End If
End With
End Sub

Private Sub cboClientesNota_GotFocus()
Me.KeyPreview = False
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
  If cboClientesNota.Text = "" Then Exit Sub
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboClientesNota.Text & "'"
  If .Recordset.NoMatch = False Then
    cboClientesNota.Text = .Recordset("nome")
    With dbCarros
      .RecordSource = "select *from clientescarros where codigocliente=" & dbClientes.Recordset!CodigoCliente
      .Refresh
    End With
  End If
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

Private Sub cboProduto_LostFocus()
Me.KeyPreview = True
With dbProdutos2
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProduto.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodProduto.Text = .Recordset("codigo")
  End If
End With
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

Private Sub cboResponsavel_LostFocus()
With dbResponsavel
  .Refresh
  If cboResponsavel.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "nome='" & cboResponsavel.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboResponsavel.Text = .Recordset("nome")
End With
End Sub

Private Sub cboTanque_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboTanque_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboTanque_LostFocus()
Me.KeyPreview = True
With dbTanques
  .Refresh
  If cboTanque.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "tanque=" & cboTanque.Text
  If .Recordset.NoMatch = False Then
    cboTanque.Text = .Recordset("tanque")
  End If
End With
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
  If .Recordset.NoMatch = True Then Exit Sub
  cboTurno.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdAnterior_Click()
For i = 0 To Tela.Count - 1
  Tela(i).Visible = True
Next i
intTela = Paginas.Tab
If intTela < 0 Then
  Exit Sub
Else
  intTela = intTela - 1
  Paginas.Tab = intTela
  cmdProximo.Enabled = True
  If intTela = 0 Then
    cmdAnterior.Enabled = False
  End If
End If

cmdFinalizar.Enabled = False
cmdResumo.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
Dim Resposta As Integer
With cmdCancelar
  If .Caption = "&Sair" Then
    Unload Me
  Else
    If cmdFinalizar.Visible = True Then
      Resposta = MsgBox("Deseja cancelar o fechamento atual?", vbYesNo + vbDefaultButton2, "Cancela fechamento!")
      If Resposta = vbYes Then
        dbFechamento.Recordset.Edit
        dbFechamento.Recordset("cancelado") = True
        dbFechamento.Recordset.Update
      End If
    End If
    cmdProximo.Enabled = False
    cmdAnterior.Enabled = False
    cmdFinalizar.Enabled = False
    cmdCancelar.Caption = "&Sair"
    Paginas.Visible = False
    cmdFinalizar.Visible = True
    cmdResumo.Enabled = False
    
    cmdProximo.Enabled = False
    cmdInlueBomba.Enabled = True
    cboResponsavel.Enabled = True
    txtData.Enabled = True
    cboTurno.Enabled = True
    cboResponsavel.SetFocus
    With QTemp
      .Connect = Conectar
      .DatabaseName = Caminho
      .RecordSource = "Select *from fechamentodiario where cancelado=0 order by codigofechamento"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        lblUltimoLancado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
        On Error Resume Next
        If .Recordset!Confirmado = 0 Then
          Do While .Recordset!Confirmado = 0
            If .Recordset.BOF = True Then Exit Do
            .Recordset.MovePrevious
          Loop
        End If
        lblUltimoConfirmado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
      End If
      On Error GoTo 0
    End With
  End If
End With
End Sub

Private Sub cmdFinalizar_Click()
Dim Resposta As Integer, LucroVenda As Currency, TempValor As Currency
Dim DifEstoque As Double, ValorDiferenca As Currency
Dim Dias As Double, ReceberData As Date, StrTemp As String
Dim Mes As Boolean, Intervalo As String, Pontos As Double

'Primeiro verifica se não existe outro fechamento pendente
With QTemp
  .RecordSource = "Select *from fechamentodiario where cancelado=0 order by codigofechamento"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    .Recordset.FindFirst "codigofechamento=" & CodigoFechamento
    If .Recordset.NoMatch = False Then
      .Recordset.MovePrevious
      If .Recordset.BOF = False Then
        If .Recordset!Confirmado = 0 Then
          Do While .Recordset!Confirmado = 0
            .Recordset.MovePrevious
          Loop
          MsgBox "O último caixa confirmado foi: " & .Recordset!Data & " / Turno: " & .Recordset!Turno
          Permissao = False
          frmPermissao.Show vbModal
          If Permissao = False Then Exit Sub
        End If
      End If
    End If
  End If
  If IsNumeric(txtPontos.Text) = True Then
    If IsNull(.Recordset!Pontos) = False Then
      If .Recordset!Pontos <> 0 Then
        Pontos = CDbl(txtPontos.Text) - .Recordset!Pontos
        If Pontos > CDbl(lblTotalVendas.Caption) Then
          Resposta = MsgBox("O volume de vendas é menor que o total de pontos distribuidos!" & Chr(13) & "Deseja continuar?", vbYesNo)
          If Resposta = vbNo Then Exit Sub
        End If
      End If
    End If
  End If
End With

'verifica se deseja fechar mesmo
Resposta = MsgBox("Deseja finalizar o fechamento agora?!", vbYesNo, "Fechamento")
If Resposta = vbNo Then Exit Sub

Screen.MousePointer = vbHourglass

'Atualiza todos os dados do lançamento
AbreFechamento CodigoFechamento, dbPosto.Recordset("codigoposto")
Totaliza
TempValor = 0
If dbFormaDePgRecebido.Recordset.RecordCount = 0 Then
  Permissao = False
  MsgBox "Não foi lançado nenhums tipo de recebimento!"
  frmPermissao.Show vbModal
  If Permissao = False Then
    Exit Sub
  End If
End If

'calcula o movimento de bico, registrando o contador da bomba, o lucro de venda
'e o estoque do tanque
With dbBicoMovimento
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbBico.Recordset.FindFirst "codigobico=" & .Recordset("codigobico")
      If .Recordset.EOF = True Then
        MsgBox "O bico " & .Recordset("bico") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
      End If
      dbBico.Recordset.Edit
      dbBico.Recordset("ultimonumero") = .Recordset("valorfinal")
      dbBico.Recordset("ultimomecanico") = .Recordset("mecanicoFinal")
      dbBico.Recordset.Update
      
      QTemp.RecordSource = "select *from bicoencerrantes where codigofechamento=" & CodigoFechamento & " and tanque=" & .Recordset("tanque")
      QTemp.Refresh
      If QTemp.Recordset.RecordCount <> 0 Then
        dbTanques.Refresh
        If dbTanques.Recordset.RecordCount = 0 Then
          MsgBox "Cadastro de tanques vazio!"
        Else
          dbTanques.Recordset.FindFirst "tanque=" & dbBicoMovimento.Recordset("tanque")
          If .Recordset.NoMatch = True Then
            MsgBox "O tanque " & dbBicoMovimento.Recordset("tanque") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
          Else
            dbProdutos.Refresh
            dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
            LucroVenda = (.Recordset!precounitario - dbProdutos.Recordset!precocompra) * .Recordset!Vendas
            dbTanques.Recordset.Edit
            dbTanques.Recordset("estoque") = dbTanques.Recordset("estoque") - QTemp.Recordset("vendido")
            dbTanques.Recordset.Update
          End If
        End If
        dbProdutos.Refresh
        If dbProdutos.Recordset.RecordCount <> 0 Then
          dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
          If dbProdutos.Recordset.NoMatch = False Then
            dbProdutos.Recordset.Edit
            dbProdutos.Recordset("estoque") = dbProdutos.Recordset("estoque") - .Recordset("vendas")
            dbProdutos.Recordset("acumulativo") = dbProdutos.Recordset("acumulativo") + .Recordset("vendas")
            dbProdutos.Recordset!LucroVenda = dbProdutos.Recordset!LucroVenda + LucroVenda
            dbProdutos.Recordset.Update
          End If
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With QTemp
  .RecordSource = "select sum(valorvendido) as total, codigoproduto from bicomovimento where codigofechamento=" & CodigoFechamento & " group by codigoproduto"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If IsNull(.Recordset!Total) = False Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
        If dbProdutos.Recordset.NoMatch = False Then
          dbProdutos.Recordset.Edit
          If IsNull(dbProdutos.Recordset!totalvendido) = True Then dbProdutos.Recordset!totalvendido = 0
          dbProdutos.Recordset!totalvendido = dbProdutos.Recordset!totalvendido + .Recordset!Total
          dbProdutos.Recordset.Update
        End If
        .Recordset.MoveNext
      Loop
    End If
  End If
End With

If IsNumeric(txtJuros.Text) Then
  With QTemp
    .RecordSource = "select *from status"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Edit
      .Recordset!Juros = .Recordset!Juros + CCur(txtJuros.Text)
      .Recordset.Update
    End If
  End With
End If

With dbVendas
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      dbProdutos2.Refresh
      dbProdutos2.Recordset.FindFirst "codigoproduto=" & .Recordset("codigoproduto")
      If dbProdutos2.Recordset.NoMatch = True Then
        MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
      Else
        LucroVenda = (.Recordset("valorunitario") * .Recordset("quantidade")) - (dbProdutos2.Recordset("precocompra") * .Recordset("quantidade")) - .Recordset("ValorDesconto") - .Recordset!ValorComissao
        dbProdutos2.Recordset.Edit
        dbProdutos2.Recordset("estoque") = dbProdutos2.Recordset("estoque") - .Recordset("quantidade")
        dbProdutos2.Recordset!LucroVenda = dbProdutos2.Recordset!LucroVenda + LucroVenda
        dbProdutos2.Recordset!acumulativo = dbProdutos2.Recordset!acumulativo + .Recordset("quantidade")
        If IsNull(dbProdutos2.Recordset!totalvendido) = True Then dbProdutos2.Recordset!totalvendido = 0
        dbProdutos2.Recordset!totalvendido = dbProdutos2.Recordset!totalvendido + .Recordset!ValorTotal
        dbProdutos2.Recordset.Update
      End If
      If ComissaoAcumulativa = True Then
        If .Recordset!ValorComissao <> 0 Then
          If .Recordset!CodigoVendedor <> 0 Then
            dbResponsavel.Refresh
            If dbResponsavel.Recordset.RecordCount <> 0 Then
              dbResponsavel.Recordset.FindFirst "codigo=" & .Recordset!CodigoVendedor
              If dbResponsavel.Recordset.NoMatch = False Then
                dbResponsavel.Recordset.Edit
                dbResponsavel.Recordset!comissaosaldo = dbResponsavel.Recordset!comissaosaldo + .Recordset!ValorComissao
                dbResponsavel.Recordset.Update
                dbResponsavel.Refresh
              End If
            End If
          End If
        End If
      Else
        .Recordset!Pago = True
        .Recordset.Update
      End If
      On Error Resume Next
      .Recordset.Update
      On Error GoTo 0
      .Recordset.MoveNext
    Loop
  End If
End With

'Registra a compra de produtos
With dbProdutoEntra
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      dbProdutos.Refresh
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = True Then
        MsgBox "Erro na tabela de produtos! Codigo produto: " & .Recordset!CodigoProduto
      End If
      TempValor = (.Recordset!PrecoNovo - dbProdutos.Recordset!precocompra) * dbProdutos.Recordset!Estoque
      dbProdutos.Recordset.Edit
      dbProdutos.Recordset!precocompra = .Recordset!PrecoNovo
      dbProdutos.Recordset!Variacao = dbProdutos.Recordset!Variacao + TempValor
      dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque + .Recordset!Quantidade
      dbProdutos.Recordset.Update
      With dbDespesasLanc
        .Recordset.AddNew
        .Recordset!CodigoFechamento = CodigoFechamento
        .Recordset!Origem = "Fechamento"
        .Recordset!Data = Date
        .Recordset!Hora = Now
        .Recordset!Vencimento = Date
        .Recordset!CodigoDespesa = -1
        .Recordset!Descri = "Compra de Produto"
        .Recordset!obs = dbProdutos.Recordset!Descri & " " & dbProdutoEntra.Recordset!Quantidade & " - " & dbProdutoEntra.Recordset!valornota
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


'Totaliza as despesas para lançar no status e no saldo das contas
With dbDespesasLanc
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

'Totaliza os recebimentos e lança nas contas
With QTemp
  .RecordSource = "select *from formadepagamentorecebido where codigofechamento=" & CodigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  .RecordSource = "select *from QFormadePgContasRecebido where codigofechamento=" & CodigoFechamento
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      LucroVenda = LucroVenda + .Recordset("valor")
      Dias = .Recordset("reembolso")
      Mes = .Recordset("mes")
      txtDataBordero.Value = .Recordset!Data
      If Mes = True Then
        Intervalo = "m"
      Else
        Intervalo = "d"
      End If
      If Dias > 0 Then
        ReceberData = DateAdd(Intervalo, Dias, txtDataBordero.Value)
      Else
        Dias = .Recordset("diadomes")
        If Dias > 0 Then
          If Dias >= txtData.Day Then
            If Dias < 28 Then
              StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
            Else
              StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
              Do While IsDate(StrTemp) = False
                Dias = Dias - 1
                If Dias <= 0 Then Dias = 31
                StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
              Loop
            End If
            ReceberData = CDate(StrTemp)
          Else
            If Dias < 28 Then
              StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
            Else
              StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
              Do While IsDate(StrTemp) = False
                Dias = Dias - 1
                If Dias <= 0 Then Dias = 31
                StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
              Loop
            End If
            ReceberData = CDate(StrTemp)
          End If
        End If
      End If
      If Dias > 0 Then
        Select Case Weekday(ReceberData)
          Case 1 'domingo
            ReceberData = DateAdd("d", 1, ReceberData)
          Case 7 'sábado
            ReceberData = DateAdd("d", 2, ReceberData)
        End Select
        dbCartoes.Refresh
        If dbCartoes.Recordset.EOF = False Then
          dbCartoes.Recordset.FindFirst "codigoformapg=" & .Recordset!CodigoFormadePg & " and datalanc=#" & DataInglesa(Trim(Str(txtDataBordero.Value))) & "# and confirmado=0"
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
        dbCartoes.Recordset!codigoconta = .Recordset("contas.codigoconta")
        dbCartoes.Recordset!conta = .Recordset("contas.descri")
        dbCartoes.Recordset!CodigoFormaPg = .Recordset!CodigoFormadePg
        dbCartoes.Recordset!Grupo = .Recordset!Grupo
        dbCartoes.Recordset!Descri = .Recordset("formadepagamento.descri")
        dbCartoes.Recordset!datalanc = txtDataBordero.Value
        dbCartoes.Recordset!DataPrevista = ReceberData
        dbCartoes.Recordset!ValorBruto = dbCartoes.Recordset!ValorBruto + .Recordset!ValorBruto
        dbCartoes.Recordset!valorliquido = dbCartoes.Recordset!valorliquido + .Recordset!Valor
        dbCartoes.Recordset.Update
        
      Else
        ReceberData = txtDataBordero.Value
        Select Case Weekday(ReceberData)
          Case 1 'domingo
            ReceberData = DateAdd("d", 1, ReceberData)
          Case 7 'sábado
            ReceberData = DateAdd("d", 2, ReceberData)
        End Select
        
        With dbConciliaNova
          .Recordset.AddNew
          .Recordset!codigoconta = QTemp.Recordset("contas.codigoconta")
          .Recordset!datalanc = Now
          .Recordset!compensado = True
          .Recordset!Data = Date
          .Recordset!tipo = "Fechamento"
          .Recordset!Codigo = 999999998
          .Recordset!Descri = Left("Caixa - " & Format(txtData.Value, "short date") & " - " & cboTurno.Text & " - " & QTemp.Recordset("formadepagamento.descri"), 50)
          .Recordset!nrdocumento = Format(txtDataBordero.Value, "short date")
          .Recordset!Valor = QTemp.Recordset!Valor
          .Recordset.Update
        End With
        dbContas.Refresh
        dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("contas.codigoconta")
        If dbContas.Recordset.NoMatch = True Then
          MsgBox "Conta " & .Recordset("contas.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
        Else
          TempValor = .Recordset("valor")
          dbContas.Recordset.Edit
          dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + TempValor
          dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
          dbContas.Recordset.Update
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

'fecha nota de clientes
With dbClientesNota
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With


'fecha cheques
With dbCheques
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!CodigoSoma = "1"
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

'Lança comissão como despesa
If ComissaoAcumulativa = False Then
  If IsNumeric(lblComissao.Caption) = True Then
    With dbDespesasLanc
      .Recordset.AddNew
      .Recordset("codigofechamento") = CodigoFechamento
      .Recordset!Origem = "Fechamento"
      .Recordset("data") = txtData.Value
      .Recordset!Vencimento = txtData.Value
      .Recordset("hora") = Now
      .Recordset("codigoconta") = 0
      .Recordset("conta") = "Comissão"
      .Recordset("codigodespesa") = 0
      .Recordset("descri") = "Comissões"
      .Recordset("obs") = " "
      .Recordset!compensado = True
      .Recordset("valor") = -CCur(lblComissao.Caption)
      .Recordset!valorpago = -CCur(lblComissao.Caption)
      .Recordset.Update
      .Refresh
    End With
  End If
End If
With dbFechamento
  .Refresh
  .Recordset.FindFirst "codigofechamento=" & CodigoFechamento
  If .Recordset.EOF = True Then
    MsgBox "Erro na tabela de fechamento!"
  Else
    .Recordset.Edit
    .Recordset("totalVendas") = CCur(lblTotalVendas.Caption)
    .Recordset("TotalDespesa") = CCur(lblTotalDespesas.Caption)
    If IsNumeric(lblTotalRecebido.Caption) Then
      .Recordset("Totalrecebimento") = CCur(lblTotalRecebido.Caption)
    Else
      .Recordset!TotalRecebimento = 0
    End If
    .Recordset("diferenca") = CCur(lblDiferenca.Caption)
    .Recordset!Juros = CCur(txtJuros.Text)
    .Recordset("confirmado") = True
    .Recordset!fechames = False
    If IsNumeric(lblTotalChequeResumo.Caption) = True Then
      .Recordset!chequeavista = CCur(lblTotalChequeResumo.Caption)
    End If
    .Recordset.Update
  End If
End With


With QTemp
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from fechamentodiario where cancelado=0 order by codigofechamento"
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    lblUltimoLancado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
    If .Recordset!Confirmado = 0 Then
      Do While .Recordset!Confirmado = 0
        .Recordset.MovePrevious
      Loop
    End If
    lblUltimoConfirmado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
  End If
End With


cboResponsavel.Enabled = True
txtData.Enabled = True
cboTurno.Enabled = True
cmdAnterior.Enabled = False
cmdProximo.Enabled = False
cmdFinalizar.Enabled = False
cmdInlueBomba.Enabled = True
cmdCancelar.Caption = "&Sair"
cmdResumo.Enabled = False
cboResponsavel.SetFocus
txtJuros = Format(0, "Currency")
Paginas.Visible = False
intTela = 0
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdInclueBico_Click()
Dim ValorUnitario As Currency, ValorInicial As Double, ValorFinal As Double
Dim Quantidade As Double, Mecanico As Double, Retorno As Double, LucroVenda As Currency
Dim Zerou As Boolean

If cboBico.Text <> dbBico.Recordset("bico") Then
  MsgBox "Bico inválido!", vbCritical, "Erro!"
  cboBico.SetFocus
  Exit Sub
End If
If IsNumeric(txtBicoEncerra.Text) = False Then
  MsgBox "Encerramento inválido!"
  txtBicoEncerra.SetFocus
  Exit Sub
End If
With dbBicoMovimento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codigobico=" & dbBico.Recordset!codigobico
    If .Recordset.NoMatch = False Then
      MsgBox "Bico já lançado!"
      cboBico.SetFocus
      Exit Sub
    End If
  End If
End With
ValorFinal = CDbl(txtBicoEncerra.Text)
If ValorFinal < dbBico.Recordset("ultimonumero") Then
  If ValorFinal < 3000 Then
    If dbBico.Recordset("ultimonumero") > 996000 Then
      Resposta = MsgBox("Este lançamento está acusando que a numeração do bico ultrapassou o número 999999, isto está correto?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      ValorInicial = dbBico.Recordset("ultimoNumero")
    Else
      MsgBox "Encerramento inválido!", vbCritical, "Erro!"
      txtBicoEncerra.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "Encerramento inválido!", vbCritical, "Erro!"
    txtBicoEncerra.SetFocus
    Exit Sub
  End If
Else
  ValorInicial = dbBico.Recordset("ultimoNumero")
End If

'Veirfica de existe outro caixa em lançamento
With dbBico
  If .Recordset!UltimoNumero <> .Recordset!provisorionumero Then
    ValorInicial = .Recordset!provisorionumero
  End If
End With

If IsNumeric(txtMecanico.Text) = False Then
  txtMecanico.Text = txtBicoEncerra.Text
End If
Mecanico = CDbl(txtMecanico.Text)
Zerou = False
If Mecanico < dbBico.Recordset("ultimomecanico") Then
  If Mecanico < 3000 Then
    If dbBico.Recordset("ultimomecanico") > 996000 Then
      Resposta = MsgBox("Este lançamento está acusando que a numeração do bico ultrapassou o número 999999, isto está correto?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      Zerou = True
    Else
      MsgBox "Número mecânico inválido!"
      txtMecanico.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "Número mecânico inválido!"
    txtMecanico.SetFocus
    Exit Sub
  End If
End If
If Zerou = True Then
  TempValor = (Mecanico + 1000000 - dbBico.Recordset("ultimomecanico")) - (ValorFinal - dbBico.Recordset("ultimonumero"))
Else
  TempValor = (Mecanico - dbBico.Recordset("ultimomecanico")) - (ValorFinal - dbBico.Recordset("ultimonumero"))
End If
If TempValor > 5 Or TempValor < -5 Then
  MsgBox "Discordância de valores!"
  Permissao = False
  frmPermissao.Show vbModal
  If Permissao = False Then
    txtMecanico.SetFocus
    Exit Sub
  End If
End If
Retorno = 0
If IsNumeric(txtRetorno.Text) = True Then
  Retorno = CDbl(txtRetorno.Text)
End If

With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Tabela de Produtos vazia!", vbCritical, "Erro!"
    Exit Sub
  End If
  .Recordset.FindFirst "codigoproduto=" & dbBico.Recordset("codigoproduto")
  If .Recordset.NoMatch = True Then
    MsgBox "Produto da bomba não encontrado!", vbCritical, "Erro!"
    Exit Sub
  End If
  ValorUnitario = dbBico.Recordset("precovenda")
End With


If ValorFinal < ValorInicial Then
  Quantidade = (999999 - ValorInicial) + ValorFinal - Retorno
Else
  Quantidade = ValorFinal - ValorInicial - Retorno
End If

With dbBicoMovimento
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Refresh
    .Recordset.FindFirst "bico=" & cboBico.Text
    If .Recordset.NoMatch = True Then
      .Recordset.AddNew
    Else
      Resposta = MsgBox("Bico já lançado! Deseja alterar?", vbYesNo + vbDefaultButton2, "Bico lançado!")
      If Resposta = vbNo Then Exit Sub
      .Recordset.Edit
    End If
  End If
  .Recordset("codigoFechamento") = CodigoFechamento
  .Recordset("Data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("codigobico") = dbBico.Recordset("codigobico")
  .Recordset("bico") = dbBico.Recordset("bico")
  .Recordset("valorinicial") = ValorInicial
  .Recordset("valorfinal") = ValorFinal
  .Recordset("mecanicoInicial") = dbBico.Recordset("ultimomecanico")
  .Recordset("mecanicofinal") = Mecanico
  .Recordset("precocompra") = dbProdutos.Recordset!precocompra
  .Recordset("precounitario") = ValorUnitario
  .Recordset("vendas") = Quantidade
  .Recordset("retorno") = Retorno
  .Recordset("valorvendido") = Quantidade * ValorUnitario
  .Recordset("tanque") = dbBico.Recordset("tanque")
  .Recordset("codigoProduto") = dbProdutos.Recordset("Codigoproduto")
  .Recordset!prazo = dbBico.Recordset!prazo
  .Recordset.Update
  .Refresh
End With
With dbBico
  .Recordset.Edit
  .Recordset!provisorionumero = ValorFinal
  .Recordset.Update
  .Refresh
End With

DBGrid1.Refresh
Totaliza
txtBicoEncerra.Text = ""
txtMecanico.Text = ""
txtRetorno.Text = "0"
cboBico.SetFocus

End Sub

Private Sub cmdInclueNota_Click()
Dim DataPrevista As Date

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
If IsNumeric(txtCupom.Text) = False Then
  MsgBox "Informe o número do cupom válido!"
  txtCupom.SetFocus
  Exit Sub
End If
If dbClientes.Recordset("diapagamento") <> 0 Then
  DataPrevista = CDate(Format(dbClientes.Recordset("diapagamento"), "00") & "/" & txtData.Month & "/" & txtData.Year)
Else
  DataPrevista = DateAdd("m", 1, txtData.Value)
End If
If DataPrevista < Date Then
  DataPrevista = DateAdd("m", 1, DataPrevista)
End If

With dbClientesNota
  .Recordset.AddNew
  .Recordset("codigofechamento") = CodigoFechamento
  .Recordset("codigocliente") = dbClientes.Recordset("codigoCliente")
  .Recordset("nome") = dbClientes.Recordset("nome")
  .Recordset("datalanc") = Now
  .Recordset("dataprevista") = DataPrevista
  .Recordset("valorprevisto") = CCur(txtNotaValor.Text)
  .Recordset!Data = txtData.Value
  .Recordset!placa = cboPlaca.Text
  .Recordset!cupom = txtCupom.Text
  .Recordset.Update
  .Refresh
End With
'With dbClientes
'  .Recordset.Edit
'  .Recordset!Saldo = .Recordset!Saldo - CCur(txtNotaValor.Text)
'  .Recordset.Update
'  .Refresh
'End With
Totaliza
cboClientesNota.Text = ""
txtNotaValor.Text = ""
cboPlaca.Text = ""
txtCupom.Text = ""
cboClientesNota.SetFocus
End Sub

Private Sub cmdIncluirDespesa_Click()

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
With dbDespesasLanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = CodigoFechamento
  .Recordset!Origem = "Fechamento"
  .Recordset("data") = txtData.Value
  .Recordset!Vencimento = txtData.Value
  .Recordset("hora") = Now
  .Recordset("codigoconta") = -1
  .Recordset("conta") = "Fechamento de Caixa"
  .Recordset("codigodespesa") = dbDespesas.Recordset("codigodespesa")
  .Recordset("descri") = dbDespesas.Recordset("descri")
  .Recordset("obs") = Left(txtDespesaObs.Text, 50)
  .Recordset!compensado = True
  .Recordset("valor") = -CCur(txtDespesaValor.Text)
  .Recordset!valorpago = -CCur(txtDespesaValor.Text)
  .Recordset.Update
  .Refresh
End With

Totaliza
cboDespesa.Text = ""
txtDespesaValor.Text = ""
'txtDespesaObs.Text = ""


cboDespesa.SetFocus
End Sub

Private Sub cmdIncluirRecebimento_Click()
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency, DescontoPorcento As Currency

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

With dbFormaDePgRecebido
  .Recordset.AddNew
  .Recordset("codigofechamento") = CodigoFechamento
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
  .Refresh
End With

Totaliza

cboRecebimento.Text = ""
txtValorRecebe.Text = ""
txtOperacoes.Text = ""
cboRecebimento.SetFocus

End Sub

Private Sub cmdIncluirTanque_Click()
Dim CodigoProduto As Double, Quantidade As Double, Diferenca As Double
Dim Descri As String, Estoque As Double, Vendido As Double

If dbTanques.Recordset.EOF = True Then
  MsgBox "Tanque inválido!", vbCritical, "Erro!"
  cboTanque.SetFocus
  Exit Sub
End If
If CDbl(cboTanque.Text) <> dbTanques.Recordset("tanque") Then
  MsgBox "Tanque inválido!", vbCritical, "Erro!"
  cboTanque.SetFocus
  Exit Sub
End If
If IsNumeric(txtRegua.Text) = False Then
  MsgBox "Número inválido!", vbCritical, "Erro!"
  txtRegua.SetFocus
  Exit Sub
End If
CodigoProduto = dbTanques.Recordset!CodigoProduto
If IsNumeric(txtReposicao.Text) = False Then
  txtReposicao.Text = "0"
End If
With dbProdutos
  .RecordSource = "select *from produtos where codigoproduto=" & CodigoProduto
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Produto do tanque não encontrado na tabela de produtos!"
    Exit Sub
  Else
    Descri = .Recordset!Descri
    Estoque = .Recordset!Estoque
  End If
  .RecordSource = "Select *from produtos"
  .Refresh
End With

With dbTanquesMovimento
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Recordset.FindFirst "tanque=" & dbTanques.Recordset("tanque")
    If .Recordset.NoMatch = True Then
      .Recordset.AddNew
    Else
      MsgBox "Tanque já lançado!"
      Exit Sub
    End If
  End If
  .Recordset("codigofechamento") = CodigoFechamento
  .Recordset("codigoposto") = dbPosto.Recordset("codigoposto")
  .Recordset("tanque") = dbTanques.Recordset("tanque")
  .Recordset("data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("quantidade") = CDbl(txtRegua.Text)
  .Recordset("reposicao") = CDbl(txtReposicao.Text)
  .Recordset("estoqueantes") = dbTanques.Recordset("estoque")
  .Recordset("estoquedepois") = 0
  .Recordset.Update
  .Refresh
End With
Quantidade = CDbl(txtRegua.Text)
If qGalonagem.Recordset.RecordCount <> 0 Then
  With qGalonagem
    .Recordset.MoveFirst
    .Recordset.FindFirst "produtos.codigoproduto=" & CodigoProduto
    If .Recordset.NoMatch = False Then
      Vendido = .Recordset!Vendido
    Else
      Vendido = 0
    End If
  End With
End If

With dbDifComb
  .RecordSource = "select *from DiferencaCombustivel where codigofechamento=" & CodigoFechamento & " and codigoproduto=" & CodigoProduto
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Recordset.Edit
  End If
  If IsNull(.Recordset!Diferenca) = False Then
    Quantidade = Quantidade + .Recordset!Tanque
  End If
  Diferenca = Quantidade - Estoque
  If dbPosto.Recordset!medetanqueantes = False Then
    Diferenca = Diferenca + Vendido
  End If
  .Recordset!CodigoFechamento = CodigoFechamento
  .Recordset!CodigoProduto = CodigoProduto
  .Recordset!Descri = Descri
  .Recordset!Tanque = Quantidade
  .Recordset!Estoque = Estoque
  .Recordset!Diferenca = Diferenca
  .Recordset!Vendido = Vendido
  .Recordset.Update
  .RecordSource = "select *from DiferencaCombustivel where codigofechamento=" & CodigoFechamento
  .Refresh
End With
Totaliza
txtRegua.Text = ""
txtReposicao.Text = ""
cboTanque.SetFocus

End Sub

Private Sub cmdIncluirVendas_Click()
Dim CodigoProduto As Double, codigoPosto As Double, Desconto As Currency
Dim Descricao As String, Qtd As Double, ValorUnitario As Currency
Dim ValorTotal As Currency, CodigoVendedor As Double, ValorDesconto As Currency
Dim Comissao As Double, ValorComissao As Currency
Dim CodProduto As Double
If dbProdutos2.Recordset!Descri <> cboProduto.Text Then
  MsgBox "Produto inválido!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If txtCodProduto.Text = "" Then
  MsgBox "Escolha um Produto a ser incluído na lista de Vendidos!", vbCritical, "Erro!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If cboProduto.Text = "" Then
  MsgBox "Escolha um Produto a ser incluído na lista de Vendidos!", vbCritical, "Erro!"
  cboProduto.SetFocus
  Exit Sub
End If
If IsNumeric(txtProdutoQuantidade.Text) = False Then
  MsgBox "Informe uma quantidade correta!"
  txtProdutoQuantidade.SetFocus
  Exit Sub
End If
If IsNumeric(txtCodVendedor.Text) = False Then
  txtCodVendedor.Text = 0
End If
Desconto = 0
If IsNumeric(txtDesconto.Text) = True Then
  Desconto = CDbl(txtDesconto.Text)
End If
Qtd = CDbl(txtProdutoQuantidade.Text)
With dbProdutos2
  If .Recordset!comissaovalor <> 0 Then
    If txtCodVendedor.Text = "0" Then
      If Desconto = 0 Then
        MsgBox "Informe o código do funcionário!"
        txtCodVendedor.SetFocus
        Exit Sub
      End If
    End If
  End If
  If .Recordset!Comissao <> 0 Then
    If txtCodVendedor.Text = "0" Then
      If Desconto = 0 Then
        MsgBox "Informe o código do funcionário!"
        txtCodVendedor.SetFocus
        Exit Sub
      End If
    End If
  End If
  CodigoProduto = .Recordset("codigoproduto")
  Descricao = .Recordset("descri")
  ValorUnitario = .Recordset("precovenda")
  ValorTotal = ValorUnitario * Qtd
  If IsNull(.Recordset("comissao")) = False Then
    Comissao = .Recordset("comissao")
  End If
  If IsNumeric(txtDesconto.Text) = True Then
    If CCur(txtDesconto.Text) <> 0 Then
      Desconto = CCur(txtDesconto.Text)
      ValorComissao = 0
      ValorTotal = ValorTotal - Desconto
    Else
      Desconto = 0
      ValorComissao = ValorTotal * Comissao
      ValorComissao = ValorComissao + (.Recordset!comissaovalor * Qtd)
    End If
  Else
    Desconto = 0
    ValorComissao = ValorTotal * Comissao
    ValorComissao = ValorComissao + (.Recordset!comissaovalor * Qtd)
  End If
  CodProduto = .Recordset("codigo")
End With
CodigoVendedor = txtCodVendedor.Text
If dbProdutos2.Recordset!Comissao = 0 And dbProdutos2.Recordset!comissaovalor = 0 Then
  CodigoVendedor = 0
End If

codigoPosto = dbPosto.Recordset("codigoposto")

With dbVendas
  .Recordset.AddNew
  .Recordset("codigoposto") = codigoPosto
  .Recordset("codigofechamento") = CodigoFechamento
  .Recordset("data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("codigoproduto") = CodigoProduto
  .Recordset("codproduto") = CodProduto
  .Recordset("descri") = Descricao
  .Recordset("quantidade") = Qtd
  .Recordset("valorunitario") = ValorUnitario
  .Recordset("valortotal") = ValorTotal
  .Recordset("codigovendedor") = CodigoVendedor
  .Recordset("comissao") = Comissao
  .Recordset("valorcomissao") = ValorComissao
  .Recordset!CodigoVendedor = CodigoVendedor
  .Recordset!ValorDesconto = Desconto
  .Recordset.Update
  .Refresh
End With

Totaliza
txtCodProduto.Text = ""
cboProduto.Text = ""
txtProdutoQuantidade.Text = ""
lblProdutoTotal.Caption = ""
txtCodProduto.SetFocus

End Sub

Private Sub cmdInlueBomba_Click()

If IsDate(txtData.Value) = False Then
  MsgBox "Data inválida!", vbCritical, "Erro!"
  txtData.SetFocus
  Exit Sub
End If
If cboTurno.Text <> dbTurno.Recordset!Descri Then
  MsgBox "Turno inválido!"
  cboTurno.SetFocus
  Exit Sub
End If
Screen.MousePointer = vbHourglass
With dbFechamento
  .RecordSource = "select *from FechamentoDiario where codigoposto=" & dbPosto.Recordset("codigoPosto") & " and data=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and cancelado=0 and codigoturno=" & dbTurno.Recordset!codigoturno
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If .Recordset("confirmado") = True Then
      CodigoFechamento = .Recordset!CodigoFechamento
      Confirmado
    Else
      If dbResponsavel.Recordset.EOF = True Then
        MsgBox "Responsável inválido!", vbCritical, "Erro!"
        cboResponsavel.SetFocus
        Exit Sub
      End If
      If dbResponsavel.Recordset("nome") <> cboResponsavel.Text Then
        MsgBox "Responsável inválido!", vbCritical, "Erro!"
        cboResponsavel.SetFocus
        Exit Sub
      End If
      NaoConfirmado
    End If
  End If
  .RecordSource = "select *from FechamentoDiario where codigoposto=" & dbPosto.Recordset("codigoPosto") & " and  data=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and cancelado=0 and codigoturno=" & dbTurno.Recordset!codigoturno
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset("codigoposto") = dbPosto.Recordset("codigoposto")
    .Recordset("CodigoResponsavel") = dbResponsavel.Recordset("codigovendedor")
    .Recordset("data") = txtData.Value
    .Recordset("hora") = Now
    .Recordset!codigoturno = dbTurno.Recordset!codigoturno
    .Recordset!Turno = dbTurno.Recordset!Descri
    .Recordset.Update
    .Refresh
  End If
  
  CodigoFechamento = .Recordset("codigofechamento")
  
End With
AbreFechamento CodigoFechamento, dbPosto.Recordset("codigoposto")


intTela = 0
Totaliza
SomaDeclarados
Paginas.Visible = True
Paginas.Tab = 0
cmdProximo.Enabled = True
cmdInlueBomba.Enabled = False
cmdCancelar.Caption = "&Cancelar"
With dbPosto
  If .Recordset!leituramecanica = True Then
    txtMecanico.Enabled = True
  Else
    txtMecanico.Enabled = False
  End If
End With
With dbFechamento
  If IsNumeric(lblTotalChequeResumo.Caption) = True Then
    .Recordset.Edit
    .Recordset!chequeavista = CCur(lblTotalChequeResumo.Caption)
    On Error Resume Next
    .Recordset.Update
    On Error GoTo 0
  End If
End With

cboResponsavel.Enabled = False
txtData.Enabled = False
cboTurno.Enabled = False
Screen.MousePointer = vbDefault
If cmdFinalizar.Visible = True Then
  On Error Resume Next
  cboBico.SetFocus
End If
End Sub

Private Sub cmdProdutoEntra_Click()
Dim PrecoAntigo As Currency, PrecoNovo As Currency, TempValor As Currency
Dim Variacao As Currency

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
TempValor = PrecoNovo + dbProdutos.Recordset!comissaovalor + (dbProdutos.Recordset!PrecoVenda * (dbProdutos.Recordset!Comissao / 100))

If TempValor >= (dbProdutos.Recordset!PrecoVenda / 2) Then
  MsgBox "Margem de lucro abaixo de 50%! Custo= " & Format(TempValor, "Currency") & " / Venda= " & Format(dbProdutos.Recordset!PrecoVenda, "Currency"), vbCritical
End If

With dbProdutoEntra
  .Recordset.AddNew
  .Recordset!CodigoFechamento = CodigoFechamento
  .Recordset!Data = txtData.Value
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
With QProdutoEntra
  .Refresh
End With
txtCod.SetFocus
End Sub

Private Sub cmdProximo_Click()
For i = 0 To Tela.Count - 1
  Tela(i).Visible = True
Next i
intTela = Paginas.Tab
If intTela < 0 Then
  Exit Sub
Else
  intTela = intTela + 1
  Paginas.Tab = intTela
  cmdAnterior.Enabled = True
  If intTela = Paginas.Tabs - 1 Then
    cmdProximo.Enabled = False
    cmdFinalizar.Enabled = True
    cmdResumo.Enabled = True
  End If
End If
Screen.MousePointer = vbHourglass
Totaliza
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRelaciona_Click()
If IsDate(MaskEdBox1(5).Text) = False Then
  MsgBox "Data inválida!"
  MaskEdBox1(5).SetFocus
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
If cboClienteCheque.Text <> dbClientesCheques.Recordset!Nome Then
  MsgBox "Não foi cadastrado o cliente!"
  txtCodCliente.SetFocus
  Exit Sub
End If
With dbCheques
  .Refresh
  .Recordset.FindFirst "comp='" & MaskEdBox1(0).Text & "' and banco='" & MaskEdBox1(1).Text & "' and agencia='" & MaskEdBox1(2).Text & "' and conta='" & MaskEdBox1(3).Text & "' and chequeNr='" & MaskEdBox1(4).Text & "'"
  If .Recordset.NoMatch = False Then
    MsgBox "Cheque já cadastrado!"
    Exit Sub
  End If
  .Recordset.AddNew
  .Recordset!CodigoFechamento = CodigoFechamento
  .Recordset!cmc7 = CodBar
  .Recordset!comp = MaskEdBox1(0).Text
  .Recordset!banco = MaskEdBox1(1).Text
  .Recordset!agencia = MaskEdBox1(2).Text
  .Recordset!conta = MaskEdBox1(3).Text
  .Recordset!chequenr = MaskEdBox1(4).Text
  .Recordset!datalanc = Now
  .Recordset!datacheque = MaskEdBox1(5).Text
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset!CodigoSoma = "2"
  .Recordset!valornabomba = CCur(lblValorNaBomba.Caption)
  .Recordset!diaspre = lblJurosTabelado.Caption
  If IsNull(dbClientesCheques.Recordset!CIC) = False Then
    If dbClientesCheques.Recordset!CIC = "" Then
      .Recordset!cpf = dbClientesCheques.Recordset!CNPJ
    Else
      .Recordset!cpf = dbClientesCheques.Recordset!CIC
    End If
  Else
    .Recordset!cpf = dbClientesCheques.Recordset!CNPJ
  End If
  .Recordset.Update
  .Refresh
End With

With qCheques
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      Label54.Caption = Format(.Recordset!Total, "Currency")
    Else
      Label54.Caption = Format(0, "Currency")
    End If
  Else
    Label54.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
txtValor.Text = ""
lblNome.Caption = ""
lblStatus = ""
txtCodCliente.Text = ""
cboClienteCheque.Text = ""
MaskEdBox1(0).SetFocus

End Sub

Private Sub cmdRemover_Click()
If dbBicoMovimento.Recordset.RecordCount = 0 Then Exit Sub
If dbBicoMovimento.Recordset.EOF = True Then Exit Sub
dbBicoMovimento.Recordset.Delete
End Sub

Private Sub cmdResumo_Click()
Dim StrTemp As String, Largura As Double
Dim X1 As Double, X2 As Double
Dim Y1 As Double, Y2 As Double
Dim Total As Double, Estoque As Double, Tanque As Double, Diferenca As Double

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Largura = 195

Printer.ScaleMode = vbMillimeters
Printer.DrawWidth = 2

Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Demonstrativo de Movimento"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp
Printer.FontBold = False

Printer.FontSize = 8
Printer.Line (0, 9)-(Largura, 9)
Printer.Line (0, 21)-(Largura, 21)
Printer.Line (78, 9)-(78, 21)
Printer.Line (137, 9)-(137, 21)
Printer.Line (170, 9)-(170, 21)

Printer.CurrentY = 10
Printer.CurrentX = 0
Printer.Print "Nome do Posto";
Printer.CurrentX = 80
Printer.Print "Nome do Responsável";
Printer.CurrentX = 139
Printer.Print "Data do Movimento";
Printer.CurrentX = 172
Printer.Print "Turno"

Printer.CurrentX = 0
Printer.Print NomePosto;
Printer.CurrentX = 80
Printer.Print cboResponsavel.Text;
Printer.CurrentX = 139
Printer.Print Format(txtData.Value, "Short Date");
Printer.CurrentX = 172
Printer.Print cboTurno.Text

Printer.Line (0, 23)-(Largura, 23)
Printer.Line (0, 29)-(Largura, 29)

Printer.CurrentY = 24
Printer.CurrentX = 0
Printer.Print "Bicos";
Printer.CurrentX = 16
Printer.Print "Inicial";
Printer.CurrentX = 53
Printer.Print "Final";
Printer.CurrentX = 99
Printer.Print "Vendas";
Printer.CurrentX = 119
Printer.Print "Valor";
Printer.CurrentX = 142
Printer.Print "Vendas";
Printer.CurrentX = 174
Printer.Print "Retorno"

Printer.CurrentY = 30
With dbBicoMovimento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.CurrentX = 0
      Printer.Print .Recordset!Bico;
      StrTemp = Format(.Recordset!ValorInicial, "#,##0.00")
      Printer.CurrentX = 49 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!ValorFinal, "#,##0.00")
      Printer.CurrentX = 92 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Vendas, "#,##0.00")
      Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!precounitario, "#,##0.000")
      Printer.CurrentX = 137 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valorvendido, "#,##0.000")
      Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Retorno, "#,##0.00")
      Printer.CurrentX = Largura - 2 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      
      .Recordset.MoveNext
    Loop
    Y2 = Printer.CurrentY
    Printer.Line (14, 23)-(14, Y2)
    Printer.Line (51, 23)-(51, Y2)
    Printer.Line (97, 23)-(97, Y2)
    Printer.Line (117, 23)-(117, Y2)
    Printer.Line (140, 23)-(140, Y2)
    Printer.Line (172, 23)-(172, Y2)
  End If
  Printer.CurrentY = Y2 + 2
  Y2 = Printer.CurrentY
  
  Printer.Line (0, Y2)-(Largura, Y2)
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.CurrentX = 0
  Printer.Print "Movimento de Combustível"
  Printer.CurrentY = Printer.CurrentY + 0.5
  Y1 = Printer.CurrentY
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  
  Printer.CurrentX = 0
  Printer.Print "Combustível";
  Printer.CurrentX = 38.5
  Printer.Print "Vendido";
  Printer.CurrentX = 77
  Printer.Print "Estoque Sist.";
  Printer.CurrentX = 115.5
  Printer.Print "Estoque Posto";
  Printer.CurrentX = 154
  Printer.Print "Diferença"
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  Y2 = Printer.CurrentY
   
  With qGalonagem
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        X1 = 0
        StrTemp = .Recordset!Descri
        Printer.CurrentX = X1
        Printer.Print StrTemp;
        Total = Total + .Recordset!Vendido
        StrTemp = Format(.Recordset!Vendido, "#,##0.00")
        X1 = X1 + 38.5
        X1 = X1 + 38.5
        X2 = X1 - 2 - Printer.TextWidth(StrTemp)
        Printer.CurrentX = X2
        Printer.Print StrTemp;
        If dbDifComb.Recordset.RecordCount <> 0 Then
          dbDifComb.Recordset.MoveFirst
          dbDifComb.Recordset.FindFirst "codigoproduto=" & .Recordset("produtos.CodigoProduto")
          If dbDifComb.Recordset.NoMatch = False Then
            Estoque = Estoque + dbDifComb.Recordset!Estoque
            StrTemp = Format(dbDifComb.Recordset!Estoque, "#,##0.00")
            X1 = X1 + 38.5
            X2 = X1 - 2 - Printer.TextWidth(StrTemp)
            Printer.CurrentX = X2
            Printer.Print StrTemp;
            Tanque = Tanque + dbDifComb.Recordset!Tanque
            StrTemp = Format(dbDifComb.Recordset!Tanque, "#,##0.00")
            X1 = X1 + 38.5
            X2 = X1 - 2 - Printer.TextWidth(StrTemp)
            Printer.CurrentX = X2
            Printer.Print StrTemp;
            Diferenca = Diferenca + dbDifComb.Recordset!Diferenca
            StrTemp = Format(dbDifComb.Recordset!Diferenca, "#,##0.00")
            X1 = X1 + 38.5
            X2 = X1 - 2 - Printer.TextWidth(StrTemp)
            Printer.CurrentX = X2
            Printer.Print StrTemp;
          End If
        End If
        Printer.Print ""
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.5
        .Recordset.MoveNext
      Loop
      Y2 = Printer.CurrentY - 1
      X1 = 37.5
      For i = 1 To 4
        Printer.Line (X1, Y1)-(X1, Y2)
        X1 = X1 + 38.5
      Next i
    End If
    X1 = 0
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 0
    Printer.Print "Total";
    StrTemp = Format(Total, "#,##0.00")
    X1 = X1 + 38.5
    X1 = X1 + 38.5
    X2 = X1 - 2 - Printer.TextWidth(StrTemp)
    Printer.CurrentX = X2
    Printer.Print StrTemp;
    StrTemp = Format(Estoque, "#,##0.00")
    X1 = X1 + 38.5
    X2 = X1 - 2 - Printer.TextWidth(StrTemp)
    Printer.CurrentX = X2
    Printer.Print StrTemp;
    StrTemp = Format(Tanque, "#,##0.00")
    X1 = X1 + 38.5
    X2 = X1 - 2 - Printer.TextWidth(StrTemp)
    Printer.CurrentX = X2
    Printer.Print StrTemp;
    StrTemp = Format(Diferenca, "#,##0.00")
    X1 = X1 + 38.5
    X2 = X1 - 2 - Printer.TextWidth(StrTemp)
    Printer.CurrentX = X2
    Printer.Print StrTemp
  End With
  
  With dbVendas
    .RecordSource = "select *from venda where codigofechamento=" & CodigoFechamento & " order by codigoproduto"
    .Refresh
    Printer.CurrentY = Printer.CurrentY + 1
    StrTemp = "Venda de Produtos"
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    Y1 = Printer.CurrentY
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    X1 = 0
    For i = 1 To 3
      Printer.CurrentX = 1 + X1
      Printer.Print "Cod.";
      Printer.CurrentX = 10 + X1
      Printer.Print "Qtd.";
      Printer.CurrentX = 20 + X1
      Printer.Print "Valor";
      Printer.CurrentX = 46 + X1
      Printer.Print "Cod.Func.";
      
      X1 = X1 + 65
    Next i
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    X1 = 0
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        For i = 1 To 3
          If .Recordset.EOF = True Then Exit For
          StrTemp = .Recordset!CodProduto
          Printer.CurrentX = 8 + X1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          StrTemp = .Recordset!Quantidade
          Printer.CurrentX = 18 + X1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          StrTemp = Format(.Recordset!ValorTotal, "Currency")
          Printer.CurrentX = 43 + X1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          TempValor = TempValor + .Recordset!ValorTotal
          StrTemp = .Recordset!CodigoVendedor
          Printer.CurrentX = 63 + X1 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          .Recordset.MoveNext
          
          X1 = X1 + 65
        Next i
        X1 = 0
        Printer.Print ""
        'If .Recordset.EOF = False Then .Recordset.MoveNext
      Loop
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Y2 = Printer.CurrentY
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      
      For i = 1 To 3
        Printer.Line (9 + X1, Y1)-(9 + X1, Y2)
        Printer.Line (19 + X1, Y1)-(19 + X1, Y2)
        Printer.Line (44 + X1, Y1)-(44 + X1, Y2)
        If i < 3 Then
          Printer.Line (64 + X1, Y1)-(64 + X1, Y2)
        End If
        X1 = X1 + 65
      Next i
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      StrTemp = "Total"
      Printer.CurrentX = 173 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(TempValor, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = Printer.CurrentY + 0.5
    End If
  End With
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Y1 = Printer.CurrentY
  With dbFormaDePgRecebido
    Printer.Line (0, Y1)-(97, Y1)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 0
    Printer.Print "Conta";
    StrTemp = "Valor"
    Printer.CurrentX = 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(97, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.CurrentX = 0
        Printer.Print .Recordset!Descri;
        StrTemp = Format(.Recordset!ValorBruto, "Currency")
        Printer.CurrentX = 95 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.Line (0, Printer.CurrentY)-(97, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.5
        
        .Recordset.MoveNext
      Loop
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.CurrentX = 0
      Printer.Print "Notas de Clientes";
      StrTemp = Format(lblClientesNota.Caption, "Currency")
      Printer.CurrentX = 95 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(97, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.CurrentX = 0
      Printer.Print "Cheques";
      StrTemp = Format(lblTotalChequeResumo.Caption, "Currency")
      Printer.CurrentX = 95 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(97, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
'      Printer.CurrentY = Printer.CurrentY + 0.5
'      Printer.CurrentX = 0
'      Printer.Print "Cheques Pré-Datado";
'      StrTemp = Format(txtChequesPre.Text, "Currency")
'      Printer.CurrentX = 95 - Printer.TextWidth(StrTemp)
'      Printer.Print StrTemp
'      Printer.CurrentY = Printer.CurrentY + 0.5
'      Printer.Line (0, Printer.CurrentY)-(97, Printer.CurrentY)
'      Printer.CurrentY = Printer.CurrentY + 0.5
    End If
    
    Y2 = Printer.CurrentY
    Printer.CurrentY = Y1
    
    Printer.Line (98, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 99
    Printer.Print "Produto";
    StrTemp = "Valor"
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (98, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 99
    Printer.Print "Venda de Combustível";
    StrTemp = Format(lblVendasCombustivel.Caption, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (98, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 99
    Printer.Print "Venda de Produtos";
    StrTemp = Format(lblVendasProdutos.Caption, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (98, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 99
    Printer.Print "Juros";
    If IsNumeric(txtJuros.Text) = True Then
      StrTemp = Format(txtJuros.Text, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (98, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    If Printer.CurrentY > Y2 Then
      Y2 = Printer.CurrentY
    End If
    Printer.Line (0, Y2)-(97, Y2)
    Printer.Line (98, Y2)-(Largura, Y2)
    
    Printer.Line (46.5, Y1)-(46.5, Y2)
    Printer.Line (97.5, Y1)-(97.5, Y2)
    Printer.Line (148, Y1)-(148, Y2)
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y1 = Printer.CurrentY
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 0
    Printer.Print "Total de Vendas";
    Printer.CurrentX = 62
    Printer.Print "Total das Despesas";
    Printer.CurrentX = 107
    Printer.Print "Total dos Recebimentos";
    Printer.CurrentX = 168
    Printer.Print "Diferença"
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentX = 0
    Printer.Print lblTotalVendas.Caption;
    Printer.CurrentX = 62
    Printer.Print lblTotalDespesas.Caption;
    Printer.CurrentX = 107
    Printer.Print lblRecebimentos.Caption;
    Printer.CurrentX = 168
    Printer.Print lblDiferenca.Caption;
    
    Printer.CurrentY = Printer.CurrentY + 5
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Y2 = Printer.CurrentY
    
    Printer.Line (60, Y1)-(60, Y2)
    Printer.Line (105, Y1)-(105, Y2)
    Printer.Line (166, Y1)-(166, Y2)
    
  End With
End With

With QComissoes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Comissões"
    Printer.CurrentX = 0
    Printer.Print StrTemp
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    B1 = Printer.CurrentY
    Printer.Line (0, B1)-(110, B1)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    StrTemp = "Cod."
    Printer.CurrentX = 1
    Printer.Print StrTemp;
    
    StrTemp = "Nome"
    Printer.CurrentX = 16
    Printer.Print StrTemp;
    
    StrTemp = "Comissão"
    Printer.CurrentX = 81
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Line (0, Printer.CurrentY)-(110, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Do While .Recordset.EOF = False
      dbResponsavel.Refresh
      If dbResponsavel.Recordset.RecordCount <> 0 Then
        dbResponsavel.Recordset.FindFirst "codigo=" & .Recordset!CodigoVendedor
        If dbResponsavel.Recordset.NoMatch = False Then
          
          StrTemp = .Recordset!CodigoVendedor
          Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          
          StrTemp = dbResponsavel.Recordset!Nome
          Printer.CurrentX = 16
          Printer.Print StrTemp;
          
          StrTemp = Format(.Recordset!Comissao, "Currency")
          Printer.CurrentX = 109 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp
          
        End If
      End If
      .Recordset.MoveNext
    Loop
    B2 = Printer.CurrentY
    If B2 < B1 Then B1 = 0
    Printer.Line (0, B2)-(110, B2)
    Printer.Line (0, B1)-(0, B2)
    Printer.Line (15, B1)-(15, B2)
    Printer.Line (80, B1)-(80, B2)
    Printer.Line (110, B1)-(110, B2)
  End If
End With

Printer.Print ""

StrTemp = "Encerrante de Pontuação: " & txtPontos.Text
Printer.Print StrTemp

Printer.EndDoc

NaoImprime:

End Sub


Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
Dim Resposta As Integer, UltimoNumero As Double
Dim CodigoProduto As Double, LucroVenda As Currency
Dim Quantidade As Double

With dbBicoMovimento
  Resposta = MsgBox("Deseja remover o lançamento atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then
    Cancel = True
    Exit Sub
  End If
  LucroVenda = (.Recordset!precounitario - .Recordset!precocompra) * .Recordset!Vendas
  Quantidade = .Recordset!Vendas
  CodigoProduto = .Recordset!CodigoProduto
  UltimoNumero = .Recordset!ValorFinal
  dbBico.Refresh
  dbBico.Recordset.FindFirst "codigobico=" & .Recordset!codigobico
  If dbBico.Recordset.NoMatch = False Then
    If dbBico.Recordset!provisorionumero <> UltimoNumero Then
      MsgBox "Não pode remover o lançamento atual pois existe um lançamento posterior!"
      Cancel = True
      Exit Sub
    Else
      dbBico.Recordset.Edit
      dbBico.Recordset!provisorionumero = .Recordset!ValorInicial
      dbBico.Recordset.Update
      dbBico.Refresh
    End If
  End If
End With
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyDelete
    Dim Quantidade As Double, CodigoProduto As Double
    Quantidade = dbTanquesMovimento.Recordset!Quantidade
    dbTanques.Recordset.MoveFirst
    dbTanques.Recordset.FindFirst "tanque=" & dbTanquesMovimento.Recordset!Tanque
    If dbTanques.Recordset.NoMatch = False Then
      CodigoProduto = dbTanques.Recordset!CodigoProduto
    Else
      MsgBox "Erro na tabela de tanques"
      Cancel = True
      Exit Sub
    End If
    With dbDifComb
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveFirst
        .Recordset.FindFirst "codigoproduto=" & CodigoProduto
        If .Recordset.NoMatch = False Then
          .Recordset.Edit
          .Recordset!Tanque = .Recordset!Tanque - Quantidade
          Quantidade = .Recordset!Tanque - .Recordset!Estoque
          If dbPosto.Recordset!medetanqueantes = False Then
            Quantidade = Quantidade + Vendido
          End If
          .Recordset!Diferenca = Quantidade
          .Recordset.Update
          .Refresh
        End If
      End If
    End With
    dbTanquesMovimento.Recordset.Delete
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select
End Sub

Private Sub Form_Load()
StrTemp = GetSetting(App.EXEName, "Base", "COM")

strTemp2 = GetSetting(App.EXEName, "Base", "Baud", "9600")
strTemp2 = strTemp2 & "," & GetSetting(App.EXEName, "Base", "Paridade", "n")
strTemp2 = strTemp2 & "," & GetSetting(App.EXEName, "Base", "DataBit", "8")
strTemp2 = strTemp2 & "," & GetSetting(App.EXEName, "Base", "StopBit", "1")

MSComm1.Settings = strTemp2

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
txtJuros.Text = Format(0, "Currency")
txtDataBordero.Value = Date
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConciliaNova
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
With QTemp
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from fechamentodiario where cancelado=0 order by codigofechamento"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    lblUltimoLancado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
    If .Recordset!Confirmado = 0 Then
      On Error Resume Next
      Do While .Recordset!Confirmado = 0
        If .Recordset.BOF = True Then Exit Do
        .Recordset.MovePrevious
      Loop
    End If
    lblUltimoConfirmado.Caption = .Recordset!Data & " - Turno: " & .Recordset!Turno
  End If
  On Error GoTo 0
End With


AbreFechamento 0, 0
CodigoFechamento = 0
Totaliza
intTela = -1
'For i = 0 To Tela.Count - 1
'  Tela(i).Visible = False
'Next i
txtData.Value = Date
End Sub

Private Sub Image1_DblClick()
On Error Resume Next
With MSComm1
  If .PortOpen = True Then
    .PortOpen = False
  Else
    .PortOpen = True
  End If
  If .PortOpen = False Then
    Image1.Picture = LoadResPicture(102, vbResBitmap)
  Else
    Image1.Picture = LoadResPicture(101, vbResBitmap)
  End If
End With
End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)

With MaskEdBox1(Index)
  If Index = 5 Then
    .SelStart = 0
    .SelLength = 2
  Else
    .SelStart = 0
    .SelLength = Len(.Text)
  End If
End With

End Sub

Private Sub Paginas_Click(PreviousTab As Integer)
Call txtDinheiro_LostFocus
Select Case Paginas.Tab
  Case Paginas.Tabs - 1
    cmdAnterior.Enabled = True
    cmdProximo.Enabled = False
    cmdResumo.Enabled = True
    cmdFinalizar.Enabled = True
  Case 0
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = True
    cmdResumo.Enabled = False
    cmdFinalizar.Enabled = False
  Case Else
    cmdAnterior.Enabled = True
    cmdProximo.Enabled = True
    cmdResumo.Enabled = False
    cmdFinalizar.Enabled = False
End Select
Select Case Paginas.Tab
  Case 0
    If Tela(0).Enabled = True Then cboBico.SetFocus
  Case 1
    If Tela(0).Enabled = True Then cboTanque.SetFocus
  Case 2
    If Tela(0).Enabled = True Then txtCodProduto.SetFocus
  Case 3
    If Tela(0).Enabled = True Then cboDespesa.SetFocus
  Case 4
    If Tela(0).Enabled = True Then txtDinheiro.SetFocus
  Case 5
    If Tela(0).Enabled = True Then cboClientesNota.SetFocus
  Case 6
    If Tela(0).Enabled = True Then MaskEdBox1(0).SetFocus
  Case 7
    If Tela(0).Enabled = True Then txtCod.SetFocus
  Case 8
    cmdResumo.SetFocus
End Select
Screen.MousePointer = vbHourglass
Totaliza
Screen.MousePointer = vbDefault
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
      
      txtCodCliente.SetFocus
      
      With dbCheques
        .Refresh
        If .Recordset.RecordCount <> 0 Then
          .Recordset.FindFirst "comp='" & MaskEdBox1(0) & "' and banco='" & MaskEdBox1(1) & "' and agencia='" & MaskEdBox1(2) & "' and conta='" & MaskEdBox1(3) & "' and chequenr='" & MaskEdBox1(4) & "'"
          If .Recordset.NoMatch = False Then
            MaskEdBox1(5).Text = .Recordset("datacheque")
            txtValor.Text = Format(.Recordset("valor"), "Currency")
          End If
        End If
      End With
    End If
  End If
  Timer1.Enabled = True
End If

End Sub

Private Sub txtBicoEncerra_GotFocus()
With txtBicoEncerra
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtBicoEncerra_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtCartoes_GotFocus()
With txtCartoes
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCartoes_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtCartoes_LostFocus()
SomaDeclarados
End Sub

Private Sub txtChequeRecebido_GotFocus()
With txtChequeRecebido
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtChequeRecebido_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtChequeRecebido_LostFocus()
SomaDeclarados
End Sub

Private Sub txtChequesPre_GotFocus()
With txtChequesPre
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtChequesPre_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtChequesPre_LostFocus()
SomaDeclarados
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
    If .Recordset!posicao = True Then
      lblStatus.ForeColor = vbBlack
      lblStatus.Caption = "Ativo"
    Else
      lblStatus.ForeColor = vbRed
      lblStatus.Caption = "Inativo"
    End If
  End If
End With
End Sub

Private Sub txtCodProduto_GotFocus()
With txtCodProduto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos2
  If txtCodProduto.Text = "" Then Exit Sub
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto.Text
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset("descri")
  End If
End With
End Sub

Private Sub txtCodVendedor_GotFocus()
txtCodVendedor.SelStart = 0
txtCodVendedor.SelLength = Len(txtCodVendedor.Text)
End Sub

Private Sub txtCupom_GotFocus()
With txtCupom
  .SelStart = 0
  .SelLength = Len(.Text)
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
txtDataBordero.Value = txtData.Value
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

Private Sub txtDesconto_GotFocus()
With txtDesconto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDesconto_LostFocus()
With txtDesconto
  If IsNumeric(.Text) = False Then
    .Text = 0
  End If
  .Text = Format(.Text, "currency")
End With
TotalizaVenda
End Sub

Private Sub txtDespesaObs_GotFocus()
txtDespesaObs.SelStart = 0
txtDespesaObs.SelLength = Len(txtDespesaObs.Text)
End Sub

Private Sub txtDespesas_GotFocus()
With txtDespesas
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDespesas_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtDespesas_LostFocus()
SomaDeclarados
End Sub

Private Sub txtDespesaValor_KeyPress(KeyAscii As Integer)
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

Private Sub txtDinheiro_GotFocus()
With txtDinheiro
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDinheiro_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtDinheiro_LostFocus()
SomaDeclarados
End Sub

Private Sub txtMecanico_GotFocus()
With txtMecanico
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMecanico_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtNotas_GotFocus()
With txtNotas
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtNotas_LostFocus()
SomaDeclarados
End Sub

Private Sub txtNotaValor_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtNotaValor_LostFocus()
If IsNumeric(txtNotaValor.Text) = False Then Exit Sub
txtNotaValor.Text = Format(txtNotaValor.Text, "Currency")
End Sub

Private Sub txtProdutoQuantidade_GotFocus()
With txtProdutoQuantidade
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtProdutoQuantidade_LostFocus()
TotalizaVenda
End Sub

Private Sub txtQtdEntra_GotFocus()
With txtQtdEntra
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRetorno_GotFocus()
txtRetorno.SelStart = 0
txtRetorno.SelLength = Len(txtRetorno.Text)
End Sub

Private Sub txtRetorno_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtTickets_GotFocus()
With txtTickets
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTickets_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtTickets_LostFocus()
SomaDeclarados
End Sub

Private Sub txtTotalEntra_GotFocus()
With txtTotalEntra
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotalEntra_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtTotalEntra_LostFocus()
Dim Qtd As Double, ICMS As Double, IPI As Double, Unitario As Currency
Dim Total As Currency

'lblUnitarioCalc.Caption = ""

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

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
Dim Dias As Double, Taxa As Double
Dim Valor As Currency
With txtValor
  lblValorNaBomba.Caption = ""
  lblJurosTabelado.Caption = "ERR"
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
  Dias = DateDiff("d", txtData.Value, CDate(MaskEdBox1(5).Text))
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

Private Sub txtValorRecebe_KeyPress(KeyAscii As Integer)
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

Private Sub txtVT_GotFocus()
With txtVT
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVT_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtVT_LostFocus()
SomaDeclarados
End Sub
