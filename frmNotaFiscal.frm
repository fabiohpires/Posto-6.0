VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClienteNotaFiscal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de nota fiscal"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   Icon            =   "frmNotaFiscal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3135
      Left            =   1680
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Data dbConfigNota 
         Caption         =   "dbConfigNota"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ConciliaNova"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Produtos order by descri"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Data dbNotasCorpo 
         Caption         =   "dbNotasCorpo"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NotasCorpo"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Data dbNotas 
         Caption         =   "dbNotas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Notas"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Data dbCFOP 
         Caption         =   "dbCFOP"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CFOP"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Data dbNaturezaOp 
         Caption         =   "dbNaturezaOp"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NaturezaOP"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdNova 
      Caption         =   "Nova"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   7560
      Picture         =   "frmNotaFiscal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Imprimir"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Editar Atual"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Destinatário"
      TabPicture(0)   =   "frmNotaFiscal.frx":0EC4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label41"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label40"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label39"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtHoraSaida"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDataSaida"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboCFOP2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboNatureza2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboCliente"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDataEmissao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkEntrada"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNotaNr2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "optSaida2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "optEntrada2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCodCliente"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDadosFatura"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Corpo"
      TabPicture(1)   =   "frmNotaFiscal.frx":0EE0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label29"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label33"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label30"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label32"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label34"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label35"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label36"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label37"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label38"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "DBGrid1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cboProdutos"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdCalcular"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdProdutoPreencher"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCodProduto1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Frame3"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtBaseICMS"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtValorICMS"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtBaseICMSSubst"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtTotalProdutos"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtValorICMSSubst"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtValorFrete"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtValorSeguro"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtOutrasDesp"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtTotalNota"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtValorIPI2"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Transportador"
      TabPicture(2)   =   "frmNotaFiscal.frx":0EFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label42"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtFretePorConta"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   102
         Top             =   420
         Width           =   8055
         Begin VB.TextBox txtPesoLiquido 
            DataField       =   "PesoLiquido"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   6840
            MaxLength       =   20
            TabIndex        =   118
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtPesoBruto 
            DataField       =   "PesoBruto"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   5760
            MaxLength       =   20
            TabIndex        =   117
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtNumero 
            DataField       =   "Numero"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   116
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtMarca 
            DataField       =   "Marca"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   3600
            MaxLength       =   20
            TabIndex        =   115
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtEspecie 
            DataField       =   "Especie"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   2520
            MaxLength       =   20
            TabIndex        =   114
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtUF3 
            DataField       =   "UF3"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   113
            Top             =   1080
            Width           =   375
         End
         Begin VB.OptionButton optDestinatario 
            Caption         =   "Destinatário"
            Height          =   255
            Left            =   4800
            TabIndex        =   112
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optEmitente 
            Caption         =   "Emitente"
            Height          =   255
            Left            =   3600
            TabIndex        =   111
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtNome 
            DataField       =   "Nome2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   110
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtEndereco2 
            DataField       =   "Endereco2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   109
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtIE2 
            DataField       =   "IE2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   108
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtQtd2 
            DataField       =   "Quantidade2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   107
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtMunicipio2 
            DataField       =   "Municipio2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   5400
            MaxLength       =   30
            TabIndex        =   106
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox TxtPlaca 
            DataField       =   "Placa"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   6000
            MaxLength       =   20
            TabIndex        =   105
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtCNPJ2 
            DataField       =   "CNPJ2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   104
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtUF2 
            DataField       =   "UF2"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   103
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label55 
            Caption         =   "Peso Líquido:"
            Height          =   255
            Left            =   6840
            TabIndex        =   133
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "Peso Bruto:"
            Height          =   255
            Left            =   5760
            TabIndex        =   132
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Número:"
            Height          =   255
            Left            =   4680
            TabIndex        =   131
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label58 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   3600
            TabIndex        =   130
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label59 
            Caption         =   "Espécie:"
            Height          =   255
            Left            =   2520
            TabIndex        =   129
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label60 
            Caption         =   "UF:"
            Height          =   255
            Left            =   7440
            TabIndex        =   128
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label61 
            Caption         =   "Frete por conta:"
            Height          =   255
            Left            =   3600
            TabIndex        =   127
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label62 
            Caption         =   "Nome/Razão Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label63 
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   1920
            TabIndex        =   125
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label64 
            Caption         =   "I.E.:"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   1440
            TabIndex        =   123
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label66 
            Caption         =   "Município:"
            Height          =   255
            Left            =   5400
            TabIndex        =   122
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label67 
            Caption         =   "Placa:"
            Height          =   255
            Left            =   6000
            TabIndex        =   121
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label68 
            Caption         =   "C.N.P.J.:"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label69 
            Caption         =   "UF:"
            Height          =   255
            Left            =   7440
            TabIndex        =   119
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   99
         Top             =   2700
         Width           =   7935
         Begin VB.TextBox txtDadosAdicionais 
            DataField       =   "DadosAdicionais"
            DataSource      =   "dbNotas"
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   100
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label Label70 
            Caption         =   "Dados Adicionais:"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtFretePorConta 
         DataField       =   "FretePorConta"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -73560
         TabIndex        =   98
         Top             =   4980
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtValorIPI2 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorIPI"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -70200
         TabIndex        =   83
         Top             =   4980
         Width           =   1455
      End
      Begin VB.TextBox txtTotalNota 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalDaNota"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -68520
         TabIndex        =   82
         Top             =   4980
         Width           =   1455
      End
      Begin VB.TextBox txtOutrasDesp 
         Alignment       =   1  'Right Justify
         DataField       =   "OutrasDespesas"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -71760
         TabIndex        =   81
         Top             =   4980
         Width           =   1455
      End
      Begin VB.TextBox txtValorSeguro 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorSeguro"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -73320
         TabIndex        =   80
         Top             =   4980
         Width           =   1455
      End
      Begin VB.TextBox txtValorFrete 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorFrete"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -74880
         TabIndex        =   79
         Top             =   4980
         Width           =   1455
      End
      Begin VB.TextBox txtValorICMSSubst 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMSSubst"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -70200
         TabIndex        =   78
         Top             =   4380
         Width           =   1455
      End
      Begin VB.TextBox txtTotalProdutos 
         Alignment       =   1  'Right Justify
         DataField       =   "TotalDosProdutos"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -68520
         TabIndex        =   77
         Top             =   4380
         Width           =   1455
      End
      Begin VB.TextBox txtBaseICMSSubst 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMSSubst"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -71760
         TabIndex        =   76
         Top             =   4380
         Width           =   1455
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMS"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -73320
         TabIndex        =   75
         Top             =   4380
         Width           =   1455
      End
      Begin VB.TextBox txtBaseICMS 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMS"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   -74880
         TabIndex        =   74
         Top             =   4380
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   49
         Top             =   780
         Width           =   8055
         Begin VB.CommandButton cmdIncluirProduto 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   5520
            TabIndex        =   62
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtValorIPI1 
            Height          =   285
            Left            =   4080
            TabIndex        =   61
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtIPI 
            Height          =   285
            Left            =   3360
            TabIndex        =   60
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtICMS 
            Height          =   285
            Left            =   2640
            TabIndex        =   59
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtValorTotal 
            Height          =   285
            Left            =   1320
            TabIndex        =   58
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtValorUnitario 
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtQtd 
            Height          =   285
            Left            =   6480
            TabIndex        =   56
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtUnidade 
            Height          =   285
            Left            =   5760
            TabIndex        =   55
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtSitTrib 
            Height          =   285
            Left            =   5040
            TabIndex        =   54
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtClasFisc 
            Height          =   285
            Left            =   4320
            TabIndex        =   53
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtProduto 
            Height          =   285
            Left            =   840
            TabIndex        =   52
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtCodProduto2 
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton cmdRemover 
            Caption         =   "Remover"
            Height          =   375
            Left            =   6840
            TabIndex        =   50
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Valor IPI:"
            Height          =   255
            Left            =   4080
            TabIndex        =   73
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "IPI:"
            Height          =   255
            Left            =   3360
            TabIndex        =   72
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label26 
            Caption         =   "ICMS:"
            Height          =   255
            Left            =   2640
            TabIndex        =   71
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "Valor Total:"
            Height          =   255
            Left            =   1320
            TabIndex        =   70
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "Valor Unitário:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   6480
            TabIndex        =   68
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Unid.:"
            Height          =   255
            Left            =   5760
            TabIndex        =   67
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "Sit.:"
            Height          =   255
            Left            =   5040
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Clas:"
            Height          =   255
            Left            =   4320
            TabIndex        =   65
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "Produto:"
            Height          =   255
            Left            =   840
            TabIndex        =   64
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Código:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtCodProduto1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74280
         TabIndex        =   48
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmdProdutoPreencher 
         Caption         =   "Preencher"
         Height          =   375
         Left            =   -68640
         TabIndex        =   47
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   3780
         Width           =   1455
      End
      Begin VB.TextBox txtDadosFatura 
         DataField       =   "DadosFatura"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Top             =   4980
         Width           =   7695
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   2460
         Width           =   8055
         Begin VB.TextBox txtUF 
            DataField       =   "UF"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   5640
            MaxLength       =   2
            TabIndex        =   21
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtIE 
            DataField       =   "Ie"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   6120
            MaxLength       =   30
            TabIndex        =   20
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtFone 
            DataField       =   "Fone"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   4200
            MaxLength       =   20
            TabIndex        =   19
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtMunicipio 
            DataField       =   "Municipio"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   18
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox txtCEP 
            DataField       =   "Cep"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtBairro 
            DataField       =   "Bairro"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtEndereco 
            DataField       =   "Endereco"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   15
            Top             =   1080
            Width           =   4455
         End
         Begin VB.TextBox txtCNPJ 
            DataField       =   "CNPJ"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   5760
            MaxLength       =   30
            TabIndex        =   14
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtCliente 
            DataField       =   "Nome"
            DataSource      =   "dbNotas"
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   13
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label Label13 
            Caption         =   "UF:"
            Height          =   255
            Left            =   5640
            TabIndex        =   30
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "I.E.:"
            Height          =   255
            Left            =   6120
            TabIndex        =   29
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Fone/Fax:"
            Height          =   255
            Left            =   4200
            TabIndex        =   28
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Município:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "CEP:"
            Height          =   255
            Left            =   6600
            TabIndex        =   26
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Bairro/Distrito:"
            Height          =   255
            Left            =   4680
            TabIndex        =   25
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "C.N.P.J. / C.P.F.:"
            Height          =   255
            Left            =   5760
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Nome/Razão Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2100
         Width           =   495
      End
      Begin VB.OptionButton optEntrada2 
         Caption         =   "Entrada"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton optSaida2 
         Caption         =   "Saída"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtNotaNr2 
         Alignment       =   1  'Right Justify
         DataField       =   "NotaNr"
         DataSource      =   "dbNotas"
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   780
         Width           =   1215
      End
      Begin VB.CheckBox chkEntrada 
         Caption         =   "Entrada"
         DataField       =   "Entrada"
         DataSource      =   "dbNotas"
         Height          =   255
         Left            =   6960
         TabIndex        =   6
         Top             =   540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker txtDataEmissao 
         DataField       =   "DataEmissao"
         DataSource      =   "dbNotas"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   79167489
         CurrentDate     =   39007
      End
      Begin MSDBCtls.DBCombo cboCliente 
         Bindings        =   "frmNotaFiscal.frx":0F18
         Height          =   315
         Left            =   960
         TabIndex        =   32
         Top             =   2100
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo cboNatureza2 
         Bindings        =   "frmNotaFiscal.frx":0F31
         DataField       =   "NaturezaOP"
         DataSource      =   "dbNotas"
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo cboCFOP2 
         Bindings        =   "frmNotaFiscal.frx":0F4C
         DataField       =   "CFOP"
         DataSource      =   "dbNotas"
         Height          =   315
         Left            =   3120
         TabIndex        =   34
         Top             =   780
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txtDataSaida 
         DataField       =   "DataSaida"
         DataSource      =   "dbNotas"
         Height          =   300
         Left            =   1560
         TabIndex        =   35
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   79167489
         CurrentDate     =   39007
      End
      Begin MSComCtl2.DTPicker txtHoraSaida 
         DataField       =   "HoraSaida"
         DataSource      =   "dbNotas"
         Height          =   300
         Left            =   3000
         TabIndex        =   36
         Top             =   1380
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         Format          =   79167490
         CurrentDate     =   39007
      End
      Begin MSDBCtls.DBCombo cboProdutos 
         Bindings        =   "frmNotaFiscal.frx":0F61
         Height          =   315
         Left            =   -72960
         TabIndex        =   84
         Top             =   420
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmNotaFiscal.frx":0F7A
         Height          =   1335
         Left            =   -74880
         OleObjectBlob   =   "frmNotaFiscal.frx":0F95
         TabIndex        =   85
         Top             =   2340
         Width           =   7935
      End
      Begin VB.Label Label42 
         Caption         =   "Frete por conta:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   134
         Top             =   4980
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Valor IPI:"
         Height          =   255
         Left            =   -70200
         TabIndex        =   97
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label37 
         Caption         =   "Outras Desp.:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   96
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Label Label36 
         Caption         =   "Valor do Seguro:"
         Height          =   255
         Left            =   -73320
         TabIndex        =   95
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label35 
         Caption         =   "Valor do Frete:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "Valor Total da Nota:"
         Height          =   255
         Left            =   -68520
         TabIndex        =   93
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "Valor do ICMS Subst.:"
         Height          =   255
         Left            =   -70200
         TabIndex        =   92
         Top             =   4140
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Base ICMS Subst.:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   91
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   "Valor do ICMS:"
         Height          =   255
         Left            =   -73320
         TabIndex        =   90
         Top             =   4140
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "B. de Cálc. ICMS:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "Total dos Produtos:"
         Height          =   255
         Left            =   -68520
         TabIndex        =   88
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Produto:"
         Height          =   255
         Left            =   -73680
         TabIndex        =   87
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Código:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   86
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Dados da Fatura:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   4740
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Código:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Operação:"
         Height          =   195
         Left            =   1200
         TabIndex        =   42
         Top             =   540
         Width           =   1665
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   3120
         TabIndex        =   41
         Top             =   540
         Width           =   465
      End
      Begin VB.Label Label41 
         Caption         =   "Número da Nota:"
         Height          =   255
         Left            =   4080
         TabIndex        =   40
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data da Emissão:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Data da Saída:"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Hora da Saída:"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   1140
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmClienteNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoNota As Double, Adicionando As Boolean

Private Sub CalculaNota()
Dim BaseICMS As Currency, ValorICMS As Currency
Dim BaseICMSSub As Currency, ValorICMSSub As Currency
Dim ValorTotalProdutos As Currency, ValorTotalIPI As Currency
Dim ValorTotalDaNota As Currency
Dim ValorFrete As Currency
Dim ValorSeguro As Currency
Dim OutrasDesp As Currency

With dbNotasCorpo
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If .Recordset!aliquotaicms <> 0 Then
      BaseICMS = BaseICMS + (.Recordset!ValorTotal)
      ValorICMS = ValorICMS + ((.Recordset!aliquotaicms / 100) * .Recordset!ValorTotal)
    End If
    If IsNumeric(.Recordset!subtributaria) = True Then
      If .Recordset!subtributaria <> "0" Then
        BaseICMSSub = BaseICMSSub + (.Recordset!ValorTotal)
        ValorICMSSub = ValorICMSSub + ((CDbl(.Recordset!subtributaria) / 100) * .Recordset!ValorTotal)
      End If
    End If
    ValorTotalProdutos = ValorTotalProdutos + (.Recordset!ValorTotal)
    If IsNull(.Recordset!valoripi) = False Then
      ValorTotalIPI = ValorTotalIPI + .Recordset!valoripi
    End If
    .Recordset.MoveNext
  Loop
End With

txtBaseICMS.Text = Format(BaseICMS, "#,##0.00")
txtValorICMS.Text = Format(ValorICMS, "#,##0.00")
txtBaseICMSSubst.Text = Format(BaseICMSSub, "#,##0.00")
txtValorICMSSubst.Text = Format(ValorICMSSub, "#,##0.00")
txtTotalProdutos.Text = Format(ValorTotalProdutos, "#,##0.00")
txtValorIPI2.Text = Format(ValorTotalIPI, "#,##0.00")

If IsNumeric(txtValorFrete.Text) = True Then ValorFrete = CCur(txtValorFrete.Text)
If IsNumeric(txtValorSeguro.Text) = True Then ValorSeguro = CCur(txtValorSeguro.Text)
If IsNumeric(txtOutrasDesp.Text) = True Then OutrasDesp = CCur(txtOutrasDesp.Text)

ValorTotalDaNota = ValorTotalProdutos + ValorFrete + ValorSeguro + OutrasDesp + ValorTotalIPI + ValorICMS + ValorICMSSub

txtTotalNota.Text = Format(ValorTotalDaNota, "#,##0.00")

End Sub

Private Sub cboCliente_LostFocus()
DbClientes.Refresh
If cboCliente.Text = "" Then Exit Sub
With DbClientes
  .Recordset.FindFirst "nome='" & cboCliente.Text & "'"
  If .Recordset.NoMatch = False Then
    txtCodCliente.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cboProdutos_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProdutos.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProdutos.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  txtCodProduto1.Text = .Recordset!Codigo
End With
End Sub

Private Sub cmdCalcular_Click()
CalculaNota
End Sub

Public Sub ClientePreenche()
With DbClientes
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  On Error Resume Next
  If .Recordset!Nome <> cboCliente.Text Then Exit Sub
  If IsNull(.Recordset!nome2) = False Then txtCliente.Text = .Recordset!nome2
  If IsNull(.Recordset!CNPJ) = False Then txtCNPJ.Text = .Recordset!CNPJ
  If IsNull(.Recordset!Endereco) = False Then txtEndereco.Text = .Recordset!Endereco
  If IsNull(.Recordset!bairro) = False Then txtBairro.Text = .Recordset!bairro
  If IsNull(.Recordset!CEP) = False Then txtCEP.Text = .Recordset!CEP
  If IsNull(.Recordset!cidade) = False Then txtMunicipio.Text = .Recordset!cidade
  If IsNull(.Recordset!Telefone) = False Then txtFone.Text = .Recordset!Telefone
  If IsNull(.Recordset!Estado) = False Then txtUF.Text = .Recordset!Estado
  If IsNull(.Recordset!ie) = False Then txtIE.Text = .Recordset!ie
End With
End Sub

Private Sub cmdEdit_Click()
SSTab1.Tab = 1
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0
Printer.ScaleMode = vbCentimeters
Printer.Font = "Arial"
Printer.FontSize = 8
With dbConfigNota
  .RecordSource = "select *from confignota"
  .Refresh
  'On Error Resume Next
  StrTemp = txtNotaNr2.Text
  Printer.CurrentX = .Recordset!nrnotatopox
  Printer.CurrentY = .Recordset!nrnotatopoy
  Printer.Print StrTemp
  
  StrTemp = "X"
  If optSaida2.Value = True Then
    Printer.CurrentX = .Recordset!saidax
    Printer.CurrentY = .Recordset!saiday
  Else
    Printer.CurrentX = .Recordset!entradax
    Printer.CurrentY = .Recordset!entraday
  End If
  Printer.Print StrTemp
  
  StrTemp = cboNatureza2.Text
  Printer.CurrentX = .Recordset!naturezaoperacaox
  Printer.CurrentY = .Recordset!naturezaoperacaoy
  Printer.Print StrTemp
  
  StrTemp = cboCFOP2.Text
  Printer.CurrentX = .Recordset!cfopx
  Printer.CurrentY = .Recordset!cfopy
  Printer.Print StrTemp
  
  StrTemp = Format(txtDataEmissao.Value, "short date")
  Printer.CurrentX = .Recordset!dataemissaox
  Printer.CurrentY = .Recordset!dataemissaoy
  Printer.Print StrTemp
  
  StrTemp = Format(txtDataSaida.Value, "Short date")
  Printer.CurrentX = .Recordset!datasaidax
  Printer.CurrentY = .Recordset!datasaiday
  Printer.Print StrTemp
  
  StrTemp = Format(txtHoraSaida.Value, "short time")
  Printer.CurrentX = .Recordset!horasaidax
  Printer.CurrentY = .Recordset!horasaiday
  Printer.Print StrTemp
  
  StrTemp = txtCliente.Text
  Printer.CurrentX = .Recordset!nomex
  Printer.CurrentY = .Recordset!nomey
  Printer.Print StrTemp
  
  StrTemp = txtCNPJ.Text
  Printer.CurrentX = .Recordset!cnpjx
  Printer.CurrentY = .Recordset!cnpjy
  Printer.Print StrTemp
  
  StrTemp = txtEndereco.Text
  Printer.CurrentX = .Recordset!enderecox
  Printer.CurrentY = .Recordset!enderecoy
  Printer.Print StrTemp
  
  StrTemp = txtBairro.Text
  Printer.CurrentX = .Recordset!bairrox
  Printer.CurrentY = .Recordset!bairroy
  Printer.Print StrTemp
  
  StrTemp = txtCEP.Text
  Printer.CurrentX = .Recordset!cepx
  Printer.CurrentY = .Recordset!cepy
  Printer.Print StrTemp
  
  StrTemp = txtMunicipio.Text
  Printer.CurrentX = .Recordset!municipiox
  Printer.CurrentY = .Recordset!municipioy
  Printer.Print StrTemp
  
  StrTemp = txtFone.Text
  Printer.CurrentX = .Recordset!fonex
  Printer.CurrentY = .Recordset!foney
  Printer.Print StrTemp
  
  StrTemp = txtUF.Text
  Printer.CurrentX = .Recordset!uf1x
  Printer.CurrentY = .Recordset!uf1y
  Printer.Print StrTemp
  
  StrTemp = txtIE.Text
  Printer.CurrentX = .Recordset!iex
  Printer.CurrentY = .Recordset!iey
  Printer.Print StrTemp
  
  StrTemp = txtDadosFatura.Text
  Printer.CurrentX = .Recordset!dadosfaturax
  Printer.CurrentY = .Recordset!dadosfaturay
  Printer.Print StrTemp
  
  
  'início do corpo da nota
  Printer.CurrentY = dbConfigNota.Recordset!iniciocorpoy
  
  With dbNotasCorpo
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        StrTemp = ""
        If IsNull(.Recordset!CodigoProduto) = False Then StrTemp = .Recordset!CodigoProduto
        Printer.CurrentX = dbConfigNota.Recordset!colunadescri - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!clasfiscal) = False Then StrTemp = .Recordset!clasfiscal
        Printer.CurrentX = dbConfigNota.Recordset!colunasubsttrib - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!subtributaria) = False Then StrTemp = .Recordset!subtributaria
        Printer.CurrentX = dbConfigNota.Recordset!colunaunidade - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!unidade) = False Then StrTemp = .Recordset!unidade
        Printer.CurrentX = dbConfigNota.Recordset!colunaunidade
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!Quantidade) = False Then StrTemp = Format(.Recordset!Quantidade, "currency")
        Printer.CurrentX = dbConfigNota.Recordset!colunavunitario - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!valorUnitario) = False Then StrTemp = Format(.Recordset!valorUnitario, "currency")
        Printer.CurrentX = dbConfigNota.Recordset!colunavtotal - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!ValorTotal) = False Then StrTemp = Format(.Recordset!ValorTotal, "currency")
        Printer.CurrentX = dbConfigNota.Recordset!colunaaliquotaicms - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!aliquotaicms) = False Then StrTemp = .Recordset!aliquotaicms
        Printer.CurrentX = dbConfigNota.Recordset!colunaaliquotaipi - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!aliquotaipi) = False Then StrTemp = .Recordset!aliquotaipi
        Printer.CurrentX = dbConfigNota.Recordset!colunavaloripi - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!valoripi) = False Then StrTemp = Format(.Recordset!valoripi, "currency")
        Printer.CurrentX = dbConfigNota.Recordset!colunalimite - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!descriproduto) = False Then StrTemp = .Recordset!descriproduto
        ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, dbConfigNota.Recordset!colunadescri, Printer.CurrentY, dbConfigNota.Recordset!colunaclasfiscal - 0.2
        
        .Recordset.MoveNext
      Loop
    End If
  End With
  
  StrTemp = Format(txtBaseICMS.Text, "currency")
  Printer.CurrentX = .Recordset!baseicmsx
  Printer.CurrentY = .Recordset!baseicmsy
  Printer.Print StrTemp
  
  StrTemp = Format(txtValorICMS.Text, "currency")
  Printer.CurrentX = .Recordset!valoricmsx
  Printer.CurrentY = .Recordset!valoricmsy
  Printer.Print StrTemp
  
  StrTemp = Format(txtBaseICMSSubst.Text, "currency")
  Printer.CurrentX = .Recordset!baseicmssubx
  Printer.CurrentY = .Recordset!baseicmssuby
  Printer.Print StrTemp
  
  StrTemp = Format(txtValorICMSSubst.Text, "currency")
  Printer.CurrentX = .Recordset!valoricmssubx
  Printer.CurrentY = .Recordset!valoricmssuby
  Printer.Print StrTemp
  
  StrTemp = Format(txtTotalProdutos.Text, "currency")
  Printer.CurrentX = .Recordset!valortotalprodutosx
  Printer.CurrentY = .Recordset!valortotalprodutosy
  Printer.Print StrTemp
  
  StrTemp = Format(txtValorFrete.Text, "currency")
  Printer.CurrentX = .Recordset!valordofretex
  Printer.CurrentY = .Recordset!valordofretey
  Printer.Print StrTemp
  
  StrTemp = Format(txtValorSeguro.Text, "currency")
  Printer.CurrentX = .Recordset!valordosegurox
  Printer.CurrentY = .Recordset!valordoseguroy
  Printer.Print StrTemp
  
  StrTemp = Format(txtOutrasDesp.Text, "currency")
  Printer.CurrentX = .Recordset!outrasdespx
  Printer.CurrentY = .Recordset!outrasdespy
  Printer.Print StrTemp
  
  StrTemp = Format(txtValorIPI2.Text, "currency")
  Printer.CurrentX = .Recordset!valortotalipix
  Printer.CurrentY = .Recordset!valortotalipiy
  Printer.Print StrTemp
  
  StrTemp = Format(txtTotalNota.Text, "currency")
  Printer.CurrentX = .Recordset!valortotalnotax
  Printer.CurrentY = .Recordset!valortotalnotay
  Printer.Print StrTemp
  
  StrTemp = txtNome.Text
  Printer.CurrentX = .Recordset!nome2x
  Printer.CurrentY = .Recordset!nome2y
  Printer.Print StrTemp
  
  StrTemp = txtFretePorConta.Text
  Printer.CurrentX = .Recordset!freteporcontax
  Printer.CurrentY = .Recordset!freteporcontay
  Printer.Print StrTemp
  
  StrTemp = txtPlaca.Text
  Printer.CurrentX = .Recordset!placax
  Printer.CurrentY = .Recordset!placay
  Printer.Print StrTemp
  
  StrTemp = txtUF2.Text
  Printer.CurrentX = .Recordset!uf2x
  Printer.CurrentY = .Recordset!uf2y
  Printer.Print StrTemp
  
  StrTemp = txtCNPJ2.Text
  Printer.CurrentX = .Recordset!cnpj2x
  Printer.CurrentY = .Recordset!cnpj2y
  Printer.Print StrTemp
  
  StrTemp = txtEndereco2.Text
  Printer.CurrentX = .Recordset!endereco2x
  Printer.CurrentY = .Recordset!endereco2y
  Printer.Print StrTemp
  
  StrTemp = txtMunicipio2.Text
  Printer.CurrentX = .Recordset!municipio2x
  Printer.CurrentY = .Recordset!municipio2y
  Printer.Print StrTemp
  
  StrTemp = txtUF3.Text
  Printer.CurrentX = .Recordset!uf3x
  Printer.CurrentY = .Recordset!uf3y
  Printer.Print StrTemp
  
  StrTemp = txtIE2.Text
  Printer.CurrentX = .Recordset!ie2x
  Printer.CurrentY = .Recordset!ie2y
  Printer.Print StrTemp
  
  StrTemp = txtQtd2.Text
  Printer.CurrentX = .Recordset!qtd2x
  Printer.CurrentY = .Recordset!qtd2y
  Printer.Print StrTemp
  
  StrTemp = txtEspecie.Text
  Printer.CurrentX = .Recordset!especiex
  Printer.CurrentY = .Recordset!especiey
  Printer.Print StrTemp
  
  StrTemp = txtMarca.Text
  Printer.CurrentX = .Recordset!marcax
  Printer.CurrentY = .Recordset!marcay
  Printer.Print StrTemp
  
  StrTemp = txtNumero.Text
  Printer.CurrentX = .Recordset!numerox
  Printer.CurrentY = .Recordset!numeroy
  Printer.Print StrTemp
  
  StrTemp = txtPesoBruto.Text
  Printer.CurrentX = .Recordset!pesobrutox
  Printer.CurrentY = .Recordset!pesobrutoy
  Printer.Print StrTemp
  
  StrTemp = txtPesoLiquido.Text
  Printer.CurrentX = .Recordset!pesoliquidox
  Printer.CurrentY = .Recordset!pesoliquidoy
  Printer.Print StrTemp
  
  ImprimeTextoJustificado Printer, txtDadosAdicionais.Text, AlinhaEsquerda, .Recordset!dadosadicionais1x, .Recordset!dadosadicionais1y, .Recordset!dadosadicionais2x
    
  StrTemp = txtNotaNr2.Text
  Printer.CurrentX = .Recordset!nrnotacanhotox
  Printer.CurrentY = .Recordset!nrnotacanhotoy
  Printer.Print StrTemp
  
  Printer.EndDoc
  
End With
NaoImprime:

End Sub

Private Sub cmdIncluirProduto_Click()
With dbNotasCorpo
  On Error Resume Next
  .Recordset.AddNew
  .Recordset!CodigoNota = CodigoNota
  .Recordset!CodigoProduto = txtCodProduto2.Text
  .Recordset!descriproduto = txtProduto.Text
  .Recordset!clasfiscal = txtClasFisc.Text
  .Recordset!subtributaria = txtSitTrib.Text
  .Recordset!unidade = txtUnidade.Text
  .Recordset!Quantidade = txtQtd.Text
  .Recordset!valorUnitario = txtValorUnitario.Text
  .Recordset!ValorTotal = txtValorTotal.Text
  .Recordset!aliquotaicms = txtICMS.Text
  .Recordset!aliquotaipi = txtIPI.Text
  .Recordset!valoripi = txtValorIPI1.Text
  .Recordset.Update
End With
CalculaNota
End Sub

Private Sub cmdNova_Click()
Adicionando = True
dbNotas.Recordset.AddNew
CodigoNota = dbNotas.Recordset!CodigoNota
dbNotas.Recordset.Update
dbNotas.Refresh
dbNotas.Recordset.FindFirst "codigonota=" & CodigoNota
txtDataEmissao.Value = Date
txtDataSaida.Value = Date
txtHoraSaida.Value = Time
SSTab1.Tab = 1
Adicionando = False
End Sub

Private Sub cmdProdutoPreencher_Click()
With dbProdutos
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  On Error Resume Next
  If cboProdutos.Text <> .Recordset!Descri Then Exit Sub
  If IsNull(.Recordset!Codigo) = False Then txtCodProduto2.Text = .Recordset!Codigo
  If IsNull(.Recordset!Descri) = False Then txtProduto.Text = .Recordset!Descri
  If IsNull(.Recordset!PrecoVenda) = False Then txtValorUnitario.Text = Format(.Recordset!PrecoVenda, "Currency")
End With
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbNotasCorpo
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o item atual?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With
CalculaNota
End Sub


Private Sub dbNotas_Reposition()
If Adicionando = True Then Exit Sub

If dbNotas.Recordset.EOF = True Then Exit Sub
If dbNotas.Recordset.BOF = True Then Exit Sub
CodigoNota = dbNotas.Recordset!CodigoNota
With dbNotasCorpo
  On Error Resume Next
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from Notascorpo where codigonota=" & CodigoNota & " order by CodigoNotaCorpo"
  .Refresh
End With
If txtFretePorConta.Text = "1" Then
  optEmitente.Value = True
ElseIf txtFretePorConta.Text = "2" Then
  optDestinatario.Value = True
End If
If chkEntrada.Value = vbChecked Then
  optEntrada2.Value = True
Else
  optSaida2.Value = True
End If
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
Dim Ws As Workspace, db As Database
Adicionando = False
With dbNaturezaOp
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from NaturezaOP order by descri"
  On Error GoTo 0
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table NaturezaOP (Descri text(20))"
  End If
End With
With dbCFOP
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from CFOP order by Codigo"
  On Error GoTo 0
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table CFOP (codigo double, Descri text(20))"
  End If
End With
With dbNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from Notascorpo order by CodigoNotaCorpo"
  On Error GoTo 0
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table NotasCorpo (CodigoNotaCorpo counter, CodigoNota double, CodigoProduto Text(10), DescriProduto Text(255), ClasFiscal Text(4), SubTributaria Text(4), Unidade Text(4), Quantidade double, ValorUnitario currency, ValorTotal currency, AliquotaICMS double, AliquotaIPI double, ValorIpi Currency)"
  End If
End With
With dbNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from Notas order by CodigoNota"
  On Error GoTo 0
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table Notas (CodigoNota counter, NaturezaOP text(20), CFOP Text(5), NotaNr double, Entrada bit, DataEmissao datetime, DataSaida DateTime, HoraSaida DateTime, Nome Text(30), CNPJ text(30), Endereco Text(30), Bairro Text(20), Cep Text(20), Municipio Text(30), Fone Text(20), UF Text(2), Ie Text(30), DadosFatura Text(50), BaseICMS Currency, ValorICMS Currency, BaseICMSSubst Currency, ValorICMSSubst Currency, TotalDosProdutos Currency, ValorFrete Currency, ValorSeguro Currency, OutrasDespesas Currency, ValorIPI currency, ValorTotalDaNota Currency, Nome2 Text(30), FretePorConta integer, Placa text(10), UF2 Text(2), CNPJ2 Text(30), Endereco2 Text(30), Municipio2 Text(30), UF3 text(2), IE2 text(30), Quantidade2 Text(20), Especie Text(20), Marca Text(20), Numero Text (20), PesoBruto Text(20), PesoLiquido Text(20), DadosAdicionais Text(255))"
  End If
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConfigNota
  .Connect = Conectar
  .DatabaseName = Caminho
  On Error GoTo 0
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    On Error GoTo 0
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table ConfigNota (NrNotaTopoX double)"
    'Topo da Nota
    db.Execute "alter table Confignota add column NrNotaTopoY double"
    db.Execute "alter table Confignota add column NrNotaCanhotoX double"
    db.Execute "alter table Confignota add column NrNotaCanhotoY double"
    db.Execute "alter table Confignota add column SaidaX double"
    db.Execute "alter table Confignota add column SaidaY double"
    db.Execute "alter table Confignota add column EntradaX double"
    db.Execute "alter table Confignota add column EntradaY double"
    db.Execute "alter table Confignota add column NaturezaOperacaoX double"
    db.Execute "alter table Confignota add column NaturezaOperacaoY double"
    db.Execute "alter table Confignota add column CFOPX double"
    db.Execute "alter table Confignota add column CFOPY double"
    db.Execute "alter table Confignota add column DataEmissaoX double"
    db.Execute "alter table Confignota add column DataEmissaoY double"
    db.Execute "alter table Confignota add column DataSaidaX double"
    db.Execute "alter table Confignota add column DataSaidaY double"
    db.Execute "alter table Confignota add column HoraSaidaX double"
    db.Execute "alter table Confignota add column HoraSaidaY double"
    db.Execute "alter table Confignota add column DadosFaturaX double"
    db.Execute "alter table Confignota add column DadosFaturaY double"
    'Destinatário
    db.Execute "alter table Confignota add column NomeX double"
    db.Execute "alter table Confignota add column NomeY double"
    db.Execute "alter table Confignota add column CNPJX double"
    db.Execute "alter table Confignota add column CNPJY double"
    db.Execute "alter table Confignota add column EnderecoX double"
    db.Execute "alter table Confignota add column EnderecoY double"
    db.Execute "alter table Confignota add column BairroX double"
    db.Execute "alter table Confignota add column BairroY double"
    db.Execute "alter table Confignota add column CEPX double"
    db.Execute "alter table Confignota add column CEPY double"
    db.Execute "alter table Confignota add column MunicipioX double"
    db.Execute "alter table Confignota add column MunicipioY double"
    db.Execute "alter table Confignota add column FoneX double"
    db.Execute "alter table Confignota add column FoneY double"
    db.Execute "alter table Confignota add column UF1X double"
    db.Execute "alter table Confignota add column UF1Y double"
    db.Execute "alter table Confignota add column IEX double"
    db.Execute "alter table Confignota add column IEY double"
    'Corpo
    db.Execute "alter table Confignota add column InicioCorpoY double"
    db.Execute "alter table Confignota add column ColunaCodigo double"
    db.Execute "alter table Confignota add column ColunaDescri double"
    db.Execute "alter table Confignota add column ColunaClasFiscal double"
    db.Execute "alter table Confignota add column ColunaSubstTrib double"
    db.Execute "alter table Confignota add column ColunaUnidade double"
    db.Execute "alter table Confignota add column ColunaQuantidade double"
    db.Execute "alter table Confignota add column ColunaVUnitario double"
    db.Execute "alter table Confignota add column ColunaVTotal double"
    db.Execute "alter table Confignota add column ColunaAliquotaICMS double"
    db.Execute "alter table Confignota add column ColunaAliquotaIPI double"
    db.Execute "alter table Confignota add column ColunaValorIPI double"
    db.Execute "alter table Confignota add column ColunaLimite double"
    db.Execute "alter table Confignota add column BaseICMSX double"
    db.Execute "alter table Confignota add column BaseICMSY double"
    db.Execute "alter table Confignota add column ValorICMSX double"
    db.Execute "alter table Confignota add column ValorICMSY double"
    db.Execute "alter table Confignota add column BaseICMSSubX double"
    db.Execute "alter table Confignota add column BaseICMSSubY double"
    db.Execute "alter table Confignota add column ValorICMSSubX double"
    db.Execute "alter table Confignota add column ValorICMSSubY double"
    db.Execute "alter table Confignota add column ValorTotalProdutosX double"
    db.Execute "alter table Confignota add column ValorTotalProdutosY double"
    db.Execute "alter table Confignota add column ValorDoFreteX double"
    db.Execute "alter table Confignota add column ValorDoFreteY double"
    db.Execute "alter table Confignota add column ValorDoSeguroX double"
    db.Execute "alter table Confignota add column ValorDoSeguroY double"
    db.Execute "alter table Confignota add column OutrasDespX double"
    db.Execute "alter table Confignota add column OutrasDespY double"
    db.Execute "alter table Confignota add column ValorTotalIPIX double"
    db.Execute "alter table Confignota add column ValorTotalIPIY double"
    db.Execute "alter table Confignota add column ValorTotalNotaX double"
    db.Execute "alter table Confignota add column ValorTotalNotaY double"
    'Transportador
    db.Execute "alter table Confignota add column Nome2X double"
    db.Execute "alter table Confignota add column Nome2Y double"
    db.Execute "alter table Confignota add column FretePorContaX double"
    db.Execute "alter table Confignota add column FretePorContaY double"
    db.Execute "alter table Confignota add column PlacaX double"
    db.Execute "alter table Confignota add column PlacaY double"
    db.Execute "alter table Confignota add column UF2X double"
    db.Execute "alter table Confignota add column UF2Y double"
    db.Execute "alter table Confignota add column CNPJ2X double"
    db.Execute "alter table Confignota add column CNPJ2Y double"
    db.Execute "alter table Confignota add column Endereco2X double"
    db.Execute "alter table Confignota add column Endereco2Y double"
    db.Execute "alter table Confignota add column Municipio2X double"
    db.Execute "alter table Confignota add column Municipio2Y double"
    db.Execute "alter table Confignota add column UF3X double"
    db.Execute "alter table Confignota add column UF3Y double"
    db.Execute "alter table Confignota add column IE2X double"
    db.Execute "alter table Confignota add column IE2Y double"
    db.Execute "alter table Confignota add column QTD2X double"
    db.Execute "alter table Confignota add column QTD2Y double"
    db.Execute "alter table Confignota add column EspecieX double"
    db.Execute "alter table Confignota add column EspecieY double"
    db.Execute "alter table Confignota add column MarcaX double"
    db.Execute "alter table Confignota add column MarcaY double"
    db.Execute "alter table Confignota add column NumeroX double"
    db.Execute "alter table Confignota add column NumeroY double"
    db.Execute "alter table Confignota add column PesoBrutoX double"
    db.Execute "alter table Confignota add column PesoBrutoY double"
    db.Execute "alter table Confignota add column PesoLiquidoX double"
    db.Execute "alter table Confignota add column PesoLiquidoY double"
    db.Execute "alter table Confignota add column DadosAdicionais1X double"
    db.Execute "alter table Confignota add column DadosAdicionais1Y double"
    db.Execute "alter table Confignota add column DadosAdicionais2X double"
    db.Execute "alter table Confignota add column DadosAdicionais2Y double"
  End If
End With
With dbConfigNota
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    For i = 0 To .Recordset.Fields.Count - 1
      .Recordset(i) = 0
    Next i
    .Recordset.Update
  End If
End With
End Sub

Private Sub Label49_Click()

End Sub

Private Sub optDestinatario_Click()
txtFretePorConta.Text = "2"
End Sub

Private Sub optEmitente_Click()
txtFretePorConta.Text = "1"
End Sub

Private Sub optEntrada2_Click()
chkEntrada.Value = vbChecked
End Sub

Private Sub optSaida2_Click()
chkEntrada.Value = vbUnchecked
End Sub

Private Sub txtCodCliente_LostFocus()
DbClientes.Refresh
If txtCodCliente.Text = "" Then Exit Sub
If IsNumeric(txtCodCliente.Text) = False Then Exit Sub
With DbClientes
  .Recordset.FindFirst "codigocliente=" & txtCodCliente.Text
  If .Recordset.NoMatch = False Then
    cboCliente.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtCodProduto1_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodProduto1.Text = "" Then Exit Sub
  If IsNumeric(txtCodProduto1.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigo=" & txtCodProduto1.Text
  If .Recordset.NoMatch = True Then Exit Sub
  cboProdutos.Text = .Recordset!Descri
End With
End Sub

Private Sub txtDadosAdicionais_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDadosAdicionais_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtIPI_LostFocus()
If IsNumeric(txtIPI.Text) = True Then
  If IsNumeric(txtValorTotal.Text) = True Then
    txtValorIPI1.Text = Format((CDbl(txtIPI.Text) / 100) * (CCur(txtValorTotal.Text)), "currency")
  End If
End If
End Sub

Private Sub txtQtd_LostFocus()
If IsNumeric(txtQtd.Text) = True Then
  If IsNumeric(txtValorUnitario.Text) = True Then
    txtValorTotal.Text = Format(CCur(txtValorUnitario.Text) * CDbl(txtQtd.Text), "currency")
  End If
End If
End Sub

Private Sub txtValorUnitario_LostFocus()
If IsNumeric(txtQtd.Text) = True Then
  If IsNumeric(txtValorUnitario.Text) = True Then
    txtValorTotal.Text = Format(CCur(txtValorUnitario.Text) * CDbl(txtQtd.Text), "currency")
  End If
End If
End Sub
