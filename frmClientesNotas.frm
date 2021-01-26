VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientesNotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes Cobrança"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmClientesNotas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   5775
      Left            =   4080
      TabIndex        =   47
      Top             =   6360
      Visible         =   0   'False
      Width           =   8295
      Begin MSAdodcLib.Adodc dbProdutos2 
         Height          =   330
         Left            =   3120
         Top             =   4200
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   $"frmClientesNotas.frx":0442
         Caption         =   "dbProdutos2"
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
      Begin VB.Data qProtestados 
         Caption         =   "qProtestados"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(valor) as total from clientescobranca where protestado=-1"
         Top             =   4680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbProtestados 
         Caption         =   "dbProtestados"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientescobranca where protestado=-1 order by datafechamento"
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data QPendencias 
         Caption         =   "QPendencias"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select sum(valor) as total from clientescobranca where pago=0"
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbPendencias 
         Caption         =   "dbPendencias"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientescobranca where pago=0 order by datafechamento"
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbClientes2 
         Caption         =   "dbClientes2"
         Connect         =   "Access"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientes"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbJurosBoleto 
         Caption         =   "dbJurosBoleto"
         Connect         =   "Access"
         DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from JurosBoleto order by inicio, final"
         Top             =   720
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
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Produtos order by descri"
         Top             =   2160
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
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NaturezaOP"
         Top             =   720
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
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CFOP"
         Top             =   1080
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
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Notas"
         Top             =   1800
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
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NotasCorpo"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Data dbConfigNota 
         Caption         =   "dbConfigNota"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\rede\dados\Atalai.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ConciliaNova"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Data dbDespesas 
         Caption         =   "dbDespesas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesasLanc2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbCartoes 
         Caption         =   "dbCartoes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Cartoes"
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbConciliaNova 
         Caption         =   "dbConciliaNova"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ConciliaNova"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbCobranca 
         Caption         =   "dbCobranca"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesCobranca order by datafechamento"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbFormaDePg 
         Caption         =   "dbFormaDePg"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from formadepagamento order by descri"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbContas 
         Caption         =   "dbContas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from contas"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
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
         RecordSource    =   "select *from clientesNota where codigocliente=0"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data QSoma 
         Caption         =   "QSoma"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientesNota where codigocliente=0"
         Top             =   2880
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data dbClientesNotas2 
         Caption         =   "dbClientesNotas2"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientesNota2 where codigocliente=0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
         Height          =   330
         Left            =   3120
         Top             =   4560
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
   End
   Begin VB.CommandButton cmdNotaDeOutroCliente 
      Caption         =   "Nota de outro cliente"
      Height          =   375
      Left            =   2160
      TabIndex        =   54
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdConfiguraNota 
      Caption         =   "Configura Nota"
      Height          =   375
      Left            =   240
      TabIndex        =   53
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   42
      Top             =   6240
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Por Cliente"
      TabPicture(0)   =   "frmClientesNotas.frx":0507
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblObs"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(34)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(12)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDias"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPrazo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboCliente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkConfirmadas"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Cobranças Pendentes"
      TabPicture(1)   =   "frmClientesNotas.frx":0523
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTotalPendente"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtProrrogar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DBGrid1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAtualiza"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdImprime2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdProrrogar"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtNrNota"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdAlteraNota"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Cobranças Protestadas"
      TabPicture(2)   =   "frmClientesNotas.frx":053F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTotalProtesto"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(3)=   "txtResgatado"
      Tab(2).Control(4)=   "DBGrid5"
      Tab(2).Control(5)=   "cmdAtualizaProtestos"
      Tab(2).Control(6)=   "cmdResgatar"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmdResgatar 
         Caption         =   "Resgatar"
         Height          =   375
         Left            =   -70200
         TabIndex        =   52
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAtualizaProtestos 
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   -74760
         TabIndex        =   51
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlteraNota 
         Height          =   495
         Left            =   1560
         Picture         =   "frmClientesNotas.frx":055B
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtNrNota 
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdProrrogar 
         Caption         =   "Prorrogar"
         Height          =   375
         Left            =   6720
         TabIndex        =   37
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CheckBox chkConfirmadas 
         Caption         =   "Notas Confirmadas"
         Height          =   255
         Left            =   -69720
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin MSDBCtls.DBCombo cboCliente 
         Bindings        =   "frmClientesNotas.frx":1225
         Height          =   315
         Left            =   -74160
         TabIndex        =   3
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         Text            =   ""
      End
      Begin VB.CommandButton cmdImprime2 
         Height          =   615
         Left            =   8040
         Picture         =   "frmClientesNotas.frx":123E
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "Imprimir"
         Top             =   5160
         Width           =   735
      End
      Begin VB.CommandButton cmdAtualiza 
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   5520
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmClientesNotas.frx":1CC0
         Height          =   4215
         Left            =   240
         OleObjectBlob   =   "frmClientesNotas.frx":1CDB
         TabIndex        =   31
         Top             =   480
         Width           =   8535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fechamento"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   46
         Top             =   1800
         Width           =   3495
         Begin VB.CommandButton cmdExibe 
            Caption         =   "Exibir"
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdImprimeNotas 
            Caption         =   "Imprime"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3360
            Width           =   975
         End
         Begin VB.CommandButton cmdFechar 
            Caption         =   "Fechar"
            Height          =   375
            Left            =   2520
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin MSComCtl2.DTPicker txtFechamento 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109772801
            CurrentDate     =   37664
         End
         Begin MSDBGrid.DBGrid DBGrid4 
            Bindings        =   "frmClientesNotas.frx":2D92
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "frmClientesNotas.frx":2DB1
            TabIndex        =   15
            Top             =   840
            Width           =   3255
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frmClientesNotas.frx":3950
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmClientesNotas.frx":396E
            TabIndex        =   43
            Top             =   840
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   1200
            TabIndex        =   17
            Top             =   3360
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cobrança"
         Height          =   4095
         Left            =   -71400
         TabIndex        =   45
         Top             =   1800
         Width           =   5295
         Begin VB.CommandButton cmdExtornar 
            Caption         =   "Extornar"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CommandButton cmdProtestar 
            Caption         =   "Protestar"
            Height          =   375
            Left            =   3120
            TabIndex        =   29
            Top             =   3600
            Width           =   1215
         End
         Begin MSDBCtls.DBCombo cboFormadePg 
            Bindings        =   "frmClientesNotas.frx":450D
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   3000
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            Text            =   ""
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3720
            TabIndex        =   25
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton cmdRecebe 
            Caption         =   "Receber"
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdImprime 
            Height          =   615
            Left            =   4320
            Picture         =   "frmClientesNotas.frx":4527
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "Imprimir"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txtVencimento 
            Height          =   315
            Left            =   240
            TabIndex        =   27
            Top             =   3600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109772801
            CurrentDate     =   37664
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "frmClientesNotas.frx":4FA9
            Height          =   2055
            Left            =   120
            OleObjectBlob   =   "frmClientesNotas.frx":4FC2
            TabIndex        =   19
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   1800
            TabIndex        =   20
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label lblTotal2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2280
            TabIndex        =   21
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pagamento:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   2760
            Width           =   1560
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   3720
            TabIndex        =   24
            Top             =   2760
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Pago/Protestado em:"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   3360
            Width           =   1515
         End
      End
      Begin MSComCtl2.DTPicker txtProrrogar 
         Height          =   300
         Left            =   5280
         TabIndex        =   36
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   109772801
         CurrentDate     =   37952
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "frmClientesNotas.frx":5B65
         Height          =   4815
         Left            =   -74880
         OleObjectBlob   =   "frmClientesNotas.frx":5B81
         TabIndex        =   48
         Top             =   480
         Width           =   8655
      End
      Begin MSComCtl2.DTPicker txtResgatado 
         Height          =   315
         Left            =   -71760
         TabIndex        =   55
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   109772801
         CurrentDate     =   37664
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Resgatado em:"
         Height          =   195
         Left            =   -72960
         TabIndex        =   56
         Top             =   5400
         Width           =   1080
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -68640
         TabIndex        =   50
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblTotalProtesto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67920
         TabIndex        =   49
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblPrazo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67560
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDias 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -68280
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dia Pg:"
         Height          =   195
         Index           =   12
         Left            =   -68280
         TabIndex        =   7
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prazo:"
         Height          =   195
         Index           =   34
         Left            =   -67560
         TabIndex        =   9
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblObs 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   -74880
         TabIndex        =   6
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label Label13 
         Caption         =   "Observações:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Nr. Nota:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Código:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Prorrogar para:"
         Height          =   195
         Left            =   5280
         TabIndex        =   35
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Label lblTotalPendente 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   39
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   6000
         TabIndex        =   38
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   2
         Top             =   360
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmClientesNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoCliente As Double
Dim XIni As Double, YIni As Double, XFim As Double, YFim As Double, ValorAPagar As Currency


Private Sub ImprimeNotaFiscal(ByVal NrNota As Double, ByVal CodigoNota As Double)

With dbNotas
  .Recordset.FindFirst "codigonota=" & CodigoNota
  If .Recordset.NoMatch = True Then
    MsgBox "Não foi localizada a nota!"
    Exit Sub
  End If
End With
With dbNotasCorpo
  .RecordSource = "select *from notascorpo where codigonota=" & dbNotas.Recordset!CodigoNota
  .Refresh
  If .Recordset.RecordCount = 0 Then
    If dbNotas.Recordset!servicototal = 0 Then
      MsgBox "A nota não possui preenchimento do corpo da nota!"
      Exit Sub
    End If
  End If
End With
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
  StrTemp = Format(dbNotas.Recordset!notanr, "000000")
  Printer.CurrentX = .Recordset!nrnotatopox
  Printer.CurrentY = .Recordset!nrnotatopoy
  Printer.Print StrTemp
  
  StrTemp = "X"
  If dbNotas.Recordset!Entrada = False Then
    Printer.CurrentX = .Recordset!saidax
    Printer.CurrentY = .Recordset!saiday
  Else
    Printer.CurrentX = .Recordset!entradax
    Printer.CurrentY = .Recordset!entraday
  End If
  Printer.Print StrTemp
  
  StrTemp = dbNotas.Recordset!NaturezaOP
  Printer.CurrentX = .Recordset!naturezaoperacaox
  Printer.CurrentY = .Recordset!naturezaoperacaoy
  Printer.Print StrTemp
  
  StrTemp = dbNotas.Recordset!cfop
  Printer.CurrentX = .Recordset!cfopx
  Printer.CurrentY = .Recordset!cfopy
  Printer.Print StrTemp
  
  StrTemp = Format(dbNotas.Recordset!dataemissao, "short date")
  Printer.CurrentX = .Recordset!dataemissaox
  Printer.CurrentY = .Recordset!dataemissaoy
  Printer.Print StrTemp
  
  StrTemp = Format(dbNotas.Recordset!datasaida, "Short date")
  Printer.CurrentX = .Recordset!datasaidax
  Printer.CurrentY = .Recordset!datasaiday
  Printer.Print StrTemp
  
  StrTemp = Format(dbNotas.Recordset!horasaida, "short time")
  Printer.CurrentX = .Recordset!horasaidax
  Printer.CurrentY = .Recordset!horasaiday
  Printer.Print StrTemp
  
  StrTemp = dbNotas.Recordset!Nome
  Printer.CurrentX = .Recordset!nomex
  Printer.CurrentY = .Recordset!nomey
  Printer.Print StrTemp
  
  If IsNull(dbNotas.Recordset!CNPJ) = False Then
    StrTemp = dbNotas.Recordset!CNPJ
    Printer.CurrentX = .Recordset!cnpjx
    Printer.CurrentY = .Recordset!cnpjy
    Printer.Print StrTemp
  End If
  
  StrTemp = dbNotas.Recordset!Endereco
  Printer.CurrentX = .Recordset!enderecox
  Printer.CurrentY = .Recordset!enderecoy
  Printer.Print StrTemp
  
  StrTemp = dbNotas.Recordset!bairro
  Printer.CurrentX = .Recordset!bairrox
  Printer.CurrentY = .Recordset!bairroy
  Printer.Print StrTemp
  
  If IsNull(dbNotas.Recordset!CEP) = False Then
    StrTemp = dbNotas.Recordset!CEP
    Printer.CurrentX = .Recordset!cepx
    Printer.CurrentY = .Recordset!cepy
    Printer.Print StrTemp
  End If
  
  If IsNull(dbNotas.Recordset!municipio) = False Then
    StrTemp = dbNotas.Recordset!municipio
    Printer.CurrentX = .Recordset!municipiox
    Printer.CurrentY = .Recordset!municipioy
    Printer.Print StrTemp
  End If
  
  If IsNull(dbNotas.Recordset!fone) = False Then
    StrTemp = dbNotas.Recordset!fone
  Else
    StrTemp = ""
  End If
  Printer.CurrentX = .Recordset!fonex
  Printer.CurrentY = .Recordset!foney
  Printer.Print StrTemp
  
  StrTemp = dbNotas.Recordset!uf
  Printer.CurrentX = .Recordset!uf1x
  Printer.CurrentY = .Recordset!uf1y
  Printer.Print StrTemp
  
  If IsNull(dbNotas.Recordset!ie) = False Then
    StrTemp = dbNotas.Recordset!ie
    Printer.CurrentX = .Recordset!iex
    Printer.CurrentY = .Recordset!iey
    Printer.Print StrTemp
  End If
  
  StrTemp = dbNotas.Recordset!dadosfatura
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
        If IsNull(.Recordset!Quantidade) = False Then StrTemp = Format(.Recordset!Quantidade, "0.000")
        Printer.CurrentX = dbConfigNota.Recordset!colunavunitario - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!valorUnitario) = False Then StrTemp = Format(.Recordset!valorUnitario, "0.000")
        Printer.CurrentX = dbConfigNota.Recordset!colunavtotal - 0.2 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!ValorTotal) = False Then StrTemp = Format(.Recordset!ValorTotal, "0.00")
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
  
  
  If IsNull(dbNotas.Recordset!servico) = False Then
    If dbNotas.Recordset!servico <> "" Then
      StrTemp = ""
      StrTemp = dbNotas.Recordset!servico
      Printer.CurrentX = .Recordset!prestacaoservicox
      Printer.CurrentY = .Recordset!prestacaoservicoy
      Printer.Print StrTemp
      
      StrTemp = ""
      StrTemp = dbNotas.Recordset!servicoiss
      Printer.CurrentX = .Recordset!prestacaoservicoissx
      Printer.CurrentY = .Recordset!prestacaoservicoissy
      Printer.Print StrTemp
      
      StrTemp = ""
      StrTemp = Format(dbNotas.Recordset!servicototal, "currency")
      Printer.CurrentX = .Recordset!prestacaoservicototalx
      Printer.CurrentY = .Recordset!prestacaoservicototaly
      Printer.Print StrTemp
    End If
  End If
  
  If dbNotas.Recordset!servicoiss <> 0 Then
    
  End If
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!BaseICMS, "currency")
  Printer.CurrentX = .Recordset!baseicmsx
  Printer.CurrentY = .Recordset!baseicmsy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!ValorICMS, "currency")
  Printer.CurrentX = .Recordset!valoricmsx
  Printer.CurrentY = .Recordset!valoricmsy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!baseicmssubst, "currency")
  Printer.CurrentX = .Recordset!baseicmssubx
  Printer.CurrentY = .Recordset!baseicmssuby
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!ValorICMSSubst, "currency")
  Printer.CurrentX = .Recordset!valoricmssubx
  Printer.CurrentY = .Recordset!valoricmssuby
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!totaldosprodutos, "currency")
  Printer.CurrentX = .Recordset!valortotalprodutosx
  Printer.CurrentY = .Recordset!valortotalprodutosy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!ValorFrete, "currency")
  Printer.CurrentX = .Recordset!valordofretex
  Printer.CurrentY = .Recordset!valordofretey
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!ValorSeguro, "currency")
  Printer.CurrentX = .Recordset!valordosegurox
  Printer.CurrentY = .Recordset!valordoseguroy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!OutrasDespesas, "currency")
  Printer.CurrentX = .Recordset!outrasdespx
  Printer.CurrentY = .Recordset!outrasdespy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!valoripi, "currency")
  Printer.CurrentX = .Recordset!valortotalipix
  Printer.CurrentY = .Recordset!valortotalipiy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!ValorTotalDaNota, "currency")
  Printer.CurrentX = .Recordset!valortotalnotax
  Printer.CurrentY = .Recordset!valortotalnotay
  Printer.Print StrTemp
  
  On Error Resume Next
  StrTemp = ""
  StrTemp = dbNotas.Recordset!nome2
  Printer.CurrentX = .Recordset!nome2x
  Printer.CurrentY = .Recordset!nome2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!FretePorConta
  Printer.CurrentX = .Recordset!freteporcontax
  Printer.CurrentY = .Recordset!freteporcontay
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Placa
  Printer.CurrentX = .Recordset!placax
  Printer.CurrentY = .Recordset!placay
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!UF2
  Printer.CurrentX = .Recordset!uf2x
  Printer.CurrentY = .Recordset!uf2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!CNPJ2
  Printer.CurrentX = .Recordset!cnpj2x
  Printer.CurrentY = .Recordset!cnpj2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Endereco2
  Printer.CurrentX = .Recordset!endereco2x
  Printer.CurrentY = .Recordset!endereco2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Municipio2
  Printer.CurrentX = .Recordset!municipio2x
  Printer.CurrentY = .Recordset!municipio2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!UF3
  Printer.CurrentX = .Recordset!uf3x
  Printer.CurrentY = .Recordset!uf3y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!IE2
  Printer.CurrentX = .Recordset!ie2x
  Printer.CurrentY = .Recordset!ie2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Qtd2
  Printer.CurrentX = .Recordset!qtd2x
  Printer.CurrentY = .Recordset!qtd2y
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Especie
  Printer.CurrentX = .Recordset!especiex
  Printer.CurrentY = .Recordset!especiey
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Marca
  Printer.CurrentX = .Recordset!marcax
  Printer.CurrentY = .Recordset!marcay
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!Numero
  Printer.CurrentX = .Recordset!numerox
  Printer.CurrentY = .Recordset!numeroy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!PesoBruto
  Printer.CurrentX = .Recordset!pesobrutox
  Printer.CurrentY = .Recordset!pesobrutoy
  Printer.Print StrTemp
  
  StrTemp = ""
  StrTemp = dbNotas.Recordset!PesoLiquido
  Printer.CurrentX = .Recordset!pesoliquidox
  Printer.CurrentY = .Recordset!pesoliquidoy
  Printer.Print StrTemp
  
  ImprimeTextoJustificado Printer, dbNotas.Recordset!dadosadicionais, AlinhaEsquerda, .Recordset!dadosadicionais1x, .Recordset!dadosadicionais1y, .Recordset!dadosadicionais2x
  
  StrTemp = ""
  StrTemp = Format(dbNotas.Recordset!notanr, "000000")
  Printer.CurrentX = .Recordset!nrnotacanhotox
  Printer.CurrentY = .Recordset!nrnotacanhotoy
  Printer.Print StrTemp;
  
  If IsNull(dbNotas.Recordset!nota) = False Then
    StrTemp = ""
    StrTemp = "NF. " & dbNotas.Recordset!nota
    Printer.CurrentX = 9
    Printer.CurrentY = .Recordset!nrnotacanhotoy
    Printer.Print StrTemp;
  End If
  
  Printer.EndDoc
  
End With
NaoImprime:

End Sub

Private Sub AtualizaNotas()
Dim Ws As Workspace, db As Database
Adicionando = False
With dbConfigNota
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from confignota"
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
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    For i = 0 To .Recordset.Fields.Count - 1
      .Recordset(i) = 0
    Next i
    .Recordset.Update
  End If
End With
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
    db.Execute "create table NotasCorpo (CodigoNotaCorpo counter, CodigoNota double, CodigoProduto Text(20), DescriProduto Text(255), ClasFiscal Text(4), SubTributaria Text(4), Unidade Text(10), Quantidade double, ValorUnitario currency, ValorTotal currency, AliquotaICMS double, AliquotaIPI double, ValorIpi Currency)"
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
    db.Execute "create table Notas (CodigoNota counter, NaturezaOP text(30), CFOP Text(10), NotaNr double, Entrada bit, DataEmissao datetime, DataSaida DateTime, HoraSaida DateTime, Nome Text(130), CNPJ text(30), Endereco Text(130), Bairro Text(80), Cep Text(20), Municipio Text(50), Fone Text(20), UF Text(2), Ie Text(30), DadosFatura Text(250), BaseICMS Currency, ValorICMS Currency, BaseICMSSubst Currency, ValorICMSSubst Currency, TotalDosProdutos Currency, ValorFrete Currency, ValorSeguro Currency, OutrasDespesas Currency, ValorIPI currency, ValorTotalDaNota Currency, Nome2 Text(130), FretePorConta integer, Placa text(20), UF2 Text(2), CNPJ2 Text(30), Endereco2 Text(130), Municipio2 Text(50), UF3 text(2), IE2 text(30), Quantidade2 Text(20), Especie Text(20), Marca Text(30), Numero Text (20), PesoBruto Text(20), PesoLiquido Text(20), DadosAdicionais Text(255))"
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

Private Sub AtivaCliente()
  dbClientes.Recordset.Edit
  dbClientes.Recordset!mensalista = True
  dbClientes.Recordset.Update
End Sub

Private Sub CabecaNotas(ByVal Dia As Date, Largura As Double)
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relação de Documentos para Cobrança"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  
  Printer.FontSize = 10
  
  StrTemp = "Cliente: " & cboCliente.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
  
  StrTemp = "Impresso em: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.Print ""
  
  YIni = Printer.CurrentY
  XIni = 0
  XFim = Largura
  
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  StrTemp = "Data"
  Printer.CurrentX = 1
  Printer.Print StrTemp;
  
  StrTemp = "Nr. Documento"
  Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Valor"
  Printer.CurrentX = 104 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
End Sub

Private Sub ImprimeBoletosBB()
Dim StrTemp As String, DataDoc As Date, DataVenc As Date
Dim Praso As Double, Nome As String, Endereco As String
Dim Instrucao As String, CEP As String, NrDoc As Double
Dim Valor As Currency


With dbPendencias
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  .Recordset.FindFirst "datafechamento>=#" & DataInglesa(Trim(Str(txtDataBoleto.Value))) & "#"
  
  If .Recordset.NoMatch = True Then
    MsgBox "A data informada é maior que o vencimento dos boletos em aberto!"
    Exit Sub
  End If
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  On Error Resume Next
  Printer.PaperSize = vbPRPSFanfoldStdGerman
  On Error GoTo 0
  B = 0
  Do While .Recordset.EOF = False
    A = A + 1
    If IsNull(.Recordset!NrNota) = True Then
      NrDoc = .Recordset!CodigoCobranca
    Else
      NrDoc = .Recordset!NrNota
    End If
    Valor = .Recordset!Valor
    DataVenc = .Recordset!DataFechamento
    With dbClientes
      .Refresh
      .Recordset.FindFirst "codigocliente=" & dbPendencias.Recordset!CodigoCliente
      'Praso = .Recordset!Praso
      If IsNull(.Recordset!nome2) = True Then
        Nome = .Recordset!Nome
      Else
        Nome = .Recordset!nome2
      End If
      Endereco = .Recordset!Endereco & "  " & .Recordset!complemento
      CEP = Format(.Recordset!CEP, "00000-000") & "   " & .Recordset!bairro & "   " & .Recordset!cidade & " - " & .Recordset!Estado
      Instrucao = .Recordset!Instrucoes
      DataDoc = Date
      'DataVenc = DateAdd("d", Praso, DataDoc)
    End With
    StrTemp = Format(DataVenc, "dd/mm/yyyy")
    Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
    Printer.CurrentY = B + 10
    Printer.Print StrTemp
    
    StrTemp = Format(DataDoc, "dd/mm/yyyy")
    Printer.CurrentX = 23 - Printer.TextWidth(StrTemp)
    Printer.CurrentY = B + 24
    Printer.Print StrTemp;
    
    StrTemp = NrDoc
    Printer.CurrentX = 72 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    StrTemp = Format(Valor, "Currency")
    Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
    Printer.CurrentY = B + 30
    Printer.Print StrTemp
    
    StrTemp = Instrucao
    Printer.CurrentX = 0
    Printer.CurrentY = B + 37
    Printer.Print StrTemp
    
    StrTemp = Nome
    Printer.CurrentX = 0
    Printer.CurrentY = B + 68
    Printer.Print StrTemp
    
    
    StrTemp = Endereco
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 0.25
    Printer.Print StrTemp
    
    StrTemp = CEP
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 0.25
    Printer.Print StrTemp
    
    If A >= 3 Then
      A = 0
      B = 0
      Printer.NewPage
    Else
      B = B + 101
    End If
    .Recordset.MoveNext
  Loop
End With
Printer.EndDoc
NaoImprime:

End Sub

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
  Dim StrTemp As String
  
  Printer.FontName = "Arial"
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 16
  
  StrTemp = "Relatório de Contas a Receber"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontSize = 9
  
  StrTemp = Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página:" & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Vencimento"
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Nr. Doc."
  Printer.CurrentX = 25
  Printer.Print StrTemp;
  
  StrTemp = "Cod."
  Printer.CurrentX = 53 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Cliente"
  Printer.CurrentX = 55
  Printer.Print StrTemp;
  
  StrTemp = "Valor"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  
End Sub

Private Sub cboCliente_LostFocus()
With dbClientes
  .Refresh
  .Recordset.FindFirst "nome='" & cboCliente.Text & "'"
  If .Recordset.NoMatch = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
    If IsNull(.Recordset!Obs) = False Then
      lblObs.Caption = .Recordset!Obs
    Else
      lblObs.Caption = ""
    End If
    lblDias.Caption = .Recordset!diapagamento
    lblPrazo.Caption = .Recordset!Praso
  End If
End With
Call cmdExibe_Click
End Sub

Private Sub cboFormaDePg_LostFocus()
With dbFormaDePg
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboFormadePg.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFormadePg.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub chkConfirmadas_Click()
  If chkConfirmadas.Value = vbChecked Then
    cmdFechar.Enabled = False
  Else
    cmdFechar.Enabled = True
  End If
  Call cmdExibe_Click
End Sub

Private Sub cmdAlteraNota_Click()
Dim CodigoNota As Double, Total As Currency, TotalProdutos As Currency
Dim db As Database, Ws As Workspace
Dim ValorICMS As Currency, TotalICMS As Currency
Dim BaseDeCalculoICMS As Currency

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)


With dbPendencias
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma cobrança primeiro!"
    Exit Sub
  End If
  If txtNrNota.Text = "" Then
    MsgBox "Informe o número da nota!"
    txtNrNota.SetFocus
    Exit Sub
  End If
  .Recordset.Edit
  .Recordset!NrNota = txtNrNota.Text
  .Recordset.Update
End With
dbClientes2.Recordset.FindFirst "codigocliente=" & dbPendencias.Recordset!CodigoCliente
If dbClientes2.Recordset.NoMatch = True Then
  MsgBox "Cliente não encontrado!"
  Exit Sub
End If
With dbNotas
  .RecordSource = "Select *from notas where codigoboleto=" & dbPendencias.Recordset!CodigoCobranca
  .Refresh
  Do While .Recordset.RecordCount <> 0
    CodigoNota = .Recordset!CodigoNota
    db.Execute "delete *from notas where codigonota=" & CodigoNota
    db.Execute "delete *from notascorpo where codigonota=" & CodigoNota
    .Refresh
    .Refresh
  Loop
  
  .Recordset.AddNew
  CodigoNota = .Recordset!CodigoNota
  .Recordset!notanr = txtNrNota.Text
  .Recordset!cfop = "5929"
  .Recordset!NaturezaOP = "Venda"
  .Recordset!Entrada = False
  .Recordset!dataemissao = Date
  .Recordset!datasaida = Date
  .Recordset!horasaida = Time
  .Recordset!Nome = dbClientes2.Recordset!nome2
  .Recordset!CNPJ = dbClientes2.Recordset!CNPJ
  .Recordset!codmunicipio = dbClientes2.Recordset("codigo")
  .Recordset!Endereco = dbClientes2.Recordset!Endereco
  .Recordset!bairro = dbClientes2.Recordset!bairro
  .Recordset!CEP = dbClientes2.Recordset!CEP
  .Recordset!municipio = dbClientes2.Recordset("municipios.nome")
  .Recordset!fone = dbClientes2.Recordset!Telefone
  .Recordset!uf = dbClientes2.Recordset!Estado
  .Recordset!ie = dbClientes2.Recordset!ie
  .Recordset!dadosfatura = "Fatura número " & txtNrNota.Text
  .Recordset!codigoboleto = dbPendencias.Recordset!CodigoCobranca
  .Recordset!BaseICMS = 0
  .Recordset!ValorICMS = 0
  .Recordset!baseicmssubst = 0
  .Recordset!valoripi = 0
  .Recordset!dadosadicionais = "Conforme demonstrativo de venda em anexo."
  .Recordset!servicototal = 0
  If IsNull(dbClientes2.Recordset!nota) = False Then .Recordset!nota = dbClientes2.Recordset!nota
  .Recordset.Update
End With
With dbNotasCorpo
  .RecordSource = "Select *from notascorpo where codigonota=" & CodigoNota
  .Refresh
  Do While .Recordset.RecordCount <> 0
    .Recordset.Delete
    .Refresh
  Loop
End With
With dbProdutos2
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade, valorunitariodif, codigo, descri, cfop, cst, codigoOrigem, AliquotaICMS from qclientesnota2produtos where codigosoma='" & dbPendencias.Recordset!codigoSoma & "' group by valorunitariodif, codigo, descri, cfop, cst, codigoorigem, AliquotaICMS order by codigo"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    TotalICMS = 0
    Do While .Recordset.EOF = False
      With dbNotasCorpo
        dbProdutos.Recordset.FindFirst "codigo=" & dbProdutos2.Recordset!Codigo
        If dbProdutos.Recordset!servico = False Then
          'falta preencher o corpo da nota e calcular os tributos
          .Recordset.AddNew
          .Recordset!CodigoNota = CodigoNota
          .Recordset!CodigoProduto = dbProdutos2.Recordset!Codigo
          .Recordset!descriproduto = dbProdutos2.Recordset!Descri
          If IsNull(dbProdutos.Recordset!unidade) = False Then
            .Recordset!unidade = dbProdutos.Recordset!unidade
          Else
            .Recordset!unidade = "LT"
          End If
          .Recordset!Quantidade = dbProdutos2.Recordset!Quantidade
          .Recordset!valorUnitario = dbProdutos2.Recordset!ValorUnitarioDif
          .Recordset!ValorTotal = dbProdutos2.Recordset!Total
          Total = Total + .Recordset!ValorTotal
          TotalProdutos = TotalProdutos + .Recordset!ValorTotal
          .Recordset!cfop = dbProdutos2.Recordset!cfop
          .Recordset!clasfiscal = dbProdutos2.Recordset!cst
          .Recordset!Origem = dbProdutos2.Recordset!CodigoOrigem
          .Recordset!aliquotaicms = dbProdutos2.Recordset!aliquotaicms
          ValorICMS = 0
          If IsNull(dbProdutos2.Recordset!aliquotaicms) = False Then
            If dbProdutos2.Recordset!aliquotaicms <> 0 Then
              ValorICMS = dbProdutos2.Recordset!Total * (dbProdutos2.Recordset!aliquotaicms / 100)
            End If
          End If
          .Recordset!ValorICMS = ValorICMS
          TotalICMS = TotalICMS + ValorICMS
          BaseDeCalculoICMS = BaseDeCalculoICMS + dbProdutos2.Recordset!Total
          .Recordset.Update
'        Else
'          dbNotas.Recordset.FindFirst "codigonota=" & CodigoNota
'          If dbNotas.Recordset.NoMatch = False Then
'            dbNotas.Recordset.Edit
'            dbNotas.Recordset!servico = dbProdutos.Recordset!descriservico
'            dbNotas.Recordset!servicoiss = 0
'            dbNotas.Recordset!servicototal = dbNotas.Recordset!servicototal + dbProdutos2.Recordset!Total
'            Total = Total + dbProdutos2.Recordset!Total
'            dbNotas.Recordset.Update
'          End If
        End If
      End With
      .Recordset.MoveNext
      dbConfigNota.RecordSource = "select *from confignota"
      dbConfigNota.Refresh
      If IsNull(dbConfigNota.Recordset!linhascorpo) = True Then
        dbConfigNota.Recordset.Edit
        dbConfigNota.Recordset!linhascorpo = 999
        dbConfigNota.Recordset.Update
      End If
      
      If dbNotasCorpo.Recordset.RecordCount >= dbConfigNota.Recordset!linhascorpo Then
        dbNotas.Recordset.FindFirst "codigonota=" & CodigoNota
        If dbNotas.Recordset.NoMatch = False Then
          dbNotas.Recordset.Edit
          dbNotas.Recordset!BaseICMS = BaseDeCalculoICMS
          dbNotas.Recordset!ValorICMS = TotalICMS
          dbNotas.Recordset!totaldosprodutos = TotalProdutos
          dbNotas.Recordset!ValorTotalDaNota = Total
          dbNotas.Recordset.Update
        End If
        ImprimeNotaFiscal txtNrNota.Text, CodigoNota
        BaseDeCalculoICMS = 0
        ValorICMS = 0
        Total = 0
        TotalProdutos = 0
        If .Recordset.EOF = False Then
          txtNrNota.Text = InputBox("Informe o número da próxima nota!", "Próxima nota!", CLng(txtNrNota.Text) + 1)
        Else
          GoTo FimDaNota
        End If
        If txtNrNota.Text = "" Then
          MsgBox "Esta operação será cancelada sem concluir os itens a serem impressos!"
          Exit Sub
        End If
        With dbNotas
          .Recordset.AddNew
          CodigoNota = .Recordset!CodigoNota
          .Recordset!notanr = txtNrNota.Text
          .Recordset!cfop = "5929"
          .Recordset!NaturezaOP = "Venda"
          .Recordset!Entrada = False
          .Recordset!dataemissao = Date
          .Recordset!datasaida = Date
          .Recordset!horasaida = Time
          .Recordset!Nome = dbClientes2.Recordset!nome2
          .Recordset!CNPJ = dbClientes2.Recordset!CNPJ
          .Recordset!Endereco = dbClientes2.Recordset!Endereco
          .Recordset!bairro = dbClientes2.Recordset!bairro
          .Recordset!CEP = dbClientes2.Recordset!CEP
          .Recordset!codmunicipio = dbClientes2.Recordset("codigo")
          .Recordset!municipio = dbClientes2.Recordset("municipios.nome")
          .Recordset!fone = dbClientes2.Recordset!Telefone
          .Recordset!uf = dbClientes2.Recordset!Estado
          .Recordset!ie = dbClientes2.Recordset!ie
          .Recordset!dadosfatura = "Fatura número " & txtNrNota.Text
          .Recordset!codigoboleto = dbPendencias.Recordset!CodigoCobranca
          .Recordset!BaseICMS = 0
          .Recordset!ValorICMS = 0
          .Recordset!baseicmssubst = 0
          .Recordset!valoripi = 0
          .Recordset!dadosadicionais = "Conforme demonstrativo de venda em anexo."
          .Recordset.Update
          dbNotasCorpo.RecordSource = "Select *from Notascorpo where codigonota=" & CodigoNota
          dbNotasCorpo.Refresh
        End With
      End If
      
    Loop
  End If
End With
dbNotas.Recordset.FindFirst "codigonota=" & CodigoNota
If dbNotas.Recordset.NoMatch = False Then
  dbNotas.Recordset.Edit
  dbNotas.Recordset!BaseICMS = BaseDeCalculoICMS
  dbNotas.Recordset!ValorICMS = TotalICMS
  dbNotas.Recordset!totaldosprodutos = TotalProdutos
  dbNotas.Recordset!ValorTotalDaNota = Total
  dbNotas.Recordset.Update
End If

If IsNumeric(txtNrNota.Text) = False Then
    MsgBox "Informe o número da nota fiscal correto!"
    Exit Sub
End If
If NotaExiste(txtNrNota.Text, dbNotas.Recordset!CodigoNota) = True Then
    MsgBox "Nota fiscal já emitida!"
    Exit Sub
End If


ImprimeNotaFiscal txtNrNota.Text, CodigoNota
FimDaNota:
txtNrNota.Text = ""
End Sub

Private Sub cmdAtualiza_Click()
dbPendencias.Refresh
QPendencias.Refresh
If IsNull(QPendencias.Recordset!Total) = False Then
  lblTotalPendente.Caption = Format(QPendencias.Recordset!Total, "Currency")
Else
  lblTotalPendente.Caption = Format(0, "Currency")
End If
End Sub

Private Sub cmdAtualizaProtestos_Click()
Dim TempValor As Currency

With dbProtestados
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qProtestados
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblTotalProtesto.Caption = Format(TempValor, "Currency")
End With
End Sub

Private Sub cmdConfiguraNota_Click()
frmConfigNota.Show vbModal
End Sub

Private Sub cmdExibe_Click()
Dim Confirmado As Double
If chkConfirmadas.Value = vbChecked Then
  Confirmado = -1
Else
  Confirmado = 0
End If
With dbClientes
  If .Recordset.EOF = True Then
    MsgBox "Cliente não encontrado!"
    Exit Sub
  End If
  CodigoCliente = .Recordset!CodigoCliente
End With
With dbClientesNotas2
  .Connect = Conectar
  If Confirmado = 0 Then
    .RecordSource = "select *from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=" & Confirmado & " and data<=#" & DataInglesa(Trim(Str(txtFechamento.Value))) & "#  order by data"
  Else
    .RecordSource = "select *from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=" & Confirmado & " and data>=#" & DataInglesa(Trim(Str(txtFechamento.Value))) & "# order by data"
  End If
  .DatabaseName = Caminho
  .Refresh
End With
With QSoma
  .Connect = Conectar
  .DatabaseName = Caminho
  If Confirmado = 0 Then
    .RecordSource = "select sum(valorprevisto) as total from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=" & Confirmado & " and data<=#" & DataInglesa(Trim(Str(txtFechamento.Value))) & "#"
  Else
    .RecordSource = "select sum(valorprevisto) as total from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=" & Confirmado & " and data>=#" & DataInglesa(Trim(Str(txtFechamento.Value))) & "#"
  End If
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  End If
End With
lblTotal.Caption = Format(TempValor, "Currency")
With dbCobranca
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from ClientesCobranca where codigocliente=" & CodigoCliente & " and pago=0 order by datafechamento"
  .Refresh
End With
End Sub

Private Sub cmdExtornar_Click()
Dim Resposta As Integer
Dim db As New ADODB.Connection

With dbCobranca
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma fatura primeiro"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja cancelar o fechamento atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  db.Open CaminhoADO
  db.Execute "update clientesnota2 set confirmado=0 where codigosoma='" & .Recordset!codigoSoma & "' and codigocliente=" & .Recordset!CodigoCliente
  db.Close
  With dbClientes
    .Recordset.Edit
    .Recordset!TotalBoleto = .Recordset!TotalBoleto - dbCobranca.Recordset!Valor
    .Recordset!TotalNotas = .Recordset!TotalNotas + dbCobranca.Recordset!Valor
    .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
    .Recordset.Update
  End With
  .Recordset.Delete
  .Refresh
End With
End Sub

Private Sub cmdFechar_Click()
Dim ValorTotal As Currency, codigoSoma As String
Dim Resposta As Integer, Vencimento As Date, Dias As Double
Dim Confirmado As Double, CodigoCobranca As Double

If chkConfirmadas.Value = vbChecked Then
  MsgBox "Não pode fechar com a opção 'Notas confirmadas' ativado!"
  Exit Sub
End If
'If DateDiff("d", Date, txtFechamento.Value) > 30 Then
'  MsgBox "Data muito futura!"
'  Exit Sub
'End If
'If DateDiff("d", Date, txtFechamento.Value) < -30 Then
'  If Usuarios.Grupo.AdmEstatus = 2 Then
'    Resposta = MsgBox("Data muito antiga! Deseja continuar?", vbYesNo + vbDefaultButton2)
'    If Resposta = vbNo Then
'      Exit Sub
'    End If
'  Else
'    MsgBox "Data muito antiga!"
'    Exit Sub
'  End If
'End If

Resposta = MsgBox("Deseja realmente fazer o fechamento do dia atual?", vbInformation + vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

Resposta = MsgBox("Deseja Imprimir a listagem antes de fazer o fechamento dos clientes?", vbYesNoCancel)
Select Case Resposta
  Case vbYes
    Call cmdImprimeNotas_Click
  Case vbCancel
    Exit Sub
End Select

With QSoma
  .RecordSource = "select sum(valorprevisto) as total from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=0 and data<=#" & DataInglesa(Trim(Str(txtFechamento.Value))) & "#"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    ValorTotal = ValorTotal + .Recordset!Total
  End If
  
End With
With dbClientesNotas2
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe nota para ser finalizada!"
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If .Recordset!Autorizar = True Then
      If .Recordset!Autorizado = False Then
        MsgBox "Existe nota não autorizada! Não será possível finalizar!"
        Exit Sub
      End If
    End If
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
End With

codigoSoma = Str(CDbl(Now))
Dias = dbClientes.Recordset!Praso
Vencimento = DateAdd("d", Dias, txtFechamento.Value)
With dbCobranca
  .Recordset.AddNew
  CodigoCobranca = .Recordset!CodigoCobranca
  .Recordset!datasoma = Now
  .Recordset!DataFechamento = Vencimento
  .Recordset!codigoSoma = codigoSoma
  .Recordset!CodigoCliente = CodigoCliente
  .Recordset!Cliente = cboCliente.Text
  .Recordset!Valor = ValorTotal
  .Recordset!Origem = "Fiado"
  .Recordset!Descri = "Fechamento até " & Format(txtFechamento.Value, "Short date")
  .Recordset.Update
  .Refresh
End With
With dbClientesNotas2
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!codigoSoma = codigoSoma
      .Recordset!Confirmado = True
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  .Refresh
End With
With dbClientes
  .Recordset.Edit
  .Recordset!TotalBoleto = .Recordset!TotalBoleto + ValorTotal
  .Recordset!TotalNotas = .Recordset!TotalNotas - ValorTotal
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With
Call cboCliente_LostFocus
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, DataDoc As Date, DataVenc As Date
Dim Praso As Double, Nome As String, Endereco As String
Dim Instrucao As String, CEP As String, NrDoc As Double
Dim Valor As Currency


With dbCobranca
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Escolha uma cobrança primeiro!"
    Exit Sub
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  If IsNull(.Recordset!NrNota) = False Then
    NrDoc = .Recordset!NrNota
  Else
    NrDoc = .Recordset!CodigoCobranca
  End If
  Valor = .Recordset!Valor
  DataVenc = .Recordset!DataFechamento
  With dbClientes
    'Praso = .Recordset!Praso
    If IsNull(.Recordset!nome2) = True Then
      Nome = .Recordset!Nome
    Else
      Nome = .Recordset!nome2
    End If
    Endereco = .Recordset!Endereco & "  " & .Recordset!complemento
    CEP = Format(.Recordset!CEP, "00000-000") & "   " & .Recordset!bairro & "   " & .Recordset!cidade & " - " & .Recordset!Estado
    Instrucao = .Recordset!Instrucoes
    DataDoc = Date
    'DataVenc = DateAdd("d", Praso, DataDoc)
  End With
  StrTemp = Format(DataVenc, "dd/mm/yyyy")
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.CurrentY = 10
  Printer.Print StrTemp
  
  StrTemp = Format(DataDoc, "dd/mm/yyyy")
  Printer.CurrentX = 23 - Printer.TextWidth(StrTemp)
  Printer.CurrentY = 24
  Printer.Print StrTemp;
  
  StrTemp = NrDoc
  Printer.CurrentX = 72 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = Format(Valor, "Currency")
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.CurrentY = 30
  Printer.Print StrTemp
  
  StrTemp = Instrucao
  Printer.CurrentX = 0
  Printer.CurrentY = 37
  Printer.Print StrTemp
  
  StrTemp = Nome
  Printer.CurrentX = 0
  Printer.CurrentY = 68
  Printer.Print StrTemp
  
  
  StrTemp = Endereco
  Printer.CurrentX = 0
  Printer.CurrentY = Printer.CurrentY - 0.25
  Printer.Print StrTemp
  
  StrTemp = CEP
  Printer.CurrentX = 0
  Printer.CurrentY = Printer.CurrentY - 0.25
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub cmdImprime2_Click()
Dim Resposta As Integer
Dim Largura As Double, StrTemp As String, Dia As Date

With dbPendencias
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  .Recordset.MoveFirst
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 9
  Largura = 190
  Dia = Now
  Cabeca Largura, Dia
  
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 35 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    On Error Resume Next
    StrTemp = .Recordset!DataFechamento
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!NrNota) = False Then
      StrTemp = .Recordset!NrNota
    Else
      StrTemp = .Recordset!CodigoCobranca
    End If
    Printer.CurrentX = 25
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoCliente
    Printer.CurrentX = 53 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Cliente
    Printer.CurrentX = 55
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Valor, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(lblTotalPendente.Caption, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With

NaoImprime:

End Sub

Private Sub cmdImprimeNotas_Click()
  Dim Dia As Date, Largura As Double, StrTemp As String
  Dim Total As Currency
  
  If cboCliente.Text <> dbClientes.Recordset!Nome Then
    MsgBox "Selecione um cliente!"
    cboCliente.SetFocus
    Exit Sub
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.DrawWidth = 3
  Total = 0
  Dia = Now
  Largura = 105
  
  CabecaNotas Dia, Largura
  
  With dbClientesNotas2
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      
      Do While .Recordset.EOF = False
        If Printer.CurrentY >= Printer.ScaleHeight - 25 Then
          YFim = Printer.CurrentY
          Printer.Line (0, YIni)-(0, YFim)
          Printer.Line (30, YIni)-(30, YFim)
          Printer.Line (65, YIni)-(65, YFim)
          Printer.Line (105, YIni)-(105, YFim)
          
          Printer.NewPage
          
          CabecaNotas Dia, Largura
        End If
        StrTemp = Format(.Recordset!Data, "short date")
        Printer.CurrentX = 1
        Printer.Print StrTemp;
        
        StrTemp = .Recordset!Cupom
        Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Total = Total + .Recordset!ValorPrevisto
        StrTemp = Format(.Recordset!ValorPrevisto, "currency")
        Printer.CurrentX = 104 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.5
        
        .Recordset.MoveNext
      Loop
    End If
  End With
  
  
  YFim = Printer.CurrentY
  Printer.Line (0, YIni)-(0, YFim)
  Printer.Line (30, YIni)-(30, YFim)
  Printer.Line (65, YIni)-(65, YFim)
  Printer.Line (105, YIni)-(105, YFim)
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  StrTemp = "Total=" & Format(Total, "currency")
  Printer.CurrentX = 104 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.Print ""
  Printer.Print ""
  Printer.Print ""
  Printer.Print ""
  
  Printer.Line (Largura - 70, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  StrTemp = "Assinatura do Cliente"
  Printer.CurrentX = (Largura - 35) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.EndDoc
NaoImprime:
End Sub

Private Sub cmdNotaDeOutroCliente_Click()
Dim Resposta As Integer, NovoCodigo As Double, AntigoCodigo As Double, NovoNome As String, AntigoNome As String


With dbClientesNotas2
  If .Recordset.EOF = True Then
    MsgBox "Precisa selecionar uma nota primeiro!"
    Exit Sub
  End If
  If .Recordset!Confirmado = True Then
    MsgBox "Esta nota já foi fechada em cobrança.!"
    Exit Sub
  End If
End With
Resposta = MsgBox("Deseja trocar a nota atual de cliente?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
frmClientesNotasAlteraCupom.Show vbModal

With frmClientesNotasAlteraCupom
  NovoCodigo = .CodigoCliente
  NovoNome = .Nome
End With

Unload frmClientesNotasAlteraCupom
With dbClientes
  .Recordset.Edit
  .Recordset!TotalNotas = .Recordset!TotalNotas - dbClientesNotas2.Recordset!ValorPrevisto
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With
With dbClientes2
  .Recordset.FindFirst "codigocliente=" & NovoCodigo
  .Recordset.Edit
  .Recordset!TotalNotas = .Recordset!TotalNotas + dbClientesNotas2.Recordset!ValorPrevisto
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With
With dbClientesNotas2
  If .Recordset!CodigoCliente = NovoCodigo Then Exit Sub
  .Recordset.Edit
  AntigoCodigo = .Recordset!CodigoCliente
  AntigoNome = .Recordset!Nome
  .Recordset!CodigoCliente = NovoCodigo
  .Recordset!Nome = NovoNome
  .Recordset!clienteantigo = AntigoCodigo
  .Recordset!usuariotroca = Usuarios.Nome
  .Recordset.Update
  Call cmdExibe_Click
End With

End Sub

Private Sub cmdProrrogar_Click()
Dim Resposta As Integer

If Usuarios.Grupo.AdmEstatus <> 2 Then
  MsgBox "Somente usuário administrativo pode prorrogar um boleto!"
  Exit Sub
End If


Resposta = MsgBox("Deseja prorrogar o título atual para a data indicada?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbPendencias
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  .Recordset.Edit
  .Recordset!DataFechamento = txtProrrogar.Value
  .Recordset.Update
  .Refresh
End With
End Sub

Private Sub cmdProtestar_Click()
Dim TempValor As Currency, StrDescri As String, CodigoCobranca As Double
Dim Resposta As Integer, DataDocumento As Date, ValorRecebido As Currency
Dim StrTemp As String, CodigoCliente As Double, NrNota As String

Resposta = MsgBox("Deseja Protestar o valor atual?", vbYesNo + vbDefaultButton2, App.Title)
If Resposta = vbNo Then Exit Sub

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If dbCobranca.Recordset.RecordCount = 0 Then Exit Sub
If dbCobranca.Recordset.EOF = True Then
  MsgBox "Escolha uma cobrança primeiro!"
  Exit Sub
End If
If dbCobranca.Recordset!protestado = True Then
  MsgBox "Esta cobrança já está protestada!"
  Exit Sub
End If

With dbCobranca
  ValorRecebido = -.Recordset!Valor
  CodigoCobranca = .Recordset!CodigoCobranca
  DataDocumento = .Recordset!DataFechamento
  CodigoCliente = .Recordset!CodigoCliente
  StrDescri = Left(dbCobranca.Recordset!Cliente, 25) & " vencido em:" & dbCobranca.Recordset!DataFechamento & " - Nota:" & dbCobranca.Recordset!NrNota
  NrNota = "Nota:" & dbCobranca.Recordset!NrNota & " - " & dbCobranca.Recordset!Cliente
  .Recordset.Edit
  .Recordset!Pago = True
  .Recordset!protestado = True
  .Recordset!dataprotestado = txtVencimento.Value
  .Recordset.Update
  .Refresh
End With


With dbDespesas
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Despesa"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoDespesa = 0
  .Recordset!NrDocumento = CodigoCobranca
  .Recordset!Descri = "Protesto de Boleto"
  .Recordset!Obs = StrDescri
  .Recordset!Valor = ValorRecebido
  .Recordset!valorpago = ValorRecebido
  .Recordset!compensado = True
  .Recordset!fechamentodiario = True
  .Recordset!codigoenviar = "1"
  .Recordset.Update
End With

With dbClientes
  .Refresh
  .Recordset.FindFirst "codigocliente=" & CodigoCliente
  If .Recordset.NoMatch = False Then
    Nome = .Recordset!Nome
    If IsNull(.Recordset!CNPJ) = False Then
      Documento = .Recordset!CNPJ
    End If
  End If
End With
Resposta = vbYes
Do While Resposta = vbYes
  If ProtestoAdicionaHistorico(txtVencimento.Value, DataDocumento, "Cobrança", "Protestando", NrNota, ValorRecebido, Nome, Documento, CodigoCliente) = False Then
    Resposta = MsgBox("Não foi possível criar o histórico de protesto. Deseja Tentar de novo?", vbYesNo)
  Else
    Exit Do
  End If
Loop

Call cboCliente_LostFocus
DBGrid3.SetFocus
End Sub

Private Sub cmdRecebe_Click()
Dim TempValor As Currency, Taxa As Double, Juros As Currency
Dim ValorRecebido As Currency, ValorDesconto As Currency
Dim Resposta As Integer
Dim Valor As Currency, Obs As String


With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

ValorAPagar = CalculaJurosBoleto(dbCobranca.Recordset!DataFechamento, txtVencimento.Value, dbCobranca.Recordset!Valor)
If ValorAPagar > CCur(txtValor.Text) Then
  Obs = ""
  Do While Obs = ""
    Obs = InputBox("Valor recebido está abaixo do valor com juros calculado. Justifique o motivo!", "Juros Calculado")
    If Obs = "" Then
      Resposta = MsgBox("Você deve descrever o motivo! Deseja continuar?", vbYesNo)
      If Resposta = vbNo Then Exit Sub
    End If
  Loop
End If

If DateDiff("d", Date, txtVencimento.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber boleto com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) <= -10 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber boleto com data anterior a 10 dias!"
    Exit Sub
  End If
End If
Resposta = MsgBox("Deseja receber o valor atual?", vbYesNo + vbDefaultButton2, App.Title)
If Resposta = vbNo Then Exit Sub

If dbCobranca.Recordset.RecordCount = 0 Then Exit Sub
If dbCobranca.Recordset.EOF = True Then
  MsgBox "Escolha uma cobrança primeiro!"
  Exit Sub
End If
If cboFormadePg.Text <> dbFormaDePg.Recordset!Descri Then
  MsgBox "Forma de Pagamento inválida!"
  cboFormadePg.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If

ValorRecebido = CCur(txtValor.Text)
Taxa = dbFormaDePg.Recordset!DescontoPorcento / 100
ValorDesconto = (ValorRecebido * Taxa) - dbFormaDePg.Recordset!descontovalor

With dbContas
  .Refresh
  .Recordset.FindFirst "codigoconta=" & dbFormaDePg.Recordset!CodigoConta
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de conta!"
    Exit Sub
  End If
End With

With dbConciliaNova
  .Recordset.AddNew
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  If dbFormaDePg.Recordset!reembolso > 0 Then
    .Recordset!compensado = False
  Else
    .Recordset!Data = txtVencimento.Value
    .Recordset!compensado = True
  End If
  .Recordset!Tipo = "Cobranca"
  .Recordset!Codigo = 999999997
  .Recordset!Descri = Left(dbFormaDePg.Recordset!Descri & " - " & cboCliente.Text, 50)
  .Recordset!NrDocumento = dbCobranca.Recordset!CodigoCobranca
  .Recordset!Valor = ValorRecebido
  .Recordset.Update
End With

With dbContas
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo + ValorRecebido
  .Recordset.Update
End With

With dbCobranca
  .Recordset.Edit
  .Recordset!Pago = True
  .Recordset!valorpago = ValorRecebido
  .Recordset!DataPagamento = txtVencimento.Value
  .Recordset!CodigoFormadePg = dbFormaDePg.Recordset!CodigoPagamento
  .Recordset!Descri = dbFormaDePg.Recordset!Descri
  Juros = ValorRecebido - .Recordset!Valor
  Valor = .Recordset!Valor
  .Recordset!Juros = ValorRecebido - .Recordset!Valor
  .Recordset!JurosDevido = JurosValor
  If IsNull(.Recordset!Obs) = False Then
    .Recordset!Obs = .Recordset!Obs & Obs
  Else
    .Recordset!Obs = Obs
  End If
  .Recordset!fechames = False
  .Recordset.Update
  .Refresh
  If dbClientes.Recordset!protestado = False Then
    If .Recordset.RecordCount = 0 Then
      If dbClientes.Recordset!mensalista = False Then
        AtivaCliente
      Else
        If dbClientes.Recordset!mensalista = False Then
          .Recordset.MoveFirst
          If .Recordset!DataFechamento > DateAdd("d", -2, Date) Then
            AtivaCliente
          End If
        End If
      End If
    Else
      .Recordset.MoveFirst
      If .Recordset!DataFechamento > DateAdd("d", -2, Date) Then
        AtivaCliente
      End If
    End If
  End If
End With
With dbClientes
  .Recordset.Edit
  .Recordset!TotalBoleto = .Recordset!TotalBoleto - Valor
  .Recordset!Saldo = .Recordset!Limite - .Recordset!TotalNotas - .Recordset!TotalBoleto
  .Recordset.Update
End With
Call cboCliente_LostFocus
DBGrid3.SetFocus

End Sub

Private Sub cmdResgatar_Click()
Dim TempValor As Currency, StrDescri As String, CodigoCobranca As Double
Dim Resposta As Integer, DataDocumento As Date, ValorRecebido As Currency, CodigoCliente As Double
Dim Nome As String, Documento As String, NrNota As String

Resposta = MsgBox("Deseja resgatar o protesto atual?", vbYesNo + vbDefaultButton2, App.Title)
If Resposta = vbNo Then Exit Sub

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If dbProtestados.Recordset.RecordCount = 0 Then Exit Sub
If dbProtestados.Recordset.EOF = True Then
  MsgBox "Escolha um protesto primeiro!"
  Exit Sub
End If
If dbProtestados.Recordset!protestado = False Then
  MsgBox "Esta cobrança já está resgatada!"
  Exit Sub
End If

With dbProtestados
  ValorRecebido = .Recordset!Valor
  CodigoCobranca = .Recordset!CodigoCobranca
  DataDocumento = .Recordset!DataFechamento
  CodigoCliente = .Recordset!CodigoCliente
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Escolha um protesto primeiro!"
    Exit Sub
  End If
  StrDescri = Left(.Recordset!Cliente, 25) & " vencido em:" & .Recordset!DataFechamento & " - Nota:" & .Recordset!NrNota
  NrNota = "Nota:" & .Recordset!NrNota & " - " & .Recordset!Cliente
  .Recordset.Edit
  .Recordset!Pago = False
  .Recordset!protestado = False
  .Recordset.Update
  .Refresh
End With


With dbDespesas
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Despesa"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoDespesa = 0
  .Recordset!NrDocumento = CodigoCobranca
  .Recordset!Descri = "Protesto Resgatado de Boleto"
  .Recordset!Obs = StrDescri
  .Recordset!Valor = ValorRecebido
  .Recordset!valorpago = ValorRecebido
  .Recordset!compensado = True
  .Recordset!fechamentodiario = True
  .Recordset.Update
End With

With dbClientes
  .Refresh
  .Recordset.FindFirst "codigocliente=" & CodigoCliente
  If .Recordset.NoMatch = False Then
    Nome = .Recordset!Nome
    If IsNull(.Recordset!CNPJ) = False Then
      Documento = .Recordset!CNPJ
    End If
  End If
End With

Resposta = vbYes
Do While Resposta = vbYes
  If ProtestoAdicionaHistorico(txtResgatado.Value, DataDocumento, "Cobrança", "Resgatando", NrNota, ValorRecebido, Nome, Documento, CodigoCliente) = False Then
    Resposta = MsgBox("Não foi possível criar o histórico de protesto. Deseja Tentar de novo?", vbYesNo)
  Else
    Exit Do
  End If
Loop

Call cmdAtualizaProtestos_Click

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbCobranca_Reposition()
txtValor.Text = ""
With dbCobranca
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  txtValor.Text = Format(.Recordset!Valor, "Currency")
  If IsNull(.Recordset!DataFechamento) = False Then
    txtVencimento.Value = .Recordset!DataFechamento
  End If
End With
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If dbPendencias.RecordSource = "Select *from clientescobranca where pago=0 order by " & DBGrid1.Columns(ColIndex).DataField & ", Cliente" Then
  dbPendencias.RecordSource = "Select *from clientescobranca where pago=0 order by " & DBGrid1.Columns(ColIndex).DataField & " desc, Cliente"
Else
  dbPendencias.RecordSource = "Select *from clientescobranca where pago=0 order by " & DBGrid1.Columns(ColIndex).DataField & ", Cliente"
End If
dbPendencias.Refresh
End Sub

Private Sub dbPendencias_Reposition()
On Error Resume Next
txtProrrogar.Value = dbPendencias.Recordset!DataFechamento
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select

End Sub

Private Sub Form_Load()
CodigoCliente = 0

lblTotal.Caption = Format(0, "Currency")
lblTotal2.Caption = Format(0, "Currency")
  
txtFechamento.Value = Date
txtVencimento.Value = Date
txtResgatado.Value = Date



With dbPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "select *from clientescobranca where pago=0 order by datafechamento"
      QPendencias.RecordSource = "select sum(valor) as total from clientescobranca where pago=0"
    Case ""
      .RecordSource = "select *from clientescobranca where pago=0 order by datafechamento"
      QPendencias.RecordSource = "select sum(valor) as total from clientescobranca where pago=0"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "Select *from clientes  order by nome"
      .RecordSource = "select *from clientescobranca where pago=0 and planodeconta in (" & StrTemp & ") order by datafechamento"
      QPendencias.RecordSource = "select sum(valor) as total from clientescobranca where pago=0 and planodeconta in (" & StrTemp & ")"
  End Select
  
  .Refresh
End With
With QPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
If IsNull(QPendencias.Recordset!Total) = False Then
  lblTotalPendente.Caption = Format(QPendencias.Recordset!Total, "Currency")
Else
  lblTotalPendente.Caption = Format(0, "Currency")
End If

With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "Select *from clientes order by nome"
    Case ""
      .RecordSource = "Select *from clientes order by nome"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "Select *from clientes where planodeconta in (" & StrTemp & ") order by nome"
  End Select
  .Refresh
End With

With dbClientesNotas2
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbFormaDePg
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCobranca
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from clientescobranca where codigocliente=0 and pago=0"
  .Refresh
End With
With QSoma
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valorprevisto) as total from clientesNota2 where codigocliente=0 and confirmado=0"
  .Refresh
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProtestados
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "select *from clientescobranca where protestado=-1 order by datafechamento"
    Case ""
      .RecordSource = "select *from clientescobranca where protestado=-1 order by datafechamento"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "select *from clientescobranca where protestado=-1 and planodeconta in (" & StrTemp & ") order by datafechamento"
  End Select
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qProtestados
  Debug.Print .RecordSource
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "select sum(valor) as total from clientescobranca where protestado=-1"
    Case ""
      .RecordSource = "select sum(valor) as total from clientescobranca where protestado=-1"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "select sum(valor) as total from clientescobranca where protestado=-1 and planodeconta in (" & StrTemp & ")"
  End Select
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos2
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade, valorunitario, codigo, descri from qclientesnota2produtos group by valorunitario, codigo, descri order by codigo"
  .Refresh
End With
With dbJurosBoleto
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientes2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select clientes.*, municipios.* from clientes, municipios where clientes.municipio=municipios.codigo order by clientes.nome"
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from bloqueiafechamento"
  .Refresh
End With
With dbConfigNota
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbNaturezaOp
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCFOP
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbNotasCorpo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

If Usuarios.Nome = "Usuário Master" Then
  cmdProrrogar.Enabled = True
Else
  cmdProrrogar.Enabled = False
End If

Select Case Usuarios.Grupo.ControleNotas
  Case 1 'Somente leitura
    cmdFechar.Enabled = False
    cmdRecebe.Enabled = False
    cmdProtestar.Enabled = False
    cmdExtornar.Visible = False
  Case 2 'Liberado
    cmdFechar.Enabled = True
    cmdRecebe.Enabled = True
    cmdProtestar.Enabled = True
    cmdExtornar.Visible = True
End Select

End Sub


Private Sub txtCodigo_LostFocus()
With dbClientes
  .Refresh
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigocliente=" & txtCodigo.Text
  If .Recordset.NoMatch = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
    If IsNull(.Recordset!Obs) = False Then
      lblObs.Caption = .Recordset!Obs
    Else
      lblObs.Caption = ""
    End If
    lblDias.Caption = .Recordset!diapagamento
    lblPrazo.Caption = .Recordset!Praso
  End If
End With
End Sub

Private Sub txtFechamento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtFechamento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtFechamento_LostFocus()
Me.KeyPreview = True
Call cmdExibe_Click
End Sub

Private Sub txtResgatado_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtResgatado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtResgatado_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
If IsNumeric(txtValor.Text) = False Then Exit Sub
txtValor.Text = Format(txtValor.Text, "Currency")
End Sub

Private Sub txtVencimento_Change()
If dbCobranca.Recordset.EOF = False And dbCobranca.Recordset.BOF = False Then
  ValorAPagar = CalculaJurosBoleto(dbCobranca.Recordset!DataFechamento, txtVencimento.Value, dbCobranca.Recordset!Valor)
  txtValor.Text = Format(ValorAPagar, "Currency")
End If
End Sub

Private Sub txtVencimento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtVencimento_LostFocus()
Me.KeyPreview = True
End Sub
