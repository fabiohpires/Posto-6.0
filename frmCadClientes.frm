VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmCadClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPlanosDeConta 
      Caption         =   "Planos De Conta"
      Height          =   375
      Left            =   7800
      TabIndex        =   127
      Top             =   120
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo cboCliente 
      Bindings        =   "frmCadClientes.frx":0442
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   4575
      Left            =   5760
      TabIndex        =   97
      Top             =   6960
      Visible         =   0   'False
      Width           =   3495
      Begin MSAdodcLib.Adodc dbClientesProdutos 
         Height          =   330
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ClientesProdutos"
         Caption         =   "dbClientesProdutos"
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
      Begin MSAdodcLib.Adodc dbCobranca2 
         Height          =   330
         Left            =   240
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ClientesCobranca"
         Caption         =   "dbCobranca2"
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
      Begin MSAdodcLib.Adodc qCobranca2 
         Height          =   330
         Left            =   240
         Top             =   1680
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ClientesCobranca"
         Caption         =   "qCobranca2"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   240
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from ClientesCarros"
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc dbCobranca 
         Height          =   330
         Left            =   240
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select codigocliente, sum(valor) as total from ClientesCobranca where pago=0 group by codigocliente"
         Caption         =   "dbCobranca"
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
      Begin MSAdodcLib.Adodc dbNotasPendentes 
         Height          =   330
         Left            =   240
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select codigocliente, sum(valorprevisto) as total from ClientesNota2 where confirmado=0 group by codigocliente"
         Caption         =   "dbNotasPendentes"
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
      Begin MSAdodcLib.Adodc dbUltimoAbastecimento 
         Height          =   330
         Left            =   240
         Top             =   2400
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ClientesNota2"
         Caption         =   "dbUltimoAbastecimento"
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
      Begin MSAdodcLib.Adodc dbProdutos 
         Height          =   330
         Left            =   240
         Top             =   2760
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Produtos"
         Caption         =   "dbProdutos"
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
      Begin MSAdodcLib.Adodc dbTurnos 
         Height          =   330
         Left            =   240
         Top             =   3120
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "turnos"
         Caption         =   "dbTurnos"
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
      Begin MSAdodcLib.Adodc dbClientesTipo 
         Height          =   330
         Left            =   240
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from ClientesTipo order by tipocliente"
         Caption         =   "dbClientesTipo"
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
      Begin MSAdodcLib.Adodc dbMunicipios 
         Height          =   330
         Left            =   240
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from Municipios order by nome"
         Caption         =   "dbMunicipios"
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
      Begin MSAdodcLib.Adodc dbClientesPlanos 
         Height          =   330
         Left            =   240
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientesplanodeconta order by descri"
         Caption         =   "dbClientesPlanos"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   94
      Top             =   600
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmCadClientes.frx":0457
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmCadClientes.frx":0473
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdImprimeHistorico"
      Tab(1).Control(1)=   "DataGrid3"
      Tab(1).Control(2)=   "cmdExibir"
      Tab(1).Control(3)=   "txtDataIni"
      Tab(1).Control(4)=   "txtDataFim"
      Tab(1).Control(5)=   "lblTotalJuros"
      Tab(1).Control(6)=   "lblTotal"
      Tab(1).Control(7)=   "lblTotalPago"
      Tab(1).Control(8)=   "Label17"
      Tab(1).Control(9)=   "Label15"
      Tab(1).Control(10)=   "Label13"
      Tab(1).Control(11)=   "Label12"
      Tab(1).Control(12)=   "Label19"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Veículos"
      TabPicture(2)   =   "frmCadClientes.frx":048F
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(29)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(30)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(31)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DataGrid4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtVeiculo"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtPlaca"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cboCombustivel"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdIncluir"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdRemove"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Lista"
      TabPicture(3)   =   "frmCadClientes.frx":04AB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExporta"
      Tab(3).Control(1)=   "DataGrid2"
      Tab(3).Control(2)=   "cmdImprime"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Produtos"
      TabPicture(4)   =   "frmCadClientes.frx":04C7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboTurno"
      Tab(4).Control(1)=   "cboProduto"
      Tab(4).Control(2)=   "DataGrid1"
      Tab(4).Control(3)=   "cmdImprimeProdutos"
      Tab(4).Control(4)=   "cmdProdutoRemover"
      Tab(4).Control(5)=   "cmdProdutoAlterar"
      Tab(4).Control(6)=   "cmdProdutoIncluir"
      Tab(4).Control(7)=   "txtDataValidade"
      Tab(4).Control(8)=   "mskPreco"
      Tab(4).Control(9)=   "txtCodProduto"
      Tab(4).Control(10)=   "mskValorSomar"
      Tab(4).Control(11)=   "mskPorcento"
      Tab(4).Control(12)=   "Label26"
      Tab(4).Control(13)=   "Label25"
      Tab(4).Control(14)=   "Label24"
      Tab(4).Control(15)=   "Label23"
      Tab(4).Control(16)=   "Label22"
      Tab(4).Control(17)=   "Label21"
      Tab(4).Control(18)=   "Label20"
      Tab(4).ControlCount=   19
      Begin VB.CommandButton cmdExporta 
         Caption         =   "Exportar para Pista"
         Height          =   495
         Left            =   -73920
         TabIndex        =   71
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprimeHistorico 
         Height          =   615
         Left            =   -74880
         Picture         =   "frmCadClientes.frx":04E3
         Style           =   1  'Graphical
         TabIndex        =   123
         Tag             =   "Imprimir"
         Top             =   5160
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cboTurno 
         Bindings        =   "frmCadClientes.frx":0F65
         Height          =   315
         Left            =   -73320
         TabIndex        =   84
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboProduto 
         Bindings        =   "frmCadClientes.frx":0F7C
         Height          =   315
         Left            =   -73920
         TabIndex        =   119
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmCadClientes.frx":0F95
         Height          =   4095
         Left            =   -74880
         TabIndex        =   117
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "DataSoma"
            Caption         =   "Fechado Em"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d/M/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DataFechamento"
            Caption         =   "Vencimento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d/M/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "NrNota"
            Caption         =   "Nota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "DataPagamento"
            Caption         =   "Pago Em"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d/M/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Valor"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ValorPago"
            Caption         =   "Valor Pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Juros"
            Caption         =   "Juros"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Descri"
            Caption         =   "Forma de Pg."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Pago"
            Caption         =   "Pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Sim"
               FalseValue      =   "Não"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Protestado"
            Caption         =   "Protestado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Sim"
               FalseValue      =   "Não"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   1
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1275,024
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604,976
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   854,929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmCadClientes.frx":0FAF
         Height          =   4575
         Left            =   -74880
         TabIndex        =   69
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "CodigoCliente"
            Caption         =   "Cod."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nome"
            Caption         =   "Nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Bomba"
            Caption         =   "Bomba"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Mensalista"
            Caption         =   "Ativo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Sim"
               FalseValue      =   "Não"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "FormaDePagamento"
            Caption         =   "Forma de Pg"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "TotalBoleto"
            Caption         =   "Total Boleto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "totalNotas"
            Caption         =   "Total Notas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Saldo"
            Caption         =   "Total Devido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   1
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3569,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1260,284
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCadClientes.frx":0FC4
         Height          =   2655
         Left            =   -74880
         TabIndex        =   115
         Top             =   1680
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "codProduto"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descri"
            Caption         =   "Produto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Preco"
            Caption         =   "Preço"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ValorASomar"
            Caption         =   "Somar"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Porcento"
            Caption         =   "%"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Validade"
            Caption         =   "Validade"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d/M/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Turno"
            Caption         =   "Turno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3030,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   659,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdImprimeProdutos 
         Height          =   615
         Left            =   -74880
         Picture         =   "frmCadClientes.frx":0FE5
         Style           =   1  'Graphical
         TabIndex        =   114
         Tag             =   "Imprimir"
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmdProdutoRemover 
         Caption         =   "Remover"
         Height          =   375
         Left            =   -66240
         TabIndex        =   87
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdProdutoAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   -70800
         TabIndex        =   86
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdProdutoIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -72000
         TabIndex        =   85
         Top             =   1200
         Width           =   975
      End
      Begin MSComCtl2.DTPicker txtDataValidade 
         Height          =   300
         Left            =   -74880
         TabIndex        =   82
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   130875393
         CurrentDate     =   58806
      End
      Begin MSMask.MaskEdBox mskPreco 
         Height          =   300
         Left            =   -69720
         TabIndex        =   76
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.TextBox txtCodProduto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         TabIndex        =   73
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdExibir 
         Caption         =   "Exibir"
         Height          =   375
         Left            =   -70080
         TabIndex        =   103
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker txtDataIni 
         Height          =   300
         Left            =   -73320
         TabIndex        =   100
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   130875393
         CurrentDate     =   39231
      End
      Begin VB.CommandButton cmdImprime 
         Height          =   615
         Left            =   -74880
         Picture         =   "frmCadClientes.frx":1A67
         Style           =   1  'Graphical
         TabIndex        =   70
         Tag             =   "Imprimir"
         Top             =   5160
         Width           =   735
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remover"
         Height          =   375
         Left            =   6480
         TabIndex        =   68
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   5280
         TabIndex        =   67
         Top             =   660
         Width           =   975
      End
      Begin VB.ComboBox cboCombustivel 
         Height          =   315
         Left            =   2520
         TabIndex        =   64
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtPlaca 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   66
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtVeiculo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   62
         Top             =   660
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   5415
         Left            =   -74880
         TabIndex        =   95
         Top             =   360
         Width           =   9615
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "frmCadClientes.frx":24E9
            DataField       =   "PlanoDeConta"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   4920
            TabIndex        =   129
            Top             =   4920
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   "CodigoPlano"
            Text            =   "DataCombo3"
         End
         Begin VB.ComboBox cboNotaFiscal 
            DataField       =   "Nota"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmCadClientes.frx":2508
            Left            =   2520
            List            =   "frmCadClientes.frx":2515
            TabIndex        =   126
            Top             =   4920
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmCadClientes.frx":2538
            DataField       =   "Municipio"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   120
            TabIndex        =   124
            Top             =   1680
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmCadClientes.frx":2564
            DataField       =   "TipoCliente"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   120
            TabIndex        =   122
            Top             =   4920
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "TipoCliente"
            Text            =   ""
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "FormaDePagamento"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmCadClientes.frx":2581
            Left            =   5400
            List            =   "frmCadClientes.frx":2591
            TabIndex        =   112
            Top             =   2880
            Width           =   2295
         End
         Begin VB.TextBox txtFields 
            DataField       =   "CodigoNoPosto"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   15
            Left            =   720
            MaxLength       =   50
            TabIndex        =   7
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame Frame2 
            Caption         =   "Autorizado"
            Height          =   1215
            Left            =   7920
            TabIndex        =   98
            Top             =   2760
            Width           =   1575
            Begin VB.CheckBox Check7 
               Caption         =   "Óleo"
               DataField       =   "PodeOleo"
               DataSource      =   "Adodc1"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   720
               Width           =   1215
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Lavagem"
               DataField       =   "PodeLavagem"
               DataSource      =   "Adodc1"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Combustível"
               DataField       =   "PodeCombustivel"
               DataSource      =   "Adodc1"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Protestado / Desativação Manual"
            DataField       =   "Protestado"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   3600
            TabIndex        =   53
            Top             =   3360
            Width           =   1935
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            DataField       =   "Limite"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   120
            TabIndex        =   51
            Top             =   3480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   " "
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Limitar o cliente atual"
            DataField       =   "Limitar"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   2160
            TabIndex        =   52
            Top             =   3360
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Usa Boleto"
            DataField       =   "Boleto"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   5640
            TabIndex        =   49
            Top             =   3360
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "Bomba"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmCadClientes.frx":25BF
            Left            =   3600
            List            =   "frmCadClientes.frx":25C9
            TabIndex        =   48
            Top             =   2880
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "Tipo"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmCadClientes.frx":25DF
            Left            =   1800
            List            =   "frmCadClientes.frx":25EF
            TabIndex        =   46
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            DataField       =   "Obs"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Top             =   4080
            Width           =   4455
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            DataField       =   "Saldo"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   120
            TabIndex        =   44
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Celular"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   9
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   34
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Fax"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   8
            Left            =   120
            MaxLength       =   50
            TabIndex        =   32
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Telefone"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   7
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "CEP"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   4
            Left            =   8400
            MaxLength       =   50
            TabIndex        =   19
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Praso"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   31
            Left            =   9000
            MaxLength       =   50
            TabIndex        =   42
            Top             =   2280
            Width           =   375
         End
         Begin VB.TextBox Text1 
            DataField       =   "Instrucoes"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Top             =   4080
            Width           =   4695
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Nome2"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   30
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   11
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CNPJ"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   14
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "IE"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   13
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   24
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ativo"
            DataField       =   "Mensalista"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   5400
            TabIndex        =   37
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Estado"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   6
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Bairro"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   3
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Complemento"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   2
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Endereco"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Nome"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   9
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Contato"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   10
            Left            =   6360
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Email"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   11
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   36
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "DiaPagamento"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   12
            Left            =   7920
            MaxLength       =   50
            TabIndex        =   40
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label29 
            Caption         =   "Plano de Conta:"
            Height          =   255
            Left            =   4920
            TabIndex        =   128
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Nota Fiscal:"
            Height          =   255
            Left            =   2520
            TabIndex        =   125
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Tipo de Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Desativado"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   6360
            TabIndex        =   116
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Forma de Pagamento:"
            Height          =   255
            Left            =   5400
            TabIndex        =   113
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "UltimoAbastecimento"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   7920
            TabIndex        =   110
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cod. no Posto:"
            Height          =   195
            Index           =   15
            Left            =   720
            TabIndex        =   6
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label11 
            Caption         =   "Valor Limite:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "Abastece em Bomba:"
            Height          =   255
            Left            =   3600
            TabIndex        =   47
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Data desativação:"
            Height          =   255
            Left            =   6360
            TabIndex        =   38
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Obs:"
            Height          =   255
            Left            =   4920
            TabIndex        =   59
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label lblCodigo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "CodigoCliente"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Cód.:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   195
            Index           =   34
            Left            =   9000
            TabIndex        =   41
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label Label9 
            Caption         =   "Instruções:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   33
            Left            =   4800
            TabIndex        =   10
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   43
            Top             =   2640
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C.N.P.J:"
            Height          =   195
            Index           =   14
            Left            =   4560
            TabIndex        =   25
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "I.E.:"
            Height          =   195
            Index           =   13
            Left            =   3000
            TabIndex        =   23
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Index           =   7
            Left            =   8160
            TabIndex        =   29
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Index           =   6
            Left            =   2520
            TabIndex        =   21
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Index           =   4
            Left            =   8400
            TabIndex        =   18
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   3
            Left            =   5760
            TabIndex        =   16
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Index           =   2
            Left            =   3840
            TabIndex        =   14
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantazia:"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   8
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   2040
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cel:"
            Height          =   195
            Index           =   9
            Left            =   1560
            TabIndex        =   33
            Top             =   2040
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
            Height          =   195
            Index           =   10
            Left            =   6360
            TabIndex        =   27
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Index           =   11
            Left            =   3000
            TabIndex        =   35
            Top             =   2040
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
            Height          =   195
            Index           =   12
            Left            =   7920
            TabIndex        =   39
            Top             =   2040
            Width           =   930
         End
         Begin VB.Label Label16 
            Caption         =   "Último Abastecimento:"
            Height          =   255
            Left            =   7920
            TabIndex        =   111
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSComCtl2.DTPicker txtDataFim 
         Height          =   300
         Left            =   -71640
         TabIndex        =   102
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   130875393
         CurrentDate     =   39231
      End
      Begin MSMask.MaskEdBox mskValorSomar 
         Height          =   300
         Left            =   -68520
         TabIndex        =   78
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskPorcento 
         Height          =   300
         Left            =   -67320
         TabIndex        =   80
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmCadClientes.frx":261F
         Height          =   3735
         Left            =   120
         TabIndex        =   118
         Top             =   1200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Veiculo"
            Caption         =   "Veiculo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Combustivel"
            Caption         =   "Combustivel"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Placa"
            Caption         =   "Placa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   2550,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989,858
            EndProperty
         EndProperty
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   83
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label25 
         Caption         =   "Valido até:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Porcento:"
         Height          =   255
         Left            =   -67320
         TabIndex        =   79
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Valor a Somar:"
         Height          =   255
         Left            =   -68520
         TabIndex        =   77
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Preço:"
         Height          =   255
         Left            =   -69720
         TabIndex        =   75
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   -73920
         TabIndex        =   74
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label20 
         Caption         =   "Código:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   72
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblTotalJuros 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -66960
         TabIndex        =   108
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -71640
         TabIndex        =   104
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblTotalPago 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69360
         TabIndex        =   106
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Pago:"
         Height          =   255
         Left            =   -69840
         TabIndex        =   107
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Total:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   105
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "a"
         Height          =   255
         Left            =   -71880
         TabIndex        =   101
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Período fechado em:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Index           =   31
         Left            =   3840
         TabIndex        =   65
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Veículo:"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   61
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combustível:"
         Height          =   195
         Index           =   29
         Left            =   2520
         TabIndex        =   63
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label19 
         Caption         =   "Juros:"
         Height          =   255
         Left            =   -67440
         TabIndex        =   109
         Top             =   5400
         Width           =   495
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   10155
      TabIndex        =   96
      Top             =   6600
      Width           =   10155
      Begin VB.CommandButton cmdAtualizaSaldo 
         Caption         =   "Atualiza Saldo"
         Height          =   300
         Left            =   8280
         TabIndex        =   120
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1620
         TabIndex        =   89
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2715
         TabIndex        =   90
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3810
         TabIndex        =   91
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4905
         TabIndex        =   92
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   6000
         TabIndex        =   93
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   540
         TabIndex        =   88
         Top             =   0
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6930
      Width           =   10155
      _ExtentX        =   17912
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Clientes"
      Caption         =   "Adodc1"
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
   Begin VB.Label Label4 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoCliente As Double, strOrdem As String, strOrdemProdutos As String
Dim strOrdemHistorico As String, AtualizandoSaldo As Boolean

Private Sub AtualizaSaldo(ByVal Todos As Boolean)
Dim TotalBoleto As Currency, TotalNotas As Currency
Dim Total As Currency
Dim UltimoAbastecimento As Date
  AtualizandoSaldo = True
  With Adodc1
    If .Recordset.RecordCount <> 0 Then
      If .Recordset.BOF = False And .Recordset.EOF = False Then
        If Todos = True Then
          .Recordset.MoveLast
          .Recordset.MoveFirst
        End If
        Do While .Recordset.EOF = False
          On Error GoTo TrataErro
          DoEvents
          Total = 0
          TotalNotas = 0
          TotalBoleto = 0
          If dbNotasPendentes.Recordset.RecordCount <> 0 Then
            dbNotasPendentes.Recordset.MoveFirst
            dbNotasPendentes.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
            If dbNotasPendentes.Recordset.EOF = False Then
              Total = dbNotasPendentes.Recordset!Total
              TotalNotas = dbNotasPendentes.Recordset!Total
            End If
          End If
          If dbCobranca.Recordset.RecordCount <> 0 Then
            dbCobranca.Recordset.MoveFirst
            dbCobranca.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
            If dbCobranca.Recordset.EOF = False Then
              Total = Total + dbCobranca.Recordset!Total
              TotalBoleto = dbCobranca.Recordset!Total
            End If
          End If
          With dbUltimoAbastecimento
            .Refresh
            If .Recordset.RecordCount <> 0 Then
              .Recordset.MoveFirst
              .Recordset.Find "codigocliente=" & Adodc1.Recordset!CodigoCliente
              If .Recordset.EOF = False Then
                UltimoAbastecimento = .Recordset!Data
              Else
                UltimoAbastecimento = Date
              End If
            End If
          End With
          .Recordset!Saldo = Total
          .Recordset!TotalBoleto = TotalBoleto
          .Recordset!TotalNotas = TotalNotas
          If IsNull(.Recordset!UltimoAbastecimento) = True Then
            .Recordset!UltimoAbastecimento = UltimoAbastecimento
          End If
          If UltimoAbastecimento < DateAdd("m", -3, Date) Then
            If .Recordset!mensalista = -1 Then
              .Recordset!mensalista = 0
              .Recordset!desativado = Now
            End If
          End If
          .Recordset.Update
          If Todos = False Then Exit Do
          .Recordset.MoveNext
        Loop
      End If
    End If
  End With
  
TrataErro:
  AtualizandoSaldo = False
End Sub



Private Sub Atualiza()
Dim Codigo As Double
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1

If AtualizandoSaldo = True Then Exit Sub

If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
  Codigo = 0
Else
  If IsNull(Adodc1.Recordset!CodigoCliente) = False Then
    Codigo = Adodc1.Recordset!CodigoCliente
  Else
    Codigo = 0
  End If
End If
On Error Resume Next
With Adodc2
  .Recordset.Filter = "codigocliente=" & Codigo
  .Recordset.Sort = "veiculo, Placa"
End With
With dbClientesProdutos
  .Recordset.Filter = "CodigoCliente = " & Codigo
  .Recordset.Sort = strOrdemProdutos
End With

Call cmdExibir_Click
End Sub


Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

StrTemp = "Clientes de Notas - " & NomePosto
Printer.FontSize = 16
Printer.FontBold = True
Printer.CurrentY = 0
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = "Impresso em: " & UCase(Format(Dia, "long Date")) & " - " & Format(Dia, "short time")
Printer.FontSize = 8
Printer.FontBold = True
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Cod."
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = 10 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Nome"
Printer.CurrentX = 11
Printer.Print StrTemp;

StrTemp = "Tipo"
Printer.CurrentX = 80
Printer.Print StrTemp;

StrTemp = "Prazo"
Printer.CurrentX = 108 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Bomba"
Printer.CurrentX = 110
Printer.Print StrTemp;

StrTemp = "Ativo"
Printer.CurrentX = 125
Printer.Print StrTemp;

StrTemp = "Dt. Bloqueio"
Printer.CurrentX = 135
Printer.Print StrTemp;

StrTemp = "Boleto"
Printer.CurrentX = 160
Printer.Print StrTemp;

StrTemp = "Total"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5



End Sub

Private Sub CabecaProdutos(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

StrTemp = "Clientes de Notas / Produtos - " & NomePosto
Printer.FontSize = 16
Printer.FontBold = True
Printer.CurrentY = 0
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = "Impresso em: " & UCase(Format(Dia, "long Date")) & " - " & Format(Dia, "short time")
Printer.FontSize = 8
Printer.FontBold = True
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Cod."
Printer.FontSize = 10
Printer.FontBold = False
Printer.CurrentX = 10 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Nome"
Printer.CurrentX = 11
Printer.Print StrTemp;

StrTemp = "Tipo"
Printer.CurrentX = 80
Printer.Print StrTemp;

StrTemp = "Prazo"
Printer.CurrentX = 108 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Bomba"
Printer.CurrentX = 110
Printer.Print StrTemp;

StrTemp = "Ativo"
Printer.CurrentX = 125
Printer.Print StrTemp;

StrTemp = "Dt. Bloqueio"
Printer.CurrentX = 135
Printer.Print StrTemp;

StrTemp = "Boleto"
Printer.CurrentX = 160
Printer.Print StrTemp;

StrTemp = "Limite"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp


Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5



End Sub




Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Atualiza
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Save = True Then
  If QuerGravar = False Then
    On Error Resume Next
    Adodc1.Recordset.CancelUpdate
  End If
End If
End Sub

Private Sub cboCliente_LostFocus()
With Adodc1
  If cboCliente.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & cboCliente.Text & "'"
  If .Recordset.EOF = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    txtCodProduto.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If cboTurno.Text <> "" Then
      .Recordset.MoveFirst
      .Recordset.Find "Descri='" & cboTurno.Text & "'"
    End If
  End If
End With
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  Adodc1.Recordset!UltimoAbastecimento = Date
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(15).SetFocus
End Sub

Private Sub cmdAtualizaSaldo_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja atualizar só o atual?", vbYesNo)
If Resposta = vbYes Then
  AtualizaSaldo False
Else
  AtualizaSaldo True
End If
End Sub

Private Sub cmdDelete_Click()
  Dim Resposta As Integer
  
  
  With Adodc1
    If .Recordset.EOF = True Then Exit Sub
    If dbCobranca.Recordset.RecordCount <> 0 Then
      dbCobranca.Recordset.MoveFirst
      dbCobranca.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
      If dbCobranca.Recordset.EOF = False Then
        MsgBox "ESTE CLIENTE NÃO PODE SER EXCLUIDO POIS EXISTE COBRANÇA PENDENTE PARA ELE!", vbCritical
        Exit Sub
      End If
    End If
    If dbNotasPendentes.Recordset.RecordCount <> 0 Then
      dbNotasPendentes.Recordset.MoveFirst
      dbNotasPendentes.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
      If dbNotasPendentes.Recordset.EOF = False Then
        MsgBox "ESTE CLIENTE NÃO PODE SER EXCLUIDO POIS EXISTE NOTA PARA SER COBRADA!", vbCritical
        Exit Sub
      End If
    End If
  End With
  
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo, "Excluir!")
  If Resposta = vbNo Then
    Exit Sub
  End If
  
  With Adodc1.Recordset
    If .EOF = False Then
      .Delete
      If .EOF = False Then
      .MoveNext
      Else
        If .BOF = False Then .MoveLast
      End If
    End If
  End With
  
  Frame1.Enabled = False
End Sub

Private Sub cmdEditar_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Frame1.Enabled = True
txtFields(15).SetFocus
End Sub

Private Sub cmdExibir_Click()
If Adodc1.Recordset.EOF = True Then Exit Sub

Dim Total As Currency, Pago As Currency, Juros As Currency
Dim Codigo As Double

If AtualizandoSaldo = True Then Exit Sub

Codigo = 0
If Adodc1.Recordset.EOF = False Then
  If IsNull(Adodc1.Recordset!CodigoCliente) = False Then
    Codigo = Adodc1.Recordset!CodigoCliente
  Else
    Codigo = 0
  End If
End If
With dbCobranca2
  .RecordSource = "select *from clientescobranca where codigocliente=" & Codigo & " and datasoma between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  .Refresh
  .Recordset.Sort = strOrdemHistorico
End With
With qCobranca2
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(valor) as Total, sum(ValorPago) as recebido, sum(juros) as juro from clientescobranca where codigocliente=" & Codigo & " and datasoma between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    Total = .Recordset!Total
  Else
    Total = 0
  End If
  If IsNull(.Recordset!recebido) = False Then
    Pago = .Recordset!recebido
  Else
    Pago = 0
  End If
  If IsNull(.Recordset!juro) = False Then
    Juros = .Recordset!juro
  Else
    Juros = 0
  End If
  lblTotal.Caption = Format(Total, "Currency")
  lblTotalJuros.Caption = Format(Juros, "Currency")
  lblTotalPago.Caption = Format(Pago, "Currency")
End With

End Sub

Private Sub cmdExporta_Click()
ExportaClienteMicrosffer
End Sub

Private Sub cmdImprime_Click()
Dim Largura As Double, StrTemp As String
Dim Total As Currency, Dia As Date
Dim Ws As Workspace, db As Database, dbCobraPendente As Recordset
Dim SaldoLimite As Currency, YUltimoAbastece As Double, A As Double

cmdImprime.Enabled = False

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then
  cmdImprime.Enabled = True
  Exit Sub
End If
On Error GoTo 0

Dia = Now
Largura = 190

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
Set dbCobraPendente = db.OpenRecordset("select *from clientescobranca where pago=0 order by codigocliente, datafechamento")


With Adodc1
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  A = .Recordset.AbsolutePosition
  Cabeca Largura, Dia
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 10
  Printer.FontBold = False
  
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If Printer.CurrentY + 40 >= Printer.ScaleHeight Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      StrTemp = Format(Total, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.NewPage
      Cabeca Largura, Dia
      
    End If
    'If .Recordset!protestado = False Then
      If .Recordset!mensalista = False Then
        Printer.FontBold = True
      Else
        Printer.FontBold = False
      End If
      StrTemp = .Recordset!CodigoCliente
      Printer.FontSize = 8
      Printer.CurrentX = 10 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      On Error Resume Next
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 11
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Tipo) = False Then
        StrTemp = .Recordset!Tipo
        Printer.CurrentX = 80
        Printer.Print StrTemp;
      End If
      
      If IsNull(.Recordset!Praso) = False Then
        StrTemp = .Recordset!Praso
        Printer.CurrentX = 108 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
      End If
      
      If IsNull(.Recordset!Bomba) = False Then
        StrTemp = .Recordset!Bomba
        Printer.CurrentX = 110
        Printer.Print StrTemp;
      End If
      
      If .Recordset!mensalista = True And .Recordset!protestado = False Then
        StrTemp = "Sim"
      Else
        StrTemp = "Não"
      End If
      Printer.CurrentX = 125
      Printer.Print StrTemp;
      
      If .Recordset!mensalista = False Or .Recordset!protestado = True Then
        StrTemp = Format(.Recordset!desativado, "short date")
        Printer.CurrentX = 135
        Printer.Print StrTemp;
      End If
      
      If .Recordset!Boleto = True Then
        StrTemp = "Sim"
      Else
        StrTemp = "Não"
      End If
      Printer.CurrentX = 160
      Printer.Print StrTemp
      
      StrTemp = "Autorizado consumo de:"
      If .Recordset!podecombustivel = True Then
        StrTemp = StrTemp & " Combustível"
      End If
      If .Recordset!podelavagem = True Then
        StrTemp = StrTemp & " Lavagem"
      End If
      If .Recordset!podeoleo = True Then
        StrTemp = StrTemp & " Óleo"
      End If
      Printer.CurrentX = 0
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!UltimoAbastecimento) = False Then
        StrTemp = "Ultimo Abastecimento=" & Format(.Recordset!UltimoAbastecimento, "Short Date")
        YUltimoAbastece = Printer.CurrentY
        Printer.Print ""
        Printer.CurrentX = 0
        Printer.Print StrTemp
        Printer.CurrentY = YUltimoAbastece
      End If
      
      If dbCobraPendente.EOF = False Then
        dbCobraPendente.FindFirst "codigocliente=" & .Recordset!CodigoCliente
        If dbCobraPendente.NoMatch = False Then
          Do While dbCobraPendente!CodigoCliente = .Recordset!CodigoCliente
            StrTemp = "Vencimento: " & Format(dbCobraPendente!DataFechamento, "short date")
            Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
            Printer.Print StrTemp;
            
            StrTemp = "Valor: " & Format(dbCobraPendente!Valor, "currency")
            Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
            Printer.Print StrTemp
            
            If Printer.CurrentY + 40 >= Printer.ScaleHeight Then
              Printer.CurrentY = Printer.CurrentY + 0.5
              Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
              Printer.CurrentY = Printer.CurrentY + 0.5
              
              Printer.NewPage
              Cabeca Largura, Dia
            End If
            
            dbCobraPendente.MoveNext
            If dbCobraPendente.EOF = True Then Exit Do
          Loop
          dbCobraPendente.MoveFirst
        End If
      End If
      
      If IsNull(.Recordset!TotalNotas) = False Then
        StrTemp = "Notas a faturar= " & Format(.Recordset!TotalNotas, "currency")
      Else
        StrTemp = "Notas a faturar= " & Format(0, "currency")
      End If
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (155, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
      
      If IsNull(.Recordset!Saldo) = False Then
        Total = Total + .Recordset!Saldo
        SaldoLimite = .Recordset!Saldo
        StrTemp = "Total Pendente= " & Format(.Recordset!Saldo, "Currency")
      Else
        StrTemp = "Total Pendente= " & Format(0, "Currency")
      End If
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      If .Recordset!limitar = True Then
        StrTemp = "Limite= " & Format(.Recordset!Limite, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        If IsNull(.Recordset!Limite) = False Then
          SaldoLimite = .Recordset!Limite - SaldoLimite
        Else
          SaldoLimite = 0
        End If
        StrTemp = "Limite - Cobranças - Abastecido = Saldo: " & Format(SaldoLimite, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
      Else
        Printer.Print ""
      End If
      
      If Printer.FontBold = True Then
        For i = 0 To 2
          Printer.CurrentY = Printer.CurrentY + 5
          Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Next i
      End If
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
    'End If
     
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  
  StrTemp = Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  On Error Resume Next
  .Recordset.AbsolutePosition = A
End With

Printer.EndDoc

NaoImprime:

cmdImprime.Enabled = True

End Sub

Private Sub cmdImprimeHistorico_Click()
Dim Titulo As String, Titulo2 As String, Titulo3 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then
  Exit Sub
End If
On Error GoTo 0



Printer.ScaleMode = vbMillimeters


Titulo = NomePosto & " - Histórico de Cliente de Nota"
Titulo2 = lblCodigo.Caption & " - " & txtFields(0).Text
Titulo3 = "Período: " & txtDataIni.Value & " a " & txtDataFim.Value & Chr(vbKeyReturn) & "Impresso em: " & Format(Now, "Long Date")

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(Titulo) / 2)
Printer.Print Titulo

Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(Titulo2) / 2)
Printer.Print Titulo2

Printer.Print Titulo3

StrTemp = "Fechamento: " & Combo1.Text
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Prazo: " & txtFields(31).Text
Printer.CurrentX = 45
Printer.Print StrTemp;

StrTemp = "Limite: " & Format(MaskEdBox2.Text, "Currency")
Printer.CurrentX = 85
Printer.Print StrTemp

Printer.Print ""

ImprimeADOGrid DataGrid1, Printer, dbClientesProdutos

Printer.ScaleMode = vbMillimeters


ImprimeADOGrid DataGrid3, Printer, dbCobranca2


Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdImprimeProdutos_Click()
Dim Dia As Date, Largura As Double, A As Double
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

With Adodc1
  A = .Recordset.AbsolutePosition
  If .Recordset.RecordCount = 0 Then Exit Sub
  Dia = Date
  Largura = 190
  CabecaProdutos Largura, Dia
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 10
  Printer.FontBold = False
  
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If Printer.CurrentY + 40 >= Printer.ScaleHeight Then
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
            
      Printer.NewPage
      CabecaProdutos Largura, Dia
      
    End If
    If .Recordset!protestado = False Then
      If .Recordset!mensalista = False Then
        Printer.FontBold = True
      Else
        Printer.FontBold = False
      End If
      StrTemp = .Recordset!CodigoCliente
      Printer.FontSize = 8
      Printer.CurrentX = 10 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 11
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Tipo) = False Then
        StrTemp = .Recordset!Tipo
        Printer.CurrentX = 80
        Printer.Print StrTemp;
      End If
      
      If IsNull(.Recordset!Praso) = False Then
        StrTemp = .Recordset!Praso
        Printer.CurrentX = 108 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
      End If
      
      If IsNull(.Recordset!Bomba) = False Then
        StrTemp = .Recordset!Bomba
        Printer.CurrentX = 110
        Printer.Print StrTemp;
      End If
      
      If .Recordset!mensalista = True And .Recordset!protestado = False Then
        StrTemp = "Sim"
      Else
        StrTemp = "Não"
      End If
      Printer.CurrentX = 125
      Printer.Print StrTemp;
      
      If .Recordset!mensalista = False Or .Recordset!protestado = True Then
        StrTemp = Format(.Recordset!desativado, "short date")
        Printer.CurrentX = 135
        Printer.Print StrTemp;
      End If
      
      If .Recordset!Boleto = True Then
        StrTemp = "Sim"
      Else
        StrTemp = "Não"
      End If
      Printer.CurrentX = 160
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Limite) = False Then
        StrTemp = Format(.Recordset!Limite, "Currency")
      Else
        StrTemp = ""
      End If
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      ImprimeADOGrid DataGrid1, Printer, dbClientesProdutos
      
      Printer.ScaleMode = vbMillimeters
      
      Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 0.5
    End If
     
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 0.5
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 0.5
  On Error Resume Next
  .Recordset.AbsolutePosition = A
End With

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdIncluir_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
If Adodc1.Recordset.EOF = True Then Exit Sub
If txtVeiculo.Text = "" Then
  MsgBox "Informe um veículo!"
  txtVeiculo.SetFocus
  Exit Sub
End If
If cboCombustivel.Text = "" Then
  MsgBox "Informe o combustível!"
  cboCombustivel.SetFocus
  Exit Sub
End If
If txtPlaca.Text = "" Then
  MsgBox "Informe uma placa!"
  txtPlaca.SetFocus
  Exit Sub
End If

With Adodc2
  .Recordset.AddNew
  .Recordset!CodigoCliente = Adodc1.Recordset!CodigoCliente
  .Recordset!Veiculo = txtVeiculo.Text
  .Recordset!Combustivel = cboCombustivel.Text
  .Recordset!Placa = txtPlaca.Text
  .Recordset.Update
  .Refresh
  On Error Resume Next
End With
Atualiza
txtVeiculo.Text = ""
cboCombustivel.Text = ""
txtPlaca.Text = ""
txtVeiculo.SetFocus
End Sub

Private Sub cmdPlanosDeConta_Click()
frmCadClientesPlano.Show vbModal
End Sub

Private Sub cmdProdutoAlterar_Click()
If Adodc1.Recordset.EOF = True Then
  MsgBox "Selecione um cliente primeiro!"
  Exit Sub
End If
If IsNumeric(txtCodProduto.Text) = False Then
  MsgBox "Selecione um produto primeiro!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset.EOF = True Then
  MsgBox "Selecione um turno primeiro!"
  cboTurno.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset!Descri <> cboTurno.Text Then
  MsgBox "Turno inválido!"
  cboTurno.SetFocus
  Exit Sub
End If
If IsNumeric(mskPreco.Text) = False Then mskPreco.Text = 0
If IsNumeric(mskValorSomar.Text) = False Then mskValorSomar.Text = 0
If IsNumeric(mskPorcento.Text) = False Then mskPorcento.Text = 0
With dbClientesProdutos
  If .Recordset.EOF = True Then
    MsgBox "Selecione um produto diferenciado primeiro!"
    Exit Sub
  End If
  .Recordset!CodigoCliente = Adodc1.Recordset!CodigoCliente
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!CodProduto = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!Preco = CCur(mskPreco.Text)
  .Recordset!valorasomar = CCur(mskValorSomar.Text)
  .Recordset!Porcento = CDbl(mskPorcento.Text)
  .Recordset!validade = txtDataValidade.Value
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset!Turno = dbTurnos.Recordset!Descri
  .Recordset!Grupo = 0
  .Recordset.Update
End With
txtCodProduto.Text = ""
cboProduto.Text = ""
mskPreco.Text = ""
mskValorSomar.Text = ""
mskPorcento.Text = ""
cboTurno.Text = ""
txtCodProduto.SetFocus
Atualiza
End Sub

Private Sub cmdProdutoIncluir_Click()
If Adodc1.Recordset.EOF = True Then
  MsgBox "Selecione um cliente primeiro!"
  Exit Sub
End If
If IsNumeric(txtCodProduto.Text) = False Then
  MsgBox "Selecione um produto primeiro!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset.EOF = True Then
  MsgBox "Selecione um turno primeiro!"
  cboTurno.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset!Descri <> cboTurno.Text Then
  MsgBox "Turno inválido!"
  cboTurno.SetFocus
  Exit Sub
End If
If IsNumeric(mskPreco.Text) = False Then mskPreco.Text = 0
If IsNumeric(mskValorSomar.Text) = False Then mskValorSomar.Text = 0
If IsNumeric(mskPorcento.Text) = False Then mskPorcento.Text = 0
With dbClientesProdutos
  .Recordset.AddNew
  .Recordset!CodigoCliente = Adodc1.Recordset!CodigoCliente
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!CodProduto = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!Preco = CCur(mskPreco.Text)
  .Recordset!valorasomar = CCur(mskValorSomar.Text)
  .Recordset!Porcento = CDbl(mskPorcento.Text)
  .Recordset!validade = txtDataValidade.Value
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset!Turno = dbTurnos.Recordset!Descri
  .Recordset!HoraIni = dbTurnos.Recordset!HoraIni
  .Recordset!Grupo = 0
  .Recordset.Update
End With
txtCodProduto.Text = ""
cboProduto.Text = ""
mskPreco.Text = ""
mskValorSomar.Text = ""
mskPorcento.Text = ""
cboTurno.Text = ""
txtCodProduto.SetFocus
Atualiza
End Sub

Private Sub cmdProdutoRemover_Click()
Dim Resposta As Integer, A As Double

Resposta = MsgBox("Deseja remover o produto diferenciado atual?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With dbClientesProdutos
  If .Recordset.EOF = True Then
    MsgBox "Selecione um produto diferenciado primeiro!"
    Exit Sub
  End If
  .Recordset.Delete adAffectCurrent
  On Error Resume Next
  Atualiza
End With
End Sub

Private Sub cmdRefresh_Click()

  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdRemove_Click()
Dim Resposta As Integer

With Adodc2
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Escolha um registro primeiro!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
  .Refresh
End With
Atualiza
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    A = .Recordset.AbsolutePosition
    .Recordset.Update
    .Recordset.AbsolutePosition = A
  End With
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Frame1.Enabled = False

End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DataGrid2.Columns(ColIndex).DataField & ", Nome" Then
  strOrdem = DataGrid2.Columns(ColIndex).DataField & " desc, Nome"
Else
  strOrdem = DataGrid2.Columns(ColIndex).DataField & ", Nome"
End If

With Adodc1
  .Recordset.Sort = strOrdem
End With
End Sub

Private Sub DataGrid3_HeadClick(ByVal ColIndex As Integer)
If strOrdemHistorico = DataGrid3.Columns(ColIndex).DataField Then
  strOrdemHistorico = DataGrid3.Columns(ColIndex).DataField & " desc"
Else
  strOrdemHistorico = DataGrid3.Columns(ColIndex).DataField
End If
With dbCobranca2
  .Recordset.Sort = strOrdemHistorico
End With
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
txtDataIni.Value = DateAdd("m", -3, Date)
txtDataFim.Value = DateAdd("m", 1, Date)
txtDataValidade.Value = Date

strOrdemProdutos = "codproduto, validade, horaini"
strOrdem = "Nome, Nome"
strOrdemHistorico = "DataSoma"
With dbClientesPlanos
  .ConnectionString = CaminhoADO
  .Refresh
End With

With dbClientesProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from clientesprodutos"
  .Refresh
  .Recordset.Filter = "codigocliente=0"
  .Recordset.Sort = strOrdemProdutos
End With

With dbNotasPendentes
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCobranca
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCobranca2
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from clientescobranca"
  .Refresh
End With
With qCobranca2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With Adodc2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbUltimoAbastecimento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from clientesnota2 order by data desc"
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from produtos order by descri"
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from turnos order by descri"
  .Refresh
End With
With cboCombustivel
  .Clear
  .AddItem "Álcool"
  .AddItem "Diesel"
  .AddItem "Gasolina"
End With
With dbClientesTipo
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbMunicipios
  .ConnectionString = CaminhoADO
  .Refresh
End With


With Adodc1
  .ConnectionString = CaminhoADO
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
  .Recordset.Sort = strOrdem
End With


Call cmdRefresh_Click
Select Case Usuarios.Grupo.CadCliente
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdIncluir.Enabled = False
    cmdRemove.Enabled = False
  Case 2 'Liberado
    
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub mskPorcento_LostFocus()
With mskPorcento
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "0.000")
  End If
End With
End Sub

Private Sub mskPreco_LostFocus()
With mskPreco
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "0.000")
  End If
End With
End Sub

Private Sub mskValorSomar_LostFocus()
With mskValorSomar
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "0.000")
  End If
End With
End Sub

Private Sub Text1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub Text1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtCodigo_LostFocus()
With Adodc1
  If txtCodigo.Text = "" Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "codigocliente=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    If IsNull(.Recordset!Nome) = False Then
      cboCliente.Text = .Recordset!Nome
    End If
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  .Recordset.Find "codigo=" & txtCodProduto.Text
  If .Recordset.EOF = False Then
    txtCodProduto.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub txtDataValidade_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataValidade_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataValidade_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
  Case 14
    If Fu_consistir_CgcCpf(txtFields(Index).Text) = False Then
      MsgBox "CNPJ ou CPF inválido!"
    End If
End Select
End Sub
