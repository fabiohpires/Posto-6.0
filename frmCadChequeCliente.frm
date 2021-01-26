VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadChequeCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes de Cheques"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   Icon            =   "frmCadChequeCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboRelatorio 
      Height          =   315
      Left            =   120
      TabIndex        =   75
      Top             =   5880
      Width           =   4575
   End
   Begin VB.CommandButton cmdLocalizaCheque 
      Caption         =   "Localizar Cheque"
      Height          =   255
      Left            =   2280
      TabIndex        =   106
      Top             =   0
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc qPendentes2 
      Height          =   375
      Left            =   1080
      Top             =   4200
      Visible         =   0   'False
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
      RecordSource    =   "select *from cheques"
      Caption         =   "qPendentes2"
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
      Height          =   375
      Left            =   1080
      Top             =   3840
      Visible         =   0   'False
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
      RecordSource    =   "select *from ChequesContas"
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
   Begin MSAdodcLib.Adodc dbCarros 
      Height          =   375
      Left            =   1080
      Top             =   3480
      Visible         =   0   'False
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
      RecordSource    =   "select *from chequescarros"
      Caption         =   "dbCarros"
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
      Height          =   375
      Left            =   1080
      Top             =   3120
      Visible         =   0   'False
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
      RecordSource    =   "select *from ChequesClientesCobraHistorico"
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
   Begin MSComCtl2.DTPicker txtDataLista 
      Height          =   300
      Left            =   1920
      TabIndex        =   77
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72613889
      CurrentDate     =   38949
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      OutBufferSize   =   1024
      RTSEnable       =   -1  'True
      BaudRate        =   115200
   End
   Begin VB.TextBox txtProcuraCNPJ 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtProcuraCIC 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtProcuraNome 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   3840
      Picture         =   "frmCadChequeCliente.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   78
      Tag             =   "Imprimir"
      Top             =   6240
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   4800
      TabIndex        =   47
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "frmCadChequeCliente.frx":0EC4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdContar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contas"
      TabPicture(1)   =   "frmCadChequeCliente.frx":0EE0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid5"
      Tab(1).Control(1)=   "cmdIncluirConta"
      Tab(1).Control(2)=   "txtBanco"
      Tab(1).Control(3)=   "txtConta"
      Tab(1).Control(4)=   "txtAg"
      Tab(1).Control(5)=   "txtBancoNr"
      Tab(1).Control(6)=   "txtComp"
      Tab(1).Control(7)=   "Label7"
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "Label3"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Carros"
      TabPicture(2)   =   "frmCadChequeCliente.frx":0EFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid4"
      Tab(2).Control(1)=   "cmdIncluir"
      Tab(2).Control(2)=   "txtPlaca"
      Tab(2).Control(3)=   "txtCarro"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Cheques"
      TabPicture(3)   =   "frmCadChequeCliente.frx":0F18
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExibeCheques"
      Tab(3).Control(1)=   "cmdImprimeCheques"
      Tab(3).Control(2)=   "dbCheques2"
      Tab(3).Control(3)=   "DataGrid1"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Cobrança"
      TabPicture(4)   =   "frmCadChequeCliente.frx":0F34
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdIncluirCobranca"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtFields(20)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtFields(19)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtObs"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtNomeContato"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "lblLabels(13)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "lblLabels(12)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label20"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label19"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      Begin VB.CommandButton Command1 
         Caption         =   "Todos que não abastecem mais"
         Height          =   615
         Left            =   4800
         TabIndex        =   109
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdContar 
         Caption         =   "Contar Cheques"
         Height          =   495
         Left            =   4800
         TabIndex        =   108
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdExibeCheques 
         Caption         =   "Exibe Cheques"
         Height          =   375
         Left            =   -74640
         TabIndex        =   107
         Top             =   6120
         Width           =   2295
      End
      Begin VB.CommandButton cmdImprimeCheques 
         Caption         =   "Imprime Lista de cheques"
         Height          =   375
         Left            =   -71280
         TabIndex        =   105
         Top             =   6120
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmCadChequeCliente.frx":0F50
         Height          =   5055
         Left            =   -74880
         TabIndex        =   104
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8916
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Comp"
            Caption         =   "Comp"
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
            DataField       =   "Banco"
            Caption         =   "Banco"
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
            DataField       =   "BancoNumero"
            Caption         =   "Bco. Nr."
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
            DataField       =   "Agencia"
            Caption         =   "Agência"
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
            DataField       =   "Conta"
            Caption         =   "Conta"
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
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065,26
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmCadChequeCliente.frx":0F67
         Height          =   5055
         Left            =   -74760
         TabIndex        =   103
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8916
         _Version        =   393216
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Carro"
            Caption         =   "Carro"
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
            BeginProperty Column00 
               ColumnWidth     =   2745,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1409,953
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmCadChequeCliente.frx":0F7E
         Height          =   4695
         Left            =   -74760
         TabIndex        =   102
         Top             =   1680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8281
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "LancadoEm"
            Caption         =   "LancadoEm"
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
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
            DataField       =   "Contato"
            Caption         =   "Contato"
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
            DataField       =   "Obs"
            Caption         =   "Obs"
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
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1830,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc dbCheques2 
         Height          =   330
         Left            =   -72960
         Top             =   4080
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   $"frmCadChequeCliente.frx":0F97
         Caption         =   "dbCheques2"
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
      Begin VB.CommandButton cmdIncluirCobranca 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -69840
         TabIndex        =   97
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   -74760
         MaxLength       =   15
         TabIndex        =   94
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone2"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -73080
         MaxLength       =   15
         TabIndex        =   93
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   -74760
         TabIndex        =   92
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txtNomeContato 
         Height          =   285
         Left            =   -71400
         TabIndex        =   90
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdIncluirConta 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -69840
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtBanco 
         Height          =   300
         Left            =   -74280
         TabIndex        =   72
         Top             =   840
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtConta 
         Height          =   300
         Left            =   -70800
         TabIndex        =   71
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "999999-9"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtAg 
         Height          =   300
         Left            =   -71520
         TabIndex        =   69
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtBancoNr 
         Height          =   300
         Left            =   -72240
         TabIndex        =   67
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtComp 
         Height          =   300
         Left            =   -74880
         TabIndex        =   65
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -71040
         TabIndex        =   63
         Top             =   720
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtPlaca 
         Height          =   300
         Left            =   -72120
         TabIndex        =   62
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "AAA-9999"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtCarro 
         Height          =   285
         Left            =   -74760
         TabIndex        =   60
         Top             =   840
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   6135
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   6135
         Begin VB.CheckBox chkFields 
            Caption         =   "Atualizar"
            DataField       =   "Atualizar"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   5640
            Width           =   1095
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            DataField       =   "LimiteValor"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   2520
            TabIndex        =   34
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Limite"
            DataSource      =   "Adodc1"
            Height          =   300
            Index           =   18
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   32
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            DataField       =   "Obs"
            DataSource      =   "Adodc1"
            Height          =   615
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   4080
            Width           =   5895
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Telefone2"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   17
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   19
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Codigo"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Nome"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   50
            TabIndex        =   9
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Endereco"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "CEP"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   3
            Left            =   3600
            MaxLength       =   9
            TabIndex        =   14
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Telefone"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   4
            Left            =   120
            MaxLength       =   15
            TabIndex        =   17
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "CIC"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   5
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "RG"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   6
            Left            =   120
            MaxLength       =   20
            TabIndex        =   23
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Origem"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   7
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   25
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Origem2"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   8
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   26
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "CNPJ"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   9
            Left            =   2640
            MaxLength       =   20
            TabIndex        =   28
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "IE"
            DataSource      =   "Adodc1"
            Height          =   300
            Index           =   10
            Left            =   120
            MaxLength       =   20
            TabIndex        =   30
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Consultado"
            DataField       =   "Consultado"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   4920
            Width           =   1095
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Ativo"
            DataField       =   "Posicao"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   5280
            Width           =   855
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   1320
            TabIndex        =   49
            Top             =   4680
            Width           =   3255
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "ValorDevolvido"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   16
               Left            =   1920
               TabIndex        =   55
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "Devolvidos"
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   15
               Left            =   1080
               TabIndex        =   54
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "ValorDepositado"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """ ""#.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   14
               Left            =   1920
               TabIndex        =   53
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "Depositados"
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   13
               Left            =   1080
               TabIndex        =   52
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
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
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   12
               Left            =   1920
               TabIndex        =   51
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "NumeroDeCheques"
               DataSource      =   "Adodc1"
               Height          =   285
               Index           =   11
               Left            =   1080
               TabIndex        =   50
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Recebidos:"
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   810
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Devolvidos:"
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   57
               Top             =   960
               Width           =   840
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Depositados:"
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   56
               Top             =   600
               Width           =   930
            End
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            DataField       =   "LimiteValor2"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   120
            TabIndex        =   79
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   " "
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Cadastrado Por:"
            Height          =   195
            Left            =   4200
            TabIndex        =   99
            Top             =   2640
            Width           =   1665
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "CadastradoPor"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   4200
            TabIndex        =   98
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label lblDataCadastro 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "DataCadastro"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3600
            TabIndex        =   88
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Data Cadastro:"
            Height          =   195
            Left            =   3600
            TabIndex        =   87
            Top             =   3240
            Width           =   1065
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   2520
            TabIndex        =   84
            Top             =   3240
            Width           =   450
         End
         Begin VB.Label lblSaldo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2520
            TabIndex        =   83
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Em Aberto:"
            Height          =   195
            Left            =   1440
            TabIndex        =   82
            Top             =   3240
            Width           =   780
         End
         Begin VB.Label lblPre 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   81
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Limite R$ total:"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Limite R$ por Cheque:"
            Height          =   255
            Left            =   2520
            TabIndex        =   33
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Limite:"
            Height          =   255
            Left            =   1920
            TabIndex        =   31
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Observações:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Celular:"
            Height          =   255
            Index           =   11
            Left            =   1800
            TabIndex        =   18
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Bairro:"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Nome:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Endereco:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "CEP:"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   12
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Telefone:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "CIC:"
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   20
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "RG:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Emissão:"
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   24
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "CNPJ:"
            Height          =   255
            Index           =   9
            Left            =   2640
            TabIndex        =   27
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "IE:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   29
            Top             =   2640
            Width           =   375
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCadChequeCliente.frx":1036
         Height          =   5535
         Left            =   -74880
         TabIndex        =   100
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9763
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
         ColumnCount     =   23
         BeginProperty Column00 
            DataField       =   "Compensado"
            Caption         =   "Compensado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Comp"
            Caption         =   "Comp"
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
            DataField       =   "Banco"
            Caption         =   "Banco"
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
            DataField       =   "Agencia"
            Caption         =   "Agencia"
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
            DataField       =   "Conta"
            Caption         =   "Conta"
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
         BeginProperty Column05 
            DataField       =   "ChequeNr"
            Caption         =   "ChequeNr"
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
            DataField       =   "DataCheque"
            Caption         =   "DataCheque"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "ContaDescri"
            Caption         =   "ContaDescri"
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
         BeginProperty Column09 
            DataField       =   "Devolvido"
            Caption         =   "Devolvido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "DataDevolucao"
            Caption         =   "DataDevolucao"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "DescriDevolucao"
            Caption         =   "DescriDevolucao"
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
         BeginProperty Column12 
            DataField       =   "Cobrando"
            Caption         =   "Cobrando"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "DataCobrando"
            Caption         =   "DataCobrando"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "DataPgto"
            Caption         =   "DataPgto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "ValorPgto"
            Caption         =   "ValorPgto"
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
         BeginProperty Column16 
            DataField       =   "Protesto"
            Caption         =   "Protesto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "DataProtesto"
            Caption         =   "DataProtesto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "EmpresaDeCobranca"
            Caption         =   "Emp.Cob."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0,000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "DataEmpresaDeCobranca"
            Caption         =   "DataEmpresaDeCobranca"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "UsuarioLanc"
            Caption         =   "UsuarioLanc"
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
         BeginProperty Column21 
            DataField       =   "DataCaixa"
            Caption         =   "DataCaixa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column22 
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   569,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   750,047
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   96
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Celular:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   -73080
         TabIndex        =   95
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Resultado:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Nome do Contato:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   89
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   73
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   -70800
         TabIndex        =   70
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Ag.:"
         Height          =   255
         Left            =   -71520
         TabIndex        =   68
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bco. Nr:"
         Height          =   195
         Left            =   -72240
         TabIndex        =   66
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Comp:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Placa:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Veículo:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   59
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11325
      TabIndex        =   45
      Top             =   6885
      Width           =   11325
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   2925
         TabIndex        =   41
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   4020
         TabIndex        =   42
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   5115
         TabIndex        =   43
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   6210
         TabIndex        =   44
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   7305
         TabIndex        =   46
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   1845
         TabIndex        =   40
         Top             =   0
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   0
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from chequesClientes order by nome"
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
   End
   Begin MSComCtl2.DTPicker txtDatalistaIni 
      Height          =   300
      Left            =   120
      TabIndex        =   76
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72613889
      CurrentDate     =   38949
   End
   Begin MSAdodcLib.Adodc qPendentes 
      Height          =   375
      Left            =   1080
      Top             =   4560
      Visible         =   0   'False
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
      RecordSource    =   "select *from cheques"
      Caption         =   "qPendentes"
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
   Begin MSAdodcLib.Adodc dbCheques 
      Height          =   375
      Left            =   1080
      Top             =   2280
      Visible         =   0   'False
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
      Connect         =   "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Maria Vitoria;Data Source=temvale17"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Maria Vitoria;Data Source=temvale17"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from cheques"
      Caption         =   "dbCheques"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCadChequeCliente.frx":104F
      Height          =   4455
      Left            =   120
      TabIndex        =   101
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7858
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CodigoChequeCliente"
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
         DataField       =   "CIC"
         Caption         =   "CIC"
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
         DataField       =   "CNPJ"
         Caption         =   "CNPJ"
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
            Alignment       =   1
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label24 
      Caption         =   "Relatório de Cheques:"
      Height          =   255
      Left            =   120
      TabIndex        =   110
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "a"
      Height          =   255
      Left            =   1680
      TabIndex        =   86
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "Período de validade da lista de cheques:"
      Height          =   255
      Left            =   120
      TabIndex        =   85
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label14 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "CNPJ:"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "CIC:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCadChequeCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Verificando As Boolean, Imprimindo As Boolean, Disquete As Boolean
Public Conectado As Boolean, Desconectar As Boolean, Porta As Integer
Dim LimiteCheque As Currency, strOrdem As String

Private Sub CabecaLista(ByVal Largura As Double, ByVal Dia As Date)
Dim StrTemp As String

StrTemp = "Lista de Clientes"
Printer.FontSize = 14
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.FontSize = 12
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = "Lista válida de: " & Format(txtDatalistaIni.Value, "Short Date") & " a " & Format(txtDataLista.Value, "Short Date")
Printer.FontSize = 12
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = "Impressa em: " & Format(Dia, "long date") & " - " & Format(Dia, "short time")
Printer.FontSize = 8
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.FontSize = 8

StrTemp = "Cod."
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Nome"
Printer.CurrentX = 10
Printer.Print StrTemp;

StrTemp = "CIC / CNPJ"
Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Lim. Máximo"
Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Chqs. Pendentes"
Printer.CurrentX = 137 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total Permitido"
Printer.CurrentX = 157 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cheques Adicionados"
Printer.CurrentX = 159
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

End Sub

Private Sub ImprimeListaCheque()
Dim Resposta As Integer, Imprime As Boolean
Dim Largura As Double, Dia As Date, Saldo As Currency
Dim db As Database, Ws As Workspace


Resposta = MsgBox("A lista deve ser válida de " & txtDatalistaIni.Value & " até " & txtDataLista.Value & "?", vbYesNo)
If Resposta = vbNo Then
  txtDataLista.SetFocus
  Exit Sub
End If

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

Largura = 190
Dia = Now
CabecaLista Largura, Dia

With Adodc1
  .RecordSource = "Select *from chequesclientes where atualizar=0 and consultado=-1 and posicao=-1 and devolvidos=0" & strOrdem
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    Do While .Recordset.EOF = False
      If Printer.CurrentY + 25 > Printer.ScaleHeight Then
'        Printer.CurrentY = Printer.CurrentY + 0.5
'        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
'        Printer.CurrentY = Printer.CurrentY + 0.5
        
        Printer.Print ""
        StrTemp = "Página: " & Printer.Page
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.NewPage
        CabecaLista Largura, Dia
      End If
      
      Printer.FontSize = 8
      Imprime = True
      If .Recordset!consultado = True And .Recordset!Posicao = True Then
        Imprime = True
      Else
        If .Recordset!Devolvidos <> 0 Then
          Imprime = True
        End If
      End If
      If Imprime = True Then
        StrTemp = .Recordset!codigochequecliente
        Printer.CurrentX = 0
        Printer.Print StrTemp;
        
        StrTemp = .Recordset!Nome
        Printer.CurrentX = 10
        Printer.Print StrTemp;
        
        StrTemp = ""
        If IsNull(.Recordset!CIC) = False Then
          If .Recordset!CIC = "" Then
            If IsNull(.Recordset!CNPJ) = False Then
              If .Recordset!CNPJ <> "" Then
                StrTemp = .Recordset!CNPJ
              End If
            End If
          Else
            StrTemp = .Recordset!CIC
          End If
        Else
          If IsNull(.Recordset!CNPJ) = False Then
            If .Recordset!CNPJ <> "" Then
              StrTemp = .Recordset!CNPJ
            End If
          End If
        End If
        Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Saldo = 0
        If IsNull(.Recordset!Limitevalor2) = True Then
          StrTemp = Format(0, "Currency")
        Else
          Saldo = .Recordset!Limitevalor2
          StrTemp = Format(.Recordset!Limitevalor2, "Currency")
        End If
        Printer.CurrentX = 120 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        If IsNull(.Recordset!saldopendente) = False Then
          Saldo = Saldo - .Recordset!saldopendente
          StrTemp = Format(.Recordset!saldopendente, "Currency")
        Else
          StrTemp = Format(0, "Currency")
        End If
        Printer.CurrentX = 137 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(Saldo, "Currency")
        Printer.CurrentX = 157 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.5
      
      End If
      .Recordset.MoveNext
    Loop
    
    Printer.Print ""
    StrTemp = "Página: " & Printer.Page
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
  End If
  .Refresh
End With
Printer.EndDoc


Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
db.Execute "update cheques set datalista=#" & DataInglesa(Date) & "# where datalista=null"

NaoImprime:

End Sub

Private Sub ImprimeInativos()
Dim StrTemp As String, Largura As Double, Dia As Date
Dim Y1 As Double, Y2 As Double
Dim X1 As Double, X2 As Double

With Adodc1
  .Refresh
  .Recordset.Filter = "posicao=0"
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  Cabeca2 Largura, Dia
  
  Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 38 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = 0
        Printer.NewPage
        Cabeca2 Largura, Dia
      End If
      
      
      Y1 = Printer.CurrentY + 1
      
      Printer.ForeColor = RGB(200, 200, 200)
      Printer.Line (0, Y1)-(Largura, Y1 + 33), , BF
      
      Printer.ForeColor = vbWhite
      Printer.Line (2, Y1 + 3)-(32, Y1 + 9), , BF
      Printer.Line (33, Y1 + 3)-(188, Y1 + 9), , BF
      Printer.Line (2, Y1 + 10)-(126, Y1 + 16), , BF
      Printer.Line (127, Y1 + 10)-(188, Y1 + 16), , BF
      Printer.Line (2, Y1 + 17)-(34, Y1 + 23), , BF
      Printer.Line (35, Y1 + 17)-(72, Y1 + 23), , BF
      Printer.Line (73, Y1 + 17)-(116, Y1 + 23), , BF
      Printer.Line (117, Y1 + 17)-(160, Y1 + 23), , BF
      Printer.Line (161, Y1 + 17)-(188, Y1 + 23), , BF
      Printer.Line (2, Y1 + 24)-(57, Y1 + 30), , BF
      Printer.Line (58, Y1 + 24)-(105, Y1 + 30), , BF
      Printer.Line (106, Y1 + 24)-(159, Y1 + 30), , BF
      Printer.Line (160, Y1 + 24)-(188, Y1 + 30), , BF
      
      Printer.FontName = "Arial"
      Printer.FontSize = 7
      Printer.ForeColor = vbBlack
      On Error Resume Next
      StrTemp = "Código"
      Printer.CurrentX = 3
      Printer.CurrentY = Y1 + 3
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 3
      StrTemp = .Recordset!codigochequecliente
      Printer.Print StrTemp
      
      StrTemp = "Nome"
      Printer.CurrentX = 34
      Printer.CurrentY = Y1 + 3
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 34
      StrTemp = .Recordset!Nome
      Printer.Print StrTemp
      
      StrTemp = "Endereço"
      Printer.CurrentX = 3
      Printer.CurrentY = Y1 + 10
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 3
      StrTemp = .Recordset!Endereco
      Printer.Print StrTemp
      
      StrTemp = "Bairro"
      Printer.CurrentX = 128
      Printer.CurrentY = Y1 + 10
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 128
      StrTemp = .Recordset!Codigo
      Printer.Print StrTemp
      
      StrTemp = "CEP"
      Printer.CurrentX = 3
      Printer.CurrentY = Y1 + 17
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 3
      StrTemp = .Recordset!CEP
      Printer.Print StrTemp
      
      StrTemp = "Telefone"
      Printer.CurrentX = 36
      Printer.CurrentY = Y1 + 17
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 36
      StrTemp = Format(.Recordset!Telefone, "(###)####-####")
      Printer.Print StrTemp
      
      
      StrTemp = "CIC"
      Printer.CurrentX = 74
      Printer.CurrentY = Y1 + 17
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 74
      StrTemp = Format(.Recordset!CIC, "##,###,###,###-##")
      Printer.Print StrTemp
      
      StrTemp = "RG"
      Printer.CurrentX = 118
      Printer.CurrentY = Y1 + 17
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 118
      StrTemp = Format(.Recordset!rg, "###,###,###,###-#")
      Printer.Print StrTemp
      
      StrTemp = "Emissão"
      Printer.CurrentX = 162
      Printer.CurrentY = Y1 + 17
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 162
      StrTemp = .Recordset!Origem & " - " & .Recordset!origem2
      Printer.Print StrTemp
      
      StrTemp = "CNPJ"
      Printer.CurrentX = 3
      Printer.CurrentY = Y1 + 24
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 3
      StrTemp = Format(.Recordset!CNPJ, "##,###,###/####-##")
      Printer.Print StrTemp
      
      StrTemp = "I.E."
      Printer.CurrentX = 59
      Printer.CurrentY = Y1 + 24
      Printer.Print StrTemp
      StrTemp = ""
      Printer.CurrentX = 59
      StrTemp = Format(.Recordset!ie, "###,###,###,###")
      Printer.Print StrTemp
      
      StrTemp = "Carro"
      Printer.CurrentX = 106
      Printer.CurrentY = Y1 + 24
      Printer.Print StrTemp
      StrTemp = ""
      If dbCarros.Recordset.EOF = False Then
        Printer.CurrentX = 106
        StrTemp = dbCarros.Recordset!Carro
        Printer.Print StrTemp
      End If
      
      StrTemp = "Placa"
      Printer.CurrentX = 161
      Printer.CurrentY = Y1 + 24
      Printer.Print StrTemp
      StrTemp = ""
      If dbCarros.Recordset.EOF = False Then
        Printer.CurrentX = 161
        StrTemp = dbCarros.Recordset!Placa
        Printer.Print StrTemp
      End If
      
      Printer.CurrentY = Y1 + 33
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub Cabeca2(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Clientes Inativos"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
End Sub

Private Sub ImprimeAtivos()
Dim StrTemp As String, Largura As Double, Dia As Date

With Adodc1
  
  .RecordSource = "select *from ChequesClientes where consultado=0 and posicao=-1" & strOrdem
  .Refresh
  
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  CabecaAtivo Largura, Dia
  
  Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = 0
        Printer.NewPage
        CabecaAtivo Largura, Dia
      End If
      StrTemp = .Recordset!codigochequecliente
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      If IsNull(.Recordset!CIC) = False Then
        StrTemp = Format(.Recordset!CIC, "###,###,###,###-##")
      Else
        StrTemp = Format(.Recordset!CNPJ, "###,###,###/####-##")
      End If
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!numerodecheques, "#,###")
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Depositados, "#,###")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valordepositado, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
  
End With
NaoImprime:
Adodc1.RecordSource = "select *from ChequesClientes" & strOrdem
End Sub

Private Sub ImprimeAtivos2()
Dim StrTemp As String, Largura As Double, Dia As Date

With Adodc1
  
  .RecordSource = "select *from ChequesClientes where consultado=-1 and posicao=-1" & strOrdem
  .Refresh
  
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  CabecaAtivo2 Largura, Dia
  
  Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = 0
        Printer.NewPage
        CabecaAtivo2 Largura, Dia
      End If
      StrTemp = .Recordset!codigochequecliente
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      If IsNull(.Recordset!CIC) = False Then
        StrTemp = Format(.Recordset!CIC, "###,###,###,###-##")
      Else
        StrTemp = Format(.Recordset!CNPJ, "###,###,###/####-##")
      End If
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!numerodecheques, "#,###")
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Depositados, "#,###")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valordepositado, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
  
End With
NaoImprime:
Adodc1.RecordSource = "select *from ChequesClientes" & strOrdem
End Sub

Private Sub ImprimeInativosConsultados()
Dim StrTemp As String, Largura As Double, Dia As Date

With Adodc1
  
  .RecordSource = "select *from ChequesClientes where consultado=-1 and posicao=0" & strOrdem
  .Refresh
  
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  CabecaInativo2 Largura, Dia
  
  Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = 0
        Printer.NewPage
        CabecaInativo2 Largura, Dia
      End If
      StrTemp = .Recordset!codigochequecliente
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      If IsNull(.Recordset!CIC) = False Then
        StrTemp = Format(.Recordset!CIC, "###,###,###,###-##")
      Else
        StrTemp = Format(.Recordset!CNPJ, "###,###,###/####-##")
      End If
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!numerodecheques, "#,###")
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Depositados, "#,###")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valordepositado, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
  
End With
NaoImprime:
Adodc1.RecordSource = "select *from ChequesClientes" & strOrdem
Adodc1.Refresh
End Sub
Private Sub ImprimeTodos()
Dim StrTemp As String, Largura As Double, Dia As Date

With Adodc1
  
  .RecordSource = "select *from ChequesClientes" & strOrdem
  .Refresh
  
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
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
        
        Printer.CurrentY = 0
        Printer.NewPage
        Cabeca Largura, Dia
      End If
      StrTemp = .Recordset!codigochequecliente
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      If IsNull(.Recordset!CIC) = False Then
        StrTemp = Format(.Recordset!CIC, "###,###,###,###-##")
      Else
        StrTemp = Format(.Recordset!CNPJ, "###,###,###/####-##")
      End If
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!numerodecheques, "#,###")
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Depositados, "#,###")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valordepositado, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
  
End With
NaoImprime:
Adodc1.RecordSource = "select *from ChequesClientes" & strOrdem
Adodc1.Refresh

End Sub

Private Sub Cabeca(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Todos os Clientes Cadastrados"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontBold = False
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Cod."
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Nome"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  StrTemp = "CPF/CNPJ"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Rec."
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Comp."
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
End Sub

Private Sub CabecaAtivo(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Clientes Ativos"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontBold = False
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Cod."
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Nome"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  StrTemp = "CPF/CNPJ"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Rec."
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Comp."
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
End Sub

Private Sub CabecaAtivo2(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Clientes Ativos e Consultados"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontBold = False
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Cod."
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Nome"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  StrTemp = "CPF/CNPJ"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Rec."
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Comp."
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
End Sub
Private Sub CabecaInativo2(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Clientes Consultados e Inativos"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontBold = False
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Cod."
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Nome"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  StrTemp = "CPF/CNPJ"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Rec."
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Comp."
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
End Sub

Private Sub ImprimeDevolvidos()
Dim StrTemp As String, Largura As Double, Dia As Date

With Adodc1
  
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  Cabeca3 Largura, Dia
  
  Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = 0
        Printer.NewPage
        Cabeca3 Largura, Dia
      End If
      StrTemp = .Recordset!codigochequecliente
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = .Recordset!Nome
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      If IsNull(.Recordset!CIC) = False Then
        StrTemp = Format(.Recordset!CIC, "###,###,###,###-##")
      Else
        StrTemp = Format(.Recordset!CNPJ, "###,###,###/####-##")
      End If
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!numerodecheques, "#,###")
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Depositados, "#,###")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!valordepositado, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub Cabeca3(ByVal Largura As Double, ByVal Dia As Date)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  
  StrTemp = "Relatório de Clientes com Cheque Devolvido"
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  Printer.FontBold = False
  Printer.FontSize = 8
  StrTemp = "Data: " & Format(Dia, "Long Date")
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  
  
  StrTemp = "Cod."
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Nome"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  StrTemp = "CPF/CNPJ"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Rec."
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Comp."
  Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  StrTemp = "Total"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
End Sub

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
MsgBox ErrorNumber & " - " & Description
Response = False
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim CodCliente As Double

Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
If Imprimindo = True Then Exit Sub
If Verificando = True Then Exit Sub
If Adodc1.Recordset.EOF = True Then Exit Sub
On Error GoTo TrataErro
CodCliente = Adodc1.Recordset!codigochequecliente

With dbContas
  .Recordset.Filter = "codigocliente=" & CodCliente
End With
With dbCarros
  .Recordset.Filter = "codigocliente=" & CodCliente
End With
DataGrid1.Visible = False
With dbCobranca
  .Recordset.Filter = "codigocliente=" & CodCliente
  .Recordset.Sort = "codigohistorico desc"
End With

On Error Resume Next
Dim Saldo As Currency
lblPre.Caption = Format(Adodc1.Recordset!saldopendente, "Currency")
Saldo = Adodc1.Recordset!Limitevalor2 - Adodc1.Recordset!saldopendente
lblSaldo.Caption = Format(Saldo, "Currency")

TrataErro:
Exit Sub

End Sub

Private Sub Adodc1_WillChangeRecordset(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Save = True Then
  If QuerGravar = False Then
    Adodc1.Recordset.CancelUpdate
  End If
End If
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  Adodc1.Recordset!Limitevalor2 = LimiteCheque
  Adodc1.Recordset!datacadastro = Now
  Adodc1.Recordset!cadastradopor = Usuarios.Nome
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(18).Text = 5
  MaskEdBox1.Text = 100
  txtFields(1).SetFocus
End Sub

Private Sub cmdContar_Click()
Dim Total As Double, ValorTotal As Currency
Dim Depositados As Double, ValorDep As Currency
Dim Devolvidos As Double, ValorDev As Currency
Dim ValorEmAberto As Currency


If Adodc1.Recordset.RecordCount = 0 Then Exit Sub

Verificando = True
With dbCheques
  .RecordSource = "select comp, banco, agencia, conta, codigocliente from cheques where codigocliente=" & Adodc1.Recordset!codigochequecliente & " group by comp, banco, agencia, conta, codigocliente"
  .Refresh
End With

With dbContas
  .Recordset.Filter = "codigocliente=" & Adodc1.Recordset!codigochequecliente
End With

With dbCheques
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If dbContas.Recordset.RecordCount <> 0 Then
        dbContas.Recordset.MoveFirst
        dbContas.Recordset.Find "conta='" & .Recordset!Conta & "'"
        If dbContas.Recordset.EOF = True Then
          dbContas.Recordset.AddNew
          dbContas.Recordset!COMP = .Recordset!COMP
          dbContas.Recordset!banconumero = .Recordset!Banco
          dbContas.Recordset!Agencia = .Recordset!Agencia
          dbContas.Recordset!Conta = .Recordset!Conta
          dbContas.Recordset!CodigoCliente = Adodc1.Recordset!codigochequecliente
          dbContas.Recordset.Update
        End If
      Else
        dbContas.Recordset.AddNew
        dbContas.Recordset!COMP = .Recordset!COMP
        dbContas.Recordset!banconumero = .Recordset!Banco
        dbContas.Recordset!Agencia = .Recordset!Agencia
        dbContas.Recordset!Conta = .Recordset!Conta
        dbContas.Recordset!CodigoCliente = Adodc1.Recordset!codigochequecliente
        dbContas.Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With Adodc1
  If .Recordset.RecordCount = 0 Then Exit Sub
    Total = 0
    ValorTotal = 0
    Depositados = 0
    ValorDep = 0
    Devolvidos = 0
    ValorDev = 0
    ValorEmAberto = 0
    DoEvents
    With dbCheques2
        .RecordSource = "select *from cheques where codigocliente=" & Adodc1.Recordset!codigochequecliente & " order by datacheque"
        .Refresh
        If .Recordset.RecordCount <> 0 Then
          .Recordset.MoveLast
          Total = .Recordset.RecordCount
          .Recordset.MoveFirst
          Do While .Recordset.EOF = False
            ValorTotal = ValorTotal + .Recordset!Valor
            If .Recordset!compensado = True Then
              Depositados = Depositados + 1
              ValorDep = ValorDep + .Recordset!Valor
            Else
              ValorEmAberto = ValorEmAberto + .Recordset!Valor
            End If
            If .Recordset!devolvido = True And .Recordset!compensado = False Then
              Devolvidos = Devolvidos + 1
              ValorDev = ValorDev + .Recordset!Valor
            End If
            .Recordset.MoveNext
          Loop
        End If
    End With
    .Recordset!numerodecheques = Total
    .Recordset!Total = ValorTotal
    .Recordset!Depositados = Depositados
    .Recordset!valordepositado = ValorDep
    .Recordset!Devolvidos = Devolvidos
    .Recordset!saldopendente = ValorEmAberto
    If Devolvidos > 0 Then
      .Recordset!Posicao = False
    End If
    .Recordset!valordevolvido = ValorDev
    On Error Resume Next
    .Recordset.Update
End With
DataGrid2.SetFocus

Verificando = False
End Sub

Private Sub cmdDelete_Click()
  Dim Resposta As Integer
  
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
txtFields(1).SetFocus
End Sub

Private Sub cmdExibeCheques_Click()
Dim CodCliente As Double
With Adodc1
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  CodCliente = .Recordset!codigochequecliente
End With
With dbCheques2
  .ConnectionString = CaminhoADO
  .RecordSource = "select cheques.*, fechamentodecaixa.* from cheques, fechamentodecaixa where cheques.codigofechamento=fechamentodecaixa.codigofechamento and codigocliente=" & CodCliente & " order by datacheque"
  .Refresh
End With
DataGrid1.Visible = True
End Sub

Private Sub cmdImprime_Click()
Dim Resposta As Integer

Imprimindo = True

Select Case cboRelatorio.Text
  Case "Lista para aceitar cheque no posto"
    ImprimeListaCheque
  Case "Clientes Ativos"
    ImprimeAtivos
  Case "Clientes Ativos e Consultados"
    ImprimeAtivos2
  Case "Clientes Consultados"
    ImprimeInativosConsultados
  Case "Clientes com cheque devolvido"
    Adodc1.RecordSource = "select *from chequesclientes where devolvidos>=1" & strOrdem
    Adodc1.Refresh
    ImprimeDevolvidos
  Case "Todos os clientes cadastrados"
    ImprimeTodos
End Select

Cancelado:
Adodc1.RecordSource = "select *from chequesclientes" & strOrdem
Adodc1.Refresh
Imprimindo = False
End Sub

Private Sub cmdImprimeCheques_Click()

If dbCheques2.Recordset.RecordCount = 0 Then
  MsgBox "Não existe cheque para ser impresso"
  Exit Sub
End If

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

Printer.Orientation = vbPRORLandscape

ImprimeADOGrid DataGrid1, Printer, dbCheques2, , , , , , , "Lista de Cheques por Cliente", Adodc1.Recordset!codigochequecliente & " - " & Adodc1.Recordset!Nome, "Impresso em: " & UCase(Format(Now, "long Date") & " - " & Format(Now, "short time"))

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdIncluir_Click()
With dbCarros
  .Recordset.AddNew
  .Recordset!Carro = txtCarro.Text
  .Recordset!Placa = txtPlaca.Text
  .Recordset!CodigoCliente = Adodc1.Recordset!codigochequecliente
  .Recordset.Update
  .Refresh
  .Refresh
End With
txtCarro.Text = ""
txtPlaca.Text = "   -    "
txtCarro.SetFocus
End Sub

Private Sub cmdIncluirCobranca_Click()
If Adodc1.Recordset.EOF = True Then
  MsgBox "Selecione um cliente!"
  Exit Sub
End If
If txtNomeContato.Text = "" Then
  MsgBox "Indique o nome do Contato!"
  txtNomeContato.SetFocus
  Exit Sub
End If
If txtObs.Text = "" Then
  MsgBox "Informe o resultado da cobrança!"
  txtObs.SetFocus
  Exit Sub
End If
With dbCobranca
  .Recordset.AddNew
  .Recordset!CodigoCliente = Adodc1.Recordset!codigochequecliente
  .Recordset!lancadoem = Now
  .Recordset!Usuario = Usuarios.Nome
  .Recordset!contato = txtNomeContato.Text
  .Recordset!Obs = txtObs.Text
  .Recordset.Update
End With
End Sub

Private Sub cmdIncluirConta_Click()
With dbContas
  .Recordset.AddNew
  .Recordset!CodigoCliente = Adodc1.Recordset!codigochequecliente
  .Recordset!COMP = txtComp.Text
  If txtBanco.Text = "" Then txtBanco.Text = " "
  .Recordset!Banco = txtBanco.Text
  .Recordset!banconumero = txtBancoNr.Text
  .Recordset!Agencia = txtAg.Text
  .Recordset!Conta = txtConta.Text
  .Recordset.Update
  .Refresh
  .Refresh
End With
End Sub

Private Sub cmdLocalizaCheque_Click()
Screen.MousePointer = vbHourglass
frmCadChequesClienteLocalizar.Show
frmCadChequesClienteLocalizar.SetFocus
Screen.MousePointer = vbdefaul
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    A = .Recordset!codigochequecliente
    .Recordset.Update
    .Refresh
    .Recordset.Find "codigochequecliente=" & A
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

Private Sub Command1_Click()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim Dia As Date

With Adodc1
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Dia = DateAdd("m", -12, Date)
    db.Open CaminhoADO
    dbTemp.Open "Select *from cheques order by codigocliente, datacheque desc", db, adOpenKeyset
    If dbTemp.RecordCount <> 0 Then
      Do While .Recordset.EOF = False
        If .Recordset!Posicao = True Then
          dbTemp.MoveFirst
          dbTemp.Find "codigocliente=" & .Recordset!codigochequecliente
          If dbTemp.EOF = False Then
            If dbTemp!datacheque <= Dia Then
              .Recordset!Posicao = False
              .Recordset.Update
            End If
          End If
        End If
        DoEvents
        .Recordset.MoveNext
      Loop
    End If
    dbTemp.Close
    db.Close
  End If
End With
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DataGrid2.Columns(ColIndex).DataField
End If
Adodc1.RecordSource = "select *from chequesclientes" & strOrdem
Adodc1.Refresh
End Sub

Private Sub DBGrid4_HeadClick(ByVal ColIndex As Integer)
If dbCheques2.RecordSource = "select cheques.*, fechamentodecaixa.* from cheques, fechamentodecaixa where cheques.codigofechamento=fechamentodecaixa.codigofechamento and codigocliente=" & Adodc1.Recordset!codigochequecliente & " order by " & DBGrid4.Columns(ColIndex).DataField Then
  dbCheques2.RecordSource = "select cheques.*, fechamentodecaixa.* from cheques, fechamentodecaixa where cheques.codigofechamento=fechamentodecaixa.codigofechamento and codigocliente=" & Adodc1.Recordset!codigochequecliente & " order by " & DBGrid4.Columns(ColIndex).DataField & " desc"
Else
  dbCheques2.RecordSource = "select cheques.*, fechamentodecaixa.* from cheques, fechamentodecaixa where cheques.codigofechamento=fechamentodecaixa.codigofechamento and codigocliente=" & Adodc1.Recordset!codigochequecliente & " order by " & DBGrid4.Columns(ColIndex).DataField
End If
dbCheques2.Refresh
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
txtDataLista.Value = DateAdd("d", 3, Date)
txtDatalistaIni.Value = DateAdd("d", 1, Date)

strOrdem = " order by Nome"

With cboRelatorio
  .Clear
  .AddItem "Lista para aceitar cheque no posto"
  .AddItem "Clientes Ativos"
  .AddItem "Clientes Ativos e Consultados"
  .AddItem "Clientes Consultados"
  .AddItem "Clientes com cheque devolvido"
  .AddItem "Todos os clientes cadastrados"
End With

With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCobranca
  .ConnectionString = CaminhoADO
  .Refresh
End With
With qPendentes
  .ConnectionString = CaminhoADO
  .RecordSource = "select codigocliente, count(codigocheque) as cheques from cheques where datacheque>=#" & DataInglesa(Date) & "# group by codigocliente"
  .Refresh
End With
With qPendentes2
  .ConnectionString = CaminhoADO
  .RecordSource = "select codigocliente, sum(valor) as total from cheques where datacheque>=#" & DataInglesa(Date) & "# group by codigocliente"
  .Refresh
End With
With dbCheques
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCheques2
  .ConnectionString = CaminhoADO
  .Refresh
End With

With Adodc1
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from chequesClientes" & strOrdem
  .Refresh
End With

Select Case Usuarios.Grupo.CadClienteCheque
  Case 1 'Somente leitura
    txtFields(18).Enabled = False
    MaskEdBox1.Enabled = False
    MaskEdBox2.Enabled = False
  Case 2 'Liberado
    txtFields(18).Enabled = True
    MaskEdBox1.Enabled = True
    MaskEdBox2.Enabled = True
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub Text1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtCod_Change()
With Adodc1
  If txtCod.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "codigochequecliente=" & txtCod.Text
End With
End Sub

Private Sub txtDataLista_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataLista_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataLista_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDatalistaIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDatalistaIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDatalistaIni_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
  Case 5, 9
    If Fu_consistir_CgcCpf(txtFields(Index).Text) = False Then
      MsgBox "CNPJ ou CPF inválido!"
    End If
End Select
End Sub

Private Sub txtProcuraCIC_Change()
With Adodc1
  If txtProcuraCIC.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "cic like '" & txtProcuraCIC.Text & "*'"
End With
End Sub

Private Sub txtProcuraCNPJ_Change()
With Adodc1
  If txtProcuraCNPJ.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "cnpj like '" & txtProcuraCNPJ.Text & "*'"
End With
End Sub

Private Sub txtProcuraNome_Change()
With Adodc1
  If txtProcuraNome.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "Nome like '" & txtProcuraNome.Text & "*'"
End With
End Sub
