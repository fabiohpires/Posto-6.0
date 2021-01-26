VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFechamentoDeCaixaNovo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   Icon            =   "frmFechamentoDeCaixaNovo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar dados"
      Height          =   375
      Left            =   1560
      TabIndex        =   59
      ToolTipText     =   "Importa os dados do caixa (F5)"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   5055
      Left            =   3720
      TabIndex        =   55
      Top             =   6360
      Visible         =   0   'False
      Width           =   5895
      Begin MSAdodcLib.Adodc dbPdvs 
         Height          =   330
         Left            =   120
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from pdvs order by descri"
         Caption         =   "dbPdvs"
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
         Left            =   120
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from Turnos order by horaini"
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
      Begin MSAdodcLib.Adodc dbVendedores 
         Height          =   330
         Left            =   120
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from vendedores order by nome"
         Caption         =   "dbVendedores"
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
      Begin MSAdodcLib.Adodc dbFechamentos 
         Height          =   330
         Left            =   120
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from fechamentodecaixa where codigofechamento=0"
         Caption         =   "dbFechamentos"
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
      Begin MSAdodcLib.Adodc dbEncerrantes 
         Height          =   330
         Left            =   120
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from bicoEncerrantes where codigofechamento=0"
         Caption         =   "dbEncerrantes"
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
      Begin MSAdodcLib.Adodc dbVendas 
         Height          =   330
         Left            =   120
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from venda2 where codigofechamento=0"
         Caption         =   "dbVendas"
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
      Begin MSAdodcLib.Adodc dbDifComb 
         Height          =   330
         Left            =   120
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from diferencacombustivel where codigofechamento=0"
         Caption         =   "dbDifComb"
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
      Begin MSAdodcLib.Adodc qProdutosAltera 
         Height          =   330
         Left            =   120
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"frmFechamentoDeCaixaNovo.frx":0442
         Caption         =   "qProdutosAltera"
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
      Begin MSAdodcLib.Adodc dbErros 
         Height          =   330
         Left            =   120
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "importacaoerros"
         Caption         =   "dbErros"
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
      Begin MSAdodcLib.Adodc dbTanques2 
         Height          =   330
         Left            =   120
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tanques"
         Caption         =   "dbTanques2"
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
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmFechamentoDeCaixaNovo.frx":051E
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4620
      Left            =   240
      TabIndex        =   43
      Top             =   1560
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8149
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Bicos"
      TabPicture(0)   =   "frmFechamentoDeCaixaNovo.frx":0535
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTotalValorComb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbltotalQtdComb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Produtos"
      TabPicture(1)   =   "frmFechamentoDeCaixaNovo.frx":0551
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDesconto"
      Tab(1).Control(1)=   "cmdRemoverProduto"
      Tab(1).Control(2)=   "DataGrid2"
      Tab(1).Control(3)=   "txtCodProduto"
      Tab(1).Control(4)=   "txtQtd"
      Tab(1).Control(5)=   "txtCodFunc"
      Tab(1).Control(6)=   "cmdIncluir"
      Tab(1).Control(7)=   "cboProdutos"
      Tab(1).Control(8)=   "lblDesconto"
      Tab(1).Control(9)=   "lblTotalVenda"
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(12)=   "lblComissoes2"
      Tab(1).Control(13)=   "lblTotalProdutos2"
      Tab(1).Control(14)=   "Label15"
      Tab(1).Control(15)=   "Label1"
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(17)=   "Label8"
      Tab(1).Control(18)=   "Label11"
      Tab(1).Control(19)=   "lblEstoque"
      Tab(1).Control(20)=   "Label14"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Estoque"
      TabPicture(2)   =   "frmFechamentoDeCaixaNovo.frx":056D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdEntraCombustivel"
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).Control(2)=   "cmdTransfere"
      Tab(2).Control(3)=   "cmdExtornaTurno"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Resumo"
      TabPicture(3)   =   "frmFechamentoDeCaixaNovo.frx":0589
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExibeComissoes"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Erros de Importação"
      TabPicture(4)   =   "frmFechamentoDeCaixaNovo.frx":05A5
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid4"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69000
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdExtornaTurno 
         Caption         =   "Extorna Notas futuras"
         Height          =   375
         Left            =   -73440
         TabIndex        =   69
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton cmdExibeComissoes 
         Caption         =   "Exibe Comissões"
         Height          =   375
         Left            =   -72120
         TabIndex        =   68
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdTransfere 
         Caption         =   "Transferência"
         Height          =   375
         Left            =   -66720
         TabIndex        =   65
         Top             =   4080
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":05C1
         Height          =   3735
         Left            =   -74880
         TabIndex        =   60
         Top             =   600
         Width           =   10455
         _ExtentX        =   18441
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "DataImportado"
            Caption         =   "Importado"
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
         BeginProperty Column02 
            DataField       =   "Descri"
            Caption         =   "Descrição"
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
            DataField       =   "bico"
            Caption         =   "bico"
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
            DataField       =   "CodigoNoPosto"
            Caption         =   "Prod. Posto"
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
            DataField       =   "CodigoFuncionario"
            Caption         =   "Func."
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
            DataField       =   "CodigoProduto"
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
         BeginProperty Column07 
            DataField       =   "ValorPosto"
            Caption         =   "Valor Posto"
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
            DataField       =   "ValorSistema"
            Caption         =   "Valor Sist."
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
            DataField       =   "CodigoClienteNoPosto"
            Caption         =   "Cli. Posto"
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
         BeginProperty Column10 
            DataField       =   "CodigoClienteSistema"
            Caption         =   "Cli. Sist."
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
         BeginProperty Column11 
            DataField       =   "LimiteNadata"
            Caption         =   "Limite Data"
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
            DataField       =   "ValorBloqueado"
            Caption         =   "$ Bloq."
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
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2099,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   510,236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   824,882
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdRemoverProduto 
         Caption         =   "Remover"
         Height          =   375
         Left            =   -65520
         TabIndex        =   58
         Top             =   600
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":05D7
         Height          =   3495
         Left            =   -74880
         TabIndex        =   33
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "TanqueNr"
            Caption         =   "Tq"
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
            DataField       =   "Vendido"
            Caption         =   "Vendido"
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
         BeginProperty Column03 
            DataField       =   "Estoque"
            Caption         =   "Sistema"
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
         BeginProperty Column04 
            DataField       =   "Tanque"
            Caption         =   "Posto"
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
            DataField       =   "Diferenca"
            Caption         =   "Diferença"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4334,74
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   915,024
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":05EF
         Height          =   3015
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5318
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "CodProduto"
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
            DataField       =   "Descri"
            Caption         =   "Descrição"
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
            DataField       =   "Quantidade"
            Caption         =   "Qtd."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,###.##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ValorUnitario"
            Caption         =   "R$ Unitário"
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
         BeginProperty Column04 
            DataField       =   "ValorTotal"
            Caption         =   "R$ Total"
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
            DataField       =   "CodigoVendedor"
            Caption         =   "Vendedor"
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
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4605,166
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   854,929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":0606
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "Bico"
            Caption         =   "Bico"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Abertura"
            Caption         =   "Abertura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Encerrante"
            Caption         =   "Encerrante"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Vendas"
            Caption         =   "Vendas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Retorno"
            Caption         =   "Retorno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "ValorTotal"
            Caption         =   "R$ Total"
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
         BeginProperty Column07 
            DataField       =   "DesteCaixaQtd"
            Caption         =   "Qdt Deste"
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
         BeginProperty Column08 
            DataField       =   "DesteCaixaValor"
            Caption         =   "R$ Deste"
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
         BeginProperty Column09 
            DataField       =   "DeOutroCaixaQtd"
            Caption         =   "Qtd Outro"
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
         BeginProperty Column10 
            DataField       =   "DeOutroCaixaValor"
            Caption         =   "R$ Outro"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   510,236
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1094,74
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   " Resumo "
         Height          =   3975
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   4575
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Comissões sobre Combustíveis:"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblComissoesCombustiveis 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   66
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblDiferenca 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   54
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Diferença calculada:"
            Height          =   255
            Left            =   720
            TabIndex        =   53
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblComissoes 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   52
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Comissões sobre produtos:"
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Venda de Combustível:"
            Height          =   255
            Left            =   720
            TabIndex        =   50
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Venda de Produtos:"
            Height          =   255
            Left            =   480
            TabIndex        =   49
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Faturamento calculado:"
            Height          =   255
            Left            =   720
            TabIndex        =   48
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblTotalCombustivel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblTotalProdutos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   46
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblFaturamento 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2640
            TabIndex        =   45
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdEntraCombustivel 
         Caption         =   "Entra Tanque"
         Height          =   375
         Left            =   -74880
         TabIndex        =   34
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtCodProduto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69960
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtCodFunc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69480
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -66480
         TabIndex        =   31
         Top             =   600
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cboProdutos 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":0622
         Height          =   315
         Left            =   -74280
         TabIndex        =   20
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "produtos.Descri"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.Label lblDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Desc.:"
         Height          =   195
         Left            =   -69000
         TabIndex        =   25
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblTotalVenda 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67680
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Venda de Produtos:"
         Height          =   255
         Left            =   -68400
         TabIndex        =   64
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Comissões a pagar:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   63
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblComissoes2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -70440
         TabIndex        =   62
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lblTotalProdutos2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -66720
         TabIndex        =   61
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   -74280
         TabIndex        =   19
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lbltotalQtdComb 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6840
         TabIndex        =   57
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lblTotalValorComb 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8640
         TabIndex        =   56
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod.:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtd:"
         Height          =   195
         Left            =   -69960
         TabIndex        =   21
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Func.:"
         Height          =   195
         Left            =   -69480
         TabIndex        =   23
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label11 
         Caption         =   "Total:"
         Height          =   195
         Left            =   -66960
         TabIndex        =   29
         Top             =   480
         Width           =   405
      End
      Begin VB.Label lblEstoque 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -68280
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Estoque:"
         Height          =   255
         Left            =   -68280
         TabIndex        =   27
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Ca&ncelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<<"
      Height          =   255
      Left            =   480
      TabIndex        =   40
      ToolTipText     =   "Anterior (F2)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdPosterior 
      Caption         =   ">>"
      Height          =   255
      Left            =   840
      TabIndex        =   39
      ToolTipText     =   "Prócimo (F3)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimeiro 
      Caption         =   "|<<"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      ToolTipText     =   "Primeiro (F1)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   ">>|"
      Height          =   255
      Left            =   1200
      TabIndex        =   37
      ToolTipText     =   "Último (F4)"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Im&primir"
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   375
      Left            =   9960
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdDesconfirmar 
      Caption         =   "Cancela Confirmação"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Vulgo Papel Higiênico"
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSDataListLib.DataCombo cboPdvs 
      Bindings        =   "frmFechamentoDeCaixaNovo.frx":0640
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
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   6240
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   161
      FullHeight      =   33
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
      Format          =   181141505
      CurrentDate     =   39974
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox txtInformado 
         DataField       =   "Informado"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dbFechamentos"
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cboResponsavel 
         Bindings        =   "frmFechamentoDeCaixaNovo.frx":0655
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Con&firmar"
         Height          =   375
         Left            =   9720
         TabIndex        =   36
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Responsável:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Faturamento Informado:"
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Data:"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label10 
      Caption         =   "PDV:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmFechamentoDeCaixaNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechamentoAnterior As Double, ErroNaSoma As Boolean
Dim Abrindo As Boolean, AlteraAnterior As Double, FechandoLote As Boolean
Dim MedeAntes As Boolean, Intermitente As Boolean, Arredondamento As Double
Dim SemTabelaDePrecos As Boolean

Dim db As New ADODB.Connection
Dim dbResponsavel2 As New ADODB.Recordset

Dim Hora1 As Date, Hora2 As Date

Private Sub IncluirProduto(Optional valorUnitario As Currency = 0, Optional ValorTotal As Currency = 0)
Dim Comissao As Currency, Unitario As Currency, Desconto As Currency
If ValorTotal = 0 Then
  TotalVenda
Else
  lblTotalVenda.Caption = Format(ValorTotal, "Currency")
End If
If IsNumeric(lblTotalVenda.Caption) = False Then
  MsgBox "Erro no total da venda!"
  Exit Sub
End If
If IsNumeric(txtCodProduto.Text) = False Then
  MsgBox "Informe um código válido!"
  txtCodProduto.SetFocus
  Exit Sub
End If

With qProdutosAltera
  '.Refresh
  If SemTabelaDePrecos = False Then
    .Recordset.Find "produtos.codigo=" & txtCodProduto.Text
    If .Recordset.EOF = True Then
      MsgBox "Produto não encontrado na tabela de preços!"
      txtCodProduto.SetFocus
      Exit Sub
    End If
  Else
    .Recordset.Find "codigo=" & txtCodProduto.Text
    If .Recordset.EOF = True Then
      MsgBox "Produto não encontrado na tabela de preços!"
      txtCodProduto.SetFocus
      Exit Sub
    End If
  End If
  
'  If .Recordset!Estoque < CDbl(txtQtd.Text) Then
'    Resposta = MsgBox("A venda atual tornará o estoque negativo! Deseja continuar?", vbYesNo + vbDefaultButton2)
'    If Resposta = vbNo Then Exit Sub
'  End If

  If .Recordset!Comissao <> 0 Or .Recordset!ComissaoValor <> 0 Then
    If txtCodFunc.Text = "" Then txtCodFunc.Text = "0"
'    If IsNumeric(txtCodFunc.Text) = False Or txtCodFunc.Text = "0" Then
'      MsgBox "Informe um codigo de funcionário válido!"
'      txtCodFunc.SetFocus
'      Exit Sub
'    End If
    dbResponsavel2.MoveFirst
    dbResponsavel2.Find "codigo=" & txtCodFunc.Text
    If dbResponsavel2.EOF = True Then
'      MsgBox "Funcionário " & txtCodFunc.Text & " não encontrado!"
'      txtCodFunc.SetFocus
'      Exit Sub
    End If
  Else
    'txtCodFunc.Text = ""
  End If
End With
If IsNumeric(txtQtd.Text) = False Then
  MsgBox "Informe uma quantidade válida!"
  txtQtd.SetFocus
  Exit Sub
End If

If SemTabelaDePrecos = False Then
  Unitario = qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda")
Else
  Unitario = qProdutosAltera.Recordset("PrecoVenda")
End If

If Unitario <> valorUnitario And valorUnitario <> 0 Then
  If Configura.PrecoDiferente <> 0 Then
    txtDesconto.Text = Format(CCur((valorUnitario - Unitario) * CDbl(txtQtd.Text)), "Currency")
    TotalVenda
  End If
End If
If IsNumeric(txtDesconto.Text) = True Then
  Desconto = CCur(txtDesconto.Text)
End If

'Unitario = Unitario + Desconto


With qProdutosAltera
  If .Recordset!Comissao <> 0 Then
    Comissao = (Unitario * txtQtd.Text) * (.Recordset!Comissao)
  End If
  If .Recordset!ComissaoValor <> 0 Then
    Comissao = Comissao + (.Recordset!ComissaoValor * txtQtd.Text)
  End If
End With
On Error GoTo 0
With dbVendas
  .Recordset.AddNew
  .Recordset!CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
  .Recordset!Hora = Now
  .Recordset!Data = txtData.Value
  If SemTabelaDePrecos = False Then
    .Recordset!CodigoProduto = qProdutosAltera.Recordset("produtos.CodigoProduto")
    .Recordset!CodProduto = qProdutosAltera.Recordset("produtos.Codigo")
    .Recordset!Descri = qProdutosAltera.Recordset("produtos.Descri")
  Else
    .Recordset!CodigoProduto = qProdutosAltera.Recordset("CodigoProduto")
    .Recordset!CodProduto = qProdutosAltera.Recordset("Codigo")
    .Recordset!Descri = qProdutosAltera.Recordset("Descri")
  End If
  .Recordset!Quantidade = CDbl(txtQtd.Text)
  .Recordset!valorUnitario = Unitario
  .Recordset!ValorDesconto = Desconto
  .Recordset!ValorTotal = CCur(lblTotalVenda.Caption)
  If IsNumeric(txtCodFunc.Text) = True Then
    If txtCodFunc.Text = 0 Then
      .Recordset!codigovendedor = 0
      .Recordset!CodigoPagamento = 0
    Else
      dbResponsavel2.MoveFirst
      dbResponsavel2.Find "codigo=" & txtCodFunc.Text
      If dbResponsavel2.EOF = False Then
        .Recordset!codigovendedor = txtCodFunc.Text
        .Recordset!CodigoPagamento = dbResponsavel2!codigovendedor
      Else
        .Recordset!codigovendedor = 0
        .Recordset!CodigoPagamento = 0
      End If
    End If
  End If
  .Recordset!ValorComissao = Comissao
  .Recordset.Update
  .Refresh
End With

TotalProdutos

txtCodFunc.Text = ""
txtCodProduto.Text = ""
txtQtd.Text = ""
txtCodProduto.SetFocus
lblEstoque.Caption = ""
txtDesconto.Text = ""



End Sub

Private Sub CalculaComissaoCombustivel()

Dim QBicosComissao As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbVendasCombustivel As New ADODB.Recordset
Dim dbComissoes As New ADODB.Recordset
Dim dbFuncionarios As New ADODB.Recordset


Dim CodigoFuncionario As Double

If dbFechamentos.Recordset!fechado = True Then Exit Sub

db.Execute "delete from venda2 where combustivel=-1 and fechamentodiario=0 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento


dbProdutos.CursorLocation = adUseClient
If dbProdutos.State = adStateOpen Then
  dbProdutos.Close
End If
dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic


QBicosComissao.CursorLocation = adUseClient
If QBicosComissao.State = adStateOpen Then
  QBicosComissao.Close
End If
QBicosComissao.Open "select bico, codigoproduto, preco, sum(destecaixaqtd) as qtd, sum(destecaixavalor) as total, sum(comissao) as totalcomissao from bicoencerrantes where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento & " group by bico, codigoproduto, preco", db, adOpenKeyset, adLockOptimistic

dbComissoes.CursorLocation = adUseClient
If dbComissoes.State = adStateOpen Then
  dbComissoes.Close
End If
dbComissoes.Open "select *from comissoes where bico<>0 and codigofuncionario<>0 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic

dbFuncionarios.CursorLocation = adUseClient
If dbFuncionarios.State = adStateOpen Then
  dbFuncionarios.Close
End If
dbFuncionarios.Open "select *from vendedores", db, adOpenKeyset, adLockOptimistic


Do While QBicosComissao.EOF = False
  If QBicosComissao!totalcomissao <> 0 Then
    If dbProdutos.RecordCount <> 0 Then
      dbProdutos.MoveFirst
      dbProdutos.Find "codigoproduto=" & QBicosComissao!CodigoProduto
      If dbProdutos.EOF = False Then
        dbVendasCombustivel.CursorLocation = adUseClient
        If dbVendasCombustivel.State = adStateOpen Then
          dbVendasCombustivel.Close
        End If
        dbVendasCombustivel.Open "Select *from venda2 where codigoproduto=" & QBicosComissao!CodigoProduto & " and bico=" & QBicosComissao!Bico & " and codigovendedor=0 and combustivel=-1 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic
        If dbVendasCombustivel.RecordCount = 0 Then
          dbVendasCombustivel.AddNew
        End If
        dbVendasCombustivel!CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
        dbVendasCombustivel!Hora = Now
        dbVendasCombustivel!Data = dbFechamentos.Recordset!DataCaixa
        dbVendasCombustivel!CodigoProduto = QBicosComissao!CodigoProduto
        dbVendasCombustivel!CodProduto = dbProdutos!Codigo
        dbVendasCombustivel!Descri = dbProdutos!Descri
        dbVendasCombustivel!Quantidade = 0
        dbVendasCombustivel!valorUnitario = QBicosComissao!Preco
        dbVendasCombustivel!ValorTotal = 0
        dbVendasCombustivel!ValorComissao = 0
        dbVendasCombustivel!codigovendedor = 0
        dbVendasCombustivel!CodigoPagamento = 0
        dbVendasCombustivel!Combustivel = True
        dbVendasCombustivel!Pago = False
        dbVendasCombustivel!fechamentodiario = False
        dbVendasCombustivel!Quantidade = QBicosComissao!Qtd
        dbVendasCombustivel!ValorTotal = QBicosComissao!Total
        If IsNull(QBicosComissao!totalcomissao) = False Then
            dbVendasCombustivel!ValorComissao = QBicosComissao!totalcomissao
        End If
        dbVendasCombustivel!Bico = QBicosComissao!Bico
        dbVendasCombustivel.Update
        dbVendasCombustivel.Close
      End If
    End If
  End If
  QBicosComissao.MoveNext
Loop


Do While dbComissoes.EOF = False
  dbFuncionarios.Filter = "codigo=" & dbComissoes!Funcionario
  If dbFuncionarios.RecordCount <> 0 Then
    CodigoFuncionario = dbFuncionarios!codigovendedor
  Else
    CodigoFuncionario = 0
  End If
  
  dbVendasCombustivel.CursorLocation = adUseClient
  If dbVendasCombustivel.State = adStateOpen Then
    dbVendasCombustivel.Close
  End If
  dbVendasCombustivel.Open "Select *from venda2 where bico=" & dbComissoes!Bico & " And codigovendedor= 0 And Combustivel = -1 And CodigoFechamento = " & dbFechamentos.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic
  
  dbVendasCombustivel.Filter = "bico=" & dbComissoes!Bico
  If dbVendasCombustivel.RecordCount <> 0 Then
    dbVendasCombustivel!Quantidade = dbVendasCombustivel!Quantidade - dbComissoes!Qtd
    dbVendasCombustivel!ValorTotal = dbVendasCombustivel!ValorTotal - (dbComissoes!Qtd * dbVendasCombustivel!valorUnitario)
    dbVendasCombustivel!ValorComissao = dbVendasCombustivel!ValorComissao - dbComissoes!VlComissao
    dbVendasCombustivel.Update
  End If
  
  If dbProdutos.RecordCount <> 0 Then
    dbProdutos.MoveFirst
    dbProdutos.Find "codigo=" & dbComissoes!Codigo
    If dbProdutos.EOF = False Then
      dbVendasCombustivel.AddNew
      dbVendasCombustivel!CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
      dbVendasCombustivel!Hora = Now
      dbVendasCombustivel!Data = dbFechamentos.Recordset!DataCaixa
      dbVendasCombustivel!CodigoProduto = dbProdutos!CodigoProduto
      dbVendasCombustivel!CodProduto = dbProdutos!Codigo
      dbVendasCombustivel!Descri = dbProdutos!Descri
      dbVendasCombustivel!Quantidade = 0
      dbVendasCombustivel!valorUnitario = dbComissoes!VlUnitario
      dbVendasCombustivel!codigovendedor = dbComissoes!Funcionario
      dbVendasCombustivel!CodigoPagamento = CodigoFuncionario
      dbVendasCombustivel!Combustivel = True
      dbVendasCombustivel!Pago = False
      dbVendasCombustivel!fechamentodiario = False
      dbVendasCombustivel!Quantidade = dbComissoes!Qtd
      dbVendasCombustivel!ValorTotal = dbComissoes!VlTotal
      If IsNull(dbComissoes!VlComissao) = False Then
          dbVendasCombustivel!ValorComissao = dbComissoes!VlComissao
      End If
      dbVendasCombustivel!Bico = dbComissoes!Bico
      dbVendasCombustivel.Update
      
    End If
  End If
  
  dbVendasCombustivel.Close
  
  dbComissoes.MoveNext
Loop

dbProdutos.Close
QBicosComissao.Close
dbComissoes.Close
dbFuncionarios.Close



End Sub

Private Sub CalculaDifComb()

Dim dbProdutos As New ADODB.Recordset
Dim qDifComb As New ADODB.Recordset
Dim dbNotasCorpo As New ADODB.Recordset
Dim dbPostos As New ADODB.Recordset
Dim dbVendeTanque As New ADODB.Recordset


If Frame1.Enabled = True Then
  dbProdutos.CursorLocation = adUseClient
  If dbProdutos.State = adStateOpen Then
    dbProdutos.Close
  End If
  dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenForwardOnly, adLockReadOnly
  If Abrindo = True Then
    qDifComb.CursorLocation = adUseClient
    If qDifComb.State = adStateOpen Then
      qDifComb.Close
    End If
    qDifComb.Open "Select codigoproduto, sum(estoque) as total from tanques group by codigoproduto", db, adOpenForwardOnly, adLockReadOnly
    If IsNull(qDifComb!Total) = False Then
      If dbProdutos.EOF = False Then
        dbProdutos.MoveLast
        dbProdutos.MoveFirst
        Do While dbProdutos.EOF = False
          qDifComb.MoveFirst
          qDifComb.Find "codigoproduto=" & dbProdutos!CodigoProduto
          If qDifComb.EOF = False Then
            B = dbProdutos!Estoque - qDifComb!Total
            If B > 1 Or B < -1 Then
              dbNotasCorpo.CursorLocation = adUseClient
              If dbNotasCorpo.State = adStateOpen Then
                dbNotasCorpo.Close
              End If
              dbNotasCorpo.Open "select sum(quantidade) as total2 from produtosnotascorpo where aguardando=-1 and codigoproduto=" & dbProdutos!CodigoProduto, db, adOpenForwardOnly, adLockReadOnly
              If IsNull(dbNotasCorpo!Total2) = False Then
                B = CLng(dbProdutos!Estoque - dbNotasCorpo!Total2) - CLng(qDifComb!Total)
                If B > 1 Or B < -1 Then
                  'MsgBox "Erro na soma dos tanques de " & dbProdutos.Recordset!Descri & "!"
                  'CorrigeTanque dbProdutos!CodigoProduto
                End If
              Else
                'MsgBox "Erro na soma dos tanques de " & dbProdutos.Recordset!Descri & "!"
                'CorrigeTanque dbProdutos!CodigoProduto
              End If
              dbNotasCorpo.Close
            End If
          End If
          dbProdutos.MoveNext
        Loop
        dbProdutos.MoveFirst
      End If
    End If
    qDifComb.Close
  End If
  
  
  dbTanques2.RecordSource = "Select tanques.*, produtos.descri from tanques, produtos where produtos.codigoproduto=tanques.codigoproduto order by tanque"
  dbTanques2.Refresh
  
  dbPostos.CursorLocation = adUseClient
  If dbPostos.State = adStateOpen Then
    dbPostos.Close
  End If
  dbPostos.Open "Select *from postos", db, adOpenForwardOnly, adLockReadOnly
  
  With dbDifComb
    .Refresh
    If .Recordset.RecordCount <> dbTanques2.Recordset.RecordCount Then
      db.Execute "delete *from diferencacombustivel where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
    End If
    .Refresh
    If .Recordset.RecordCount = 0 Then
      If dbTanques2.Recordset.RecordCount <> 0 Then
        dbTanques2.Recordset.MoveFirst
        Do While dbTanques2.Recordset.EOF = False
          .Recordset.AddNew
          .Recordset!CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
          .Recordset!CodigoProduto = dbTanques2.Recordset("CodigoProduto")
          .Recordset!Descri = dbTanques2.Recordset!Descri
          .Recordset!tanquenr = dbTanques2.Recordset!Tanque
          .Recordset!Estoque = 0
          .Recordset.Update
          
          dbTanques2.Recordset.MoveNext
        Loop
      End If
    End If
    .Refresh
    
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbTanques2.Recordset.MoveFirst
        dbTanques2.Recordset.Find "tanque=" & .Recordset!tanquenr
        If dbTanques2.Recordset.EOF = False Then
          TempValor = 0
          TempValor2 = 0
          TempValor3 = 0
          VendaTanque = 0
          MedeAntes = dbPostos!MedetanqueAntes
          If MedeAntes = True Then
            If IsNull(dbFechamentos.Recordset!Sequencia) = False Then
              dbVendeTanque.CursorLocation = adUseClient
              If dbVendeTanque.State = adStateOpen Then
                dbVendeTanque.Close
              End If
              dbVendeTanque.Open "select sum(destecaixaqtd) as total from qbicoencerrantes where tanque=" & .Recordset!tanquenr & " and sequencia<" & dbFechamentos.Recordset!Sequencia & " and fechado=0", db, adOpenForwardOnly, adLockReadOnly
            Else
              dbVendeTanque.CursorLocation = adUseClient
              If dbVendeTanque.State = adStateOpen Then
                dbVendeTanque.Close
              End If
              dbVendeTanque.Open "select sum(destecaixaqtd) as total from qbicoencerrantes where tanque=" & .Recordset!tanquenr & " and sequencia<0 and fechado=0", db, adOpenForwardOnly, adLockReadOnly
            End If
            If IsNull(dbVendeTanque!Total) = False Then
              TempValor2 = dbVendeTanque!Total
            Else
              TempValor2 = 0
            End If
            dbVendeTanque.Close
          Else
            TempValor2 = 0
            If IsNull(dbFechamentos.Recordset!Sequencia) = False Then
              dbVendeTanque.CursorLocation = adUseClient
              If dbVendeTanque.State = adStateOpen Then
                dbVendeTanque.Close
              End If
              dbVendeTanque.Open "select sum(destecaixaqtd) as total from qbicoencerrantes where tanque=" & .Recordset!tanquenr & " and sequencia<=" & dbFechamentos.Recordset!Sequencia & " and fechado=0", db, adOpenForwardOnly, adLockReadOnly
              If IsNull(dbVendeTanque!Total) = False Then
                TempValor2 = dbVendeTanque!Total
              Else
                TempValor2 = 0
              End If
              dbVendeTanque.Close
            Else
              AtualizaSequenciaCaixa
            End If
          End If
          TempValor3 = 0
          If IsNull(dbFechamentos.Recordset!Sequencia) = False Then
            dbVendeTanque.CursorLocation = adUseClient
            If dbVendeTanque.State = adStateOpen Then
              dbVendeTanque.Close
            End If
            dbVendeTanque.Open "select sum(destecaixaqtd) as total from qbicoencerrantes where tanque=" & .Recordset!tanquenr & " and sequencia=" & dbFechamentos.Recordset!Sequencia
            If IsNull(dbVendeTanque!Total) = False Then
              TempValor3 = dbVendeTanque!Total
            Else
              TempValor3 = 0
            End If
            dbVendeTanque.Close
          End If
        End If
        TempValor = .Recordset!Estoque - (dbTanques2.Recordset!Estoque - TempValor2)
        
        If MedeAntes = True Then
          If .Recordset!Tanque - VendaTanque < 0 Then
            MsgBox "O tanque " & .Recordset!tanquenr & " vai ficar com o estoque negativo!"
            TanqueNegativo = True
          End If
        End If
        On Error Resume Next
        .Recordset!Estoque = dbTanques2.Recordset!Estoque - TempValor2
        If IsNull(.Recordset!Tanque) = True Then
          .Recordset!Tanque = 0
        End If
        If .Recordset!Tanque <> 0 Then
          TempValor = .Recordset!Tanque - .Recordset!Estoque
        Else
          TempValor = 0
        End If
        dbDifComb.Recordset!Diferenca = TempValor
        dbDifComb.Recordset!Vendido = TempValor3
        dbDifComb.Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
  End With
  
    dbProdutos.Close
    dbPostos.Close
    
  
End If


End Sub


Public Function FinalizaNotaDoCaixa() As Boolean

Dim dbNotas As New ADODB.Recordset
Dim dbCaixas As New ADODB.Recordset
Dim Dia As Date, DiaCaixa As Date


FinalizaNotaDoCaixa = False

dbCaixas.CursorLocation = adUseClient
If dbCaixas.State = adStateOpen Then
  dbCaixas.Close
End If
If db.State = adStateClosed Then
  db.Open
End If
dbCaixas.Open "Select *from fechamentodecaixa where fechado=0 order by datacaixa, horaini", db, adOpenForwardOnly, adLockReadOnly

If dbCaixas.RecordCount <> 0 Then
  dbCaixas.MoveFirst
  DiaCaixa = dbCaixas!DataCaixa + dbCaixas!HoraIni
  
  
  dbNotas.CursorLocation = adUseClient
  dbNotas.Open "select produtosnotas.*, turnos.* from produtosnotas inner join turnos on turnos.codigoturno=produtosnotas.codigoturno where datanota<=#" & DataInglesa(dbCaixas!DataCaixa) & "# and confirmado=0 and gravado=-1 order by datanota, horaini", db, adOpenForwardOnly, adLockReadOnly
  If dbNotas.RecordCount <> 0 Then
    Do While dbNotas.EOF = False
      Dia = dbNotas!datanota + dbNotas!HoraIni
      If MedeAntes = False Then
        If Dia <= DiaCaixa Then
          If ConfirmaNota(dbNotas!CodigoEntrada, dbNotas!datanota, dbNotas("turnos.CodigoTurno"), dbNotas!formadepg, dbNotas!NrNota) = False Then
            MsgBox "Existe Problema na confirmação de notas!"
            dbCaixas.Close
            dbNotas.Close
            
            Exit Function
          End If
        End If
      Else
        If Dia < DiaCaixa Then
          If ConfirmaNota(dbNotas!CodigoEntrada, dbNotas!datanota, dbNotas("turnos.CodigoTurno"), dbNotas!formadepg, dbNotas!NrNota) = False Then
            MsgBox "Existe Problema na confirmação de notas!"
            dbCaixas.Close
            dbNotas.Close
            
            Exit Function
          End If
        End If
      End If
      
      dbNotas.MoveNext
    Loop
  End If
  dbNotas.Close
End If

dbCaixas.Close


FinalizaNotaDoCaixa = True

End Function

Public Function ExtornaNotaDoCaixa() As Boolean

Dim dbNotas As New ADODB.Recordset
Dim dbCaixas As New ADODB.Recordset
Dim Dia As Date, DiaCaixa As Date


ExtornaNotaDoCaixa = False

dbCaixas.CursorLocation = adUseClient
dbCaixas.Open "Select *from fechamentodecaixa where fechado=0 order by datacaixa, horaini", db, adOpenForwardOnly, adLockReadOnly

If dbCaixas.RecordCount <> 0 Then
  dbCaixas.MoveFirst
  DiaCaixa = dbCaixas!DataCaixa + dbCaixas!HoraIni
  
  dbNotas.CursorLocation = adUseClient
  dbNotas.Open "select produtosnotas.*, turnos.* from produtosnotas inner join turnos on turnos.codigoturno=produtosnotas.codigoturno where datanota>=#" & DataInglesa(dbCaixas!DataCaixa) & "# and confirmado=-1 and gravado=-1 order by datanota desc, horaini desc", db, adOpenForwardOnly, adLockReadOnly
  If dbNotas.RecordCount <> 0 Then
    Do While dbNotas.EOF = False
      Dia = dbNotas!datanota + dbNotas!HoraIni
      
      If Dia >= DiaCaixa Then
        If DesconfirmaNota(dbNotas!CodigoEntrada, dbNotas!datanota, dbNotas("turnos.CodigoTurno"), dbNotas!formadepg, dbNotas!NrNota) = False Then
          MsgBox "Existe Problema na confirmação de notas!"
          dbCaixas.Close
          dbNotas.Close
          
          Exit Function
        End If
      End If
      
      
      dbNotas.MoveNext
    Loop
  End If
  dbNotas.Close
End If

dbCaixas.Close


ExtornaNotaDoCaixa = True

End Function


Public Function EncontraPdv(ByVal CodigoPdv As Double) As String
  If dbPdvs.Recordset.RecordCount = 0 Then Exit Function
  dbPdvs.Recordset.MoveFirst
  dbPdvs.Recordset.Find "codigopdv=" & CodigoPdv
  If dbPdvs.Recordset.EOF = False Then
    EncontraPdv = dbPdvs.Recordset!Descri
  End If
End Function

Public Sub VerificaConexao()
If db.State = adStateOpen Then
  If db.ConnectionString <> CaminhoADO Then
    db.Close
    db.Open CaminhoADO
  End If
End If
End Sub

Public Sub Importar()
Dim Dia As Date, strEncerrantes As String, intArquivo As Integer
Dim codFuncionario As String, TotalProduto As Currency
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
Dim Documento As String, DataBordero As Date


Dim dbSql As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset
Dim dbVendasLeituraX As New ADODB.Recordset
Dim dbImportacao As New ADODB.Recordset
Dim dbDespesasTipo As New ADODB.Recordset
Dim dbFormaDePg As New ADODB.Recordset
Dim dbClientes As New ADODB.Recordset
Dim dbClientesCarros As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbTotalNotas As New ADODB.Recordset
Dim dbTotalCobranca As New ADODB.Recordset
Dim dbClientesProdutos As New ADODB.Recordset

cmdImportar.Enabled = False
With Animation1
  .Visible = True
  .Open App.Path & "\engrenagem.avi"
  .Play
End With


db.Execute "delete *from importacaoerros where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento

dbDespesasTipo.CursorLocation = adUseClient
dbDespesasTipo.Open "select *from despesatipo", db, adOpenForwardOnly, adLockReadOnly

dbFormaDePg.CursorLocation = adUseClient
dbFormaDePg.Open "select *from formadepagamento", db, adOpenForwardOnly, adLockReadOnly

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


dbSql.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig!ftp
dbImportacao.CursorLocation = adUseClient
On Error Resume Next
  
  dbImportacao.Open "select *from caixas where datacaixa='" & txtData.Value & "' and turno='" & cboTurno.Text & "' and codigoposto='" & dbConfig!Porta & "' and planodeconta='" & dbPdvs.Recordset!Codigo & "' order by linhaexportada", dbSql, adOpenForwardOnly, adLockReadOnly
  
  If Err.Number <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
  End If
  
  On Error GoTo 0
  
  If dbImportacao.RecordCount = 0 Then
    MsgBox "O caixa atual ainda não foi exportado!"
    cmdImportar.Enabled = True
    Animation1.Visible = False
    GoTo Sair
  End If
  dbImportacao.MoveLast
  dbImportacao.MoveFirst
  
  
  
  SoPrimeira = False
  If ApagaRegistros = False Then
    'MsgBox "Este caixa não pode ser importado a segunda parte porque existe registro já gravado!"
    SoPrimeira = True
  End If
  
  DataGrid2.Visible = False
  
  Do While dbImportacao.EOF = False
    StrTemp = dbImportacao!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
      Case "000"
        codFuncionario = Mid(StrTemp, 5, 6)
        dbVendedores.Refresh
        If dbVendedores.Recordset.RecordCount <> 0 Then
          If IsNumeric(codFuncionario) = True Then
            dbVendedores.Recordset.MoveFirst
            dbVendedores.Recordset.Find "codigo=" & CInt(codFuncionario)
            If dbVendedores.Recordset.EOF = False Then
              cboResponsavel.Text = dbVendedores.Recordset!Nome
              Call cboResponsavel_LostFocus
            End If
          End If
        End If
      Case "001"
        'Grava os encerrantes
        SSTab1.Tab = 0
        Bico = CInt(Mid(StrTemp, 5, 6))
        Encerrante = CDbl(Mid(StrTemp, 29, 16))
        If Trim(Mid(StrTemp, 12, 16)) <> "" Then
          Abertura = CDbl(Mid(StrTemp, 12, 16))
        Else
          Abertura = 0
        End If
        
        'Abertura = 0
        If dbPdvs.Recordset.RecordCount > 1 Then
          DesteCaixaQtd = CDbl(Mid(StrTemp, 46, 16))
          DesteCaixaValor = CDbl(Mid(StrTemp, 63, 16))
        End If
        
        With dbEncerrantes
          If .Recordset.RecordCount <> 0 Then
            .Recordset.MoveFirst
            .Recordset.Find "bico=" & Bico
            If .Recordset.EOF = True Then
              'MsgBox "Bico " & Bico & " cadastrado no posto mas não localizado no sistema."
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,bico) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Bico','Bico não cadastrado'," & Bico & ")"
              Encontrou = False
            Else
              Encontrou = True
              Abertura = .Recordset!Abertura
              If Encerrante > 1000000 Then
                If Abertura = 0 Then
                  If Encerrante > 1005000 Then
                    Do While Encerrante > 1000000
                      Encerrante = Encerrante - 1000000
                    Loop
                  End If
                Else
                  If Abertura < Encerrante Then
                    Do While Encerrante > 1000000
                      Encerrante = Encerrante - 1000000
                    Loop
                  End If
                End If
              End If
              
              .Recordset!Encerrante = Encerrante
              If Len(StrTemp) > 47 Then
                .Recordset!DesteCaixaQtd = DesteCaixaQtd
                .Recordset!DesteCaixaValor = DesteCaixaValor
              Else
                .Recordset!DesteCaixaQtd = Encerrante - Abertura
                .Recordset!DesteCaixaValor = .Recordset!DesteCaixaQtd * .Recordset!Preco
              End If
              .Recordset.Update
              'CalculaBicos ColIndex
            End If
          End If
        End With
      Case "002"
        'Grava Venda
        SSTab1.Tab = 1
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
        If Qtd = 0 Then
          GoTo naoIncuirProduto
        End If
        StrTemp2 = Mid(StrTemp, 51, 12)
        If IsNumeric(StrTemp2) = True Then
          TotalProduto = CCur(StrTemp2)
          If Preco * Qtd <> TotalProduto / Qtd Then
            txtDesconto.Text = Format(TotalProduto - (Preco * Qtd), "currency")
          End If
        Else
          TotalProduto = 0
        End If
        If Bico <> 0 Then
          With dbEncerrantes
            .Recordset.MoveFirst
            .Recordset.Find "bico=" & Bico
            If .Recordset.EOF = False Then
              If .Recordset!Preco <> Preco Then
                'MsgBox "O preço da bomba " & Bico & " está cadastrado " & Format(.Recordset!Preco, "#,##0.000") & " mas no posto está " & Format(Preco, "#,##0.000")
                db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,bico,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Bico','Bico com preço errado'," & Bico & "," & NumeroIngles(Preco) & "," & NumeroIngles(.Recordset!Preco) & ")"
              End If
            End If
          End With
        Else
          txtCodProduto.Text = Codigo
          Call txtCodProduto_LostFocus
          txtQtd.Text = Qtd
          Call txtQtd_LostFocus
          
          If Configura.PrecoDiferente = 1 Then
            lblTotalVenda.Caption = Format(TotalProduto, "currency")
          End If
          
          If qProdutosAltera.Recordset.EOF = True Then
            'MsgBox "O produto " & Codigo & " não encontrado na tabela de preços!"
            db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Produto','Produto não cadastrado'," & Codigo & ")"
            GoTo naoIncuirProduto
          Else
            If Configura.PrecoDiferente = 0 Then
              If qProdutosAltera.Recordset.RecordCount <> 0 Then
                On Error Resume Next
                If qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda") <> Preco Then
                  'MsgBox "O produto " & Codigo & " está cadastrado " & Format(qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda"), "#,##0.000") & " mas no posto está " & Format(Preco, "#,##0.000")
                  db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda")) & ")"
                End If
                On Error GoTo 0
              End If
            End If
            
          End If
          If Funcionario <> 0 Then
            txtCodFunc.Text = Funcionario
            If dbVendedores.Recordset.RecordCount <> 0 Then
              dbVendedores.Recordset.MoveFirst
              dbVendedores.Recordset.Find "codigo=" & Funcionario
              If dbVendedores.Recordset.EOF = True Then
                db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,funcionario,qtd) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Funcionario','Funcionário não cadastrado'," & Codigo & "," & Funcionario & "," & NumeroIngles(Qtd) & ")"
                GoTo naoIncuirProduto
              End If
            End If
          Else
            If qProdutosAltera.Recordset!ComissaoValor <> 0 Or qProdutosAltera.Recordset!Comissao <> 0 Then
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigofuncionario,qtd) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Funcionario','Funcionário não informado'," & Codigo & "," & Funcionario & "," & NumeroIngles(Qtd) & ")"
              GoTo naoIncuirProduto
            Else
              txtCodFunc.Text = ""
            End If
          End If
          If qProdutosAltera.Recordset.EOF = True Then
            'MsgBox "O produto " & Codigo & " não encontrado na tabela de preços!"
            db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Produto','Produto não cadastrado'," & Codigo & ")"
            GoTo naoIncuirProduto
          Else
            On Error Resume Next
            If Configura.PrecoDiferente = 0 Then
              If qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda") <> Preco Then
                'MsgBox "O produto " & Codigo & " está cadastrado " & Format(qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda"), "#,##0.000") & " mas no posto está " & Format(Preco, "#,##0.000")
                db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera.Recordset("produtosalteradetalhe.PrecoVenda")) & ")"
              End If
            End If
            If Configura.PrecoDiferente = 0 Then
              IncluirProduto
            Else
              IncluirProduto Preco, TotalProduto
            End If
          End If
        End If
naoIncuirProduto:
      Case "003"
        'notas de clientes
        If SoPrimeira = False Then
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
                'LucroDif = Mid(StrTemp, 131, 15)
                If IsNumeric(Mid(StrTemp, 147, 15)) = True Then
                  valorUnitario = Mid(StrTemp, 147, 15)
                End If
                If IsNumeric(Mid(StrTemp, 163, 15)) = False Then
                  ValorTotal = valorUnitario * Qtd
                Else
                  ValorTotal = Mid(StrTemp, 163, 15)
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
            Autorizar = False
            Autorizado = False
            Motivo = ""
            LucroDif = ValorTotal - ValorTotalDif
            If IsNumeric(Cupom) = False Then
              Cupom = 0
            End If
            dbClientes.MoveFirst
            dbClientes.Find "codigonoposto=" & CodigoCliente
            If dbClientes.EOF = True Then
              'MsgBox "Código de cliente de nota " & CodigoCliente & " não encontrado!"
              'GravaBloqueado CodigoCliente, "Não encontrado", Cupom, ValorTotal, "Cliente não localizado"
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente não cadastrado'," & CodigoCliente & ")"
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
                  db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigonoposto) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cupom " & Cupom & " com produto não cadastrado'," & CodigoCliente & "," & CodigoProduto & ")"
                  GoTo SairDoCliente
                Else
                  If dbProdutos!Combustivel = True Then
                    dbEncerrantes.Recordset.MoveFirst
                    dbEncerrantes.Recordset.Find "codigoproduto=" & dbProdutos!CodigoProduto
                    If dbEncerrantes.Recordset.EOF = False And dbEncerrantes.Recordset.BOF = False Then
                      Preco = PrecoAtual(dbProdutos!CodigoProduto, dbFechamentos.Recordset!DataCaixa, dbFechamentos.Recordset!CodigoTurno, dbEncerrantes.Recordset!Bico)
                    Else
                      MsgBox CodigoProduto & " - " & dbProdutos!Descri & " - Não encontrado bico para este produto, vendido em nota de cliente"
                    End If
                  Else
                    Preco = PrecoAtual(dbProdutos!CodigoProduto, dbFechamentos.Recordset!DataCaixa, dbFechamentos.Recordset!CodigoTurno)
                  End If
                End If
                If dbClientes!mensalista = False Then
                  If dbClientes!desativado < dbFechamentos.Recordset!DataCaixa Then
                    If Usuarios.Grupo.admDatas < 2 Then
                      'MsgBox "O cliente " & DbClientes!Nome & " está desativado!"
                      If Configura.NotaBloqueia = 0 Then
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
                        Autorizar = True
                        Motivo = "Desativado"
                      End If
                    Else
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está desativado! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
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
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(ValorTotal) & ")"
                        Autorizar = True
                        Motivo = "Limite"
                      Else
                        'Resposta = MsgBox("O cliente " & DbClientes!Nome & " ultrapassará o limite dele! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
                        'If Resposta = vbNo Then GoTo SairDoCliente
                        Autorizar = True
                        Autorizado = False
                        Motivo = "Ultrapassou Limite"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(ValorTotal) & ")"
                      End If
                    End If
                  Else
                    'MsgBox "O cliente " & DbClientes!Nome & " esta marcado para ser limitado mas não possue valor definido!"
                    'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Marcado para limitar mas não possue valor a ser limitado"
                    Autorizar = True
                    Motivo = "Sem Limite"
                    db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente marcado para limitar mas sem limite cadastrado'," & CodigoCliente & "," & dbClientes!CodigoCliente & ")"
                  End If
                End If
                If dbClientes!diapagamento <> 0 Then
                  If dbClientes!diapagamento >= 28 Then
                    DataPrevista = CDate(Format(UltimoDiaDoMes(Month(dbFechamentos.Recordset!DataCaixa), Year(dbFechamentos.Recordset!DataCaixa)), "00") & "/" & Month(dbFechamentos.Recordset!DataCaixa) & "/" & Year(dbFechamentos.Recordset!DataCaixa))
                  Else
                    DataPrevista = CDate(Format(dbClientes!diapagamento, "00") & "/" & Month(dbFechamentos.Recordset!DataCaixa) & "/" & Year(dbFechamentos.Recordset!DataCaixa))
                  End If
                Else
                  DataPrevista = DateAdd("m", 1, dbFechamentos.Recordset!DataCaixa)
                End If
                If DataPrevista < dbFechamentos.Recordset!DataCaixa Then
                  DataPrevista = DateAdd("m", 1, DataPrevista)
                End If
                dbClientesProdutos.Filter = ""
                If dbClientesProdutos.RecordCount <> 0 Then
                  dbClientesProdutos.MoveFirst
                  dbClientesProdutos.Filter = "codigocliente=" & dbClientes!CodigoCliente & " and codproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(txtData.Value) & "#"
                  If dbClientesProdutos.EOF = False Then
                    If dbClientesProdutos!validade = txtData.Value Then
                      If dbClientesProdutos!HoraIni >= dbFechamentos.Recordset!HoraIni Then
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
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal) & "," & NumeroIngles(TempValorPagar) & ")"
                      Else
                        'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                        'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
                        'If Resposta = vbNo Then GoTo SairDoCliente
                        Autorizar = True
                        Autorizado = False
                        Motivo = "Preço Diferenciado"
                        db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal) & "," & NumeroIngles(TempValorPagar) & ")"
                      End If
                    End If
                  Else
                    'ValorUnitarioDif = Qtd * valorUnitario
                    TempDif = (ValorUnitarioDif * Qtd) - ValorTotal
                    If TempDif > 0.01 Or TempDif < -0.01 Then
                      'MsgBox "Preço unitário incorreto!"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(ValorUnitarioDif * Qtd) & ")"
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
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal / Qtd) & "," & NumeroIngles(Preco) & ")"
                    Else
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
                      'If Resposta = vbNo Then GoTo SairDoCliente
                      Autorizar = True
                      Autorizado = False
                      Motivo = "Preço incorreto!"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & dbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal / Qtd) & "," & NumeroIngles(Preco) & ")"
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
          
          StrTemp = StrTemp & dbFechamentos.Recordset!CodigoFechamento & "," & dbClientes!CodigoCliente & ",'" & dbClientes!Nome & "',#" & DataInglesa(Date) & " " & Time & "#,#" & DataInglesa(DataPrevista) & "#," & NumeroIngles(ValorTotal) & ",#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "#,"
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
          StrTemp2 = NumeroIngles(Qtd) & "," & NumeroIngles(Consumo) & "," & CodigoProduto & "," & NumeroIngles(valorUnitario) & "," & NumeroIngles(Qtd) & "," & NumeroIngles(ValorUnitarioDif) & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(LucroDif) & "," & CInt(Autorizar) & "," & CInt(Autorizado) & ",'" & Motivo & "')"
          'StrTemp = StrTemp & StrTemp2
          
          db.Execute StrTemp & StrTemp2
        
          If IsNull(dbClientes!UltimoAbastecimento) = True Then
            dbClientes!UltimoAbastecimento = dbFechamentos.Recordset!DataCaixa
          End If
          If dbClientes!UltimoAbastecimento < dbFechamentos.Recordset!DataCaixa Then
            dbClientes!UltimoAbastecimento = dbFechamentos.Recordset!DataCaixa
          End If
          db.Execute "update clientes set TotalNotas=TotalNotas+" & NumeroIngles(ValorTotal) & " where codigocliente=" & CodigoCliente
          db.Execute "update clientes set saldo=limite-totalnotas-totalboleto where codigocliente=" & CodigoCliente
        End If
SairDoCliente:
        
      Case "004"
        'grava estoque dos tanques
        SSTab1.Tab = 2
        Tanque = Mid(StrTemp, 5, 5)
        StrTemp2 = Mid(StrTemp, 11)
        For i = 1 To Len(StrTemp2)
          If Mid(StrTemp2, i, 1) <> 0 Then
            StrTemp2 = Mid(StrTemp2, i)
            Exit For
          End If
        Next i
        If IsNumeric(StrTemp2) = True Then
          Estoque = CDbl(StrTemp2)
        Else
          Estoque = 0
        End If
        
        With dbDifComb
          .Refresh
          If .Recordset.RecordCount <> 0 Then
            .Recordset.MoveFirst
            .Recordset.Find "tanquenr=" & Tanque
            If .Recordset.EOF = False Then
              .Recordset!Tanque = Estoque
              .Recordset.Update
            End If
          End If
        End With
      Case "005"
        'forma de pagamento recebido
        If SoPrimeira = False Then
          If dbFormaDePg.RecordCount <> 0 Then
            On Error GoTo semRecebido
            StrTemp2 = Trim(Mid(StrTemp, 5, 15))
            If IsNumeric(Trim(Trim(Mid(StrTemp, 5, 15)))) = True Then
              Codigo = CDbl(StrTemp2)
            Else
              Codigo = 0
            End If
            Valor = CCur(Mid(StrTemp, 37))
            Documento = Trim(Mid(StrTemp, 21, 15))
            DataBordero = DataInglesa(dbFechamentos.Recordset!DataCaixa)
            
            If Documento <> "" Then
              If Mid(Documento, Len(Documento) - 1, 1) = "/" Then
                DataBordero = DateAdd("m", CDbl(Mid(Documento, Len(Documento), 1)) - 1, DataBordero)
              End If
            End If
            
            dbFormaDePg.MoveFirst
            dbFormaDePg.Find "codigonoposto='" & Trim(Codigo) & "'"
            If dbFormaDePg.EOF = False Then
              Tarifa = dbFormaDePg!descontovalor
              Operacao = dbFormaDePg!descontoporoperacao
              Porcento = dbFormaDePg!DescontoPorcento / 100
              
              ValorBruto = Valor
              DescontoPorcento = 0
              
              If Porcento <> 0 Then
                DescontoPorcento = ValorBruto * Porcento
              End If
              
              Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao
              
              If dbFormaDePg!CodigoConta = 0 Then
                MsgBox "A forma de pagamento " & dbFormaDePg!Descri & " está sem conta destino!"
              Else
                
                db.Execute "insert into formadepagamentorecebido2 (codigofechamento,codigoformadepg,descri,valorbruto,valordescoper,valordesctarifa,valordesconto,valor,operacoes,data,hora) values (" & dbFechamentos.Recordset!CodigoFechamento & "," & dbFormaDePg!CodigoPagamento & ",'" & dbFormaDePg!Descri & "'," & NumeroIngles(ValorBruto) & "," & NumeroIngles(Operacao) & "," & NumeroIngles(Tarifa) & "," & NumeroIngles(DescontoPorcento) & "," & NumeroIngles(Liquido) & "," & TotalOper & ",#" & DataBordero & "#,#" & Now & "#)"
              End If
            End If
          End If
        End If
semRecebido:
        
      Case "006"
        'despesas
        If SoPrimeira = False Then
          If dbDespesasTipo.RecordCount <> 0 Then
            Codigo = Trim(Mid(StrTemp, 5, 15))
            Descri = Trim(Mid(StrTemp, 21, 50))
            Tipo = Trim(Mid(StrTemp, 72, 5))
            Valor = CCur(Mid(StrTemp, 78))
            
            If Tipo = "PAG" Then
              Valor = Valor * -1
            End If
            dbDespesasTipo.MoveFirst
            dbDespesasTipo.Find "codigonoposto='" & Codigo & "'"
            If dbDespesasTipo.EOF = False Then
              db.Execute "insert into despesaslanc2 (codigofechamento,origem,data,vencimento,hora,codigoconta,conta,codigodespesa,descri,obs,compensado,valor,valorpago) values (" & dbFechamentos.Recordset!CodigoFechamento & ",'Fechamento',#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "#,#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "#,#" & Now & "#,-1,'Fechamento de Caixa'," & dbDespesasTipo("codigodespesa") & ",'" & dbDespesasTipo("descri") & "','" & Descri & "',-1," & NumeroIngles(Valor) & "," & NumeroIngles(Valor) & ")"
            End If
          End If
        End If
      Case "007"
        GravaCupons2 StrTemp
      Case "008"
        GravaComissoes StrTemp, dbFechamentos.Recordset!CodigoFechamento
      Case "998"
        'GravaResultado StrTemp
        
        '998|     2100000000|1,54
        
        CodigoConta = Trim(Mid(StrTemp, 5, 15))
        Valor = CCur(Trim(Mid(StrTemp, 21)))
        
        db.Execute "insert into fechamentodecaixapista (codigofechamento,codigoconta,valor) values (" & dbFechamentos.Recordset!CodigoFechamento & "," & CodigoConta & "," & NumeroIngles(Valor) & ")"
        
    End Select
    dbImportacao.MoveNext
  Loop

Sair:

db.Execute "update importacaoerros set dataimportado=#" & DataInglesa(Date) & " " & Format(Time, "short time") & "# where dataimportado is null"

dbConfig.Close
'dbVendasLeituraX.Close
dbDespesasTipo.Close
dbFormaDePg.Close
'DbClientes.Close
dbClientesCarros.Close
dbProdutos.Close
dbTotalNotas.Close
dbTotalCobranca.Close
dbClientesProdutos.Close
dbImportacao.Close
dbSql.Close

DataGrid2.Visible = True


Animation1.Close
Animation1.Visible = False
cmdImportar.Enabled = True

SSTab1.Tab = 0
Call cmdAbrir_Click

End Sub

Public Sub AbreCaixa()

Dim dbFechamento2 As New ADODB.Recordset
Dim dbBicoEncerrantes2 As New ADODB.Recordset
Dim dbAlteracao As New ADODB.Recordset
Dim dbBicos As New ADODB.Recordset
Dim dbAlteraBico As New ADODB.Recordset
Dim dbEncerrantesNovos As New ADODB.Recordset
Dim dbPdvsTurnos As New ADODB.Recordset

Dim CodigoFechamento As Double, Abertura As Double
Dim CaixaAnterior As Double, AlteraPreco As Double
Dim UltimoEstacionamento As Double, AnteriorFechado As Boolean
Dim CaixaFechado As Boolean
Dim DataDoCaixa As String
Dim HoraIni As Date

Abrindo = True
ErroNaSoma = False
If DateDiff("d", Date, txtData.Value) >= 1 Then
  Resposta = MsgBox("Deseja criar um caixa futuro?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then
    
    Exit Sub
  End If
End If

With dbFechamentos
  .RecordSource = "select *from fechamentodecaixa order by datacaixa, horaini, codigopdv"
  .Refresh
  If dbPdvs.Recordset.RecordCount = 0 Then
    MsgBox "Ponto de venda não localizado!"
    
    Exit Sub
  End If
  If dbPdvs.Recordset.RecordCount <> 0 Then
    If cboPdvs.Text = "" Then
      dbPdvs.Recordset.MoveFirst
      cboPdvs.Text = dbPdvs.Recordset!Descri
    End If
  Else
    MsgBox "Sem pdv cadastrado"
    Exit Sub
  End If
  If cboPdvs.Text <> dbPdvs.Recordset!Descri Then
    Call cboPdvs_LostFocus
    If cboPdvs.Text <> dbPdvs.Recordset!Descri Then
      MsgBox "Ponto de venda não localizado!"
      cboPdvs.SetFocus
      
      Exit Sub
    End If
  End If
  If dbTurnos.Recordset.RecordCount = 0 Then
    MsgBox "Turno não encontrado!"
    
    Exit Sub
  End If
  If dbTurnos.Recordset.EOF = True Then
    MsgBox "Truno não encontrado!"
    
    Exit Sub
  End If
  If cboTurno.Text <> dbTurnos.Recordset!Descri Then
    Call cboTurno_LostFocus
    If dbTurnos.Recordset.EOF = True Then
      MsgBox "Erro na tabela de turnos!"
      Exit Sub
    End If
    If cboTurno.Text <> dbTurnos.Recordset!Descri Then
      MsgBox "Turno não encontrado!"
      On Error Resume Next
      cboTurno.SetFocus
      
      Exit Sub
    End If
  End If
  HoraIni = dbTurnos.Recordset!HoraIni
  Intermitente = dbPdvs.Recordset!Intermitente
  If Intermitente = True Then
    dbPdvsTurnos.CursorLocation = adUseClient
    dbPdvsTurnos.Open "select horaini from pdvsturnos where codigopdv=" & dbPdvs.Recordset!CodigoPdv & " and codigoturno=" & dbTurnos.Recordset!CodigoTurno, db, adOpenForwardOnly, adLockReadOnly
    If dbPdvsTurnos.RecordCount = 0 Then
      MsgBox "Este PDV nescessita de um cadastro de PDVs/Turnos!"
      dbPdvsTurnos.Close
      
      Exit Sub
    End If
    
    HoraIni = dbPdvsTurnos!HoraIni
    
    dbPdvsTurnos.Close
  End If
  If .Recordset.RecordCount <> 0 Then
    .Recordset.Filter = "datacaixa=#" & txtData.Value & "# and codigoturno=" & dbTurnos.Recordset!CodigoTurno & " and codigopdv=" & dbPdvs.Recordset!CodigoPdv
    If .Recordset.RecordCount = 0 Then
      'verifica se não existe caixa finalizado posterior
      .Recordset.Filter = ""
      .Recordset.MoveFirst
      .Recordset.Filter = "datacaixa>=#" & txtData.Value & "# and fechado=-1"
      If .Recordset.EOF = False Then
        If .Recordset!DataCaixa = txtData.Value Then
          If .Recordset!HoraIni >= HoraIni Then
            MsgBox "Já existe caixa posterior finalizado!"
            Exit Sub
          Else
            .Recordset.Filter = ""
            .Recordset.MoveFirst
            .Recordset.Filter = "datacaixa>#" & txtData.Value & "# and fechado=-1"
            If .Recordset.RecordCount <> 0 Then
              MsgBox "Já existe caixa posterior finalizado!"
              Exit Sub
            End If
          End If
        Else
          MsgBox "Já existe caixa posterior finalizado!"
          Exit Sub
        End If
      End If
      .Recordset.Filter = ""
      .Recordset.AddNew
      .Recordset!DataCaixa = txtData.Value
      .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
      .Recordset!Turno = dbTurnos.Recordset!Descri
      .Recordset!HoraIni = HoraIni
      .Recordset!horafim = dbTurnos.Recordset!horafim
      .Recordset!CodigoPdv = dbPdvs.Recordset!CodigoPdv
      .Recordset.Update
    End If
  Else
    .Recordset.AddNew
    .Recordset!DataCaixa = txtData.Value
    .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
    .Recordset!Turno = dbTurnos.Recordset!Descri
    .Recordset!HoraIni = HoraIni
    .Recordset!horafim = dbTurnos.Recordset!horafim
    .Recordset!CodigoPdv = dbPdvs.Recordset!CodigoPdv
    .Recordset.Update
  End If
  .Refresh
  .Recordset.Filter = "datacaixa=#" & txtData.Value & "# and codigoturno=" & dbTurnos.Recordset!CodigoTurno & " and codigopdv=" & dbPdvs.Recordset!CodigoPdv
  If .Recordset.EOF = True Then
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
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    DataGrid2.AllowDelete = False
    DataGrid2.AllowUpdate = False
    DataGrid3.AllowDelete = False
    DataGrid3.AllowUpdate = False
    cmdImportar.Visible = False
    cmdEntraCombustivel.Visible = False
    cmdExtornaTurno.Visible = False
  Else
    CaixaFechado = False
    Frame1.Enabled = True
    DataGrid1.AllowUpdate = True
    DataGrid2.AllowUpdate = True
    DataGrid3.AllowDelete = True
    DataGrid3.AllowUpdate = True
    cmdImportar.Visible = True
    cmdEntraCombustivel.Visible = True
    cmdExtornaTurno.Visible = True
  End If
  If IsNumeric(.Recordset!TotalCombustivel) = True Then
    lblTotalCombustivel.Caption = Format(.Recordset!TotalCombustivel, "Currency")
  End If
  If IsNumeric(.Recordset!TotalProdutos) = True Then
    lblTotalProdutos.Caption = Format(.Recordset!TotalProdutos, "Currency")
  End If
End With

'*****************************************************************************************
'*****************************************************************************************
dbFechamento2.CursorLocation = adUseClient
dbFechamento2.Open "select *from fechamentodecaixa order by datacaixa, horaini, codigopdv", db, adOpenForwardOnly, adLockReadOnly

dbBicoEncerrantes2.CursorLocation = adUseClient

FechamentoAnterior = -1
If dbFechamento2.RecordCount <> 0 Then
  dbFechamento2.MoveFirst
  dbFechamento2.Find "codigofechamento=" & CodigoFechamento
  If dbFechamento2.EOF = False Then
    DataCaixa = dbFechamento2!DataCaixa
    Do
    dbFechamento2.MovePrevious
    If dbFechamento2.BOF = False Then
      If dbFechamento2!DataCaixa >= DataCaixa Then
        If dbFechamento2!HoraIni <> HoraIni Then
          FechamentoAnterior = dbFechamento2!CodigoFechamento
          AnteriorFechado = dbFechamento2!fechado
          dbBicoEncerrantes2.Open "select *from BicoEncerrantes where codigofechamento=" & FechamentoAnterior & " order by bico", db, adOpenForwardOnly, adLockReadOnly
          Exit Do
        End If
      Else
        FechamentoAnterior = dbFechamento2!CodigoFechamento
        AnteriorFechado = dbFechamento2!fechado
        dbBicoEncerrantes2.Open "select *from BicoEncerrantes where codigofechamento=" & FechamentoAnterior & " order by bico", db, adOpenForwardOnly, adLockReadOnly
        Exit Do
      End If
    End If
    Loop Until dbFechamento2.BOF = True
  End If
End If

dbFechamento2.Close

AlteraPreco = -1

If CaixaFechado = False Then
  dbAlteracao.CursorLocation = adUseClient
  dbAlteracao.Open "select alteracoes.*, turnos.* from alteracoes, turnos where turnos.codigoturno=alteracoes.codigoturno order by dataalteracao, horaini", db, adOpenKeyset, adLockReadOnly
  If dbAlteracao.RecordCount <> 0 Then
    dbAlteracao.MoveLast
    Do While dbAlteracao.BOF = False
      If dbAlteracao!dataalteracao <= dbFechamentos.Recordset!DataCaixa Then
        If dbAlteracao!dataalteracao < dbFechamentos.Recordset!DataCaixa Then
          AlteraPreco = dbAlteracao!codalteracao
          Exit Do
        Else
          If dbAlteracao!HoraIni <= dbFechamentos.Recordset!HoraIni Then
            AlteraPreco = dbAlteracao!codalteracao
            Exit Do
          Else
            AlteraPreco = -1
          End If
        End If
      End If
      dbAlteracao.MovePrevious
    Loop
  End If
  dbAlteracao.Close
End If

If AlteraPreco >= 0 Then
  dbAlteraBico.CursorLocation = adUseClient
  dbAlteraBico.Open "select *from alterabico where codalteracao=" & AlteraPreco & " order by bico", db, adOpenForwardOnly, adLockReadOnly
End If

With dbEncerrantes
  .RecordSource = "select *from BicoEncerrantes where codigofechamento=" & CodigoFechamento & " order by bico"
  .Refresh
  If CaixaFechado = False Then
    dbBicos.CursorLocation = adUseClient
    dbBicos.Open "select *from bicos order by bico", db, adOpenForwardOnly, adLockReadOnly
    If dbBicos.RecordCount <> 0 Then
        dbBicos.MoveLast
        dbBicos.MoveFirst
        If .Recordset.RecordCount <> 0 And .Recordset.RecordCount < dbBicos.RecordCount Then
            Do While dbBicos.EOF = False
                .Recordset.MoveFirst
                .Recordset.Find "bico=" & dbBicos!Bico
                If .Recordset.EOF = True Then
                    .Recordset.AddNew
                    .Recordset!CodigoFechamento = CodigoFechamento
                    .Recordset!Bico = dbBicos!Bico
                    If FechamentoAnterior >= 0 And AnteriorFechado = False Then
                      dbBicoEncerrantes2.Find "bico=" & dbBicos!Bico
                      If dbBicoEncerrantes2.EOF = False Then
                        Abertura = dbBicoEncerrantes2!Encerrante
                      End If
                    Else
                      Abertura = dbBicos!ultimonumero
                    End If
                    
                    If Abertura > 1000000 Then
                      Do While Abertura > 1000000
                        Abertura = Abertura - 1000000
                      Loop
                    End If
                    .Recordset!Abertura = Abertura
                    
                    If AlteraPreco >= 0 Then
                      dbAlteraBico.MoveFirst
                      dbAlteraBico.Find "bico=" & .Recordset!Bico
                      If .Recordset.EOF = False Then
                        If IsNull(dbAlteraBico!Preco) = True Then
                            .Recordset!Preco = dbAlteraBico!Preco
                        Else
                            .Recordset!Preco = dbBicos!PrecoVenda
                        End If
                      Else
                        .Recordset!Preco = dbBicos!PrecoVenda
                      End If
                    Else
                      .Recordset!Preco = dbBicos!PrecoVenda
                    End If
                    .Recordset!CodigoProduto = dbBicos!CodigoProduto
                    .Recordset!Tanque = dbBicos!Tanque
                    If IsNull(.Recordset!Encerrante) = True Then
                      .Recordset!Encerrante = 0
                    End If
                    If Intermitente = True Then
                      PegaEncerranteIntermitente .Recordset
                      DataGrid1.Columns(2).Locked = True
                      DataGrid1.Columns(3).Locked = True
                      DataGrid1.Columns(4).Locked = True
                    Else
                      DataGrid1.Columns(2).Locked = False
                      DataGrid1.Columns(3).Locked = False
                      DataGrid1.Columns(4).Locked = False
                    End If
                    If .Recordset!Encerrante = 0 Then .Recordset!Encerrante = Abertura
                    .Recordset!DesteCaixaQtd = 0
                    .Recordset!DesteCaixaValor = 0
                    .Recordset!deoutrocaixaqtd = 0
                    .Recordset!deoutrocaixavalor = 0
                    .Recordset.Update
                End If
                dbBicos.MoveNext
            Loop
        End If
            
    End If
    
    If .Recordset.RecordCount = 0 Then
      If dbBicos.RecordCount <> 0 Then
        dbBicos.MoveLast
        dbBicos.MoveFirst
        Do While dbBicos.EOF = False
          .Recordset.AddNew
          .Recordset!CodigoFechamento = CodigoFechamento
          .Recordset!Bico = dbBicos!Bico
          If FechamentoAnterior >= 0 And AnteriorFechado = False Then
            dbBicoEncerrantes2.Find "bico=" & dbBicos!Bico
            If dbBicoEncerrantes2.EOF = False Then
              Abertura = dbBicoEncerrantes2!Encerrante
            End If
          Else
            Abertura = dbBicos!ultimonumero
          End If
          
          If Abertura > 1000000 Then
            Do While Abertura > 1000000
              Abertura = Abertura - 1000000
            Loop
          End If
          .Recordset!Abertura = Abertura
          
          If AlteraPreco >= 0 Then
            dbAlteraBico.MoveFirst
            dbAlteraBico.Find "bico=" & .Recordset!Bico
            If .Recordset.EOF = False Then
              .Recordset!Preco = dbAlteraBico!Preco
            Else
              .Recordset!Preco = dbBicos!PrecoVenda
            End If
          Else
            .Recordset!Preco = dbBicos!PrecoVenda
          End If
          .Recordset!CodigoProduto = dbBicos!CodigoProduto
          .Recordset!Tanque = dbBicos!Tanque
          If IsNull(.Recordset!Encerrante) = True Then
            .Recordset!Encerrante = 0
          End If
          If Intermitente = True Then
            PegaEncerranteIntermitente .Recordset
            DataGrid1.Columns(2).Locked = True
            DataGrid1.Columns(3).Locked = True
            DataGrid1.Columns(4).Locked = True
          Else
            DataGrid1.Columns(2).Locked = False
            DataGrid1.Columns(3).Locked = False
            DataGrid1.Columns(4).Locked = False
          End If
          If .Recordset!Encerrante = 0 Then .Recordset!Encerrante = Abertura
          .Recordset!DesteCaixaQtd = 0
          .Recordset!DesteCaixaValor = 0
          .Recordset!deoutrocaixaqtd = 0
          .Recordset!deoutrocaixavalor = 0
          .Recordset.Update
          dbBicos.MoveNext
        Loop
        .Recordset.MoveFirst
      End If
      
    Else
      If Intermitente = True Then
        PegaEncerranteIntermitenteTodos .Recordset
        .Recordset.MoveFirst
        DataGrid1.Columns(2).Locked = True
        DataGrid1.Columns(3).Locked = True
        DataGrid1.Columns(4).Locked = True
      Else
        DataGrid1.Columns(2).Locked = False
        DataGrid1.Columns(3).Locked = False
        DataGrid1.Columns(4).Locked = False
      End If
      If AlteraPreco >= 0 Then
        .Refresh
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          dbAlteraBico.MoveFirst
          dbAlteraBico.Find "bico=" & .Recordset!Bico
          If .Recordset.EOF = False Then
            On Error Resume Next
            .Recordset!Preco = dbAlteraBico!Preco
            On Error GoTo 0
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
If AlteraPreco >= 0 Then
  dbAlteraBico.Close
End If

With dbVendas
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from venda2 where combustivel=0 and codigofechamento=" & CodigoFechamento & " order by codproduto"
  .Refresh
End With

With dbDifComb
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from diferencacombustivel where codigofechamento=" & CodigoFechamento
  .Refresh
End With

With dbErros
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from importacaoerros where codigofechamento=" & CodigoFechamento
  .Refresh
End With

If CaixaFechado = False Then
  If FechamentoAnterior > 0 Then
    dbEncerrantesNovos.CursorLocation = adUseClient
    dbEncerrantesNovos.Open "select *from bicosencerrantesnovo where datacaixa=#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "# and horaini=#" & dbFechamentos.Recordset!HoraIni & "#", db, adOpenForwardOnly, adLockReadOnly
    
    If dbBicoEncerrantes2.RecordCount <> 0 Then
      dbBicoEncerrantes2.MoveLast
      dbBicoEncerrantes2.MoveFirst
      Do While dbBicoEncerrantes2.EOF = False
        dbEncerrantes.Recordset.MoveFirst
        dbEncerrantes.Recordset.Find "bico=" & dbBicoEncerrantes2!Bico
        If dbEncerrantes.Recordset.EOF = False Then
          If Intermitente = False Then
            If dbEncerrantesNovos.RecordCount <> 0 Then
                dbEncerrantesNovos.MoveFirst
                dbEncerrantesNovos.Find "Bico=" & dbEncerrantes.Recordset!Bico
                If dbEncerrantesNovos.EOF = False Then
                    dbEncerrantes.Recordset!Abertura = dbEncerrantesNovos!inicial
                Else
                    dbEncerrantes.Recordset!Abertura = dbBicoEncerrantes2!Encerrante
                End If
            Else
                dbEncerrantes.Recordset!Abertura = dbBicoEncerrantes2!Encerrante
                If dbBicoEncerrantes2!Encerrante > 1000000 Then
                    Do While dbEncerrantes.Recordset!Abertura > 1000000
                        dbEncerrantes.Recordset!Abertura = dbEncerrantes.Recordset!Abertura - 1000000
                    Loop
                Else
                    dbEncerrantes.Recordset!Abertura = dbBicoEncerrantes2!Encerrante
                End If
            End If
            If dbEncerrantes.Recordset!Encerrante < dbEncerrantes.Recordset!Abertura Then dbEncerrantes.Recordset!Encerrante = dbEncerrantes.Recordset!Abertura
            If dbEncerrantes.Recordset!Abertura > 1000000 Then
              Do While dbEncerrantes.Recordset!Abertura > 1000000
                dbEncerrantes.Recordset!Abertura = dbEncerrantes.Recordset!Abertura - 1000000
              Loop
            End If
            dbEncerrantes.Recordset.Update
          End If
        End If
        dbBicoEncerrantes2.MoveNext
      Loop
    End If
    dbBicoEncerrantes2.Close
  End If
End If

DataDoCaixa = dbFechamentos.Recordset!DataCaixa & " " & dbFechamentos.Recordset!HoraIni

If SemTabelaDePrecos = False Then
  With qProdutosAltera
    .RecordSource = "select CodigoProdutoAltera, (datacaixa+horaini) as Data from produtosaltera group by CodigoProdutoAltera, (datacaixa+horaini) order by (datacaixa+horaini) desc"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      .Recordset.Find "data<=#" & DataDoCaixa & "#"
      If .Recordset.EOF = True Then
        AlteraAnterior = 0
      Else
        AlteraAnterior = .Recordset!codigoprodutoaltera
      End If
    Else
      AlteraAnterior = 0
    End If
    .RecordSource = "select produtosalteradetalhe.*, produtos.* from produtosalteradetalhe, produtos where codigoprodutoaltera=" & AlteraAnterior & " and produtosalteradetalhe.codigoproduto=produtos.codigoproduto order by produtos.codigo"
    .Refresh
  End With
End If
Call cmdCalcular_Click

If IsNull(dbFechamentos.Recordset!responsavel) = False Then
  cboResponsavel.Text = dbFechamentos.Recordset!responsavel
End If
dbEncerrantes.Refresh
dbVendas.Refresh
dbDifComb.Refresh

cmdRemover.Enabled = True
Frame1.Visible = True
SSTab1.Visible = True
cboResponsavel.Visible = True
cboPdvs.Enabled = False
DataGrid1.Visible = True
DataGrid2.Visible = True
DataGrid3.Visible = True

If Frame1.Enabled = True Then
  cboResponsavel.SetFocus
End If

If dbFechamentos.Recordset!fechado = False Then
  cmdConfirmar.Visible = True
  cmdCalcular.Visible = True
Else
  cmdConfirmar.Visible = False
  cmdCalcular.Visible = False
End If


DataGrid1.Columns(4).Locked = False
If Usuarios.Grupo.AdmEstatus = 2 And Intermitente = False Then
  DataGrid1.Columns(4).Locked = False
End If

With dbFechamentos
  If IsNumeric(lblTotalCombustivel.Caption) = True Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
  End If
  If IsNumeric(lblTotalProdutos.Caption) = True Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
  End If
  .Recordset.Update
End With


Abrindo = False

If SSTab1.TabVisible(0) = True Then
  SSTab1.Tab = 0
Else
  SSTab1.Tab = 1
End If
End Sub

Private Function Desconfirmar() As Boolean

Dim dbProdutos As New ADODB.Recordset
Dim dbBicos As New ADODB.Recordset
Dim dbDespesasLanc2 As New ADODB.Recordset
Dim dbStatus As New ADODB.Recordset

Dim Resposta As Integer, LucroVenda As Currency, StrTemp As String
Dim Estacionamento As Currency, ValorEstoque As Currency
Dim Vendas As Double, LucroMedio As Currency, PrecoMedio As Currency

Desconfirmar = False
Call cmdAbrir_Click

dbTanques2.RecordSource = "select *from tanques"
dbTanques2.Refresh
dbBicos.CursorLocation = adUseClient
dbBicos.Open "Select *from bicos", db, adOpenKeyset, adLockOptimistic

dbDespesasLanc2.CursorLocation = adUseClient
dbDespesasLanc2.Open "Select *from despesaslanc2 where descri='Comissões paga no caixa' and fechamento=0", db, adOpenKeyset, adLockOptimistic

dbStatus.CursorLocation = adUseClient
dbStatus.Open "Select *from status", db, adOpenKeyset, adLockOptimistic

AtualizaSequenciaCaixa

CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento

If FechandoLote = False Then
  Resposta = MsgBox("Deseja cancelar o fechamento atual?", vbYesNo, "Fechamento de Caixa!")
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

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic

With dbEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If .Recordset!apurado = True Then
        If IsNull(.Recordset!DesteCaixaQtd) = True Then
          Vendas = .Recordset!Vendas
        Else
          Vendas = .Recordset!DesteCaixaQtd
        End If
        
        dbTanques2.Recordset.MoveFirst
        dbTanques2.Recordset.Find "tanque=" & .Recordset!Tanque
        If dbTanques2.Recordset.EOF = False Then
          dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque + Vendas
          dbTanques2.Recordset.Update
        Else
          MsgBox "Tanque '" & .Recordset!Tanque & "' não encontrado!"
        End If
        
        dbBicos.MoveFirst
        dbBicos.Find "bico=" & .Recordset!Bico
        If dbBicos.EOF = True Then
          MsgBox "O bico " & .Recordset!Bico & " não foi encontrado no cadastro!", vbCritical, "Erro!"
        End If
        dbBicos!ultimonumero = .Recordset!Abertura
        If dbBicos!PrecoVenda <> .Recordset!Preco Then
          dbBicos!PrecoVenda = .Recordset!Preco
        End If
        dbBicos.Update
        
        If dbProdutos.RecordCount <> 0 Then
          dbProdutos.MoveFirst
          dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
          If dbProdutos.EOF = False Then
            If IsNull(dbProdutos!ValorEstoque) = True Then
              dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
            End If
            If IsNull(dbProdutos!PrecoMedio) = True Then
              dbProdutos!PrecoMedio = dbProdutos!precocompra
            End If
            If IsNull(dbProdutos!DifEstoque) = True Then
              dbProdutos!DifEstoque = 0
            End If
            If IsNull(dbProdutos!valordifestoque) = True Then
              dbProdutos!valordifestoque = 0
            End If
            If IsNull(dbProdutos!LucroMedio) = True Then
              dbProdutos!LucroMedio = 0
            End If
            PrecoMedio = .Recordset!PrecoMedio
            LucroMedio = .Recordset!LucroMedio
            dbProdutos!LucroMedio = dbProdutos!LucroMedio - LucroMedio
            ValorEstoque = PrecoMedio * Vendas
            dbProdutos!ValorEstoque = dbProdutos!ValorEstoque + ValorEstoque
            LucroVenda = (.Recordset!Preco - dbProdutos!precocompra) * Vendas
            dbProdutos!PrecoVenda = .Recordset!Preco
            dbProdutos!TotalVendido = dbProdutos!TotalVendido - .Recordset!DesteCaixaValor
            dbProdutos!Estoque = dbProdutos!Estoque + Vendas
            dbProdutos!acumulativo = dbProdutos!acumulativo - Vendas
            dbProdutos!LucroVenda = dbProdutos!LucroVenda - LucroVenda
            dbProdutos.Update
            .Recordset!apurado = False
            .Recordset.Update
          End If
        End If
        
        'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, .Recordset!CodigoProduto, .Recordset!Tanque, , Vendas
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

dbProdutos.Close
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos where combustivel=0", db, adOpenKeyset, adLockOptimistic

With dbVendas
  .RecordSource = "select *from venda2 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = True Then
        If .Recordset!Combustivel = 0 Then
          dbProdutos.MoveFirst
          dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
          If dbProdutos.EOF = True Then
            MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
          Else
            LucroVenda = (.Recordset!valorUnitario * .Recordset!Quantidade) - (dbProdutos!precocompra * .Recordset!Quantidade) - .Recordset!ValorComissao + .Recordset!ValorDesconto
            
            If IsNull(dbProdutos!ValorEstoque) = True Then
              dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
            End If
            If IsNull(dbProdutos!PrecoMedio) = True Then
              dbProdutos!PrecoMedio = dbProdutos!precocompra
            End If
            If IsNull(dbProdutos!DifEstoque) = True Then
              dbProdutos!DifEstoque = 0
            End If
            If IsNull(dbProdutos!valordifestoque) = True Then
              dbProdutos!valordifestoque = 0
            End If
            If IsNull(dbProdutos!LucroMedio) = True Then
              dbProdutos!LucroMedio = 0
            End If
            
            If dbProdutos!ValorEstoque <> 0 And dbProdutos!Estoque <> 0 Then
              PrecoMedio = dbProdutos!ValorEstoque / dbProdutos!Estoque
            Else
              PrecoMedio = 0
            End If
            LucroMedio = (.Recordset!valorUnitario * .Recordset!Quantidade) - (PrecoMedio * .Recordset!Quantidade) - .Recordset!ValorComissao + .Recordset!ValorDesconto
            dbProdutos!LucroMedio = dbProdutos!LucroMedio - LucroMedio
            ValorEstoque = PrecoMedio * .Recordset!Quantidade
            dbProdutos!ValorEstoque = dbProdutos!ValorEstoque + ValorEstoque
            
            
            EstoqueAnterior = dbProdutos!Estoque
            dbProdutos!Estoque = dbProdutos!Estoque + .Recordset!Quantidade
            dbProdutos!LucroVenda = dbProdutos!LucroVenda - LucroVenda
            dbProdutos!acumulativo = dbProdutos!acumulativo - .Recordset!Quantidade
            If IsNull(dbProdutos!TotalVendido) = True Then dbProdutos!TotalVendido = 0
            dbProdutos!TotalVendido = dbProdutos!TotalVendido - .Recordset!ValorTotal
            dbProdutos.Update
          End If
          
          StrTemp = "Extorno do Caixa: " & dbFechamentos.Recordset!DataCaixa & " turno: " & dbFechamentos.Recordset!Turno
          db.Execute "insert into produtoshistorico (lancadoem,dataalteracao,codigoproduto,codigo,descriproduto,descrioperacao,precocompra,precovenda,estoqueanterior,quantidade,estoquefinal) values " & _
                      "(#" & DataInglesa(Date) & " " & Time & "#,#" & DataInglesa(Date) & "#," & dbProdutos!CodigoProduto & "," & dbProdutos!Codigo & "," & _
                      "'" & dbProdutos!Descri & "','" & StrTemp & "'," & NumeroIngles(dbProdutos!precocompra) & "," & NumeroIngles(dbProdutos!PrecoVenda) & "," & NumeroIngles(EstoqueAnterior) & "," & NumeroIngles(dbVendas.Recordset!Quantidade) & "," & _
                      NumeroIngles(EstoqueAnterior - dbVendas.Recordset!Quantidade) & ")"
          
          RegistraEstoque dbFechamentos.Recordset!DataCaixa, dbFechamentos.Recordset!CodigoTurno, dbFechamentos.Recordset!Turno, dbFechamentos.Recordset!HoraIni, dbProdutos!CodigoProduto, , , -dbVendas.Recordset!Quantidade
          
        End If
        
        On Error Resume Next
        If ComissaoAcumulativa = False Then
          With dbDespesasLanc2
            If dbDespesasLanc2.RecordCount <> 0 Then
              dbDespesasLanc2.MoveFirst
              dbDespesasLanc2.Find "Descri='Comissões paga no caixa' and fechamento=0"
              If dbDespesasLanc2.EOF = True Then
                dbDespesasLanc2.AddNew
                dbDespesasLanc2!Valor = 0
              End If
            Else
              dbDespesasLanc2.AddNew
              dbDespesasLanc2!Valor = 0
            End If
            
            dbDespesasLanc2!CodigoFechamento = -1
            dbDespesasLanc2!Origem = "Despesa"
            dbDespesasLanc2!Data = dbFechamentos.Recordset!DataCaixa
            dbDespesasLanc2!Hora = Now
            dbDespesasLanc2!Vencimento = dbFechamentos.Recordset!DataCaixa
            dbDespesasLanc2!CodigoConta = 0
            dbDespesasLanc2!CodigoDespesa = 0
            dbDespesasLanc2!Descri = "Comissões paga no caixa"
            dbDespesasLanc2!Obs = dbFechamentos.Recordset!DataCaixa & " Turno " & dbFechamentos.Recordset!Turno
            dbDespesasLanc2!Valor = dbDespesasLanc2!Valor + dbVendas.Recordset!ValorComissao
            dbDespesasLanc2!valorpago = dbDespesasLanc2!valorpago + dbVendas.Recordset!ValorComissao
            dbDespesasLanc2!Fechamento = False
            dbDespesasLanc2!compensado = True
            dbDespesasLanc2!distribuido = True
            dbDespesasLanc2!codigoenviar = "1"
            dbDespesasLanc2!fechamentodiario = True
            dbDespesasLanc2.Update
          End With
          .Recordset!Pago = False
        End If
        .Recordset!fechamentodiario = False
        .Recordset.Update
        On Error GoTo 0
        
      End If
      
      .Recordset.MoveNext
    Loop
  End If
End With
dbProdutos.Close

'*******************************************************************************************
'Registra diferença de estoque no estatus
'*******************************************************************************************
With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic
    Do While .Recordset.EOF = False
      If .Recordset!apurado = True Then
        If .Recordset!Tanque <> 0 Then
          If .Recordset!Diferenca <> 0 Then
            dbTanques2.Recordset.MoveFirst
            dbTanques2.Recordset.Find "tanque=" & .Recordset!tanquenr
            dbProdutos.MoveFirst
            dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
            If dbProdutos.EOF = False And dbProdutos.BOF = False Then
              ValorEstoque = (.Recordset!ValorDiferenca)
              dbProdutos!DifEstoque = dbProdutos!DifEstoque - .Recordset!Diferenca
              dbProdutos!valordifestoque = dbProdutos!valordifestoque - ValorEstoque
              dbProdutos!Estoque = dbProdutos!Estoque - .Recordset!Diferenca
              dbProdutos!ValorEstoque = dbProdutos!ValorEstoque - ValorEstoque
              dbProdutos.Update
            Else
              ValorEstoque = 0
            End If
            dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque - .Recordset!Diferenca
            dbTanques2.Recordset.Update
                        
            .Recordset!ValorDiferenca = ValorEstoque
          Else
            .Recordset!ValorDiferenca = 0
          End If
          'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, .Recordset!tanquenr, , , .Recordset!Diferenca
        End If
        .Recordset!apurado = False
        .Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

On Error Resume Next
dbProdutos.Close
On Error GoTo 0

dbProdutos.Open "select *from produtos where combustivel=0", db, adOpenKeyset, adLockOptimistic

If SemTabelaDePrecos = False Then
  With qProdutosAltera
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutos.MoveFirst
        dbProdutos.Find "codigoproduto=" & .Recordset("produtos.CodigoProduto")
        If dbProdutos.EOF = False Then
          If .Recordset("produtosalteradetalhe.PrecoVenda") <> dbProdutos!PrecoVenda Then
            dbProdutos!PrecoVenda = .Recordset("produtosalteradetalhe.PrecoVenda")
            dbProdutos.Update
          End If
        End If
        .Recordset.MoveNext
      Loop
    End If
  End With
End If
With dbFechamentos
  If IsNumeric(lblTotalCombustivel.Caption) = True Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
  End If
  If IsNumeric(lblTotalProdutos.Caption) = True Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
  End If
  .Recordset!responsavel = cboResponsavel.Text
  .Recordset!fechado = False
  .Recordset!finalizadopor = Usuarios.Nome
  If IsNull(.Recordset!Arredondamento) = True Then .Recordset!Arredondamento = 0
  Arredondamento = .Recordset!Arredondamento
  .Recordset.Update
End With


dbStatus!Arredondamento = dbStatus!Arredondamento - Arredondamento
dbStatus.Update

'dbProdutos.Close
'dbProdutos.Open "select *from produtosnotascorpo where codigocaixa=" & CodigoFechamento & " and aguardando=0", db, adOpenKeyset, adLockOptimistic
'If dbProdutos.RecordCount <> 0 Then
'  Do While dbProdutos.EOF = False
'    If dbTanques2.Recordset.RecordCount <> 0 Then
'      dbTanques2.Recordset.MoveFirst
'      dbTanques2.Recordset.Find "tanque=" & dbProdutos!Tanque
'      If dbTanques2.Recordset.EOF = False Then
'        dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque - dbProdutos!Quantidade
'        dbProdutos!Aguardando = True
'      End If
'    End If
'    dbProdutos.MoveNext
'  Loop
'  dbProdutos.UpdateBatch adAffectAllChapters
'  dbTanques2.Recordset.UpdateBatch adAffectAllChapters
'End If

dbProdutos.Close

On Error Resume Next
dbBicos.Close
dbDespesasLanc2.Close
dbStatus.Close


On Error GoTo 0
'Dim Estatus As New frmEstatus2
'Load Estatus
'Unload Estatus

Animation1.Visible = False


Call cmdAbrir_Click
Call cmdExtornaTurno_Click
Desconfirmar = True

End Function

Private Function FecharCaixa() As Boolean

Dim dbBloqueiaFechamento As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbBicos As New ADODB.Recordset
Dim dbDespesasLanc2 As New ADODB.Recordset
Dim dbEntraTanque As New ADODB.Recordset
Dim dbStatus As New ADODB.Recordset
Dim dbVendasCombustivel As New ADODB.Recordset

Dim Resposta As Integer, LucroVenda As Currency, StrTemp As String
Dim Estacionamento As Currency, ValorEstoque As Currency
Dim Vendas As Double, LucroMedio As Currency, PrecoMedio As Currency

FecharCaixa = False
Call cmdAbrir_Click

If dbBloqueiaFechamento.State = adStateOpen Then
  dbBloqueiaFechamento.Close
End If
dbBloqueiaFechamento.CursorLocation = adUseClient
dbBloqueiaFechamento.Open "select bloqueiafechamento.*, Turnos.* from bloqueiafechamento, turnos where bloqueiafechamento.codigoturno1=turnos.codigoturno ", db, adOpenForwardOnly, adLockReadOnly

dbTanques2.RecordSource = "select *from tanques"
dbTanques2.Refresh

If dbBicos.State = adStateOpen Then
  dbBicos.Close
End If
dbBicos.CursorLocation = adUseClient
dbBicos.Open "Select *from bicos", db, adOpenKeyset, adLockOptimistic

If dbStatus.State = adStateOpen Then
  dbStatus.Close
End If
dbStatus.CursorLocation = adUseClient
dbStatus.Open "Select *from status", db, adOpenKeyset, adLockOptimistic


AtualizaSequenciaCaixa


If dbBloqueiaFechamento.RecordCount <> 0 Then
  If dbBloqueiaFechamento!bloqueia1 = True Then
    If dbBloqueiaFechamento!Data1 <= dbFechamentos.Recordset!DataCaixa And dbBloqueiaFechamento!bloqueia1 = True Then
      If dbBloqueiaFechamento!HoraIni <= dbFechamentos.Recordset!HoraIni Then
        MsgBox "Caixa não pode ser confirmado por estar bloqueado pelo administrador!"
        Exit Function
      End If
    End If
  End If
End If

Call cmdEntraCombustivel_Click

'dbEntraTanque.CursorLocation = adUseClient
'dbEntraTanque.Open "select *from produtosnotascorpo where codigocaixa=" & dbFechamentos.Recordset!CodigoFechamento & " and aguardando=-1", db, adOpenKeyset, adLockOptimistic
'If dbEntraTanque.RecordCount <> 0 Then
'  Do While dbEntraTanque.EOF = False
'    If dbTanques2.Recordset.RecordCount <> 0 Then
'      dbTanques2.Recordset.MoveFirst
'      dbTanques2.Recordset.Find "tanque=" & dbEntraTanque!Tanque
'      If dbTanques2.Recordset.EOF = False Then
'        dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque + dbEntraTanque!Quantidade
'        dbEntraTanque!Aguardando = False
'      End If
'    End If
'    dbEntraTanque.MoveNext
'  Loop
'  dbEntraTanque.UpdateBatch adAffectAllChapters
'  dbTanques2.Recordset.UpdateBatch adAffectAllChapters
'  Call cmdCalcular_Click
'  Call cmdCalcular_Click
'  DoEvents
'End If



With dbDifComb
  .Refresh
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
CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento

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

On Error Resume Next
If db.State = adStateOpen Then
  db.Close
End If
db.Open
On Error GoTo 0

If dbProdutos.State = adStateOpen Then
  dbProdutos.Close
End If
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic

If dbBicos.State = adStateOpen Then
  dbBicos.Close
End If
dbBicos.CursorLocation = adUseClient
dbBicos.Open "Select *from bicos order by bico", db, adOpenKeyset, adLockOptimistic

If dbVendasCombustivel.State = adStateOpen Then
  dbVendasCombustivel.Close
End If
dbVendasCombustivel.CursorLocation = adUseClient
dbVendasCombustivel.Open "Select *from venda2 where combustivel=-1 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenKeyset, adLockOptimistic

With dbEncerrantes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      If .Recordset!apurado = False Then
        Vendas = .Recordset!DesteCaixaQtd
        dbTanques2.Refresh
        dbTanques2.Recordset.MoveFirst
        dbTanques2.Recordset.Find "tanque=" & .Recordset!Tanque
        If dbTanques2.Recordset.EOF = False Then
          dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque - Vendas
          dbTanques2.Recordset.Update
        Else
          MsgBox "Tanque '" & .Recordset!Tanque & "' não encontrado!"
        End If
        
        dbBicos.MoveFirst
        dbBicos.Find "bico=" & .Recordset!Bico
        If dbBicos.EOF = True Then
          MsgBox "O bico " & .Recordset!Bico & " não foi encontrado no cadastro!", vbCritical, "Erro!"
        End If
        dbBicos!ultimonumero = .Recordset!Encerrante
        .Recordset!Tanque = dbBicos!Tanque
        .Recordset!CodigoProduto = dbBicos!CodigoProduto
        .Recordset.Update
        If dbBicos!PrecoVenda <> .Recordset!Preco Then
          dbBicos!PrecoVenda = .Recordset!Preco
        End If
        dbBicos.Update
        
        If dbProdutos.RecordCount <> 0 Then
          dbProdutos.MoveFirst
          dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
          If dbProdutos.EOF = False Then
            If IsNull(dbProdutos!ValorEstoque) = True Then
              dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
            End If
            If IsNull(dbProdutos!PrecoMedio) = True Then
              dbProdutos!PrecoMedio = dbProdutos!precocompra
            End If
            If IsNull(dbProdutos!DifEstoque) = True Then
              dbProdutos!DifEstoque = 0
            End If
            If IsNull(dbProdutos!valordifestoque) = True Then
              dbProdutos!valordifestoque = 0
            End If
            If IsNull(dbProdutos!LucroMedio) = True Then
              dbProdutos!LucroMedio = 0
            End If
            If dbProdutos!ValorEstoque <> 0 And dbProdutos!Estoque <> 0 Then
              PrecoMedio = dbProdutos!ValorEstoque / dbProdutos!Estoque
            Else
              PrecoMedio = dbProdutos!precocompra
            End If
            If IsNull(.Recordset!Comissao) = False Then
                LucroMedio = ((.Recordset!Preco - PrecoMedio) * Vendas) - .Recordset!Comissao
            Else
                LucroMedio = (.Recordset!Preco - PrecoMedio) * Vendas
            End If
            dbProdutos!LucroMedio = dbProdutos!LucroMedio + LucroMedio
            ValorEstoque = PrecoMedio * Vendas
            dbProdutos!ValorEstoque = dbProdutos!ValorEstoque - ValorEstoque
            LucroVenda = ((.Recordset!Preco - dbProdutos!precocompra) * Vendas) - .Recordset!Comissao
            dbProdutos!PrecoVenda = .Recordset!Preco
            dbProdutos!TotalVendido = dbProdutos!TotalVendido + .Recordset!DesteCaixaValor
            dbProdutos!Estoque = dbProdutos!Estoque - Vendas
            dbProdutos!acumulativo = dbProdutos!acumulativo + Vendas
            dbProdutos!LucroVenda = dbProdutos!LucroVenda + LucroVenda
            dbProdutos.Update
            .Recordset!LucroMedio = LucroMedio
            .Recordset!PrecoMedio = PrecoMedio
            .Recordset!apurado = True
            .Recordset.Update
          End If
        End If
        
        
        On Error GoTo 0
                
        'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, .Recordset!CodigoProduto, .Recordset!Tanque, , Vendas
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

dbProdutos.Close
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos where combustivel=0", db, adOpenKeyset, adLockOptimistic

With dbVendas
  .RecordSource = "Select *from venda2 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!fechamentodiario = False Then
        dbProdutos.MoveFirst
        dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
        If .Recordset!Combustivel = 0 Then
          If dbProdutos.EOF = True Then
            MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
          Else
            LucroVenda = (.Recordset!valorUnitario * .Recordset!Quantidade) - (dbProdutos!precocompra * .Recordset!Quantidade) - .Recordset!ValorComissao + .Recordset!ValorDesconto
            
            If IsNull(dbProdutos!ValorEstoque) = True Then
              dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
            End If
            If IsNull(dbProdutos!PrecoMedio) = True Then
              dbProdutos!PrecoMedio = dbProdutos!precocompra
            End If
            If IsNull(dbProdutos!DifEstoque) = True Then
              dbProdutos!DifEstoque = 0
            End If
            If IsNull(dbProdutos!valordifestoque) = True Then
              dbProdutos!valordifestoque = 0
            End If
            If IsNull(dbProdutos!LucroMedio) = True Then
              dbProdutos!LucroMedio = 0
            End If
            
            If dbProdutos!ValorEstoque <> 0 And dbProdutos!Estoque <> 0 Then
              PrecoMedio = dbProdutos!ValorEstoque / dbProdutos!Estoque
            Else
              PrecoMedio = dbProdutos!precocompra
            End If
            LucroMedio = 0
            LucroMedio = (.Recordset!valorUnitario * .Recordset!Quantidade) - (PrecoMedio * .Recordset!Quantidade) - .Recordset!ValorComissao + .Recordset!ValorDesconto
            dbProdutos!LucroMedio = dbProdutos!LucroMedio + LucroMedio
            ValorEstoque = PrecoMedio * .Recordset!Quantidade
            dbProdutos!ValorEstoque = dbProdutos!ValorEstoque - ValorEstoque
            
            
            EstoqueAnterior = dbProdutos!Estoque
            dbProdutos!Estoque = dbProdutos!Estoque - .Recordset!Quantidade
            dbProdutos!LucroVenda = dbProdutos!LucroVenda + LucroVenda
            dbProdutos!acumulativo = dbProdutos!acumulativo + .Recordset!Quantidade
            If IsNull(dbProdutos!TotalVendido) = True Then dbProdutos!TotalVendido = 0
            dbProdutos!TotalVendido = dbProdutos!TotalVendido + .Recordset!ValorTotal
            dbProdutos!ultimavenda = dbFechamentos.Recordset!DataCaixa
            dbProdutos.Update
          End If
          
          StrTemp = "Venda no Caixa: " & dbFechamentos.Recordset!DataCaixa & " turno: " & dbFechamentos.Recordset!Turno
          db.Execute "insert into produtoshistorico (lancadoem,dataalteracao,codigoproduto,codigo,descriproduto,descrioperacao,precocompra,precovenda,estoqueanterior,quantidade,estoquefinal) values " & _
                      "(#" & DataInglesa(Date) & " " & Time & "#,#" & DataInglesa(Date) & "#," & dbProdutos!CodigoProduto & "," & dbProdutos!Codigo & "," & _
                      "'" & dbProdutos!Descri & "','" & StrTemp & "'," & NumeroIngles(dbProdutos!precocompra) & "," & NumeroIngles(dbProdutos!PrecoVenda) & "," & NumeroIngles(EstoqueAnterior) & "," & NumeroIngles(dbVendas.Recordset!Quantidade) & "," & _
                      NumeroIngles(EstoqueAnterior - dbVendas.Recordset!Quantidade) & ")"
          
          RegistraEstoque dbFechamentos.Recordset!DataCaixa, dbFechamentos.Recordset!CodigoTurno, dbFechamentos.Recordset!Turno, dbFechamentos.Recordset!HoraIni, dbProdutos!CodigoProduto, , , dbVendas.Recordset!Quantidade
          
        End If
        'On Error Resume Next
        If .Recordset!codigovendedor <> 0 Then
          If ComissaoAcumulativa = False Then
            With dbDespesasLanc2
              If dbDespesasLanc2.State = adStateOpen Then
                dbDespesasLanc2.Close
              End If
              If dbDespesasLanc2.State = adStateOpen Then
                dbDespesasLanc2.Close
              End If
              dbDespesasLanc2.CursorLocation = adUseClient
              dbDespesasLanc2.Open "Select *from despesaslanc2 where descri='Comissões paga no caixa' and fechamento=0", db, adOpenKeyset, adLockOptimistic

              If dbDespesasLanc2.RecordCount <> 0 Then
                'dbDespesasLanc2.MoveFirst
                'dbDespesasLanc2.Find "Descri='Comissões paga no caixa' and fechamento=0"
                If dbDespesasLanc2.EOF = True Then
                  dbDespesasLanc2.AddNew
                  dbDespesasLanc2!Valor = 0
                  dbDespesasLanc2!valorpago = 0
                End If
              Else
                dbDespesasLanc2.AddNew
                dbDespesasLanc2!Valor = 0
                dbDespesasLanc2!valorpago = 0
              End If
              
              dbDespesasLanc2!CodigoFechamento = -1
              dbDespesasLanc2!Origem = "Despesa"
              dbDespesasLanc2!Data = dbFechamentos.Recordset!DataCaixa
              dbDespesasLanc2!Hora = Now
              dbDespesasLanc2!Vencimento = dbFechamentos.Recordset!DataCaixa
              dbDespesasLanc2!CodigoConta = 0
              dbDespesasLanc2!CodigoDespesa = 0
              dbDespesasLanc2!Descri = "Comissões paga no caixa"
              dbDespesasLanc2!Obs = dbFechamentos.Recordset!DataCaixa & " Turno " & dbFechamentos.Recordset!Turno
              If IsNull(dbDespesasLanc2!Valor) = True Then dbDespesasLanc2!Valor = 0
              If IsNull(dbDespesasLanc2!valorpago) = True Then dbDespesasLanc2!valorpago = 0
              dbDespesasLanc2!Valor = dbDespesasLanc2!Valor - dbVendas.Recordset!ValorComissao
              dbDespesasLanc2!valorpago = dbDespesasLanc2!valorpago - dbVendas.Recordset!ValorComissao
              dbDespesasLanc2!Fechamento = False
              dbDespesasLanc2!compensado = True
              'dbDespesasLanc2!distribuido = True
              dbDespesasLanc2!codigoenviar = "1"
              dbDespesasLanc2!fechamentodiario = True
              dbDespesasLanc2!Produto = True
              dbDespesasLanc2.Update
            End With
            .Recordset!Pago = True
          End If
        End If
        .Recordset!fechamentodiario = True
        .Recordset.Update
        On Error GoTo 0
      End If
      
      
      
      .Recordset.MoveNext
    Loop
  End If
End With
dbProdutos.Close

'*******************************************************************************************
'Registra diferença de estoque no estatus
'*******************************************************************************************
With dbDifComb
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    
    If dbProdutos.State = adStateOpen Then
      dbProdutos.Close
    End If
    dbProdutos.Open "Select *from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic
    Do While .Recordset.EOF = False
      If .Recordset!apurado = False Then
        If .Recordset!Tanque <> 0 Then
          If .Recordset!Diferenca <> 0 Then
            dbTanques2.Refresh
            dbTanques2.Recordset.MoveFirst
            dbTanques2.Recordset.Find "tanque=" & .Recordset!tanquenr
            dbProdutos.MoveFirst
            dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
            If dbProdutos.EOF = False And dbProdutos.BOF = False Then
              If IsNull(dbProdutos!precocompra) = False Then
                ValorEstoque = dbProdutos!precocompra
              Else
                ValorEstoque = 0
              End If
              If dbProdutos!ValorEstoque <> 0 And dbProdutos!Estoque <> 0 Then
                  ValorEstoque = (dbProdutos!ValorEstoque / dbProdutos!Estoque)
                  ValorEstoque = ValorEstoque * .Recordset!Diferenca
                  dbProdutos!DifEstoque = dbProdutos!DifEstoque + .Recordset!Diferenca
                  dbProdutos!valordifestoque = dbProdutos!valordifestoque + ValorEstoque
                  dbProdutos!Estoque = dbProdutos!Estoque + .Recordset!Diferenca
                  dbProdutos!ValorEstoque = dbProdutos!ValorEstoque + ValorEstoque
                  dbProdutos.Update
                  dbTanques2.Recordset!Estoque = dbTanques2.Recordset!Estoque + .Recordset!Diferenca
                  dbTanques2.Recordset.Update
              End If
            End If
            .Recordset!ValorDiferenca = ValorEstoque
            .Recordset.Update
          Else
            .Recordset!ValorDiferenca = 0
            .Recordset.Update
          End If
          'RegistraEstoque dbFechamento.Recordset!DataCaixa, dbFechamento.Recordset!CodigoTurno, dbFechamento.Recordset!Turno, dbFechamento.Recordset!HoraIni, dbProdutos.Recordset!CodigoProduto, .Recordset!tanquenr, , , .Recordset!Diferenca
        End If
        .Recordset!apurado = True
        .Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

On Error Resume Next
dbProdutos.Close
On Error GoTo 0

dbProdutos.Open "select *from produtos where combustivel=0", db, adOpenKeyset, adLockOptimistic
If SemTabelaDePrecos = False Then
  With qProdutosAltera
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutos.MoveFirst
        dbProdutos.Find "codigoproduto=" & .Recordset("produtos.CodigoProduto")
        If dbProdutos.EOF = False Then
          If .Recordset("produtosalteradetalhe.PrecoVenda") <> dbProdutos!PrecoVenda Then
            dbProdutos!PrecoVenda = .Recordset("produtosalteradetalhe.PrecoVenda")
            dbProdutos.Update
          End If
        End If
        .Recordset.MoveNext
      Loop
    End If
  End With
End If
With dbFechamentos
  If IsNumeric(lblTotalCombustivel.Caption) = True Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
  End If
  If IsNumeric(lblTotalProdutos.Caption) = True Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
  End If
  .Recordset!responsavel = cboResponsavel.Text
  .Recordset!fechado = True
  .Recordset!finalizadopor = Usuarios.Nome
  .Recordset!Arredondamento = Arredondamento
  .Recordset!ComissaoAcumulativa = ComissaoAcumulativa
  .Recordset.Update
End With

On Error Resume Next
dbStatus!Arredondamento = dbStatus!Arredondamento + Arredondamento
dbStatus.Update

dbBloqueiaFechamento.Close
dbBicos.Close
dbDespesasLanc2.Close
dbStatus.Close
dbVendasCombustivel.Close

On Error GoTo 0

Animation1.Visible = False


Call cmdAbrir_Click
Call cmdEntraCombustivel_Click

FecharCaixa = True
End Function

Public Function PodeFechar() As Boolean
Dim dbResultado As New ADODB.Recordset
Dim dbConfereBico As New ADODB.Recordset
Dim TotalVendas As Double, TotalAssumido As Double
Dim BicoAtual As Integer
Dim TempValor As Currency, Tolerancia As Double
Dim TotalNumerarios As Currency, TotalVendidoArredonda As Currency, TotalAssumidoArredonda As Currency
Dim BloqueiaNaoImportado As Boolean

PodeFechar = False

Tolerancia = 0.1

On Error Resume Next
db.Open CaminhoADO
On Error GoTo 0

If dbResultado.State = adStateOpen Then
  dbResultado.Close
End If
dbResultado.CursorLocation = adUseClient
dbResultado.Open "select *from fechamentodecaixapista where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly

If dbConfereBico.State = adStateOpen Then
  dbConfereBico.Close
End If
dbConfereBico.CursorLocation = adUseClient
dbConfereBico.Open "select bicoencerrantes.*, fechamentodecaixa.datacaixa, fechamentodecaixa.horaini from bicoencerrantes, fechamentodecaixa where bicoencerrantes.codigofechamento=fechamentodecaixa.codigofechamento and datacaixa=#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "# and horaini=#" & dbFechamentos.Recordset!HoraIni & "# order by bico", db, adOpenForwardOnly, adLockReadOnly

If dbResultado.RecordCount <> 0 Then
  TotalNumerarios = 0
      
  dbResultado.MoveFirst
  dbResultado.Find "codigoconta='5000000000'"
  If dbResultado.EOF = False Then
    TotalNumerarios = dbResultado!Valor
  End If
  
  dbResultado.MoveFirst
  dbResultado.Find "codigoconta='4100000000'"
  If dbResultado.EOF = False Then
    TotalNumerarios = TotalNumerarios - dbResultado!Valor
  End If
  
  dbResultado.MoveFirst
  dbResultado.Find "codigoconta='3000000000'"
  If dbResultado.EOF = False Then
    TotalNumerarios = TotalNumerarios + dbResultado!Valor
  End If
  If CCur(txtInformado.Text) = 0 Then
    txtInformado.Text = Format(TotalNumerarios, "currency")
    Call cmdCalcular_Click
  End If
  If TotalNumerarios <> CCur(txtInformado.Text) Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Total informado está incorreto? Deveria ser " & Format(TotalNumerarios, "Currency") & " Deseja continuar?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      End If
    Else
      MsgBox "Total informado está incorreto? Deveria ser " & Format(TotalNumerarios, "Currency") & "! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  End If
  
End If

If dbConfereBico.RecordCount <> 0 Then
  dbConfereBico.MoveFirst
  BicoAtual = 0
  TotalVendas = 0
  TotalAssumido = 0
  TotalVendidoArredonda = 0
  TotalAssumidoArredonda = 0
  Arredondamento = 0
  Do While dbConfereBico.EOF = False
    If BicoAtual <> dbConfereBico!Bico Then
      BicoAtual = dbConfereBico!Bico
      TotalVendas = dbConfereBico!Vendas
      TotalVendidoArredonda = TotalVendidoArredonda + dbConfereBico!ValorTotal
      TotalAssumido = 0
    End If
    If IsNull(dbConfereBico!DesteCaixaQtd) = False Then
      TotalAssumido = TotalAssumido + dbConfereBico!DesteCaixaQtd
    End If
    If IsNull(dbConfereBico!DesteCaixaValor) = False Then
      TotalAssumidoArredonda = TotalAssumidoArredonda + dbConfereBico!DesteCaixaValor
    End If
    
    dbConfereBico.MoveNext
    If dbConfereBico.EOF = False Then
      If BicoAtual <> dbConfereBico!Bico Then
        TempValor = TotalVendas - TotalAssumido
        If TempValor > 0.1 Or TempValor < -0.1 Then
          MsgBox "O bico " & BicoAtual & " não foi completado a venda em todos os caixas! Diferenca de " & Format(TempValor, "Currency")
          PodeFechar = False
          Exit Function
        End If
      End If
    End If
    
    
  Loop
  TempValor = TotalVendas - TotalAssumido
  If TempValor > 0.1 Or TempValor < -0.1 Then
    MsgBox "O bico " & Bico & " não foi completado a venda em todos os caixas!"
    PodeFechar = False
    Exit Function
  End If
  
  Arredondamento = TotalVendidoArredonda - TotalAssumidoArredonda
  
End If
StrTemp = ReadINI("Fechamento", "NaoImportado", "1", App.Path & "\Posto.ini")
If Trim(StrTemp) = "0" Then
  BloqueiaNaoImportado = True
Else
  BloqueiaNaoImportado = False
End If

If dbResultado.RecordCount = 0 And BloqueiaNaoImportado = True Then
    
    If Usuarios.Grupo.AdmEstatus = 2 Then
  '      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
  '      If Resposta = vbNo Then
  '        Exit Function
  '      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
Else
  
  dbResultado.MoveFirst
  'encontra venda de combustiveis
  StrTemp = ReadINI("Fechamento", "VendasCombustivel", "1100000000", App.Path & "\Posto.ini")
  
  dbResultado.Find "codigoconta='" & StrTemp & "'"
  If dbResultado.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      End If
    Else
      MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
      Exit Function
    End If
  Else
    If IsNumeric(lblTotalCombustivel.Caption) = True Then
      TempValor = CCur(lblTotalCombustivel.Caption) - dbResultado!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como venda de combustiveis " & Format(dbResultado!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          End If
        Else
          MsgBox "Este caixa deveria ter como venda de combustiveis " & Format(dbResultado!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
  dbResultado.MoveFirst
  'encontra venda de Produtos
  StrTemp = ReadINI("Fechamento", "VendasProdutos", "1200000000", App.Path & "\Posto.ini")
  dbResultado.Find "codigoconta='" & StrTemp & "'"
  If dbResultado.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      If BloqueiaNaoImportado = True Then
        MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
        Exit Function
      End If
    End If
  Else
    If IsNumeric(lblTotalProdutos.Caption) = True Then
      TempValor = CCur(lblTotalProdutos.Caption) - dbResultado!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como venda de produtos " & Format(dbResultado!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          Else
            PodeFechar = True
          End If
        Else
          MsgBox "Este caixa deveria ter como venda de produtos " & Format(dbResultado!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
  
  dbResultado.MoveFirst
  'encontra Diferença de Caixa
  StrTemp = ReadINI("Fechamento", "Diferenca", "2100000000", App.Path & "\Posto.ini")
  dbResultado.Find "codigoconta='" & StrTemp & "'"
  If dbResultado.EOF = True Then
    If Usuarios.Grupo.AdmEstatus = 2 Then
      Resposta = MsgBox("Este caixa não foi importado! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Function
      Else
        PodeFechar = True
        Exit Function
      End If
    Else
      If BloqueiaNaoImportado = True Then
        MsgBox "Este caixa não foi importado! Somente usuário administrativo pode confirmar!"
        Exit Function
      End If
    End If
  Else
    If IsNumeric(lblDiferenca.Caption) = True Then
      TempValor = CCur(lblDiferenca.Caption) - dbResultado!Valor
      If TempValor > Tolerancia Or TempValor < -Tolerancia Then
        If Usuarios.Grupo.AdmEstatus = 2 Then
          Resposta = MsgBox("Este caixa deveria ter como diferença de caixa " & Format(dbResultado!Valor, "Currency") & "! Deseja confirmar assim mesmo?", vbYesNo + vbDefaultButton2)
          If Resposta = vbNo Then
            Exit Function
          Else
            PodeFechar = True
          End If
        Else
          MsgBox "Este caixa deveria ter como diferença de caixa " & Format(dbResultado!Valor, "Currency") & "! Somente usuário administrativo pode confirmar!"
          Exit Function
        End If
      End If
    End If
  End If
  
End If

  

dbConfereBico.Close
dbResultado.Close
db.Close

PodeFechar = True

End Function


Private Sub RemoveEspecial()

Dim dbTemp As New ADODB.Recordset
Dim Resposta As Integer, CodigoFechamento As Double



If dbFechamentos.Recordset.EOF = False Then
  StrTemp = InputBox("Código do Fechamento a ser removido", , dbFechamentos.Recordset!CodigoFechamento)
Else
  Exit Sub
End If
If IsNumeric(StrTemp) = False Then
  Exit Sub
End If
CodigoFechamento = CDbl(StrTemp)
If Usuarios.Nome <> "Usuário Master" Then
  MsgBox "Você não tem permissão para remover um caixa!"
  Exit Sub
End If
With dbFechamentos
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    .Recordset.Find "codigofechamento=" & CodigoFechamento
    If .Recordset.EOF = False Then
      If Frame1.Visible = False Then Exit Sub
      If .Recordset!fechado = True Then
        MsgBox "Não é possível remover um caixa já finalizado!"
        Exit Sub
      End If
    Else
      Exit Sub
    End If
  End If
  
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  
  Resposta = MsgBox("Deseja remover o caixa atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  
'  Permissao = False
'  frmPermissao.Show vbModal
'  If Permissao = False Then
'    Exit Sub
'  End If
  
  If ApagaRegistros(True) = False Then
    MsgBox "O Caixa não pode ser removido pois já possui registro gravado."
    Exit Sub
  End If
  dbFechamentos.Recordset.MoveFirst
  
  db.Execute "delete *from fechamentodecaixa where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from BicoEncerrantes where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from TanqueEstoque where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from venda2 where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from diferencacombustivel where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from estacionamentocaixa where codigocaixa=" & CodigoFechamento
  db.Execute "delete *from formadepagamentorecebido2 where codigofechamento=" & CodigoFechamento
  db.Execute "delete *from clientesnota2 where codigofechamento=" & CodigoFechamento
  
End With
Call cmdUltimo_Click
End Sub

Private Sub CalculaBicos(ByVal ColIndex As Integer)
With dbEncerrantes
  If IsNull(.Recordset!deoutrocaixaqtd) = True Then
    .Recordset!deoutrocaixaqtd = 0
  End If
  If IsNull(.Recordset!deoutrocaixavalor) = True Then
    .Recordset!deoutrocaixavalor = 0
  End If
  If IsNull(.Recordset!DesteCaixaQtd) = True Then
    .Recordset!DesteCaixaQtd = 0
  End If
  If IsNull(.Recordset!DesteCaixaValor) = True Then
    .Recordset!DesteCaixaValor = 0
  End If
End With

Select Case ColIndex
  Case 0 'bico não pode mudar
    Cancel = True
  Case 1 'Abertura não pode mudar
    Cancel = True
  Case 2
    'está mudando o encerrante
    With dbEncerrantes
      .Recordset!Encerrante = CDbl(DataGrid1.Columns(ColIndex).Text)
      If .Recordset!Encerrante <> 0 Then
        If .Recordset!Encerrante > 1000000 Then
          If .Recordset!Abertura > 1000000 Then
            Do While .Recordset!Encerrante > 1000000
              .Recordset!Encerrante = .Recordset!Encerrante - 1000000
            Loop
          End If
          If .Recordset!Abertura < .Recordset!Encerrante Then
            Do While .Recordset!Encerrante > 1000000
              .Recordset!Encerrante = .Recordset!Encerrante - 1000000
            Loop
          End If
        End If
        If .Recordset!Abertura > 1000000 Then
            Do While .Recordset!Abertura > 1000000
              .Recordset!Abertura = .Recordset!Abertura - 1000000
            Loop
        End If
        .Recordset!Vendas = .Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno
        .Recordset!ValorTotal = (.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno) * .Recordset!Preco
      Else
        If .Recordset!Abertura > 1000000 Then .Recordset!Abertura = .Recordset!Abertura - 1000000
        .Recordset!Encerrante = .Recordset!Abertura
        .Recordset!Vendas = 0
        .Recordset!ValorTotal = 0
      End If
    End With
  Case 3
    'está mudando as vendas
    With dbEncerrantes
      .Recordset!Vendas = CDbl(DataGrid1.Columns(ColIndex).Text)
      .Recordset!Encerrante = .Recordset!Abertura + .Recordset!Vendas
      .Recordset!ValorTotal = .Recordset!Preco * .Recordset!Vendas
      .Recordset!DesteCaixaQtd = .Recordset!Vendas - .Recordset!deoutrocaixaqtd
      .Recordset!DesteCaixaValor = .Recordset!ValorTotal - .Recordset!deoutrocaixavalor
    End With
  Case 4
    'está mudando o retorno
    With dbEncerrantes
      If IsNumeric(DataGrid1.Columns(ColIndex).Text) = True Then
        .Recordset!Retorno = CDbl(DataGrid1.Columns(ColIndex).Text)
        If .Recordset!Encerrante <> 0 Then
          If .Recordset!Encerrante > 1000000 Then
            If .Recordset!Abertura > 1000000 Then
              Do While .Recordset!Encerrante > 1000000
                .Recordset!Encerrante = .Recordset!Encerrante - 1000000
              Loop
            End If
            If .Recordset!Abertura < .Recordset!Encerrante Then
              Do While .Recordset!Encerrante > 1000000
                .Recordset!Encerrante = .Recordset!Encerrante - 1000000
              Loop
            End If
          End If
          If .Recordset!Abertura > 1000000 Then
              Do While .Recordset!Abertura > 1000000
                .Recordset!Abertura = .Recordset!Abertura - 1000000
              Loop
          End If
          .Recordset!Vendas = .Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno
          .Recordset!ValorTotal = (.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno) * .Recordset!Preco
        Else
          If .Recordset!Abertura > 1000000 Then .Recordset!Abertura = .Recordset!Abertura - 1000000
          .Recordset!Encerrante = .Recordset!Abertura
          .Recordset!Vendas = 0
          .Recordset!ValorTotal = 0
        End If
      End If
    End With
  Case 5
  Case 6
    'está mudando o valor total de venda
    With dbEncerrantes
      .Recordset!ValorTotal = CDbl(DataGrid1.Columns(ColIndex).Text)
      If .Recordset!ValorTotal = 0 Then
        .Recordset!Vendas = 0
        .Recordset!Encerrante = .Recordset!Abertura
      Else
        .Recordset!Vendas = (.Recordset!ValorTotal / .Recordset!Preco) - .Recordset!Retorno
        .Recordset!Encerrante = .Recordset!Abertura + .Recordset!Vendas
      End If
    End With
  Case 7
    'está mudando a quantidade deste caixa
    With dbEncerrantes
      .Recordset!DesteCaixaQtd = CDbl(DataGrid1.Columns(ColIndex).Text)
      .Recordset!DesteCaixaValor = .Recordset!DesteCaixaQtd * .Recordset!Preco
    End With
  Case 8
    'está mudando o valor deste caixa
    With dbEncerrantes
      .Recordset!DesteCaixaValor = CDbl(DataGrid1.Columns(ColIndex).Text)
      If .Recordset!DesteCaixaValor <> 0 Then
        .Recordset!DesteCaixaQtd = .Recordset!DesteCaixaValor / .Recordset!Preco
      Else
        .Recordset!DesteCaixaQtd = 0
      End If
    End With
  Case 9
    
  Case 10
  
End Select
With dbEncerrantes
    .Recordset!deoutrocaixaqtd = (.Recordset!Encerrante - .Recordset!Abertura - .Recordset!Retorno) - .Recordset!DesteCaixaQtd
    .Recordset!deoutrocaixavalor = .Recordset!ValorTotal - .Recordset!DesteCaixaValor
    .Recordset.Update
End With
TotalCombustivel
TotalCombustivel
DataGrid1.Refresh
End Sub

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


Private Function ApagaRegistros(Optional RemovendoCaxa As Boolean = False) As Boolean
Dim CodigoFechamento As Double
Dim SoPrimeira As Boolean

Dim dbClientesNotas As New ADODB.Recordset
Dim dbFormaDePgRecebido As New ADODB.Recordset
Dim dbDespesasLanc As New ADODB.Recordset
Dim dbClientes As New ADODB.Recordset

CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento


SoPrimeira = False
ApagaRegistros = False
If dbFechamentos.Recordset!notaconferida = True Then
  SoPrimeira = True
End If

If dbClientesNotas.State = adStateOpen Then
  dbClientesNotas.Close
End If
dbClientesNotas.CursorLocation = adUseClient
dbClientesNotas.Open "select *from clientesnota2 where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic
dbClientesNotas.Filter = "confirmado=-1"
If dbClientesNotas.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbClientesNotas.Filter = ""

If dbFormaDePgRecebido.State = adStateOpen Then
  dbFormaDePgRecebido.Close
End If
dbFormaDePgRecebido.CursorLocation = adUseClient
dbFormaDePgRecebido.Open "select fechamentodiario from formadepagamentorecebido2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If dbFormaDePgRecebido.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbFormaDePgRecebido.Close

If dbDespesasLanc.State = adStateOpen Then
  dbDespesasLanc.Close
End If
dbDespesasLanc.CursorLocation = adUseClient
dbDespesasLanc.Open "select fechamentodiario from despesaslanc2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If dbDespesasLanc.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbDespesasLanc.Close

If SoPrimeira = False Then
  ApagaRegistros = True
End If


If RemovendoCaxa = True Then
  If SoPrimeira = True Then
    ApagaRegistros = False
    Exit Function
  End If
End If

db.Execute "delete from venda2 where codigofechamento=" & CodigoFechamento
db.Execute "delete from comissoes where codigofechamento=" & CodigoFechamento

If SoPrimeira = False Then
  With dbClientesNotas
    If dbClientes.State = adStateOpen Then
      dbClientes.Close
    End If
    dbClientes.CursorLocation = adUseClient
    dbClientes.Open "select *from clientes", db, adOpenKeyset, adLockOptimistic
    
    If dbClientesNotas.RecordCount <> 0 Then
      Do While dbClientesNotas.EOF = False
        dbClientes.MoveFirst
        dbClientes.Find "codigocliente=" & dbClientesNotas!CodigoCliente
        dbClientes!TotalNotas = dbClientes!TotalNotas - dbClientesNotas!ValorPrevisto
        dbClientes!Saldo = dbClientes!Limite - dbClientes!TotalNotas - dbClientes!TotalBoleto
        dbClientes.Update
        
        dbClientesNotas.MoveNext
      Loop
      db.Execute "delete *from clientesnota2 where codigofechamento=" & CodigoFechamento
    End If
  End With
  dbClientesNotas.Close
  dbClientes.Close
  
  db.Execute "delete *from formadepagamentorecebido2 where codigofechamento=" & CodigoFechamento
  
  db.Execute "delete *from despesaslanc2 where codigofechamento=" & CodigoFechamento
  
End If

db.Execute "delete *from fechamentodecaixapista where codigofechamento=" & CodigoFechamento



End Function


Private Sub PegaEncerranteIntermitente(ByRef Bico As ADODB.Recordset)

Dim dbEncerrantes As New ADODB.Recordset
If dbEncerrantes.State = adStateOpen Then
  dbEncerrantes.Close
End If
dbEncerrantes.CursorLocation = adUseClient
dbEncerrantes.Open "select bicoencerrantes.*, fechamentodecaixa.datacaixa, fechamentodecaixa.horaini, pdvs.intermitente from bicoencerrantes, fechamentodecaixa, pdvs where bicoencerrantes.codigofechamento=fechamentodecaixa.codigofechamento and pdvs.codigopdv=fechamentodecaixa.codigopdv and pdvs.intermitente=0 and fechamentodecaixa.datacaixa=#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "# and fechamentodecaixa.horaini=#" & dbFechamentos.Recordset!HoraIni & "# and bico=" & Bico!Bico, db, adOpenForwardOnly, adLockReadOnly
If dbEncerrantes.RecordCount = 0 Then
  Bico!Encerrante = 0
  Bico!Retorno = 0
  Bico!Vendas = 0
  Bico!ValorTotal = 0
Else
  Bico!Abertura = dbEncerrantes!Abertura
  Bico!Encerrante = dbEncerrantes!Encerrante
  Bico!Retorno = dbEncerrantes!Retorno
  Bico!Vendas = dbEncerrantes!Vendas
  Bico!ValorTotal = dbEncerrantes!ValorTotal
  If IsNull(Bico!DesteCaixaQtd) = True And IsNull(Bico!DesteCaixaValor) = True And IsNull(dbEncerrantes!DesteCaixaQtd) = False And IsNull(dbEncerrantes!DesteCaixaValor) = False Then
    Bico!DesteCaixaQtd = dbEncerrantes!Vendas - dbEncerrantes!DesteCaixaQtd
    Bico!DesteCaixaValor = dbEncerrantes!ValorTotal - dbEncerrantes!DesteCaixaValor
  ElseIf Bico!DesteCaixaQtd = 0 And Bico!DesteCaixaValor = 0 And IsNull(dbEncerrantes!DesteCaixaQtd) = False And IsNull(dbEncerrantes!DesteCaixaValor) = False Then
    Bico!DesteCaixaQtd = dbEncerrantes!Vendas - dbEncerrantes!DesteCaixaQtd
    Bico!DesteCaixaValor = dbEncerrantes!ValorTotal - dbEncerrantes!DesteCaixaValor
  End If
End If

dbEncerrantes.Close


End Sub

Private Sub PegaEncerranteIntermitenteTodos(ByRef Bico As ADODB.Recordset)

Dim dbEncerrantes As New ADODB.Recordset

dbEncerrantes.CursorLocation = adUseClient
Debug.Print "select bicoencerrantes.*, fechamentodecaixa.datacaixa, fechamentodecaixa.horaini, pdvs.intermitente from bicoencerrantes, fechamentodecaixa, pdvs where bicoencerrantes.codigofechamento=fechamentodecaixa.codigofechamento and pdvs.codigopdv=fechamentodecaixa.codigopdv and pdvs.intermitente=0 and fechamentodecaixa.datacaixa=#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "# and fechamentodecaixa.horaini=#" & dbFechamentos.Recordset!HoraIni & "#"

If dbEncerrantes.State = adStateOpen Then
  dbEncerrantes.Close
End If
dbEncerrantes.Open "select bicoencerrantes.*, fechamentodecaixa.datacaixa, fechamentodecaixa.horaini, pdvs.intermitente from bicoencerrantes, fechamentodecaixa, pdvs where bicoencerrantes.codigofechamento=fechamentodecaixa.codigofechamento and pdvs.codigopdv=fechamentodecaixa.codigopdv and pdvs.intermitente=0 and fechamentodecaixa.datacaixa=#" & DataInglesa(dbFechamentos.Recordset!DataCaixa) & "# and fechamentodecaixa.horaini=#" & dbFechamentos.Recordset!HoraIni & "#", db, adOpenForwardOnly, adLockReadOnly
If dbEncerrantes.RecordCount <> 0 Then
  Bico.MoveLast
  Bico.MoveFirst
  
  Do While Bico.EOF = False
    dbEncerrantes.MoveFirst
    dbEncerrantes.Find "bico=" & Bico!Bico
    If dbEncerrantes.EOF = False Then
      Bico!Abertura = dbEncerrantes!Abertura
      Bico!Encerrante = dbEncerrantes!Encerrante
      Bico!Retorno = dbEncerrantes!Retorno
      Bico!Vendas = dbEncerrantes!Vendas
      Bico!ValorTotal = dbEncerrantes!ValorTotal
'      If IsNull(Bico!DesteCaixaQtd) = True And IsNull(Bico!DesteCaixaValor) = True And IsNull(dbEncerrantes!DesteCaixaQtd) = False And IsNull(dbEncerrantes!DesteCaixaValor) = False Then
'        Bico!DesteCaixaQtd = dbEncerrantes!Vendas - dbEncerrantes!DesteCaixaQtd
'        Bico!DesteCaixaValor = dbEncerrantes!ValorTotal - dbEncerrantes!DesteCaixaValor
'        Bico.Update
'      ElseIf Bico!DesteCaixaQtd = 0 And Bico!DesteCaixaValor = 0 And IsNull(dbEncerrantes!DesteCaixaQtd) = False And IsNull(dbEncerrantes!DesteCaixaValor) = False Then
'        Bico!DesteCaixaQtd = dbEncerrantes!Vendas - dbEncerrantes!DesteCaixaQtd
'        Bico!DesteCaixaValor = dbEncerrantes!ValorTotal - dbEncerrantes!DesteCaixaValor
'        Bico.Update
'      Else
'        Bico.Update
'      End If
        Bico.Update
    End If
    Bico.MoveNext
  Loop
End If

dbEncerrantes.Close


End Sub

Private Sub PegaPdv()
dbPdvs.Refresh
If dbPdvs.Recordset.RecordCount = 0 Then Exit Sub
dbPdvs.Recordset.MoveFirst
If IsNull(dbFechamentos.Recordset!CodigoPdv) = False Then
dbPdvs.Recordset.Find "codigopdv=" & dbFechamentos.Recordset!CodigoPdv
If dbPdvs.Recordset.EOF = False Then
  cboPdvs.Text = dbPdvs.Recordset!Descri
End If
End If
End Sub

Private Sub TotalProdutos()

Dim dbProdutos As New ADODB.Recordset
Dim qTotalVendas As New ADODB.Recordset

Dim Produtos As Currency, Comissoes As Currency

If qTotalVendas.State = adStateOpen Then
  qTotalVendas.Close
End If
qTotalVendas.CursorLocation = adUseClient
qTotalVendas.Open "select sum(valortotal) as total, sum(valorcomissao) as comissao from venda2 where combustivel=0 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If IsNull(qTotalVendas!Total) = False Then
  Produtos = qTotalVendas!Total
Else
  Produtos = 0
End If
If IsNull(qTotalVendas!Comissao) = False Then
  Comissoes = qTotalVendas!Comissao
Else
  Comissoes = 0
End If
qTotalVendas.Close


lblComissoes.Caption = Format(Comissoes, "Currency")
lblComissoes2.Caption = Format(Comissoes, "Currency")

lblTotalProdutos.Caption = Format(Produtos, "Currency")
lblTotalProdutos2.Caption = Format(Produtos, "Currency")


If ComissaoAcumulativa = True Then
  lblFaturamento.Caption = Format(CCur(lblTotalCombustivel.Caption) + Produtos, "Currency")
Else
  lblFaturamento.Caption = Format(CCur(lblTotalCombustivel.Caption) + Produtos - Comissoes, "Currency")
End If


If txtInformado.Text <> "" Then
  If IsNumeric(txtInformado.Text) = False Then
    txtInformado.Text = Format(0, "Currency")
  End If
Else
  txtInformado.Text = Format(0, "Currency")
End If

lblDiferenca.Caption = Format(CCur(txtInformado.Text) - CCur(lblFaturamento.Caption), "Currency")

With dbFechamentos
  If .Recordset!fechado = False Then
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
    .Recordset.Update
  End If
End With



End Sub

Private Sub TotalCombustivel()

Dim qTotalCombustivel As New ADODB.Recordset
Dim qTotalVendas As New ADODB.Recordset


Dim Combustivel As Currency, QtdCombustivel




If qTotalCombustivel.State = adStateOpen Then
  qTotalCombustivel.Close
End If
qTotalCombustivel.CursorLocation = adUseClient
qTotalCombustivel.Open "select sum(destecaixaqtd) as QTD, sum(destecaixavalor) as Total from BicoEncerrantes where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly

If IsNull(qTotalCombustivel!Total) = False Then
  Combustivel = qTotalCombustivel!Total
Else
  Combustivel = 0
End If
If IsNull(qTotalCombustivel!Qtd) = False Then
  QtdCombustivel = qTotalCombustivel!Qtd
Else
  QtdCombustivel = 0
End If


qTotalCombustivel.Close


lblTotalCombustivel.Caption = Format(Combustivel, "Currency")
lblTotalValorComb.Caption = Format(Combustivel, "Currency")
lbltotalQtdComb.Caption = Format(QtdCombustivel, "0.00")

If ComissaoAcumulativa = True Then
  lblFaturamento.Caption = Format(Combustivel + CCur(lblTotalProdutos.Caption), "Currency")
Else
  lblFaturamento.Caption = Format(Combustivel + CCur(lblTotalProdutos.Caption) - CCur(lblComissoes.Caption), "Currency")
End If

If txtInformado.Text <> "" Then
  If IsNumeric(txtInformado.Text) = False Then
    txtInformado.Text = Format(0, "Currency")
  End If
Else
  txtInformado.Text = Format(0, "Currency")
End If

lblDiferenca.Caption = Format(CCur(txtInformado.Text) - CCur(lblFaturamento.Caption), "Currency")

With dbFechamentos
  If .Recordset!fechado = False Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
    .Recordset.Update
  End If
End With



End Sub

Private Sub CarregaAdos()



With dbPdvs
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from pdvs order by descri"
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from Turnos order by horaini"
  .Refresh
End With
With dbVendedores
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from vendedores order by nome"
  .Refresh
End With
With dbFechamentos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from fechamentodecaixa where codigofechamento=0"
  .Refresh
End With
With dbEncerrantes
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from bicoEncerrantes where codigofechamento=0"
  .Refresh
End With
With dbVendas
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from venda2 where codigofechamento=0"
  .Refresh
End With
With dbDifComb
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from diferencacombustivel where codigofechamento=0"
  .Refresh
End With
With qProdutosAltera
  .ConnectionString = CaminhoADO
  .RecordSource = "select produtosalteradetalhe.*, produtos.* from produtosalteradetalhe, produtos where produtosalteradetalhe.codigoproduto=produtos.codigoproduto order by produtos.codigo"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .RecordSource = "select *from produtos"
    .Refresh
    cboProdutos.ListField = "Descri"
    SemTabelaDePrecos = True
  Else
    SemTabelaDePrecos = False
  End If
End With
With dbErros
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from importacaoerros where codigofechamento=0"
  .Refresh
End With
On Error Resume Next
dbResponsavel2.CursorLocation = adUseClient
dbResponsavel2.Open "Select *from vendedores order by codigo", db
On Error GoTo 0

End Sub

Private Sub TotalVenda()
Dim Unitario As Currency, Qtd As Double, Desconto As Currency, Total As Currency

With qProdutosAltera
  .Refresh
  If IsNumeric(txtQtd.Text) = False Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  
  If SemTabelaDePrecos = False Then
    .Recordset.Find "produtos.codigo=" & txtCodProduto.Text
    If .Recordset.EOF = True Then Exit Sub
    If .Recordset("produtos.Codigo") <> txtCodProduto.Text Then Exit Sub
    Unitario = .Recordset("produtosalteradetalhe.PrecoVenda")
  Else
    .Recordset.Find "codigo=" & txtCodProduto.Text
    If .Recordset.EOF = True Then Exit Sub
    If .Recordset("Codigo") <> txtCodProduto.Text Then Exit Sub
    Unitario = .Recordset("PrecoVenda")
  End If
  If IsNumeric(txtDesconto.Text) = True Then
    Desconto = CCur(txtDesconto.Text)
  End If
End With
Qtd = CDbl(txtQtd.Text)
Total = (Qtd * Unitario) + Desconto

lblTotalVenda.Caption = Format(Total, "Currency")

End Sub


Private Sub cboPdvs_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboPdvs_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
End Select
End Sub

Private Sub cboPdvs_LostFocus()
Me.KeyPreview = True
With dbPdvs
  .Refresh
  On Error Resume Next
  If cboPdvs.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboPdvs.Text & "'"
  If .Recordset.EOF = False Then
    cboPdvs.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboProdutos_LostFocus()
lblEstoque.Caption = ""
With qProdutosAltera
  .Refresh
  If cboProdutos.Text = "" Then Exit Sub
  If SemTabelaDePrecos = False Then
    .Recordset.Find "produtos.descri='" & cboProdutos.Text & "'"
    If .Recordset.EOF = False Then
      lblEstoque.Caption = .Recordset!Estoque
      txtCodProduto.Text = .Recordset("produtos.codigo")
    End If
  Else
      .Recordset.Find "descri='" & cboProdutos.Text & "'"
      If .Recordset.EOF = False Then
      lblEstoque.Caption = .Recordset!Estoque
      txtCodProduto.Text = .Recordset("codigo")
    End If
  End If
  
End With
End Sub

Private Sub cboResponsavel_LostFocus()
If cboResponsavel.Text = "" Then Exit Sub
With dbVendedores
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboResponsavel.Text & "'"
  If .Recordset.EOF = False Then
    dbFechamentos.Recordset!responsavel = .Recordset!Nome
    dbFechamentos.Recordset!Codigoresponsavel = .Recordset!codigovendedor
    dbFechamentos.Recordset.Update
  End If
End With
End Sub

Private Sub cboTurno_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboTurno_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    On Error Resume Next
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
End Select
End Sub

Private Sub cboTurno_LostFocus()
Me.KeyPreview = True
With dbTurnos
  .Refresh
  If cboTurno.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = False Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdAbrir_Click()
'Hora1 = Now

AbreCaixa

'Hora2 = Now
'
'Debug.Print DateDiff("s", Hora1, Hora2)

End Sub

Private Sub cmdAnterior_Click()
With dbFechamentos
  On Error Resume Next
  A = .Recordset!Sequencia
  .Recordset.Filter = ""
  .Recordset.Find "sequencia=" & A
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.EOF = False And .Recordset.BOF = False Then
      .Recordset.MovePrevious
      If .Recordset.BOF = False Then
        txtData.Value = .Recordset!DataCaixa
        cboTurno = .Recordset!Turno
        PegaPdv
        Call cmdAbrir_Click
      End If
    End If
  End If
End With
End Sub

Private Sub cmdCalcular_Click()


Dim dbProdutos As New ADODB.Recordset
Dim qTotalCombustivel As New ADODB.Recordset
Dim qTotalVendas As New ADODB.Recordset
Dim qDifComb As New ADODB.Recordset
Dim dbNotasCorpo As New ADODB.Recordset
Dim dbPostos As New ADODB.Recordset
Dim dbVendeTanque As New ADODB.Recordset


Dim TanqueNegativo As Boolean
Dim Combustivel As Currency, QtdCombustivel, Produtos As Currency, Comissoes As Currency
Dim TempValor As Double, TempValor2 As Double, VendaTanque As Double
Dim TempValor3 As Double, B As Double
Dim TempComissao As Currency, ComissaoCombustivel As Currency


On Error GoTo 0



If dbFechamentos.Recordset.EOF = True Then Exit Sub
If dbFechamentos.Recordset.BOF = True Then Exit Sub

If dbFechamentos.Recordset!fechado = False Then

  db.Execute "update bicoencerrantes set retorno=0 where retorno is null"
  
  db.Execute "update bicoencerrantes set vendas=(encerrante-abertura-retorno), valortotal=(encerrante-abertura-retorno)*preco where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
  
  If dbPdvs.Recordset.RecordCount = 1 Then
    db.Execute "update bicoencerrantes set destecaixaqtd=vendas, destecaixavalor=valortotal where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
    db.Execute "update bicoencerrantes set deoutrocaixaqtd=0, deoutrocaixavalor=0 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
  Else
    'db.Execute "update bicoencerrantes set deoutrocaixaqtd=0 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento & " and deoutrocaixaqtd is null"
    'db.Execute "update bicoencerrantes set deoutrocaixavalor=0 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento & " and deoutrocaixavalor is null"
    db.Execute "update bicoencerrantes set destecaixaqtd=0 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento & " and destecaixaqtd is null"
    db.Execute "update bicoencerrantes set destecaixavalor=0 where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento & " and destecaixavalor is null"
    db.Execute "update bicoencerrantes set deoutrocaixavalor=valortotal-destecaixavalor where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
    db.Execute "update bicoencerrantes set destecaixavalor=destecaixaqtd*preco where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
    db.Execute "update bicoencerrantes set deoutrocaixaqtd=vendas-destecaixaqtd where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
  End If
  
  
  
End If
dbEncerrantes.Refresh

If dbProdutos.State = adStateOpen Then
  dbProdutos.Close
End If
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "select codigoproduto, codigo, descri, comissao, comissaovalor from produtos where combustivel=-1", db, adOpenKeyset, adLockOptimistic

ComissaoCombustivel = 0

With dbEncerrantes
    If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
            TempComissao = 0
            If dbProdutos.RecordCount <> 0 Then
                dbProdutos.MoveFirst
                dbProdutos.Find "codigoproduto=" & .Recordset!CodigoProduto
                If dbProdutos.EOF = False Then
                    If dbProdutos!Comissao <> 0 Then
                        TempComissao = .Recordset!DesteCaixaValor * dbProdutos!Comissao
                    End If
                    If dbProdutos!ComissaoValor <> 0 Then
                        TempComissao = TempComissao + (.Recordset!DesteCaixaQtd * dbProdutos!ComissaoValor)
                    End If
                End If
            End If
            .Recordset!Comissao = TempComissao
            .Recordset.Update
            .Recordset.MoveNext
        Loop
    End If
End With

CalculaComissaoCombustivel
If qTotalCombustivel.State = adStateOpen Then
  qTotalCombustivel.Close
End If
qTotalCombustivel.CursorLocation = adUseClient
qTotalCombustivel.Open "select sum(valorcomissao) as Total from venda2 where combustivel=-1 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly

If IsNull(qTotalCombustivel!Total) = False Then
  ComissaoCombustivel = qTotalCombustivel!Total
End If
qTotalCombustivel.Close

qTotalCombustivel.CursorLocation = adUseClient
qTotalCombustivel.Open "select sum(destecaixaqtd) as QTD, sum(destecaixavalor) as Total from BicoEncerrantes where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly

If IsNull(qTotalCombustivel!Total) = False Then
  Combustivel = qTotalCombustivel!Total
Else
  Combustivel = 0
End If
If IsNull(qTotalCombustivel!Qtd) = False Then
  QtdCombustivel = qTotalCombustivel!Qtd
Else
  QtdCombustivel = 0
End If


qTotalCombustivel.Close

'Verifica se o último número do bico está correto
TanqueNegativo = False
If qTotalVendas.State = adStateOpen Then
  qTotalVendas.Close
End If
qTotalVendas.CursorLocation = adUseClient
qTotalVendas.Open "select sum(valortotal) as total, sum(valorcomissao) as comissao from venda2 where combustivel=0 and codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If IsNull(qTotalVendas!Total) = False Then
  Produtos = qTotalVendas!Total
Else
  Produtos = 0
End If
If IsNull(qTotalVendas!Comissao) = False Then
  Comissoes = qTotalVendas!Comissao
Else
  Comissoes = 0
End If
qTotalVendas.Close



CalculaDifComb



dbDifComb.RecordSource = "select *from DiferencaCombustivel where codigofechamento=" & dbFechamentos.Recordset!CodigoFechamento
dbDifComb.Refresh

lblComissoesCombustiveis.Caption = "R$ " & Format(ComissaoCombustivel, "0.0000")
lblComissoes.Caption = Format(Comissoes, "Currency")
lblComissoes2.Caption = Format(Comissoes, "Currency")
lblTotalCombustivel.Caption = Format(Combustivel, "Currency")
lblTotalValorComb.Caption = Format(Combustivel, "Currency")
lbltotalQtdComb.Caption = Format(QtdCombustivel, "0.00")
lblTotalProdutos.Caption = Format(Produtos, "Currency")
lblTotalProdutos2.Caption = Format(Produtos, "Currency")
If ComissaoAcumulativa = True Then
  lblFaturamento.Caption = Format(Combustivel + Produtos, "Currency")
Else
  lblFaturamento.Caption = Format(Combustivel + Produtos + Comissoes + ComissaoCombustivel, "Currency")
End If
If txtInformado.Text <> "" Then
  If IsNumeric(txtInformado.Text) = False Then
    txtInformado.Text = Format(0, "Currency")
  End If
Else
  txtInformado.Text = Format(0, "Currency")
End If
If ComissaoAcumulativa = True Then
    lblDiferenca.Caption = Format(CCur(txtInformado.Text) - CCur(lblFaturamento.Caption), "Currency")
Else
    lblDiferenca.Caption = Format(CCur(txtInformado.Text) - CCur(lblFaturamento.Caption) + Comissoes + ComissaoCombustivel, "Currency")
End If
If TanqueNegativo = False Then
  'ErroNaSoma = False
End If
With dbFechamentos
  If .Recordset!fechado = False Then
    .Recordset!TotalCombustivel = CCur(lblTotalCombustivel.Caption)
    .Recordset!TotalProdutos = CCur(lblTotalProdutos.Caption)
    .Recordset.Update
  End If
End With

dbDifComb.Refresh



End Sub

Public Sub cmdCancelar_Click()

With dbFechamentos
  If .Recordset.BOF = False And .Recordset.EOF = False Then
    On Error Resume Next
    .Recordset!responsavel = cboResponsavel.Text
    .Recordset.Update
    On Error GoTo 0
  End If
End With

Frame1.Visible = False
SSTab1.Visible = False
cboResponsavel.Visible = False
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False

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

Dim dbFechados As New ADODB.Recordset
Dim Estatus As New frmEstatus2

FinalizaNotaDoCaixa

If dbFechados.State = adStateOpen Then
  dbFechados.Close
End If
dbFechados.CursorLocation = adUseClient
dbFechados.Open "select codigofechamento, datacaixa, turno, horaini, fechado, codigopdv from fechamentodecaixa order by datacaixa, horaini, codigopdv", db, adOpenKeyset, adLockOptimistic

With dbFechamentos
  CodigoFechamento = .Recordset!CodigoFechamento
  If dbFechados.RecordCount <> 0 Then
    dbFechados.MoveFirst
    dbFechados.Find "codigofechamento=" & CodigoFechamento
    dbFechados.MovePrevious
    If dbFechados.BOF = False Then
      If dbFechados!fechado = False Then
        Resposta = MsgBox("Existe fechamento anterior para ser confirmado! Deseja confirmar em lote?", vbYesNo + vbDefaultButton2)
        If Resposta = vbNo Then Exit Sub
        FechandoLote = True
        dbFechados.MoveFirst
        dbFechados.Find "fechado=0"
        TempFechamento = CodigoFechamento
        If dbFechados.EOF = True Then Exit Sub
        Do While TempFechamento <> dbFechados!CodigoFechamento
          DoEvents
          txtData.Value = dbFechados!DataCaixa
          cboTurno.Text = dbFechados!Turno
          cboPdvs.Text = EncontraPdv(dbFechados!CodigoPdv)
          Call cmdAbrir_Click
          Me.Refresh
          If FecharCaixa = False Then
            FechandoLote = False
            Exit Sub
          End If
          Call cmdCancelar_Click
          DoEvents
          Load Estatus
          Unload Estatus
          mdiPosto.StatusBar1.Refresh
          DoEvents
          dbFechados.MoveNext
        Loop
        FechandoLote = False
        Exit Sub
      End If
    End If
  End If
  .Recordset.MoveFirst
  .Recordset.Find "codigofechamento=" & CodigoFechamento
  txtData.Value = .Recordset!DataCaixa
  cboTurno.Text = .Recordset!Turno
  FecharCaixa
End With

End Sub

Private Sub cmdDesconfirmar_Click()
Dim Resposta As Integer, TempFechamento As Double

Dim dbFechados As New ADODB.Recordset
Dim Estatus As New frmEstatus2


CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
TempFechamento = dbFechamentos.Recordset!CodigoFechamento

With dbFechamentos
  If .Recordset!fechames = True Then
    MsgBox "Este caixa pertence a mês já fechado!"
    Exit Sub
  End If
  .Recordset.Filter = ""
  .Recordset.Find "codigofechamento=" & CodigoFechamento
  .Recordset.MoveNext
  If .Recordset.EOF = False Then
    If .Recordset!fechado = True Then
      Resposta = MsgBox("Existe fechamento posterior confirmado! Deseja cancelar o fechamento em lote?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      Call cmdCancelar_Click
      FechandoLote = True
      
      dbFechados.CursorLocation = adUseClient
      If dbFechados.State = adStateOpen Then
        dbFechados.Close
      End If
      dbFechados.Open "select codigofechamento, datacaixa, turno, horaini, fechado, CodigoPdv from fechamentodecaixa order by datacaixa desc, horaini desc, codigopdv desc", db, adOpenKeyset, adLockOptimistic
      If dbFechados.RecordCount = 0 Then Exit Sub
      dbFechados.Find "fechado=-1"
      If dbFechados.EOF = True Then Exit Sub
      Do While TempFechamento <> dbFechados!CodigoFechamento
        DoEvents
        txtData.Value = dbFechados!DataCaixa
        cboTurno.Text = dbFechados!Turno
        dbPdvs.Recordset.MoveFirst
        dbPdvs.Recordset.Find "codigopdv=" & dbFechados!CodigoPdv
        If dbPdvs.Recordset.EOF = True Then
          MsgBox "Erro na tabela de Pdvs!"
          Exit Sub
        End If
        cboPdvs.Text = dbPdvs.Recordset!Descri
        Call cmdAbrir_Click
        DoEvents
        Me.Refresh
        mdiPosto.StatusBar1.Refresh
        If Desconfirmar() = False Then Exit Sub
        Call cmdCancelar_Click
        Load Estatus
        Unload Estatus
        Me.Refresh
        mdiPosto.StatusBar1.Refresh
        dbFechados.MoveNext
      Loop
      'If Desconfirmar() = False Then Exit Sub
      Exit Sub
    End If
  End If
  .Recordset.MoveFirst
  .Recordset.Find "codigofechamento=" & CodigoFechamento
  'If .Recordset!fechado = False Then
  '  MsgBox "Este caixa não está finalizado!"
  '  Exit Sub
  'End If
End With


Desconfirmar

FechandoLote = False
Load Estatus
Unload Estatus

End Sub

Private Sub cmdEntraCombustivel_Click()

FinalizaNotaDoCaixa

'Load frmFechamentoConfirmaEntrada
'frmFechamentoConfirmaEntrada.CodigoCaixa = dbFechamentos.Recordset!CodigoFechamento
'frmFechamentoConfirmaEntrada.Show vbModal
Call cmdCalcular_Click

End Sub

Private Sub cmdExibeComissoes_Click()
Load frmFechamentoDeCaixaComissoes
With frmFechamentoDeCaixaComissoes
  .CodigoFechamento = dbFechamentos.Recordset!CodigoFechamento
  .Filtra
  .Show vbModal
End With
End Sub

Private Sub cmdExtornaTurno_Click()

ExtornaNotaDoCaixa

'Load frmFechamentoConfirmaEntrada
'frmFechamentoConfirmaEntrada.CodigoCaixa = dbFechamentos.Recordset!CodigoFechamento
'frmFechamentoConfirmaEntrada.Show vbModal
Call cmdCalcular_Click

End Sub

Private Sub cmdImportar_Click()
frmFechamentoDeCaixaImportacao.Show
End Sub

Private Sub cmdIncluir_Click()
IncluirProduto
End Sub

Private Sub cmdPosterior_Click()
With dbFechamentos
  On Error Resume Next
  A = .Recordset!Sequencia
  .Recordset.Filter = ""
  .Recordset.Find "sequencia=" & A
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.EOF = False Then
      .Recordset.MoveNext
      If .Recordset.EOF = False Then
        txtData.Value = .Recordset!DataCaixa
        cboTurno = .Recordset!Turno
        PegaPdv
        Call cmdAbrir_Click
      End If
    End If
  End If
End With
End Sub

Private Sub cmdPrimeiro_Click()
With dbFechamentos
  On Error Resume Next
  A = .Recordset!Sequencia
  .Recordset.Filter = ""
  .Recordset.Find "sequencia=" & A
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    If .Recordset.BOF = False Then
      txtData.Value = .Recordset!DataCaixa
      cboTurno = .Recordset!Turno
      PegaPdv
      Call cmdAbrir_Click
    End If
  End If
End With
End Sub

Private Sub cmdRemover_Click()
RemoveEspecial
End Sub

Private Sub cmdRemoverProduto_Click()
Dim Resposta As Integer
With dbVendas
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um produto primeiro!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover o produto atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  On Error Resume Next
  .Recordset.Delete adAffectCurrent
  .Refresh
  .Refresh
  TotalProdutos
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTransfere_Click()

frmFechamentoDeCaixaTransfere.Show vbModal

CalculaDifComb

End Sub

Private Sub cmdUltimo_Click()
With dbFechamentos
  On Error Resume Next
  A = .Recordset!Sequencia
  .Refresh
  .Recordset.Filter = ""
  .Recordset.Find "sequencia=" & A
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    If .Recordset.EOF = False Then
      txtData.Value = .Recordset!DataCaixa
      cboTurno = .Recordset!Turno
      PegaPdv
      Call cmdAbrir_Click
    End If
  End If
End With
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
CalculaBicos ColIndex
End Sub

Private Sub DataGrid1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    On Error Resume Next
    DataGrid1.Row = DataGrid1.Row + 1
    If Err.Number <> 0 Then
      dbEncerrantes.Recordset.MoveFirst
      On Error GoTo 0
      On Error Resume Next
      DataGrid1.Col = DataGrid1.Col + 1
      If Err.Number <> 0 Then
        SendKeys Chr(vbKeyTab)
      End If
    End If
End Select
End Sub

Private Sub DataGrid1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub DataGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
With dbDifComb
  .Recordset!Tanque = CDbl(DataGrid3.Columns(ColIndex).Text)
  .Recordset!Diferenca = .Recordset!Tanque - .Recordset!Estoque
  .Recordset.Update
End With
End Sub

Private Sub DataGrid3_GotFocus()
Me.KeyPreview = False
Call cmdCalcular_Click
End Sub

Private Sub DataGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    On Error Resume Next
    DataGrid3.Row = DataGrid3.Row + 1
    If Err.Number <> 0 Then
      dbDifComb.Recordset.MoveFirst
      On Error GoTo 0
      On Error Resume Next
      DataGrid3.Col = DataGrid3.Col + 1
      If Err.Number <> 0 Then
        SendKeys Chr(vbKeyTab)
      End If
    End If
End Select
End Sub

Private Sub DataGrid3_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'tratamento de teclas
Select Case KeyCode
  Case vbKeyF5
    If Shift = 1 Then
      'PegaCupons
    Else
      If cmdImportar.Visible = True Then
        'Call cmdImportar_Click
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
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyF5
    'Call cmdImportar_Click
End Select
End Sub

Private Sub Form_Load()
Dim PermiteDesconto As String
Dim dbBicos As New ADODB.Recordset

If Usuarios.Grupo.AdmEstatus = 2 Then
  cmdDesconfirmar.Visible = True
Else
  cmdDesconfirmar.Visible = False
End If



On Error Resume Next
db.Open CaminhoADO
On Error GoTo 0

VerificaConexao

db.Execute "update fechamentodecaixa set codigopdv=1 where codigopdv is null"

dbBicos.CursorLocation = adUseClient
If dbBicos.State = adStateOpen Then
    dbBicos.Close
End If
dbBicos.Open "select *from bicos order by bico", db, adOpenForwardOnly, adLockReadOnly

If dbBicos.RecordCount <> 0 Then
  SSTab1.TabVisible(0) = True
  SSTab1.TabVisible(2) = True
Else
  SSTab1.TabVisible(0) = False
  SSTab1.TabVisible(2) = False
End If

dbBicos.Close


PermiteDesconto = ReadINI("Fechamento", "Desc", 0, App.Path & "\posto.ini")

If PermiteDesconto = "1" Then
  lblDesconto.Visible = True
  txtDesconto.Visible = True
Else
  lblDesconto.Visible = False
  txtDesconto.Visible = False
End If



CarregaAdos

txtData.Value = Date
FechamentoAnterior = -1

With dbTanques2
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from tanques"
  .Refresh
End With

With dbPdvs
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from pdvs order by descri"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If IsNull(.Recordset!Descri) = False Then
      cboPdvs.Text = .Recordset!Descri
    End If
  End If
End With

With dbFechamentos
  .RecordSource = "select *from fechamentodecaixa where fechado=0 order by datacaixa, horaini, codigopdv"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    txtData.Value = .Recordset!DataCaixa
    cboTurno.Text = .Recordset!Turno
    PegaPdv
  End If
End With

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
    cmdExtornaTurno.Enabled = False
    cmdConfirmar.Enabled = False
    DataGrid1.AllowUpdate = False
    DataGrid2.AllowUpdate = False
    cmdIncluir.Enabled = False
    DataGrid3.AllowDelete = False
    cboResponsavel.Enabled = False
    txtInformado.Enabled = False
  Case 2 'Liberado
    
End Select

On Error Resume Next
cboPdvs.SetFocus
On Error GoTo 0

End Sub

Private Sub Form_Terminate()
db.Close
End Sub

Private Sub txtCodProduto_GotFocus()
With txtCodProduto
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodProduto_LostFocus()
lblEstoque.Caption = ""
With qProdutosAltera
  '.Refresh
  If txtCodProduto.Text = "" Then Exit Sub
  If IsNumeric(txtCodProduto.Text) = False Then Exit Sub
  If SemTabelaDePrecos = False Then
    .Recordset.Find "produtos.codigo=" & txtCodProduto.Text
    If .Recordset.EOF = False Then
      lblEstoque.Caption = .Recordset!Estoque
      cboProdutos.Text = .Recordset("produtos.descri")
    Else
      cboProdutos.Text = ""
    End If
    
  Else
    .Recordset.Find "codigo=" & txtCodProduto.Text
    If .Recordset.EOF = False Then
      lblEstoque.Caption = .Recordset!Estoque
      cboProdutos.Text = .Recordset("descri")
    Else
      cboProdutos.Text = ""
    End If
    
  End If
End With
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDesconto_LostFocus()
TotalVenda
End Sub

Private Sub txtInformado_GotFocus()
With txtInformado
    .SelStart = 0
    .SelLength = Len(.Text)
End With
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

