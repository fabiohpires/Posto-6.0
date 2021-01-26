VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   7545
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "frmProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11535
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   145
      Top             =   7200
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "ADO"
      Height          =   6975
      Left            =   8040
      TabIndex        =   146
      Top             =   7320
      Visible         =   0   'False
      Width           =   3735
      Begin MSAdodcLib.Adodc dbCategoria 
         Height          =   330
         Left            =   120
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from produtoscategoria order by categoria"
         Caption         =   "dbCategoria"
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
      Begin MSAdodcLib.Adodc dbProdutosSubCategoria 
         Height          =   330
         Left            =   120
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from ProdutosSubCategoria order by descri"
         Caption         =   "dbProdutosSubCategoria"
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
      Begin MSAdodcLib.Adodc dbProdutosBarras 
         Height          =   330
         Left            =   120
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "ProdutosCodigos"
         Caption         =   "dbProdutosBarras"
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
      Begin MSAdodcLib.Adodc dbEstacionamento 
         Height          =   330
         Left            =   120
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "Estacionamento"
         Caption         =   "dbEstacionamento"
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
         Left            =   120
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from produtos"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   120
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from postos order by nome"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   120
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from status"
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc dbAliquotas 
         Height          =   330
         Left            =   120
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from Aliquotas"
         Caption         =   "dbAliquotas"
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
      Begin MSAdodcLib.Adodc dbProdutosHistorico 
         Height          =   330
         Left            =   120
         Top             =   3120
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from ProdutosHistorico"
         Caption         =   "dbProdutosHistorico"
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
      Begin MSAdodcLib.Adodc dbAcerto 
         Height          =   330
         Left            =   120
         Top             =   3480
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from produtosacerto"
         Caption         =   "dbAcerto"
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
      Begin MSAdodcLib.Adodc dbTanque 
         Height          =   330
         Left            =   120
         Top             =   3840
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from tanques"
         Caption         =   "dbTanque"
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
      Begin MSAdodcLib.Adodc dbProdutos2 
         Height          =   330
         Left            =   120
         Top             =   4200
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from produtos"
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
      Begin MSAdodcLib.Adodc dbVenda2 
         Height          =   330
         Left            =   120
         Top             =   4560
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from venda2 where codigoproduto=0"
         Caption         =   "dbVenda2"
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
      Begin MSAdodcLib.Adodc dbFechamento 
         Height          =   330
         Left            =   120
         Top             =   4920
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from fechamentodecaixa where fechado=-1 order by datacaixa desc, horaini desc"
         Caption         =   "dbFechamento"
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
      Begin MSAdodcLib.Adodc dbGruposIF 
         Height          =   330
         Left            =   120
         Top             =   5280
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select *from produtosgrupoif order by codigogrupo"
         Caption         =   "dbGruposIF"
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
         Top             =   5640
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "Select *from turnos order by horaini"
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
      Begin MSAdodcLib.Adodc dbCST 
         Height          =   330
         Left            =   120
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "Select *from ProdutosClassFisc order by codigoclass"
         Caption         =   "dbCST"
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
   Begin VB.ComboBox cboRelatorios 
      Height          =   315
      ItemData        =   "frmProdutos.frx":0442
      Left            =   4440
      List            =   "frmProdutos.frx":0458
      TabIndex        =   129
      Top             =   6360
      Width           =   3855
   End
   Begin VB.CommandButton cmdConferencia 
      Caption         =   "Confere Estoque"
      Height          =   735
      Left            =   9480
      TabIndex        =   131
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdTrocaCodigo 
      Caption         =   "Troca Codigo em Lote"
      Height          =   735
      Left            =   10560
      TabIndex        =   132
      Top             =   6120
      Width           =   855
   End
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   8040
      Top             =   5280
      Visible         =   0   'False
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
      RecordSource    =   "select *from contas order by descri"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmProdutos.frx":04CF
      Height          =   315
      Left            =   600
      TabIndex        =   147
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   7170
      Width           =   11535
      _ExtentX        =   20346
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
      RecordSource    =   "select *from produtos"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   " Estacionamento "
      Height          =   615
      Left            =   120
      TabIndex        =   123
      Top             =   6120
      Width           =   4095
      Begin MSMask.MaskEdBox txtPrecoEstacionamento 
         DataField       =   "Preco"
         DataSource      =   "dbEstacionamento"
         Height          =   300
         Left            =   3120
         TabIndex        =   127
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtUltimoNumero 
         Alignment       =   1  'Right Justify
         DataField       =   "UltimoNumero"
         DataSource      =   "dbEstacionamento"
         Height          =   285
         Left            =   1440
         TabIndex        =   125
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Preço:"
         Height          =   255
         Left            =   2520
         TabIndex        =   126
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Ultimo Número:"
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   7560
      TabIndex        =   143
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFonte 
      Caption         =   "Fonte"
      Height          =   375
      Left            =   4920
      TabIndex        =   141
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimeTabela 
      Caption         =   "Imprime Tabela"
      Height          =   375
      Left            =   6000
      TabIndex        =   142
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3600
      TabIndex        =   140
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8400
      Picture         =   "frmProdutos.frx":04E4
      Style           =   1  'Graphical
      TabIndex        =   130
      Tag             =   "Imprimir"
      Top             =   6120
      Width           =   735
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11535
      TabIndex        =   144
      Top             =   6840
      Width           =   11535
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   3608
         TabIndex        =   134
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4695
         TabIndex        =   135
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   5798
         TabIndex        =   136
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   6893
         TabIndex        =   137
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   7988
         TabIndex        =   138
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   2528
         TabIndex        =   133
         Top             =   0
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmProdutos.frx":0F66
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados"
      TabPicture(1)   =   "frmProdutos.frx":0F82
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dados Fiscais"
      TabPicture(2)   =   "frmProdutos.frx":0F9E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Acerto de Estoque"
      TabPicture(3)   =   "frmProdutos.frx":0FBA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label11"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblLabels(10)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblTanqueAcerto"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cboTurnos"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtDataCaixa"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Text2"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cmdAlteraPreco"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtPrecoNovo"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cmdAcerto"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "txtAcerto"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtTanqueAcerto"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Categorias / Barras"
      TabPicture(4)   =   "frmProdutos.frx":0FD6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label8"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "DataGrid3"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "DataGrid2"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdImportarGrupos"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtCodigoBarra"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdIncluirBarras"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cmdRemoverBarras"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtCodSubCategoria"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtSubCategoria"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cmdIncluirSubCategoria"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "cmdRemoverSubCategoria"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).ControlCount=   13
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   4815
         Left            =   -74880
         TabIndex        =   149
         Top             =   360
         Width           =   11055
         Begin VB.TextBox txtFields 
            DataField       =   "Categoria"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   8
            Left            =   120
            MaxLength       =   15
            TabIndex        =   27
            Top             =   1680
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "LucroMinimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   9
            Left            =   6600
            TabIndex        =   151
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "UnCaixa"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   7
            Left            =   9960
            TabIndex        =   12
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Combustível"
            DataField       =   "Combustivel"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   5160
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Comissao"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   5
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   5
            Left            =   4440
            TabIndex        =   22
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataField       =   "PrecoVenda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   4
            Left            =   3600
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PrecoCompra"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2760
            TabIndex        =   18
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Estoque"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   9000
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Descri"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   5
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            DataField       =   "Codigo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   15
            TabIndex        =   3
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            DataField       =   "DescriAbreviada"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   10
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   8
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Permite entrada no Caixa"
            DataField       =   "PermiteNoCaixa"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "DuracaoEstoque"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   6
            Left            =   120
            MaxLength       =   15
            TabIndex        =   14
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "EstoqueIdeal"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   12
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ComissaoValor"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$ ""#.##0,0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   20
            Left            =   5400
            TabIndex        =   24
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Gerar LMC"
            DataField       =   "LMC"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   2400
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "frmProdutos.frx":0FF2
            DataField       =   "Categoria"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   120
            TabIndex        =   150
            Top             =   1680
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            ListField       =   "Categoria"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmProdutos.frx":100C
            DataField       =   "SubCategoria"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   3000
            TabIndex        =   29
            Top             =   1680
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            ListField       =   "Descri"
            BoundColumn     =   "codigoSubCategoria"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Lucro Min:"
            Height          =   195
            Index           =   12
            Left            =   6600
            TabIndex        =   25
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Categoria:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Un./Caixa:"
            Height          =   195
            Index           =   8
            Left            =   9960
            TabIndex        =   11
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Comissão %:"
            Height          =   195
            Index           =   5
            Left            =   4440
            TabIndex        =   21
            Top             =   840
            Width           =   885
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "$ Venda:"
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   19
            Top             =   840
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "$ Compra:"
            Height          =   195
            Index           =   3
            Left            =   2760
            TabIndex        =   17
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estoque:"
            Height          =   195
            Index           =   2
            Left            =   9000
            TabIndex        =   9
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Sub Categoria:"
            Height          =   195
            Index           =   6
            Left            =   3000
            TabIndex        =   28
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Abreviado:"
            Height          =   195
            Index           =   9
            Left            =   6480
            TabIndex        =   7
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Dias para Compra:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estoque ideal:"
            Height          =   195
            Index           =   14
            Left            =   1560
            TabIndex        =   15
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Comissão $:"
            Height          =   195
            Index           =   23
            Left            =   5400
            TabIndex        =   23
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   4815
         Left            =   -74880
         TabIndex        =   148
         Top             =   360
         Width           =   11055
         Begin VB.TextBox txtFields 
            DataField       =   "DescriServico"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   11
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   37
            ToolTipText     =   "Descrição do Serviço para impressão em nota"
            Top             =   480
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Serviço"
            DataField       =   "Servico"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   4680
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "codEAN"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   13
            Left            =   120
            MaxLength       =   15
            TabIndex        =   43
            ToolTipText     =   "Código de Barras do Produto"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "Origem"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmProdutos.frx":1031
            Left            =   120
            List            =   "frmProdutos.frx":103E
            TabIndex        =   95
            Top             =   4080
            Width           =   3255
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "IPI"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   14
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   73
            ToolTipText     =   "Alíquota do IPI para o item"
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PIS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   15
            Left            =   120
            MaxLength       =   15
            TabIndex        =   75
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "COFINS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   16
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   83
            ToolTipText     =   "Alíquota do COFINS para o item"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ISS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   17
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   39
            ToolTipText     =   "Alíquota do ISS para o serviço."
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ReducaoIcms"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   18
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   63
            ToolTipText     =   "Percentual de redução de base de cálculo para o ICMS"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "AliquotaICMS"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   19
            Left            =   120
            MaxLength       =   15
            TabIndex        =   55
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            DataField       =   "Unidade"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   45
            ToolTipText     =   "Unidade de medida utilizada na quantificação de estoques. "
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00C0FFFF&
            DataField       =   "TipoItem"
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "frmProdutos.frx":1086
            Left            =   3120
            List            =   "frmProdutos.frx":10AE
            TabIndex        =   47
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoNCM"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   21
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   49
            ToolTipText     =   "Código da Nomenclatura Comum do MERCOSUL"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoNCM"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   22
            Left            =   8040
            MaxLength       =   15
            TabIndex        =   51
            ToolTipText     =   "Código EX, conforme a TIPI"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoNCM"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   23
            Left            =   9360
            MaxLength       =   15
            TabIndex        =   53
            ToolTipText     =   "Código do gênero do item, conforme a Tabela 4.2.1."
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoSEFAZ"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   24
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   59
            ToolTipText     =   "Código da mercadoria para a SEFAZ de São Paulo para geração do arquivo GRF-CBT."
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoLST"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   25
            Left            =   9480
            MaxLength       =   15
            TabIndex        =   41
            ToolTipText     =   "Código do serviço conforme lista do Anexo I da Lei Complementar Federal nº 116/03."
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CSOSN"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   26
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   61
            ToolTipText     =   "Código da Situação da Operação no Simples Nacional"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "BC_ICMS_ST"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   27
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   65
            ToolTipText     =   "Valor unitário de base de cálculo do ICMS ST (valor fixado pela SEFAZ)"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "IPIEntrada"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   28
            Left            =   5640
            MaxLength       =   15
            TabIndex        =   69
            ToolTipText     =   "Código da Situação Tributária do IPI nas entradas"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "IPIEntrada"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   29
            Left            =   6960
            MaxLength       =   15
            TabIndex        =   71
            ToolTipText     =   "Código da Situação Tributária do IPI nas saídas"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PISEntrada"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   30
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   77
            ToolTipText     =   "Código da Situação Tributária do PIS nas entradas"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PISSaida"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   31
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   79
            ToolTipText     =   "Código da Situação Tributária do PIS nas saídas"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "NatRecPIS"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   32
            Left            =   3360
            MaxLength       =   15
            TabIndex        =   81
            ToolTipText     =   $"frmProdutos.frx":11C5
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "COFINSEntrada"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   33
            Left            =   6360
            MaxLength       =   15
            TabIndex        =   85
            ToolTipText     =   "Código da Situação Tributária do COFINS nas entradas"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "COFINSSaida"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   34
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   87
            ToolTipText     =   "Código da Situação Tributária do COFINS nas saídas"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "COFINSSaida"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   35
            Left            =   9240
            MaxLength       =   15
            TabIndex        =   89
            ToolTipText     =   $"frmProdutos.frx":1258
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ContaContabil"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   36
            Left            =   120
            MaxLength       =   15
            TabIndex        =   91
            ToolTipText     =   "Conta Contábil do item"
            Top             =   3480
            Width           =   2415
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ContaContabil"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   37
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   93
            ToolTipText     =   "Conta Contábil do item"
            Top             =   3480
            Width           =   6375
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "frmProdutos.frx":12EE
            DataField       =   "Aliquota"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            ListField       =   "Aliquota"
            BoundColumn     =   "Impressora"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "frmProdutos.frx":1308
            DataField       =   "PlanoDeConta"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1680
            TabIndex        =   35
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            ListField       =   "Descri"
            BoundColumn     =   "CodigoConta"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "frmProdutos.frx":131F
            DataField       =   "CST"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "Código da Situação Tributária do ICMS"
            Top             =   2280
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodigoClass"
            BoundColumn     =   "CodigoClass"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "frmProdutos.frx":1333
            DataField       =   "CodigoGrupoIF"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1320
            TabIndex        =   57
            ToolTipText     =   "Descrição do código do grupo da mercadoria para o inventário."
            Top             =   1680
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            ListField       =   "Descri"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plano de Conta:"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   34
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ICMS:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código de Barras EAN/GTIN:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   42
            Tag             =   "Código de Barra do Produto"
            Top             =   840
            Width           =   2100
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Origem:"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   94
            Top             =   3840
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "CST:"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   66
            Top             =   2040
            Width           =   360
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aliquota IPI:"
            Height          =   195
            Index           =   17
            Left            =   8280
            TabIndex        =   72
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aliquota PIS:"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   74
            Top             =   2640
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aliquota COFINS:"
            Height          =   195
            Index           =   19
            Left            =   4920
            TabIndex        =   82
            Top             =   2640
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "ISS:"
            Height          =   195
            Index           =   20
            Left            =   8400
            TabIndex        =   38
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Redução ICMS:"
            Height          =   195
            Index           =   21
            Left            =   7080
            TabIndex        =   62
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aliquota ICMS:"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label9 
            Caption         =   "Unidade:"
            Height          =   255
            Left            =   2400
            TabIndex        =   44
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Tipo do Item:"
            Height          =   255
            Left            =   3120
            TabIndex        =   46
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod NCM:"
            Height          =   195
            Index           =   24
            Left            =   6720
            TabIndex        =   48
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod EX IPI:"
            Height          =   195
            Index           =   25
            Left            =   8040
            TabIndex        =   50
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod do Gênero:"
            Height          =   195
            Index           =   26
            Left            =   9360
            TabIndex        =   52
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label10 
            Caption         =   "Grupo da Mercadoria para o Inventário:"
            Height          =   255
            Left            =   1320
            TabIndex        =   56
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod SEFAZ:"
            Height          =   195
            Index           =   27
            Left            =   4440
            TabIndex        =   58
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod Serviço:"
            Height          =   195
            Index           =   28
            Left            =   9480
            TabIndex        =   40
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cod do Simples:"
            Height          =   195
            Index           =   29
            Left            =   5760
            TabIndex        =   60
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "B. C. Red. ICMS:"
            Height          =   195
            Index           =   30
            Left            =   8400
            TabIndex        =   64
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "CST IPI Entrada:"
            Height          =   195
            Index           =   31
            Left            =   5640
            TabIndex        =   68
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "CST IPI Saida:"
            Height          =   195
            Index           =   32
            Left            =   6960
            TabIndex        =   70
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "PIS Entrada:"
            Height          =   195
            Index           =   33
            Left            =   1200
            TabIndex        =   76
            Top             =   2640
            Width           =   900
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "PIS Saida:"
            Height          =   195
            Index           =   34
            Left            =   2280
            TabIndex        =   78
            Top             =   2640
            Width           =   750
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Rec. PIS:"
            Height          =   195
            Index           =   35
            Left            =   3360
            TabIndex        =   80
            Top             =   2640
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "COFINS Entrada:"
            Height          =   195
            Index           =   36
            Left            =   6360
            TabIndex        =   84
            Top             =   2640
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "COFINS Saida:"
            Height          =   195
            Index           =   37
            Left            =   7800
            TabIndex        =   86
            Top             =   2640
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nat. Rec. COFINS:"
            Height          =   195
            Index           =   38
            Left            =   9240
            TabIndex        =   88
            Top             =   2640
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Conta Contabil:"
            Height          =   195
            Index           =   39
            Left            =   120
            TabIndex        =   90
            Top             =   3240
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
            Height          =   195
            Index           =   40
            Left            =   2640
            TabIndex        =   92
            Top             =   3240
            Width           =   915
         End
      End
      Begin VB.TextBox txtTanqueAcerto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73080
         TabIndex        =   101
         Top             =   780
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAcerto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73920
         TabIndex        =   99
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton cmdAcerto 
         Caption         =   "Acerto Estoque"
         Height          =   375
         Left            =   -69480
         TabIndex        =   106
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtPrecoNovo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   108
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdAlteraPreco 
         Caption         =   "Altera Preço"
         Height          =   375
         Left            =   -73680
         TabIndex        =   109
         Top             =   1980
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74760
         TabIndex        =   97
         Top             =   780
         Width           =   735
      End
      Begin VB.CommandButton cmdRemoverSubCategoria 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -71040
         TabIndex        =   115
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton cmdIncluirSubCategoria 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -71520
         TabIndex        =   114
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtSubCategoria 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74160
         MaxLength       =   25
         TabIndex        =   113
         Top             =   780
         Width           =   2535
      End
      Begin VB.TextBox txtCodSubCategoria 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   25
         TabIndex        =   111
         Top             =   780
         Width           =   615
      End
      Begin VB.CommandButton cmdRemoverBarras 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67200
         TabIndex        =   120
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton cmdIncluirBarras 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67680
         TabIndex        =   119
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtCodigoBarra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70200
         MaxLength       =   25
         TabIndex        =   118
         Top             =   780
         Width           =   2415
      End
      Begin VB.CommandButton cmdImportarGrupos 
         Caption         =   "Importar Grupos"
         Height          =   495
         Left            =   -66000
         TabIndex        =   122
         Top             =   2700
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmProdutos.frx":134C
         Height          =   2415
         Left            =   -70320
         TabIndex        =   121
         Top             =   1140
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4260
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "CodigoBarra"
            Caption         =   "CodigoBarra"
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
               ColumnWidth     =   3119,811
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmProdutos.frx":136B
         Height          =   2415
         Left            =   -74880
         TabIndex        =   116
         Top             =   1140
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4260
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "PreCodigo"
            Caption         =   "Pré-Codigo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2715,024
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker txtDataCaixa 
         Height          =   315
         Left            =   -72360
         TabIndex        =   103
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   132775937
         CurrentDate     =   37680
      End
      Begin MSDataListLib.DataCombo cboTurnos 
         Bindings        =   "frmProdutos.frx":1390
         Height          =   315
         Left            =   -70920
         TabIndex        =   105
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmProdutos.frx":13A7
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   14
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
            DataField       =   "Codigo"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
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
            DataField       =   "DuracaoEstoque"
            Caption         =   "Duração"
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
            DataField       =   "PrecoCompra"
            Caption         =   "$ Compra"
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
            DataField       =   "LucroMinimo"
            Caption         =   "Lucro Min."
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
            DataField       =   "Sugerido"
            Caption         =   "Sugerido"
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
            DataField       =   "PrecoVenda"
            Caption         =   "$ Venda"
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
            DataField       =   "Comissao"
            Caption         =   "Comissão"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   5
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "UnCaixa"
            Caption         =   "Un/Cx"
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
            DataField       =   "Aliquota"
            Caption         =   "Alíquota"
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
            DataField       =   "Estoque"
            Caption         =   "Estoque"
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
         BeginProperty Column11 
            DataField       =   "DifEstoque"
            Caption         =   "Dif. Estoque"
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
         BeginProperty Column12 
            DataField       =   "Departamento"
            Caption         =   "Departamento"
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
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2445,166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTanqueAcerto 
         Caption         =   "Tanque:"
         Height          =   255
         Left            =   -73080
         TabIndex        =   100
         Top             =   540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtd:"
         Height          =   195
         Left            =   -73920
         TabIndex        =   98
         Top             =   540
         Width           =   300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Novo $ Compra:"
         Height          =   195
         Index           =   10
         Left            =   -74880
         TabIndex        =   107
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   104
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Data Caixa:"
         Height          =   195
         Left            =   -72360
         TabIndex        =   102
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   96
         Top             =   540
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   112
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Código:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   110
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Código-de-Barras:"
         Height          =   255
         Left            =   -70200
         TabIndex        =   117
         Top             =   540
         Width           =   1815
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Relatórios:"
      Height          =   255
      Left            =   4440
      TabIndex        =   128
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Posto:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   139
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmCadProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigoPosto As Double, strOrdem As String, ColocaProduto As Boolean
Dim strFiltroProduto As String, StrTabela As String, Imprimindo As Boolean
Dim CodigoAntigo As Double, CodigoNovo As Double, CodigoProduto As Double

Private Sub CabecaTabela(ByVal Dia As Date, ByVal Largura As Double)
Dim StrTemp As String

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.ScaleMode = vbMillimeters


StrTemp = "Tabela de Preços"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
Printer.CurrentX = 0
Printer.Print StrTemp

For i = 0 To 2
  StrTemp = "Código"
  Printer.CurrentX = Coluna + 10 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = Coluna + 11
  Printer.Print StrTemp;
  
  StrTemp = "Preco"
  Printer.CurrentX = Coluna + 64 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  Coluna = Coluna + 65
Next i
Printer.Print ""

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub ImprimeTabela()
  Dim StrTemp As String, Largura As Double, Dia As Date
  Dim Coluna As Double, Y1 As Double, Y2 As Double
  
  With Adodc1
    StrTabela = .RecordSource
    .RecordSource = "select *from produtos" & strFiltroProduto & strOrdem
    .Refresh
    If .Recordset.RecordCount = 0 Then
      .RecordSource = StrTabela
      .Refresh
      Exit Sub
    End If
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    Printer.ScaleMode = vbMillimeters
    
    Dia = Now
    
    Largura = 195
    
    CabecaTabela Dia, Largura
    
    Y1 = Printer.CurrentY
    Coluna = 0
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        If Coluna >= 130 Then
          Printer.CurrentY = Printer.CurrentY + 1
          Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
          Printer.CurrentY = Printer.CurrentY + 1
          Printer.Line (Coluna, Y1)-(Coluna, Y2)
          Coluna = 0
          Printer.NewPage
          CabecaTabela Dia, Largura
        Else
          Y2 = Printer.CurrentY
          Printer.Line (Coluna + 0.5, Y1)-(Coluna + 0.5, Y2)
          Coluna = Coluna + 65
          Printer.CurrentY = Y1
        End If
      End If
      
      StrTemp = .Recordset("codigo")
      Printer.CurrentX = Coluna + 10 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("descri")
      Printer.CurrentX = Coluna + 11
      Printer.Print StrTemp;
      
      
      StrTemp = Format(.Recordset("precoVenda"), "Currency")
      Printer.CurrentX = Coluna + 64 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      If Y2 < Printer.CurrentY Then
        Y2 = Printer.CurrentY
      End If
      .Recordset.MoveNext
    Loop
    
    Printer.CurrentY = Y2
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    If Coluna = 130 Then
      Printer.Line (Coluna, Y1)-(Coluna, Y2)
      Coluna = Coluna + 65
    Else
      Do While Coluna <= 130
        Printer.CurrentY = Y1
        Printer.Line (Coluna + 0.5, Y1)-(Coluna + 0.5, Y2)
        Coluna = Coluna + 65
        Printer.CurrentY = Y1
      Loop
    End If
    
    Printer.Line (Coluna, Y1)-(Coluna, Y2)
    .RecordSource = StrTabela
    .Refresh
  End With
  Printer.EndDoc
  Exit Sub
NaoImprime:
  Printer.KillDoc
  Exit Sub
End Sub

Private Sub CabecaTabelaICMS(ByVal Dia As Date, ByVal Largura As Double)
Dim StrTemp As String

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.ScaleMode = vbMillimeters


StrTemp = "Tabela de ICMS"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8

StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
Printer.CurrentX = 0
Printer.Print StrTemp

For i = 0 To 2
  StrTemp = "Código"
  Printer.CurrentX = Coluna + 10 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = Coluna + 11
  Printer.Print StrTemp;
  
  StrTemp = "ICMS"
  Printer.CurrentX = Coluna + 64 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  Coluna = Coluna + 65
Next i
Printer.Print ""

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub ImprimeTabelaICMS()
  Dim StrTemp As String, Largura As Double, Dia As Date
  Dim Coluna As Double, Y1 As Double, Y2 As Double
  
  
  With Adodc1
    StrTabela = .RecordSource
    .RecordSource = "select *from produtos " & strFiltroProduto & strOrdem
    .Refresh
    If .Recordset.RecordCount = 0 Then
      .RecordSource = StrTabela
      .Refresh
      Exit Sub
    End If
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    Printer.ScaleMode = vbMillimeters
    
    Dia = Now
    
    Largura = 195
    
    CabecaTabelaICMS Dia, Largura
    
    Y1 = Printer.CurrentY
    Coluna = 0
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        If Coluna >= 130 Then
          Printer.CurrentY = Printer.CurrentY + 1
          Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
          Printer.CurrentY = Printer.CurrentY + 1
          Printer.Line (Coluna, Y1)-(Coluna, Y2)
          Coluna = 0
          Printer.NewPage
          CabecaTabelaICMS Dia, Largura
        Else
          Y2 = Printer.CurrentY
          Printer.Line (Coluna + 0.5, Y1)-(Coluna + 0.5, Y2)
          Coluna = Coluna + 65
          Printer.CurrentY = Y1
        End If
      End If
      
      StrTemp = .Recordset("codigo")
      Printer.CurrentX = Coluna + 10 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("descri")
      Printer.CurrentX = Coluna + 11
      Printer.Print StrTemp;
      
      If IsNull(.Recordset!Aliquota) = False Then
        StrTemp = Mid(.Recordset!Aliquota, 1, 2) & "," & Mid(.Recordset!Aliquota, 3, 2)
      Else
        StrTemp = ""
      End If
      Printer.CurrentX = Coluna + 64 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      If Y2 < Printer.CurrentY Then
        Y2 = Printer.CurrentY
      End If
      .Recordset.MoveNext
    Loop
    
    Printer.CurrentY = Y2
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    If Coluna = 130 Then
      Printer.Line (Coluna, Y1)-(Coluna, Y2)
      Coluna = Coluna + 65
    Else
      Do While Coluna <= 130
        Printer.CurrentY = Y1
        Printer.Line (Coluna + 0.5, Y1)-(Coluna + 0.5, Y2)
        Coluna = Coluna + 65
        Printer.CurrentY = Y1
      Loop
    End If
    
    Printer.Line (Coluna, Y1)-(Coluna, Y2)
    .RecordSource = StrTabela
    .Refresh
  End With
  Printer.EndDoc
  Exit Sub
NaoImprime:
  Printer.KillDoc
  Exit Sub
End Sub


Private Sub ImprimeCompra(ByVal Preco As Boolean)
  Dim StrTemp As String, Largura As Double, strData As String
  Dim TotalCompra As Currency, TotalVenda As Currency, TempValor As Currency
  
  With Adodc1
    StrTabela = .RecordSource
    .RecordSource = "select *from produtos " & strFiltroProduto & strOrdem
    .Refresh
    If .Recordset.RecordCount = 0 Then Exit Sub
    
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    Printer.ScaleMode = vbMillimeters
    
    strData = "Data: " & Format(Now, "long date") & " - " & Format(Now, "short time")
    
    If Preco = True Then
      Largura = 170
    Else
      Largura = 150
    End If
    Cabeca Largura, strData
    
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.NewPage
        Cabeca Largura, strData
      End If
      
      Printer.FontSize = 8
      Printer.FontBold = False
      
      StrTemp = .Recordset("codigo")
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("descri")
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      
      
      StrTemp = .Recordset("Estoque")
      Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset("precoVenda"), "Currency")
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset("comissao"), "0.00%")
      Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset("comissaovalor"), "currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      TotalCompra = TotalCompra + TempValor
      
      If Preco = True Then
        StrTemp = Format(.Recordset!precocompra, "Currency")
        Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
      Else
        Printer.Print ""
      End If
      .Recordset.MoveNext
    Loop
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    .RecordSource = StrTabela
    .Refresh
  End With
  Printer.EndDoc
  Exit Sub
NaoImprime:
  Printer.KillDoc
  Exit Sub
End Sub

Private Sub ImprimeConferencia()
  Dim StrTemp As String, Largura As Double, strData As String
  Dim TotalCompra As Currency, TotalVenda As Currency, TempValor As Currency
  Dim Quebra As String, EspacoEntreLinhas As Double
  
  dbFechamento.Refresh
  
  Imprimindo = True
  
  With Adodc1
    StrTabela = .RecordSource
    If strFiltroProduto <> "" Then
      strFiltroProduto = " and " & Mid(strFiltroProduto, 7)
    End If
    .RecordSource = "select produtos.*, produtossubcategoria.* from produtos, produtossubcategoria where produtos.subcategoria=produtossubcategoria.codigosubcategoria " & strFiltroProduto & " order by produtossubcategoria.descri, produtos.descri"
    .Refresh
    If .Recordset.RecordCount = 0 Then
      .RecordSource = StrTabela
      .Refresh
      Imprimindo = False
      Exit Sub
    End If
    
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then
      Imprimindo = False
      Exit Sub
    End If
    On Error GoTo 0
    
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 2
    strData = "Data: " & Format(Now, "long date") & " - " & Format(Now, "short time")
    
    EspacoEntreLinhas = 0
    Largura = 190
    Quebra = .Recordset("produtossubcategoria.descri")
    
    Cabeca2 Largura, strData, Quebra
    
    
    Do While .Recordset.EOF = False
      
      Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
      
      If Printer.CurrentY > Printer.ScaleHeight - 30 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
        StrTemp = "Sub-Total"
        Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        Printer.CurrentY = Printer.CurrentY + 3
        Printer.Line (107, Printer.CurrentY)-(130, Printer.CurrentY)
        Printer.Line (132, Printer.CurrentY)-(155, Printer.CurrentY)
        Printer.Line (157, Printer.CurrentY)-(Largura, Printer.CurrentY)
        
        Printer.NewPage
        Cabeca2 Largura, strData, Quebra
      End If
      
      If Quebra <> .Recordset("produtossubcategoria.descri") Then
        Quebra = .Recordset("produtossubcategoria.descri")
        
        If Printer.CurrentY + Printer.TextHeight(Quebra) > Printer.ScaleHeight - 30 Then
          Printer.CurrentY = Printer.CurrentY + 1
          Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
          Printer.CurrentY = Printer.CurrentY + 1
          
          Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
          StrTemp = "Sub-Total"
          Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp;
          Printer.CurrentY = Printer.CurrentY + 3
          Printer.Line (107, Printer.CurrentY)-(130, Printer.CurrentY)
          Printer.Line (132, Printer.CurrentY)-(155, Printer.CurrentY)
          Printer.Line (157, Printer.CurrentY)-(Largura, Printer.CurrentY)
          
          Printer.NewPage
          Cabeca2 Largura, strData, Quebra
          Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
        Else
          SubCabeca2 Largura, strData, Quebra
          Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
        End If
      End If
      
      Printer.FontSize = 9
      Printer.FontBold = False
      
      StrTemp = .Recordset("codigo")
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      On Error Resume Next
      StrTemp = .Recordset("produtos.descri")
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("precocompra")
      Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("estoque")
      Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Printer.CurrentY = Printer.CurrentY + 3
      Printer.Line (107, Printer.CurrentY)-(130, Printer.CurrentY)
      
      Printer.Line (132, Printer.CurrentY)-(155, Printer.CurrentY)
      
      Printer.Line (157, Printer.CurrentY)-(Largura, Printer.CurrentY)
      
      Printer.Print ""
      
      Quebra = .Recordset("produtossubcategoria.descri")
      .Recordset.MoveNext
    Loop
    
    Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
    StrTemp = "Sub-Total"
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    Printer.CurrentY = Printer.CurrentY + 3
    Printer.Line (107, Printer.CurrentY)-(130, Printer.CurrentY)
    Printer.Line (132, Printer.CurrentY)-(155, Printer.CurrentY)
    Printer.Line (157, Printer.CurrentY)-(Largura, Printer.CurrentY)
    
    Printer.CurrentY = Printer.CurrentY + EspacoEntreLinhas
    StrTemp = "Total"
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    Printer.CurrentY = Printer.CurrentY + 3
    Printer.Line (107, Printer.CurrentY)-(130, Printer.CurrentY)
    Printer.Line (132, Printer.CurrentY)-(155, Printer.CurrentY)
    Printer.Line (157, Printer.CurrentY)-(Largura, Printer.CurrentY)
    
    .RecordSource = StrTabela
    .Refresh
  End With
  Printer.EndDoc
  Imprimindo = False
  Exit Sub
NaoImprime:
  Printer.KillDoc
  Imprimindo = False
  Exit Sub
End Sub


Private Sub ImprimeCalculoDePreco()
  Dim StrTemp As String, Largura As Double, strData As String
  Dim TotalCompra As Currency, TotalVenda As Currency, TempValor As Currency
  
  With Adodc1
    StrTabela = .RecordSource
    .RecordSource = "select *from produtos " & strFiltroProduto & strOrdem
    .Refresh
    If .Recordset.RecordCount = 0 Then Exit Sub
    
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 2
    strData = "Data: " & Format(Now, "long date") & " - " & Format(Now, "short time")
    
    Largura = 170
    
    Cabeca3 Largura, strData
    
    Do While .Recordset.EOF = False
      Printer.CurrentY = Printer.CurrentY + 5
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        Printer.NewPage
        Cabeca3 Largura, strData
      End If
      
      Printer.FontSize = 9
      Printer.FontBold = False
      
      StrTemp = .Recordset("codigo")
      Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = .Recordset("descri")
      Printer.CurrentX = 15
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset("Estoque"), "#,##0")
      Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset("PrecoCompra"), "#,##0.000")
      Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      
      Printer.CurrentY = Printer.CurrentY + 3
      Printer.Line (112, Printer.CurrentY)-(130, Printer.CurrentY)
      Printer.Line (132, Printer.CurrentY)-(150, Printer.CurrentY)
      Printer.Line (152, Printer.CurrentY)-(Largura, Printer.CurrentY)
      
      .Recordset.MoveNext
    Loop
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    .RecordSource = StrTabela
    .Refresh
  End With
  Printer.EndDoc
  Exit Sub
NaoImprime:
  Printer.KillDoc
  Exit Sub
End Sub




Private Sub Cabeca(ByVal Largura As Double, strData As String)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Printer.ForeColor = RGB(200, 200, 200)
  Printer.Line (0, 0)-(Largura, 18), , BF
  
  Printer.ForeColor = vbBlack
  Printer.FillColor = RGB(200, 200, 200)
  Printer.FontTransparent = True
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 14
  Printer.FontBold = True
  
  StrTemp = "Posição do Estoque"
  Printer.CurrentY = 2
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = Adodc2.Recordset("nome")
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  
  Printer.FontSize = 8
  Printer.FontBold = False
  
  StrTemp = strData
  Printer.CurrentX = 1
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
  Printer.Print StrTemp
  
  
  If dbFechamento.Recordset.RecordCount <> 0 Then
    dbFechamento.Recordset.MoveFirst
    StrTemp = "Último lançamento: " & dbFechamento.Recordset!DataCaixa & " - Turno: " & dbFechamento.Recordset!Turno
    Printer.CurrentX = 0
    Printer.Print StrTemp
  Else
    StrTemp = "Não existe lançamento de caixa"
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
    Printer.Print StrTemp
  End If
  Printer.CurrentY = 23
  
  
  StrTemp = "Código"
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  
  
  StrTemp = "Estoque"
  Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "$ Venda"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Comissão %"
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Comissão $"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  If Largura = 170 Then
    StrTemp = "Compra"
    Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
  Else
    Printer.Print ""
  End If
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
End Sub

Private Sub Cabeca2(ByVal Largura As Double, strData As String, ByVal Categoria As String)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Printer.ForeColor = RGB(200, 200, 200)
  Printer.Line (0, 0)-(Largura, 18), , BF
  
  Printer.ForeColor = vbBlack
  Printer.FillColor = RGB(200, 200, 200)
  Printer.FontTransparent = True
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 14
  Printer.FontBold = True
  
  StrTemp = "Conferencia do Estoque"
  Printer.CurrentY = 2
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = Adodc2.Recordset("nome")
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  
  Printer.FontSize = 9
  Printer.FontBold = False
  
  StrTemp = strData
  Printer.CurrentX = 1
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
  Printer.Print StrTemp
  
  If dbFechamento.Recordset.RecordCount <> 0 Then
    dbFechamento.Recordset.MoveFirst
    StrTemp = "Último lançamento: " & dbFechamento.Recordset!DataCaixa & " - Turno: " & dbFechamento.Recordset!Turno
    Printer.CurrentX = 0
    Printer.Print StrTemp
  Else
    StrTemp = "Não existe lançamento de caixa"
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
    Printer.Print StrTemp
  End If
  Printer.CurrentY = 23
  
  
  Printer.Print Categoria
  
  StrTemp = "Código"
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  
  StrTemp = "R$ Custo"
  Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Estoque"
  Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Posto"
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Diferença"
  Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "R$ Diferença"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
End Sub
Private Sub SubCabeca2(ByVal Largura As Double, strData As String, ByVal Categoria As String)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  
  Printer.FontSize = 9
  Printer.FontBold = False
  
  Printer.Print Categoria
  
  StrTemp = "Código"
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  
  StrTemp = "R$ Custo"
  Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Estoque"
  Printer.CurrentX = 105 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Posto"
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Diferença"
  Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "R$ Diferença"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
End Sub

Private Sub Cabeca3(ByVal Largura As Double, strData As String)
  Dim StrTemp As String
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Printer.ForeColor = RGB(200, 200, 200)
  Printer.Line (0, 0)-(Largura, 18), , BF
  
  Printer.ForeColor = vbBlack
  Printer.FillColor = RGB(200, 200, 200)
  Printer.FontTransparent = True
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontSize = 14
  Printer.FontBold = True
  
  StrTemp = "Cálculo de Preço de Venda"
  Printer.CurrentY = 2
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  StrTemp = NomePosto
  Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
  Printer.Print StrTemp
  
  
  Printer.FontSize = 9
  Printer.FontBold = False
  
  StrTemp = strData
  Printer.CurrentX = 1
  Printer.Print StrTemp;
  
  StrTemp = "Página: " & Printer.Page
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp) - 1
  Printer.Print StrTemp
  
  Printer.CurrentY = 23
  
  
  StrTemp = "Código"
  Printer.CurrentX = 14 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Produto"
  Printer.CurrentX = 15
  Printer.Print StrTemp;
  
  
  StrTemp = "Estoque"
  Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "V. Compra"
  Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "V. Venda"
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Comissão %"
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = "Comissão $"
  Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 5
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
On Error Resume Next
If Imprimindo = True Then Exit Sub
CodigoAntigo = 0
CodigoProduto = 0

CodigoAntigo = Adodc1.Recordset!Codigo
CodigoProduto = Adodc1.Recordset!CodigoProduto

If Adodc1.Recordset!Combustivel = True Then
  lblTanqueAcerto.Visible = True
  txtTanqueAcerto.Visible = True
Else
  lblTanqueAcerto.Visible = False
  txtTanqueAcerto.Visible = False
End If
With dbProdutosBarras
  If Adodc1.Recordset.EOF = True Then
    .RecordSource = "select *from produtoscodigos where codigoproduto=0 order by codigobarra"
  Else
    .RecordSource = "select *from produtoscodigos where codigoproduto=" & Adodc1.Recordset!Codigo & " order by codigobarra"
  End If
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub

Private Sub cboTurnos_LostFocus()
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboTurnos.Text & "'"
  If .Recordset.EOF = False Then
    cboTurnos.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdAcerto_Click()
Dim AcertoValor As Currency, Qtuantidade As Double
Dim EstoqueAnterior As Double, EstoquePosterior As Double
Dim PrecoMedio As Currency, ValorMedio As Currency
Dim Tanque As Integer, Quantidade As Double

Tanque = 0
If IsNumeric(txtAcerto.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtAcerto.SetFocus
  Exit Sub
End If
If IsNumeric(cboTurnos.Text) = False Then
  MsgBox "Escolha um turno correto!"
  Exit Sub
End If
dbTurnos.Refresh
If dbTurnos.Recordset.RecordCount = 0 Then
  MsgBox "Cadatre um turno primeiro!"
  Exit Sub
End If
Call cboTurnos_LostFocus
If dbTurnos.Recordset!Descri <> cboTurnos.Text Then
  MsgBox "Turno incorreto!"
  cboTurnos.SetFocus
  Exit Sub
End If
With Adodc1
  If .Recordset!Combustivel = True Then
    If IsNumeric(txtTanqueAcerto.Text) = False Then
      MsgBox "Informe um tanque correto!"
      txtTanqueAcerto.SetFocus
      Exit Sub
    End If
    Tanque = CInt(txtTanqueAcerto.Text)
    With dbTanque
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.Find "tanque=" & txtTanqueAcerto.Text
        If .Recordset.EOF = True Then
          MsgBox "Erro na tabela de tanques!"
          Exit Sub
        Else
          If .Recordset!CodigoProduto <> Adodc1.Recordset!CodigoProduto Then
            MsgBox "Tanque inválido!"
            Exit Sub
          End If
        End If
      Else
        MsgBox "Erro na tabela de tanques!"
        Exit Sub
      End If
    End With
  End If
End With
Permissao = False
ColocaProduto = True
Quantidade = CDbl(txtAcerto.Text)
frmPermissao.Show vbModal
If Permissao = False Then
  ColocaProduto = False
  Exit Sub
End If
With Adodc1
  'AcertoValor = Quantidade * .Recordset!precocompra
  EstoqueAnterior = .Recordset!Estoque
  EstoquePosterior = .Recordset!Estoque + Quantidade
  If .Recordset!ValorEstoque <> 0 And .Recordset!Estoque <> 0 Then
    If Estoque > 0 Then
        PrecoMedio = .Recordset!ValorEstoque / .Recordset!Estoque
    Else
        PrecoMedio = .Recordset!precocompra
    End If
  Else
    PrecoMedio = .Recordset!precocompra
  End If
  ValorMedio = PrecoMedio * Quantidade
  AcertoValor = ValorMedio
  .Recordset!Estoque = .Recordset!Estoque + Quantidade
  If IsNull(.Recordset!ValorEstoque) = True Then .Recordset!ValorEstoque = 0
  .Recordset!ValorEstoque = .Recordset!ValorEstoque + ValorMedio
  .Recordset.Update
  If .Recordset!Combustivel = True Then
    With dbTanque
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.Find "tanque=" & txtTanqueAcerto.Text
        If .Recordset.EOF = True Then
          MsgBox "Erro na tabela de tanques!"
        Else
          .Recordset!Estoque = .Recordset!Estoque + Quantidade
          .Recordset.Update
        End If
      Else
        MsgBox "Erro na tabela de tanques!"
      End If
    End With
  End If
End With
With Adodc3
  .Refresh
  .Recordset!acertoestoque = .Recordset!acertoestoque + AcertoValor
  .Recordset.Update
End With
With dbAcerto
  .Recordset.AddNew
  .Recordset!CodProduto = Adodc1.Recordset!CodigoProduto
  .Recordset!CodigoProduto = Adodc1.Recordset!Codigo
  .Recordset!Descri = Adodc1.Recordset!Descri
  .Recordset!Combustivel = Adodc1.Recordset!Combustivel
  .Recordset!datalancada = Date
  .Recordset!EstoqueAnterior = EstoqueAnterior
  .Recordset!EstoquePosterior = EstoquePosterior
  .Recordset!Valorutilizado = Quantidade
  If Adodc1.Recordset!ValorEstoque <> 0 And Adodc1.Recordset!Estoque <> 0 Then
    .Recordset!precocompra = Adodc1.Recordset!ValorEstoque / Adodc1.Recordset!Estoque
  Else
    .Recordset!precocompra = Adodc1.Recordset!precocompra
  End If
  .Recordset!PrecoVenda = Adodc1.Recordset!PrecoVenda
  .Recordset!ValorDiferenca = AcertoValor
  .Recordset.Update
  .Refresh
End With
With dbProdutosHistorico
  .Recordset.AddNew
  .Recordset!lancadoem = Now
  .Recordset!dataalteracao = Date
  .Recordset!CodigoProduto = Adodc1.Recordset!CodigoProduto
  .Recordset!Codigo = Adodc1.Recordset!Codigo
  .Recordset!descriproduto = Adodc1.Recordset!Descri
  .Recordset!descrioperacao = "Acerto de Estoque. Usuário=" & Usuarios.Nome
  If Adodc1.Recordset!ValorEstoque <> 0 And Adodc1.Recordset!Estoque <> 0 Then
    .Recordset!precocompra = Adodc1.Recordset!ValorEstoque / Adodc1.Recordset!Estoque
  Else
    .Recordset!precocompra = Adodc1.Recordset!precocompra
  End If
  .Recordset!PrecoVenda = Adodc1.Recordset!PrecoVenda
  .Recordset!EstoqueAnterior = EstoqueAnterior
  .Recordset!Quantidade = Quantidade
  .Recordset!estoquefinal = Adodc1.Recordset!Estoque
  .Recordset.Update
End With

RegistraEstoque txtDataCaixa.Value, dbTurnos.Recordset!CodigoTurno, dbTurnos.Recordset!Descri, dbTurnos.Recordset!HoraIni, Adodc1.Recordset!CodigoProduto, Tanque, , , Quantidade

Dim Estatus As New frmEstatus2
Load Estatus
Unload Estatus

txtAcerto.Text = ""
txtTanqueAcerto.Text = ""
ColocaProduto = False
Text2.SetFocus


End Sub

Private Sub cmdAdd_Click()
  Dim strCategoria As String, codSubCategoria As Double, Precodigo As String
  If codigoPosto = 0 Then
    MsgBox "Cadastre um Posto!"
    Exit Sub
  End If
  frmCadProdutosCategoria.Show vbModal
  strCategoria = frmCadProdutosCategoria.strCategoria
  codSubCategoria = frmCadProdutosCategoria.codSubCategoria
  Precodigo = frmCadProdutosCategoria.Precodigo
  Unload frmCadProdutosCategoria
  If strCategoria = "" Then
    MsgBox "Categoria não informada!"
    Exit Sub
  End If
  If codSubCategoria = 0 Then
    MsgBox "Sub Categoria não informada!"
    Exit Sub
  End If
  Adodc1.Recordset.AddNew
  Adodc1.Recordset("codigoPosto") = codigoPosto
  Adodc1.Recordset("Categoria") = strCategoria
  Adodc1.Recordset("subcategoria") = codSubCategoria
  Adodc1.Recordset("codigo") = Precodigo
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  Frame3.Enabled = True
  txtFields(3).Enabled = True
  txtFields(0).SetFocus
  txtFields(0).SelStart = Len(txtFields(0).Text)
  
End Sub

Private Sub cmdAlteraPreco_Click()
Dim Novo As Currency, Antigo As Currency, Diferenca As Currency
Dim Estoque As Double
Dim EstoqueMedio As Currency, EstoqueMedioNovo As Currency
Dim DiferencaMedio As Currency

With Adodc1
  If .Recordset.RecordCount = 0 Then
    MsgBox "Banco de dados vazio!"
    Exit Sub
  End If
  If IsNumeric(txtPrecoNovo.Text) = False Then
    MsgBox "Informe um preço válido!"
    txtPrecoNovo.SetFocus
    Exit Sub
  End If
  Novo = CCur(txtPrecoNovo.Text)
  If IsNull(.Recordset("precoCompra")) = False Then
    Antigo = .Recordset("precoCompra")
  End If
  If IsNull(.Recordset("estoque")) = False Then
    Estoque = .Recordset("estoque")
  End If
  EstoqueMedio = .Recordset!ValorEstoque
  EstoqueMedioNovo = .Recordset!Estoque * Novo
  DiferencaMedio = EstoqueMedioNovo - EstoqueMedio
  .Recordset!LucroMedio = .Recordset!LucroMedio + DiferencaMedio
  .Recordset!ValorEstoque = EstoqueMedioNovo
  .Recordset!PrecoMedio = Novo
  Diferenca = Novo - Antigo
  Diferenca = Diferenca * Estoque
  .Recordset("precocompra") = Novo
  .Recordset!Variacao = .Recordset!Variacao + Diferenca
  .Recordset.Update
End With
With Adodc3
  .Recordset("variacaoestoque") = .Recordset("variacaoestoque") + Diferenca
  .Recordset.Update
  .Refresh
End With

End Sub

Private Sub cmdConferencia_Click()
frmcadProdutosConfere.Show
End Sub

Private Sub cmdDelete_Click()
  Dim db As New ADODB.Connection
  Dim dbTemp As New ADODB.Recordset
  
  Dim Resposta As Integer
  
  
  With Adodc1
    If .Recordset.EOF = True Or .Recordset.BOF = True Then Exit Sub
    
    db.Open CaminhoADO
    dbTemp.CursorLocation = adUseClient
    dbTemp.Open "select produtosnotascorpo.*, produtosnotas.* from produtosnotascorpo, produtosnotas where produtosnotascorpo.codigoprodutonota=produtosnotas.codigoentrada and produtosnotas.confirmado=0 and produtosnotascorpo.codigoproduto=" & Adodc1.Recordset!CodigoProduto, db, adOpenKeyset, adLockOptimistic
    If dbTemp.RecordCount <> 0 Then
      MsgBox "Existe nota com este produto aguardando para ser confirmada a entrada!"
      Exit Sub
    End If
    dbTemp.Close
    
    
    
    If .Recordset!Estoque <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue estoque!"
      Exit Sub
    End If
    If .Recordset!acumulativo <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue venda! Aguarde o fechamento para poder excluir!"
      Exit Sub
    End If
    If .Recordset!Variacao <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue Variação de Estoque! Aguarde o fechamento para poder excluir!"
      Exit Sub
    End If
    If .Recordset!LucroVenda <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue Lucro de Venda! Aguarde o fechamento para poder excluir!"
      Exit Sub
    End If
    If .Recordset!TotalVendido <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue Acumulativo de Venda! Aguarde o fechamento para poder excluir!"
      Exit Sub
    End If
    If .Recordset!ValorEstoque <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue valor no estoque!"
      Exit Sub
    End If
    If .Recordset!valordifestoque <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue valor de diferenca no estoque!"
      Exit Sub
    End If
    
  End With
  With dbVenda2
    .ConnectionString = CaminhoADO
    .RecordSource = "Select *from venda2 where codigoproduto=" & Adodc1.Recordset!CodigoProduto & " and pago=0 and valorcomissao<>0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue Comissão a ser paga!"
      Exit Sub
    End If
    .ConnectionString = CaminhoADO
    .RecordSource = "Select *from venda2 where codigoproduto=" & Adodc1.Recordset!CodigoProduto & " and fechamentodiario=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      MsgBox "Não pode excluir o produto atual pois ele possue venda lançada no caixa sem confirmar!"
      Exit Sub
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
  Frame3.Enabled = False
End Sub

Private Sub cmdEditar_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja alterar o produto atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
If Adodc1.Recordset!Combustivel = True Then
  Resposta = MsgBox("Você está alterando um combustível! Deseja realmente fazer a alteração?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
End If
Frame1.Enabled = True
Frame3.Enabled = True
txtFields(3).Enabled = False
txtFields(0).SetFocus
End Sub


Private Sub cmdExportar_Click()
ProdutosMicrosffer
End Sub

Private Sub cmdFonte_Click()
On Error GoTo semFonte
CommonDialog1.flags = cdlCFBoth
CommonDialog1.ShowFont
With CommonDialog1
  DataGrid1.Font.Name = .FontName
  DataGrid1.Font.Size = .FontSize
  DataGrid1.Font.Bold = .FontBold
  DataGrid1.Font.Italic = .FontItalic
End With
semFonte:
End Sub

Private Sub cmdImportarGrupos_Click()
Dim dbImportar As New ADODB.Connection
Dim db As New ADODB.Connection
Dim dbGruposImportar As New ADODB.Recordset
Dim dbProdutosImportar As New ADODB.Recordset
Dim dbConfig As New ADODB.Recordset

db.Open CaminhoADO
dbConfig.CursorLocation = adUseClient
dbConfig.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly

dbImportar.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig!ftp

dbGruposImportar.CursorLocation = adUseClient
dbGruposImportar.Open "Select *from gruposecf where posto='" & dbConfig!Porta & "'", dbImportar, adOpenForwardOnly, adLockReadOnly

dbProdutosImportar.CursorLocation = adUseClient
dbProdutosImportar.Open "select *from produtosgrupo where posto='" & dbConfig!Porta & "'", dbImportar, adOpenForwardOnly, adLockReadOnly

With dbGruposIF
  .Refresh
  If dbGruposImportar.RecordCount <> 0 Then
    If .Recordset.RecordCount <> 0 Then
      db.Execute "delete *from produtosgrupoif"
    End If
    Do While dbGruposImportar.EOF = False
      db.Execute "insert into produtosgrupoif (codigogrupo,descri) values (" & dbGruposImportar!Grupo & ",'" & dbGruposImportar!Descri & "')"
      dbGruposImportar.MoveNext
    Loop
  End If
  .Refresh
  .Refresh
End With

Imprimindo = True
With Adodc1
  If .Recordset.RecordCount <> 0 Then
    If dbProdutosImportar.RecordCount <> 0 And dbGruposIF.Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbProdutosImportar.MoveFirst
        dbProdutosImportar.Find "codproduto='" & .Recordset!Codigo & "'"
        If dbProdutosImportar.EOF = False Then
            If dbProdutosImportar!grupoecf <> "" Then
                dbGruposIF.Refresh
                dbGruposIF.Recordset.MoveFirst
                dbGruposIF.Recordset.Find "codigogrupo=" & dbProdutosImportar!grupoecf
                If dbGruposIF.Recordset.EOF = False Then
                  .Recordset!codigogrupoif = dbGruposIF.Recordset!Codigo
                  .Recordset.Update
                End If
            Else
              MsgBox "Existe produto no posto sem grupo ecf!"
            End If
        Else
          'MsgBox "Produto " & .Recordset!Codigo & " - " & .Recordset!Descri & " não localizado!"
        End If
        .Recordset.MoveNext
      Loop
      .Recordset.UpdateBatch adAffectAllChapters
    End If
  End If
End With

Imprimindo = False

dbGruposImportar.Close
dbProdutosImportar.Close
dbImportar.Close
dbConfig.Close
db.Close

MsgBox "Importação finalizada!"

End Sub

Private Sub cmdImprime_Click()
Dim Resposta As Integer
strFiltroProduto = ""
Select Case cboRelatorios.Text
  Case "Tabela de Preço"
    strFiltroProduto = FiltraProdutos()
    ImprimeTabela
    Exit Sub
  Case "Conferência de Estoque"
    strFiltroProduto = FiltraProdutos()
    ImprimeConferencia
    Exit Sub
  Case "Tabela de ICMS"
    strFiltroProduto = FiltraProdutos()
    Imprimindo = True
    ImprimeTabelaICMS
    Imprimindo = False
    Exit Sub
  Case "Cálculo do Preço de Compra"
    strFiltroProduto = FiltraProdutos()
    ImprimeCalculoDePreco
    Exit Sub
  Case "Preço de Compra"
    strFiltroProduto = FiltraProdutos()
    ImprimeCompra True
  Case "Preço Médio"
    strFiltroProduto = FiltraProdutos()
    frmRelatPrecoMedio.Show
    frmRelatPrecoMedio.SetFocus
End Select

End Sub

Private Sub cmdImprimeTabela_Click()
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  ImprimeADOGrid DataGrid1, Printer, Adodc1, , True, , , , , "Tabela de Produtos", Adodc2.Recordset!Nome, Format(Now, "long Date")
  
NaoImprime:
End Sub

Private Sub cmdIncluirBarras_Click()
If txtCodigoBarra.Text = "" Then
  MsgBox "Escolha o código-de-barras a ser incluido!"
  txtCodigoBarra.SetFocus
  Exit Sub
End If
If Adodc1.Recordset.EOF = True Then
  MsgBox "Escolha um produto a ser incluido o código-de-barras!"
  Exit Sub
End If
'If EAN_13_Validar(txtCodigoBarra.Text) = False Then
'  MsgBox "Código de barras inválido!"
'  Exit Sub
'End If
With dbProdutosBarras
  .Recordset.AddNew
  .Recordset!CodigoProduto = Adodc1.Recordset!Codigo
  .Recordset!codigobarra = txtCodigoBarra.Text
  .Recordset.Update
End With
End Sub

Private Sub cmdIncluirSubCategoria_Click()
If txtCodSubCategoria.Text = "" Then
  MsgBox "Informe um código"
  txtCodSubCategoria.SetFocus
  Exit Sub
End If
If txtSubCategoria.Text = "" Then
  MsgBox "Informe uma descrição!"
  txtSubCategoria.SetFocus
  Exit Sub
End If
With dbProdutosSubCategoria
  .Recordset.AddNew
  .Recordset!Precodigo = txtCodSubCategoria.Text
  .Recordset!Descri = txtSubCategoria.Text
  .Recordset.Update
  .Refresh
End With

End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
  Frame3.Enabled = False
  txtFields(3).Enabled = False
End Sub

Private Sub cmdRemoverBarras_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja remover o codigo selecionado?", vbYesNo + vbDefaultButton2)
If respota = vbNo Then Exit Sub
With dbProdutosBarras
  If .Recordset.RecordCount = 0 Or .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um código primeiro!"
    Exit Sub
  End If
  .Recordset.Delete adAffectCurrent
  .Refresh
  .Refresh
End With
End Sub

Private Sub cmdRemoverSubCategoria_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbProdutosSubCategoria
  If .Recordset.RecordCount = 0 Or .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um item primeiro!"
    Exit Sub
  End If
  .Recordset.Delete adAffectCurrent
  .Refresh
  .Refresh
End With
End Sub

Private Sub cmdTrocaCodigo_Click()
Dim Alterar As String
Alterar = InputBox("Informe o código a ser alterado!")
If IsNumeric(Alterar) = False Then
  MsgBox "Código inválido!"
  Exit Sub
End If
With Adodc1
  If .Recordset.EOF = False And .Recordset.BOF = False Then
    .Recordset.MoveFirst
    .Recordset.Find "codigo=" & Alterar
    If .Recordset.EOF = False Then
      Alterar = InputBox("Informe o código novo!")
      If IsNumeric(Alterar) = False Then
        MsgBox "Código inválido!"
        Exit Sub
      End If
      .Recordset!Codigo = Alterar
      txtFields(0).Text = Alterar
      AlteraCodigoProduto CodigoProduto, CodigoAntigo, Alterar
      .Recordset.Update
    End If
  End If
End With
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    A = .Recordset.AbsolutePosition
    If IsNull(.Recordset!Categoria) = True Then
      MsgBox "Selecione uma categoria válida!"
      DataCombo2.SetFocus
      Exit Sub
    End If
    If Trim(.Recordset!Categoria) = "" Then
      MsgBox "Selecione uma categoria válida!"
      DataCombo2.SetFocus
      Exit Sub
    End If
    If IsNull(.Recordset!Comissao) = True Then
      Resposta = MsgBox("A comissão deste produto é Zero. Isso está correto?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        txtFields(5).SetFocus
        Exit Sub
      End If
    End If
    If .Recordset!Comissao = 0 Then
      Resposta = MsgBox("A comissão deste produto é Zero. Isso está correto?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        txtFields(5).SetFocus
        Exit Sub
      End If
    End If
    If IsNull(.Recordset!precocompra) = True Then
      MsgBox "O preco de compra não pode ser Zero."
      txtFields(3).SetFocus
      Exit Sub
    End If
    If .Recordset!precocompra < 0 Then
      MsgBox "O preco de compra está errado!"
      txtFields(3).SetFocus
      Exit Sub
    End If
    If IsNull(.Recordset!PrecoVenda) = True Then
      MsgBox "O preco de venda não pode ser Zero."
      txtFields(4).Enabled = True
      txtFields(4).SetFocus
      Exit Sub
    End If
    txtFields(4).Enabled = False
    If .Recordset!PrecoVenda <= 0 Then
      MsgBox "O preco de venda está errado!"
      txtFields(4).SetFocus
      Exit Sub
    End If
    If CodigoAntigo <> CodigoNovo Then
      If AlteraCodigoProduto(CodigoProduto, CodigoAntigo, txtFields(0).Text) = False Then
        MsgBox "Não foi possível alterar o código em todas as tabelas!" & Err.Description
      End If
    End If
    If IsNull(.Recordset!codigogrupoif) = True Then
      MsgBox "Selecione um departamento válido!"
      DataCombo6.SetFocus
      Exit Sub
    End If
    .Recordset.Update
    .Recordset.AbsolutePosition = A
  End With
  txtFields(3).Enabled = False
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Frame1.Enabled = False
  Frame3.Enabled = False

End Sub

Private Sub cmdClose_Click()
  
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Command1_Click()
With Adodc2
  If .Recordset.RecordCount = 0 Then
    codigoPosto = 0
    MsgBox "Cadastre um posto primeiro!"
    Exit Sub
  Else
    .Recordset.Find "nome='" & DataCombo1.Text & "'"
    If .Recordset.EOF = False Then
      codigoPosto = .Recordset("codigoposto")
    End If
  End If
End With
With Adodc1
  .RecordSource = "select *from produtos where codigoposto=" & codigoPosto & strOrdem
  .Refresh
End With
End Sub


Private Sub DataCombo1_LostFocus()
With Adodc2
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & DataCombo1.Text & "'"
  If .Recordset.EOF = False Then
    DataCombo1.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by  " & DataGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField
End If
With Adodc1
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos where codigoposto=" & codigoPosto & strOrdem
  .Refresh
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
Dim TabProdutos As String
Dim Comissao As Double, ComissaoValor As Double
Dim Sugerido As Double

txtDataCaixa.Value = Date

Imprimindo = False
With dbProdutos2
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If .Recordset!lucrominimo <> 0 Then
        Comissao = 0
        ComissaoValor = 0
        Sugerido = .Recordset!precocompra + (.Recordset!precocompra * (.Recordset!lucrominimo / 100))
        If .Recordset!ComissaoValor <> 0 Then
          Sugerido = Sugerido + .Recordset!ComissaoValor
        End If
        If .Recordset!Comissao <> 0 Then
          Comissao = .Recordset!Comissao
          Sugerido = Sugerido / (1 - (Comissao))
        End If
        .Recordset!Sugerido = Sugerido
        .Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With dbCST
  .ConnectionString = CaminhoADO
  .Refresh
End With

With dbProdutosSubCategoria
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from ProdutosSubCategoria order by descri"
  .Refresh
End With

With dbCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutosBarras
  .ConnectionString = CaminhoADO
  .RecordSource = "ProdutosCodigos"
  .Refresh
End With
With dbEstacionamento
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from Estacionamento"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Preco = 0
    .Recordset!ultimonumero = 0
    .Recordset.Update
    .Refresh
  End If
  If txtUltimoNumero.Text <> "0" Then
    If Usuarios.Grupo.AdmEstatus <> 2 Then
      txtUltimoNumero.Enabled = False
    Else
      txtUltimoNumero.Enabled = True
    End If
  End If
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos order by UnCaixa"
  .Refresh
End With

With Adodc2
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    codigoPosto = Adodc2.Recordset("codigoposto")
    DataCombo1.Text = Adodc2.Recordset("nome")
  Else
    codigoPosto = 0
  End If
End With
strOrdem = " order by Codigo"
With Adodc1
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos where codigoposto=" & codigoPosto & strOrdem
  .Refresh
End With
With Adodc3
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbTanque
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbAcerto
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutos2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbAliquotas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutosHistorico
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbVenda2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbFechamento
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbGruposIF
  .ConnectionString = CaminhoADO
  .Refresh
End With

Select Case Usuarios.Grupo.CadProdutos
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdAcerto.Enabled = False
    'cmdAlteraPreco.Enabled = False
    txtPrecoEstacionamento.Enabled = False
  Case 2 'Liberado
    cmdEditar.Enabled = True
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdAcerto.Enabled = True
    cmdAlteraPreco.Enabled = True
    txtPrecoEstacionamento.Enabled = True
End Select



If Usuarios.Nome = "Usuário Master" Then
  lblLabels(10).Visible = True
  txtPrecoNovo.Visible = True
  cmdAlteraPreco.Visible = True
  cmdDelete.Enabled = True
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
With Adodc1
    A = .Recordset.AbsolutePosition
    If IsNull(.Recordset!Categoria) = True Then
      MsgBox "Selecione uma categoria válida!"
      Frame1.Enabled = True
      Frame3.Enabled = True
      DataCombo2.SetFocus
      Cancel = True
      Exit Sub
    End If
    On Error Resume Next
    If Trim(.Recordset!Categoria) = "" Then
      If Err.Number <> 0 Then Exit Sub
      MsgBox "Selecione uma categoria válida!"
      Frame1.Enabled = True
      Frame3.Enabled = True
      DataCombo2.SetFocus
      Cancel = True
      Exit Sub
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Text2_LostFocus()
With Adodc1
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(Text2.Text) = False Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "codigo=" & Text2.Text
End With
End Sub

Private Sub txtDataCaixa_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataCaixa_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
  Case 3, 4, 5, 7
    On Error Resume Next
    Select Case KeyAscii
      Case Asc(".")
        KeyAscii = 0
        SendKeys ","
    End Select
  Case 0
    If IsNumeric(txtFields(0).Text) = True Then
      CodigoNovo = txtFields(0).Text
    Else
      CodigoNovo = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
  Case 0
    With dbProdutos2
      .Refresh
      If txtFields(Index).Text <> "" Then
        .Recordset.Find "codigo=" & txtFields(Index).Text
        If .Recordset.EOF = False Then
          If .Recordset!CodigoProduto <> Adodc1.Recordset!CodigoProduto Then
            MsgBox "Código já cadastrado!"
            txtFields(Index).SetFocus
            Exit Sub
          End If
        End If
      End If
    End With
End Select
End Sub

