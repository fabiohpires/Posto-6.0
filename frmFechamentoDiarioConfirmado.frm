VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFechamentoDiarioConfirmado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caixas Fechados"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   9015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "DBFs"
      Height          =   5055
      Left            =   3120
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   8535
      Begin MSAdodcLib.Adodc QVendaTotaliza 
         Height          =   330
         Left            =   5280
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QVendaTotaliza"
         Caption         =   "QVendaTotaliza"
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
      Begin MSAdodcLib.Adodc QDespesaLancTotaliza 
         Height          =   330
         Left            =   5280
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QDespesaLancTotaliza"
         Caption         =   "QDespesaLancTotaliza"
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
      Begin MSAdodcLib.Adodc QFormaDePgRecTotaliza 
         Height          =   330
         Left            =   5280
         Top             =   1440
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QFormaDePagamentoRecebidoTotaliza"
         Caption         =   "QFormaDePgRecTotaliza"
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
      Begin MSAdodcLib.Adodc QBicoMovimentaTotal 
         Height          =   330
         Left            =   5280
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QBicoMovimentoTotaliza"
         Caption         =   "QBicoMovimentaTotal"
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
      Begin MSAdodcLib.Adodc QTemp 
         Height          =   330
         Left            =   5280
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QBicoMovimentoTotalTanque"
         Caption         =   "QTemp"
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
         Left            =   2760
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from produtos order by descri"
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
      Begin MSAdodcLib.Adodc dbFormaDePg 
         Height          =   330
         Left            =   2760
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from FormaDePagamento order by descri"
         Caption         =   "dbFormaDePg"
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
      Begin MSAdodcLib.Adodc dbFormaDePgRecebido 
         Height          =   330
         Left            =   2760
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from FormaDePagamentoRecebido order by descri"
         Caption         =   "dbFormaDePgRecebido"
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
      Begin MSAdodcLib.Adodc dbDespesas 
         Height          =   330
         Left            =   2760
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from DespesaTipo order by descri"
         Caption         =   "dbDespesas"
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
      Begin MSAdodcLib.Adodc dbDespesasLanc 
         Height          =   330
         Left            =   2760
         Top             =   1440
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from DespesasLanc order by descri"
         Caption         =   "dbDespesasLanc"
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
         Left            =   2760
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from venda order by descri"
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
      Begin MSAdodcLib.Adodc dbProdutos 
         Height          =   330
         Left            =   2760
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from produtos order by descri"
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
      Begin MSAdodcLib.Adodc dbTanquesMovimento 
         Height          =   330
         Left            =   240
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from tanquesMovimento order by tanque"
         Caption         =   "dbTanquesMovimento"
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
      Begin MSAdodcLib.Adodc dbTanques 
         Height          =   330
         Left            =   240
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from tanques order by tanque"
         Caption         =   "dbTanques"
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
      Begin MSAdodcLib.Adodc dbBico 
         Height          =   330
         Left            =   240
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from bicos order by bico"
         Caption         =   "dbBico"
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
      Begin MSAdodcLib.Adodc dbBicoMovimento 
         Height          =   330
         Left            =   240
         Top             =   1440
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from bicomovimento order by bico"
         Caption         =   "dbBicoMovimento"
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
         Left            =   240
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from FechamentoDiario order by Data, Hora"
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
      Begin MSAdodcLib.Adodc dbResponsavel 
         Height          =   330
         Left            =   240
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from vendedores where gerente=-1 order by nome"
         Caption         =   "dbResponsavel"
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
      Begin MSAdodcLib.Adodc dbPosto 
         Height          =   330
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from postos order by nome"
         Caption         =   "dbPosto"
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
      Begin MSAdodcLib.Adodc dbStatus 
         Height          =   330
         Left            =   5280
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from status"
         Caption         =   "dbStatus"
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
         Height          =   330
         Left            =   5280
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
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
      Begin MSAdodcLib.Adodc dbPrevisaoRecebe 
         Height          =   330
         Left            =   5280
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from PrevisaoRecebimentos"
         Caption         =   "dbPrevisaoRecebe"
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
      Begin MSAdodcLib.Adodc dbClientes 
         Height          =   330
         Left            =   240
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientes where mensalista=-1 order by nome"
         Caption         =   "dbClientes"
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
      Begin MSAdodcLib.Adodc dbClientesNota 
         Height          =   330
         Left            =   2760
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientesnota where confirmado=0"
         Caption         =   "dbClientesNota"
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
      Begin MSAdodcLib.Adodc QClientesNota 
         Height          =   330
         Left            =   240
         Top             =   3240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QClientesNota"
         Caption         =   "QClientesNota"
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
      Begin MSAdodcLib.Adodc QGalonagem 
         Height          =   330
         Left            =   2760
         Top             =   3240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from QGalonagemProduto"
         Caption         =   "QGalonagem"
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
      Begin MSAdodcLib.Adodc dbGalonagem 
         Height          =   330
         Left            =   5280
         Top             =   3240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from galonagem"
         Caption         =   "dbGalonagem"
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
      Begin MSAdodcLib.Adodc dbCompensaPendente 
         Height          =   330
         Left            =   5280
         Top             =   3600
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from compensapendente"
         Caption         =   "dbCompensaPendente"
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
      Begin MSAdodcLib.Adodc dbTurno 
         Height          =   330
         Left            =   2760
         Top             =   3600
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from turnos order by descri"
         Caption         =   "dbTurno"
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
         Height          =   330
         Left            =   240
         Top             =   3600
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from cheques "
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
      Begin MSAdodcLib.Adodc QCheques 
         Height          =   330
         Left            =   5280
         Top             =   3960
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select sum(valor) as Total from cheques where codigofechamento=0"
         Caption         =   "QCheques"
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
      Begin MSAdodcLib.Adodc dbConcilia 
         Height          =   330
         Left            =   2760
         Top             =   3960
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from concilia"
         Caption         =   "dbConcilia"
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
      Begin MSAdodcLib.Adodc dbProdutoEntra 
         Height          =   330
         Left            =   240
         Top             =   3960
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from produtosentrada where codigofechamento=-1"
         Caption         =   "dbProdutoEntra"
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
      Begin MSAdodcLib.Adodc QProdutoEntra 
         Height          =   330
         Left            =   240
         Top             =   4320
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select sum(valornota) as total from produtosentrada where codigofechamento=-1"
         Caption         =   "QProdutoEntra"
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
   Begin VB.TextBox txtResponsavel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   14
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   5520
      Picture         =   "frmFechamentoDiarioConfirmado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "Imprimir"
      Top             =   5520
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2040
      Top             =   5400
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   5400
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
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "&Próximo >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<< &Anterior"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdInlueBomba 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24641537
      CurrentDate     =   37600
   End
   Begin MSDataListLib.DataCombo cboPosto 
      Bindings        =   "frmFechamentoDiarioConfirmado.frx":0A82
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboResponsavel 
      Bindings        =   "frmFechamentoDiarioConfirmado.frx":0A98
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmFechamentoDiarioConfirmado.frx":0AB4
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Frame Tela 
      Caption         =   " Controle de Bomba "
      Height          =   4575
      Index           =   0
      Left            =   120
      TabIndex        =   63
      Top             =   840
      Width           =   8775
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
         TabIndex        =   67
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdInclueBico 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   5040
         TabIndex        =   66
         Top             =   4080
         Visible         =   0   'False
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
         TabIndex        =   65
         Top             =   4200
         Visible         =   0   'False
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
         TabIndex        =   64
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboBico 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0ACA
         Height          =   315
         Left            =   120
         TabIndex        =   68
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Bico"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0ADF
         Height          =   3255
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "Bico"
            Caption         =   "Bico"
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
            DataField       =   "MecanicoFinal"
            Caption         =   "Mecânico Final"
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
         BeginProperty Column02 
            DataField       =   "ValorFinal"
            Caption         =   "Eletônico"
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
         BeginProperty Column03 
            DataField       =   "Vendas"
            Caption         =   "Vendas"
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
            DataField       =   "Retorno"
            Caption         =   "Retorno"
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
            DataField       =   "PrecoUnitario"
            Caption         =   "Preço"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "ValorVendido"
            Caption         =   "Valor Vendido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,00"
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
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1124,787
            EndProperty
         EndProperty
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Mecânico:"
         Height          =   195
         Left            =   1080
         TabIndex        =   75
         Top             =   3960
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Eletônico:"
         Height          =   195
         Left            =   2520
         TabIndex        =   74
         Top             =   3960
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bico:"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   3960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Retorno:"
         Height          =   195
         Left            =   3960
         TabIndex        =   72
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TotalVendido"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "QBicoMovimentaTotal"
         Height          =   255
         Left            =   5160
         TabIndex        =   71
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   4680
         TabIndex        =   70
         Top             =   3720
         Width           =   405
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Fechamento "
      Height          =   4575
      Index           =   8
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   " Resumo "
         Height          =   3255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   3735
         Begin VB.Label lblClientesNota 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Valor"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "QClientesNota"
            Height          =   255
            Left            =   1800
            TabIndex        =   53
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Notas:"
            Height          =   195
            Left            =   1170
            TabIndex        =   52
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Vendas Combustível:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblVendasCombustivel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "TotalVendido"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "QBicoMovimentaTotal"
            Height          =   255
            Left            =   1800
            TabIndex        =   50
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Vendas Produtos:"
            Height          =   195
            Left            =   375
            TabIndex        =   49
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label lblVendasProdutos 
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
            DataSource      =   "QVendaTotaliza"
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Despesas:"
            Height          =   195
            Left            =   885
            TabIndex        =   47
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lblDespesas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "TotalDespesa"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "QDespesaLancTotaliza"
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Comissão:"
            Height          =   195
            Left            =   915
            TabIndex        =   45
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lblComissao 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Comissao"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "QVendaTotaliza"
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label lblTotalRecebido 
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
            DataSource      =   "QFormaDePgRecTotaliza"
            Height          =   255
            Left            =   1800
            TabIndex        =   43
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Total Recebido:"
            Height          =   195
            Left            =   495
            TabIndex        =   42
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
            TabIndex        =   41
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Cheques:"
            Height          =   195
            Left            =   960
            TabIndex        =   40
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Compra de Produtos:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "total"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "QProdutoEntra"
            Height          =   255
            Left            =   1800
            TabIndex        =   38
            Top             =   2400
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Galonagem "
         Height          =   3255
         Left            =   4200
         TabIndex        =   33
         Top             =   240
         Width           =   4215
         Begin MSDataGridLib.DataGrid DataGrid7 
            Bindings        =   "frmFechamentoDiarioConfirmado.frx":0AFD
            Height          =   2415
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
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
            ColumnCount     =   3
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "PrecoVenda"
               Caption         =   "Preço"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   """ ""#.##0,000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   2
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Vendido"
               Caption         =   "Vendido"
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
                  ColumnWidth     =   1484,787
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   764,787
               EndProperty
            EndProperty
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
            TabIndex        =   36
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   1800
            TabIndex        =   35
            Top             =   2880
            Width           =   405
         End
      End
      Begin VB.TextBox txtJuros 
         Alignment       =   1  'Right Justify
         DataField       =   "Juros"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dbFechamento"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Vendas:"
         Height          =   195
         Left            =   240
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Despesas:"
         Height          =   195
         Left            =   2040
         TabIndex        =   60
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label lblTotalDespesas 
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
         Left            =   2040
         TabIndex        =   59
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lblRecebimentos 
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
         DataSource      =   "QFormaDePgRecTotaliza"
         Height          =   255
         Left            =   3840
         TabIndex        =   58
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Recebimentos:"
         Height          =   195
         Left            =   3840
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   5760
         TabIndex        =   55
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Juros Cobrado:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   3600
         Width           =   1065
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Compra de Produtos "
      Height          =   4575
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   8775
      Begin VB.CommandButton cmdProdutoEntra 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCod 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   975
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
         Left            =   4200
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTotalEntra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTanqueEntra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cboProdutoEntra 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B16
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid9 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B2F
         Height          =   3135
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "Codigo"
            Caption         =   "Código"
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
            DataField       =   "Qtd"
            Caption         =   "Qtd."
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
         BeginProperty Column03 
            DataField       =   "ValorNota"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,000"
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
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1214,929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3000
         TabIndex        =   30
         Top             =   4200
         Width           =   405
      End
      Begin VB.Label lblProdutoEntraTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "total"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "QProdutoEntra"
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   5280
         TabIndex        =   28
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Cod.:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   4200
         TabIndex        =   25
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblTanque 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Left            =   6720
         TabIndex        =   24
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Cheques "
      Height          =   4575
      Index           =   6
      Left            =   120
      TabIndex        =   131
      Top             =   840
      Width           =   8775
      Begin VB.CommandButton cmdRelaciona 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   6120
         TabIndex        =   133
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   132
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid8 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B4C
         Height          =   3135
         Left            =   120
         TabIndex        =   134
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "DataCheque"
            Caption         =   "DataCheque"
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
            DataField       =   "Valor"
            Caption         =   "Valor"
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
               ColumnWidth     =   569,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1184,882
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   135
         Top             =   4080
         Visible         =   0   'False
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
         TabIndex        =   136
         Top             =   4080
         Visible         =   0   'False
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
         TabIndex        =   137
         Top             =   4080
         Visible         =   0   'False
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
         TabIndex        =   138
         Top             =   4080
         Visible         =   0   'False
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
         TabIndex        =   139
         Top             =   4080
         Visible         =   0   'False
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
         Left            =   3840
         TabIndex        =   140
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   3840
         TabIndex        =   149
         Top             =   3840
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label47 
         Caption         =   "Valor:"
         Height          =   255
         Left            =   4800
         TabIndex        =   148
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   3000
         TabIndex        =   147
         Top             =   3840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   2040
         TabIndex        =   146
         Top             =   3840
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         Height          =   195
         Left            =   1320
         TabIndex        =   145
         Top             =   3840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   720
         TabIndex        =   144
         Top             =   3840
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Comp:"
         Height          =   195
         Left            =   120
         TabIndex        =   143
         Top             =   3840
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   4680
         TabIndex        =   142
         Top             =   3480
         Width           =   405
      End
      Begin VB.Label Label54 
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
         Left            =   5280
         TabIndex        =   141
         Top             =   3480
         Width           =   1815
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Notas "
      Height          =   4575
      Index           =   5
      Left            =   120
      TabIndex        =   122
      Top             =   840
      Width           =   8775
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
         Height          =   285
         Left            =   3600
         TabIndex        =   124
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdInclueNota 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   4920
         TabIndex        =   123
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cboClientesNota 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B64
         Height          =   315
         Left            =   120
         TabIndex        =   125
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   "CodigoCliente"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid6 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B7D
         Height          =   3135
         Left            =   120
         TabIndex        =   126
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         BeginProperty Column01 
            DataField       =   "ValorPrevisto"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,00"
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
               ColumnWidth     =   3119,811
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3600
         TabIndex        =   130
         Top             =   3840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   129
         Top             =   3840
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3120
         TabIndex        =   128
         Top             =   3480
         Width           =   405
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Valor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "QClientesNota"
         Height          =   255
         Left            =   3630
         TabIndex        =   127
         Top             =   3480
         Width           =   1695
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Recebimentos "
      Height          =   4575
      Index           =   4
      Left            =   120
      TabIndex        =   111
      Top             =   840
      Width           =   8775
      Begin VB.CommandButton cmdIncluirRecebimento 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   5040
         TabIndex        =   114
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
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
         Left            =   2760
         TabIndex        =   113
         Top             =   4080
         Visible         =   0   'False
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
         Left            =   4080
         TabIndex        =   112
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboRecebimento 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0B9A
         Height          =   315
         Left            =   120
         TabIndex        =   115
         Top             =   4080
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0BB4
         Height          =   3135
         Left            =   120
         TabIndex        =   116
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         BeginProperty Column01 
            DataField       =   "ValorBruto"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,00"
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
               ColumnWidth     =   2145,26
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1335,118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   121
         Top             =   3840
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2760
         TabIndex        =   120
         Top             =   3840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Operações:"
         Height          =   195
         Left            =   4080
         TabIndex        =   119
         Top             =   3840
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   2040
         TabIndex        =   118
         Top             =   3480
         Width           =   405
      End
      Begin VB.Label Label42 
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
         DataSource      =   "QFormaDePgRecTotaliza"
         Height          =   255
         Left            =   2505
         TabIndex        =   117
         Top             =   3480
         Width           =   1695
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Despesas "
      Height          =   4575
      Index           =   3
      Left            =   120
      TabIndex        =   98
      Top             =   840
      Width           =   8775
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
         TabIndex        =   101
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdIncluirDespesa 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   5280
         TabIndex        =   100
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDespesaObs 
         Height          =   285
         Left            =   120
         TabIndex        =   99
         Top             =   4200
         Visible         =   0   'False
         Width           =   4935
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0BD6
         Height          =   2655
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Valor"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboDespesa 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0BF3
         Height          =   315
         Left            =   120
         TabIndex        =   103
         Top             =   3600
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboConta 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0C0C
         Height          =   315
         Left            =   5280
         TabIndex        =   104
         Top             =   3600
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   5280
         TabIndex        =   110
         Top             =   3360
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Despesa:"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   3360
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3480
         TabIndex        =   108
         Top             =   3360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   3960
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TotalDespesa"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "QDespesaLancTotaliza"
         Height          =   255
         Left            =   5115
         TabIndex        =   106
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   105
         Top             =   3000
         Width           =   405
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Vendas "
      Height          =   4575
      Index           =   2
      Left            =   120
      TabIndex        =   85
      Top             =   840
      Width           =   8775
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
         TabIndex        =   88
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCodProduto 
         Height          =   285
         Left            =   120
         TabIndex        =   87
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdIncluirVendas 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   6720
         TabIndex        =   86
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cboProduto 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0C23
         Height          =   315
         Left            =   1200
         TabIndex        =   89
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0C3D
         Height          =   3135
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "CodProduto"
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
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ValorTotal"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """ ""#.##0,00"
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
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1214,929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   4200
         TabIndex        =   97
         Top             =   3960
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cod.:"
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   1200
         TabIndex        =   95
         Top             =   3960
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   5280
         TabIndex        =   94
         Top             =   3960
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblProdutoTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   93
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label33 
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
         DataSource      =   "QVendaTotaliza"
         Height          =   255
         Left            =   3480
         TabIndex        =   92
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3000
         TabIndex        =   91
         Top             =   3480
         Width           =   405
      End
   End
   Begin VB.Frame Tela 
      Caption         =   " Controle de Tanque "
      Height          =   4575
      Index           =   1
      Left            =   120
      TabIndex        =   76
      Top             =   840
      Width           =   8775
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
         Left            =   2520
         TabIndex        =   79
         Top             =   4080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdIncluirTanque 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   4440
         TabIndex        =   78
         Top             =   3960
         Visible         =   0   'False
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
         TabIndex        =   77
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cboTanque 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0C54
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Tanque"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmFechamentoDiarioConfirmado.frx":0C6C
         Height          =   3375
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "Tanque"
            Caption         =   "Tanque"
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
            DataField       =   "Quantidade"
            Caption         =   "Quantidade"
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
            DataField       =   "Reposicao"
            Caption         =   "Reposição"
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
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1709,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1755,213
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Reposição:"
         Height          =   195
         Left            =   2520
         TabIndex        =   84
         Top             =   3840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Régua:"
         Height          =   195
         Left            =   960
         TabIndex        =   83
         Top             =   3840
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   3840
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      Caption         =   "Turno:"
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Responsável:"
      Height          =   195
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Posto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmFechamentoDiarioConfirmado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intTela As Integer, codigoFechamento As Double, Porta As Integer
Dim CodBar As String

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String


frmImprime.ScaleMode = vbMillimeters


frmImprime.Line (0, 0)-(Largura, 13), RGB(175, 175, 175), BF


StrTemp = "Fechamento de Caixa"

frmImprime.FontName = "Arial"
frmImprime.FontSize = 14
frmImprime.CurrentX = (Largura / 2) - (frmImprime.TextWidth(StrTemp) / 2)
frmImprime.CurrentY = 2
frmImprime.Print StrTemp

frmImprime.FontSize = 8
StrTemp = Format(Dia, "Short date")
frmImprime.CurrentX = 2
frmImprime.Print StrTemp;

StrTemp = "Página: " '& Printer.Page
frmImprime.CurrentX = Largura - frmImprime.TextWidth(StrTemp) - 2
frmImprime.Print StrTemp


frmImprime.CurrentY = frmImprime.CurrentY + 2
StrTemp = "Posto: " & cboPosto.Text
frmImprime.CurrentX = 2
frmImprime.Print StrTemp;

StrTemp = "Data: " & Format(txtData.Value, "Short Date")
frmImprime.CurrentX = 90
frmImprime.Print StrTemp;

StrTemp = "Turno: " & cboTurno.Text
frmImprime.CurrentX = Largura - frmImprime.TextWidth(StrTemp) - 2
frmImprime.Print StrTemp

StrTemp = "Responsável: " & txtResponsavel.Text
frmImprime.CurrentX = Largura - frmImprime.TextWidth(StrTemp) - 2
frmImprime.Print StrTemp


End Sub

Public Sub TiraSaldo(ByVal DataInicial As String)
  Dim Ws As Workspace, Db As Database, dbSaldo As Recordset
  Dim Saldo As Currency, DiaAnterior As String
  
  Screen.MousePointer = vbHourglass
  
  Set Ws = DBEngine.Workspaces(0)
  Set Db = Ws.OpenDatabase(Caminho, , , Conectar)
  
  DiaAnterior = Str(DateAdd("d", -1, CDate(DataInicial)))
  
  Db.Execute "delete *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and tipo='Saldo' and data>=#" & DataInglesa(DataInicial) & "#"
  
  Set dbSaldo = Db.OpenRecordset("select *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and tipo='Saldo' order by Data")
  If dbSaldo.RecordCount <> 0 Then
    dbSaldo.MoveLast
    Saldo = dbSaldo!Valor
  Else
    Saldo = 0
  End If
  
  Set dbSaldo = Db.OpenRecordset("select sum (valor) as Total, Data from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and data>=#" & DataInglesa(DataInicial) & "#  group by data order by data")
  If dbSaldo.RecordCount = 0 Then Exit Sub
  Do While dbSaldo.EOF = False
    Saldo = Saldo + dbSaldo!Total
    dbConcilia.Recordset.AddNew
    dbConcilia.Recordset!codigoconta = dbContas.Recordset!codigoconta
    dbConcilia.Recordset!Data = dbSaldo!Data
    dbConcilia.Recordset!tipo = "Saldo"
    dbConcilia.Recordset!Codigo = 0
    dbConcilia.Recordset!Descri = "Saldo"
    dbConcilia.Recordset!NrDocumento = "999999999"
    dbConcilia.Recordset!Valor = Saldo
    dbConcilia.Recordset.Update
    dbSaldo.MoveNext
  Loop
End Sub

Private Sub Totaliza()
Dim Tempvalor As Currency, totalGeral As Currency

With QBicoMovimentaTotal
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from QBicoMovimentoTotaliza where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QDespesaLancTotaliza
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from QDespesaLancTotaliza where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QFormaDePgRecTotaliza
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from QFormaDePagamentoRecebidoTotaliza where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QVendaTotaliza
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from QVendaTotaliza where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QClientesNota
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from qclientesnota where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QProdutoEntra
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select sum(valornota) as total from produtosentrada where codigofechamento=" & codigoFechamento
  .Refresh
End With
With QGalonagem
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
  .RecordSource = "select *from qgalonagemproduto where codigofechamento=" & codigoFechamento
  .Refresh
  Tempvalor = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Tempvalor = Tempvalor + .Recordset("vendido")
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With


QCheques.Refresh
QCheques.Refresh
QCheques.Refresh

lblGalonagem.Caption = Tempvalor

Tempvalor = 0
If IsNumeric(lblVendasCombustivel.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblVendasCombustivel.Caption)
End If
If IsNumeric(lblVendasProdutos.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblVendasProdutos.Caption)
End If

txtJuros.Text = Format(dbFechamento.Recordset!juros, "currency")
If IsNumeric(txtJuros.Text) = True Then
  Tempvalor = Tempvalor + CCur(txtJuros.Text)
End If

totalGeral = -Tempvalor
lblTotalVendas.Caption = Format(Tempvalor, "Currency")

Tempvalor = 0
If IsNumeric(lblDespesas.Caption) = True Then
  Tempvalor = Tempvalor - CCur(lblDespesas.Caption)
End If
If IsNumeric(lblProdutoEntraTotal.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblProdutoEntraTotal.Caption)
End If

totalGeral = totalGeral + Tempvalor
lblTotalDespesas.Caption = Format(Tempvalor, "Currency")
Tempvalor = 0

If IsNumeric(lblClientesNota.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblClientesNota.Caption)
End If

If IsNumeric(lblTotalChequeResumo.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblTotalChequeResumo.Caption)
End If

If IsNumeric(lblTotalRecebido.Caption) = True Then
  Tempvalor = Tempvalor + CCur(lblTotalRecebido.Caption)
End If


totalGeral = totalGeral + Tempvalor
lblRecebimentos.Caption = Format(Tempvalor, "currency")
lblDiferenca.Caption = Format(totalGeral, "Currency")


End Sub

Private Sub AbreFechamento(ByVal Fechamento As Double, ByVal Posto As Double)
  With dbPosto
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from postos order by nome"
    .Refresh
  End With
  With dbResponsavel
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from vendedores where gerente=-1 order by nome"
    .Refresh
  End With
  With dbFechamento
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from FechamentoDiario order by Data, Hora"
    .Refresh
  End With
  With dbBicoMovimento
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from bicomovimento where codigofechamento=" & Fechamento & " order by bico"
    .Refresh
  End With
  With dbBico
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from bicos where codigoposto=" & Posto & " order by bico"
    .Refresh
  End With
  With dbTanques
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from tanques where codigoposto=" & Posto & " order by tanque"
    .Refresh
  End With
  With dbTanquesMovimento
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from tanquesMovimento where codigofechamento=" & Fechamento & " and codigoposto=" & Posto & " order by tanque"
    .Refresh
  End With
  With dbProdutos
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from produtos order by descri"
    .Refresh
  End With
  With dbProdutos2
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from produtos where combustivel=0 order by descri"
    .Refresh
  End With
  With dbVendas
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from venda where codigofechamento=" & Fechamento & " order by descri"
    .Refresh
  End With
  With dbDespesas
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from DespesaTipo order by descri"
    .Refresh
  End With
  With dbDespesasLanc
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from DespesasLanc where codigofechamento=" & Fechamento & " order by descri"
    .Refresh
  End With
  With dbFormaDePg
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from FormaDePagamento order by descri"
    .Refresh
  End With
  With dbFormaDePgRecebido
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from FormaDePagamentoRecebido where codigofechamento=" & Fechamento & " order by descri"
    .Refresh
  End With
  With dbContas
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from contas order by descri"
    .Refresh
  End With
  With dbStatus
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from status"
    .Refresh
  End With
  With QTemp
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from QBicoMovimentoTotalTanque"
    .Refresh
  End With
  With dbContas
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from contas"
    .Refresh
  End With
  With dbPrevisaoRecebe
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from PrevisaoRecebimentos"
    .Refresh
  End With
  With dbClientes
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from clientes where mensalista=-1 order by nome"
    .Refresh
  End With
  With dbClientesNota
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from clientesNota where codigofechamento=" & codigoFechamento & " order by nome"
    .Refresh
  End With
  With QClientesNota
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from qclientesnota where codigofechamento=" & codigoFechamento
    .Refresh
  End With
  With dbGalonagem
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from galonagem where codigofechamento=" & codigoFechamento
    .Refresh
  End With
  With dbCompensaPendente
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .Refresh
  End With
  With dbTurno
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .Refresh
  End With
  With dbCheques
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from cheques where codigofechamento=" & codigoFechamento
    .Refresh
  End With
  With QCheques
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select sum(valor) as Total from cheques where codigofechamento=" & codigoFechamento
    .Refresh
  End With
  With QCheques
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .Refresh
  End With
  With dbProdutoEntra
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Persist Security Info=False"
    .RecordSource = "select *from produtosentrada where codigofechamento=" & codigoFechamento
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
  .Recordset.Find "bico=" & cboBico.Text, , adSearchForward, 0
  If .Recordset.EOF = False Then
    cboBico.Text = .Recordset("bico")
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
  .Recordset.Find "nome='" & cboClientesNota.Text & "'"
  If .Recordset.EOF = False Then
    cboClientesNota.Text = .Recordset("nome")
  End If
End With
End Sub

Private Sub cboConta_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConta_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboConta_LostFocus()
Me.KeyPreview = True
With dbContas
  .Refresh
  If cboConta.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboConta.Text & "'", , adSearchForward, 0
  If .Recordset.EOF = False Then
    cboConta.Text = .Recordset("descri")
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
  .Recordset.Find "descri='" & cboDespesa.Text & "'", , adSearchForward, 0
  If .Recordset.EOF = False Then
    cboDespesa.Text = .Recordset("descri")
  End If
End With
End Sub

Private Sub cboProduto_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboProduto_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  End Select
End Sub

Private Sub cboProduto_LostFocus()
Me.KeyPreview = True
With dbProdutos2
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboProduto.Text & "'", , adSearchForward, 0
  If .Recordset.EOF = False Then
    txtCodProduto.Text = .Recordset("codigo")
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
  .Recordset.Find "descri='" & cboRecebimento.Text & "'", , adSearchForward, 0
  If .Recordset.EOF = False Then
    cboRecebimento.Text = .Recordset("descri")
  End If
End With
End Sub

Private Sub cboResponsavel_LostFocus()
With dbResponsavel
  .Refresh
  If cboResponsavel.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboResponsavel.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
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
  .Recordset.Find "tanque=" & cboTanque.Text, , adSearchForward, 0
  If .Recordset.EOF = False Then
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
  .Recordset.Find "descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  cboTurno.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdAnterior_Click()
For i = 0 To Tela.Count - 1
  Tela(i).Visible = False
Next i

If intTela < 0 Then
  Exit Sub
Else
  intTela = intTela - 1
  cmdProximo.Enabled = True
  Tela(intTela).Visible = True
  If intTela = 0 Then
    cmdAnterior.Enabled = False
  End If
End If
cmdFinalizar.Enabled = False

End Sub

Private Sub cmdCancelar_Click()
Dim Resposta As Integer
With cmdCancelar
  If .Caption = "&Sair" Then
    Unload Me
  Else
    For i = 0 To Tela.Count - 1
      Tela(i).Visible = False
    Next i
    cmdProximo.Enabled = False
    cmdAnterior.Enabled = False
    cmdFinalizar.Enabled = False
    cmdCancelar.Caption = "&Sair"
    
    cmdProximo.Enabled = False
    cmdInlueBomba.Enabled = True
    cboPosto.Enabled = True
    cboResponsavel.Enabled = True
    txtData.Enabled = True
    cboTurno.Enabled = True
    txtResponsavel.Text = ""
    cboPosto.SetFocus
  End If
End With
End Sub

Private Sub cmdFinalizar_Click()
Dim Resposta As Integer, LucroVenda As Currency, Tempvalor As Currency
Dim DifEstoque As Double, ValorDiferenca As Currency
Dim Dias As Double, ReceberData As Date, StrTemp As String
Dim Mes As Boolean, Intervalo As String

'Primeiro verifica se deseja fechar mesmo
Resposta = MsgBox("Deseja finalizar o fechamento agora?!", vbYesNo, "Fechamento")
If Resposta = vbNo Then Exit Sub

Screen.MousePointer = vbHourglass

'Atualiza todos os dados do lançamento
AbreFechamento codigoFechamento, dbPosto.Recordset("codigoposto")
Totaliza

'calcula o movimento de bico, registrando o contador da bomba, o lucro de venda
'e o estoque do tanque
With dbBicoMovimento
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbBico.Recordset.Find "codigobico=" & .Recordset("codigobico"), , adSearchForward, 0
      If .Recordset.EOF = True Then
        MsgBox "O bico " & .Recordset("bico") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
      End If
      dbBico.Recordset("ultimonumero") = .Recordset("valorfinal")
      dbBico.Recordset("ultimomecanico") = .Recordset("mecanicoFinal")
      dbBico.Recordset.Update
      
      QTemp.RecordSource = "select *from QBicoMovimentoTotalTanque where codigofechamento=" & codigoFechamento & " and tanque=" & .Recordset("tanque")
      QTemp.Refresh
      If QTemp.Recordset.RecordCount <> 0 Then
        dbTanques.Refresh
        dbTanques.Refresh
        dbTanques.Refresh
        If dbTanques.Recordset.RecordCount = 0 Then
          MsgBox "Cadastro de tanques vazio!"
        Else
          dbTanques.Recordset.Find "tanque=" & dbBicoMovimento.Recordset("tanque"), , adSearchForward, 0
          If .Recordset.EOF = True Then
            MsgBox "O tanque " & dbBicoMovimento.Recordset("tanque") & " não foi encontrado no cadastro!", vbCritical, "Erro!"
          Else
            LucroVenda = (.Recordset!precounitario - .Recordset!precocompra) * .Recordset!vendas
            dbTanques.Recordset("estoque") = dbTanques.Recordset("estoque") - QTemp.Recordset("vendido")
            dbTanques.Recordset.Update
          End If
        End If
        dbProdutos.Refresh
        dbProdutos.Refresh
        dbProdutos.Refresh
        If dbProdutos.Recordset.RecordCount <> 0 Then
          dbProdutos.Recordset.Find "codigoproduto=" & .Recordset("codigoproduto")
          If dbProdutos.Recordset.EOF = False Then
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


'calcula a diferença entre o estoque virtual e o estoque
'físico e registra no estatus
With QTemp
  .RecordSource = "select *from QTanqueMovimentoTotal where codigofechamento=" & codigoFechamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbTanques.Refresh
      dbTanques.Refresh
      dbTanques.Refresh
      dbTanques.Recordset.Find "Tanque=" & .Recordset("tanque"), , adSearchForward, 0
      If .Recordset.EOF = True Then
        MsgBox "O cadastro do tanque " & .Recordset("tanque") & " não foi encontrado!"
      Else
        dbTanques.Recordset("estoquefisico") = .Recordset("estoque")
        dbTanques.Recordset("estoque") = dbTanques.Recordset("estoque") + QTemp.Recordset("entrada")
        dbProdutos.Refresh
        If dbProdutos.Recordset.RecordCount <> 0 Then
          dbProdutos.Recordset.Find "codigoproduto=" & dbTanques.Recordset!CodigoProduto
          If dbProdutos.Recordset.EOF = False Then
            dbProdutos.Recordset!Estoque = dbProdutos.Recordset!Estoque + QTemp.Recordset("entrada")
            dbProdutos.Recordset.Update
            dbProdutos.Refresh
          End If
        End If
        If dbTanques.Recordset("estoquefisico") <> 0 Then
          dbTanques.Recordset("diferenca") = dbTanques.Recordset("estoquefisico") - dbTanques.Recordset("estoque")
        End If
        Tempvalor = dbTanques.Recordset("precocompra")
        DifEstoque = dbTanques.Recordset("diferenca")
        ValorDiferenca = ValorDiferenca + (Tempvalor * DifEstoque)
        dbTanques.Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  Else
    With dbTanques
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          If .Recordset("estoquefisico") <> 0 Then
            .Recordset("diferenca") = .Recordset("estoquefisico") - .Recordset("estoque")
          End If
          Tempvalor = .Recordset("precocompra")
          DifEstoque = .Recordset("diferenca")
          ValorDiferenca = ValorDiferenca + (Tempvalor * DifEstoque)
          .Recordset.Update
          .Recordset.MoveNext
        Loop
      End If
    End With
  End If
  dbStatus.Recordset.Resync adAffectAllChapters, adResyncUnderlyingValues
  dbStatus.Recordset("diferencacombustivel") = ValorDiferenca
  dbStatus.Recordset.Update
End With

With dbVendas
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbProdutos.Refresh
      dbProdutos2.Recordset.Find "codigoproduto=" & .Recordset("codigoproduto"), , adSearchForward, 0
      If .Recordset.EOF = True Then
        MsgBox "O produto " & .Recordset("codigoproduto") & " - " & .Recordset("descri") & " não foi encontrado no cadastro de produtos!"
      Else
        LucroVenda = (.Recordset("valorunitario") * .Recordset("quantidade")) - (dbProdutos2.Recordset("precocompra") * .Recordset("quantidade")) - .Recordset("ValorDesconto")
        dbProdutos2.Recordset("estoque") = dbProdutos2.Recordset("estoque") - .Recordset("quantidade")
        dbProdutos2.Recordset!LucroVenda = dbProdutos2.Recordset!LucroVenda + LucroVenda
        dbProdutos2.Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
  
End With

''Totaliza as despesas para lançar no status e no saldo das contas
'With dbDespesasLanc
'  .Refresh
'  LucroVenda = 0
'  If .Recordset.RecordCount <> 0 Then
'    .Recordset.MoveLast
'    .Recordset.MoveFirst
'    Do While .Recordset.EOF = False
'      TempValor = .Recordset("valor")
'      LucroVenda = LucroVenda + TempValor
'      'debita da conta
'      dbContas.Refresh
'      dbContas.Refresh
'      dbContas.Refresh
'      If dbContas.Recordset.RecordCount = 0 Then
'        MsgBox "Não foi possível encontrar nenhuma conta cadastrada!", vbCritical, "Erro!"
'      Else
'        dbContas.Recordset.Find "codigoconta=" & .Recordset("codigoconta")
'        If dbContas.Recordset.EOF = True Then
'          MsgBox "Nao foi possível encontrar o cadastro da conta " & .Recordset("conta") & "!", vbCritical, "Erro!"
'        Else
'          dbContas.Recordset("saldo") = dbContas.Recordset("saldo") - TempValor
'          dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("Previsao")
'          If TempValor < 0 Then
'            If dbContas.Recordset("temcpmf") = True Then
'              dbContas.Recordset("CPMF") = dbContas.Recordset("CPMF") + (TempValor * CPMF)
'            End If
'          End If
'          dbContas.Recordset.Update
'        End If
'      End If
'
'      With dbCompensaPendente
'        .Recordset.AddNew
'        .Recordset!codigoconta = dbDespesasLanc.Recordset!codigoconta
'        .Recordset!CodigoDespesa = dbDespesasLanc.Recordset!CodigoDespesa
'        .Recordset!Descri = dbDespesasLanc.Recordset!Descri
'        .Recordset!Valor = dbDespesasLanc.Recordset!Valor
'        .Recordset!Data = dbDespesasLanc.Recordset!Data
'        .Recordset!vencimento = dbDespesasLanc.Recordset!Data
'        .Recordset!CodigoDespesaLanc = dbDespesasLanc.Recordset!CodigoDespesaLanc
'        .Recordset.Update
'      End With
'      .Recordset.MoveNext
'    Loop
'
'    dbStatus.Recordset("despesas") = dbStatus.Recordset("despesas") - LucroVenda
'    dbStatus.Recordset.Update
'  End If
'End With

'Totaliza os recebimentos e lança nas contas
With QTemp
  .RecordSource = "select *from QFormadePgContasRecebido where codigofechamento=" & codigoFechamento
  .Refresh
  LucroVenda = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      LucroVenda = LucroVenda + .Recordset("valor")
      Dias = .Recordset("reembolso")
      Mes = .Recordset("mes")
      If Mes = True Then
        Intervalo = "m"
      Else
        Intervalo = "d"
      End If
      If Dias > 0 Then
        ReceberData = DateAdd(Intervalo, Dias, txtData.Value)
      Else
        Dias = .Recordset("diadomes")
        If Dias > 0 Then
          If Dias >= txtData.Day Then
            If Dias < 28 Then
              StrTemp = Dias & "/" & (txtData.Month + 1) & "/" & txtData.Year
            Else
              StrTemp = Dias & "/" & (txtData.Month + 1) & "/" & txtData.Year
              Do While IsDate(StrTemp) = False
                Dias = Dias - 1
                If Dias <= 0 Then Dias = 31
                StrTemp = Dias & "/" & (txtData.Month + 1) & "/" & txtData.Year
              Loop
            End If
            ReceberData = CDate(StrTemp)
          Else
            If Dias < 28 Then
              StrTemp = Dias & "/" & txtData.Month & "/" & txtData.Year
            Else
              StrTemp = Dias & "/" & txtData.Month & "/" & txtData.Year
              Do While IsDate(StrTemp) = False
                Dias = Dias - 1
                If Dias <= 0 Then Dias = 31
                StrTemp = Dias & "/" & txtData.Month & "/" & txtData.Year
              Loop
            End If
            ReceberData = CDate(StrTemp)
          End If
        End If
      End If
      If Dias > 0 Then
        Select Case Weekday(ReceberData)
          Case 1 'sabado
            ReceberData = DateAdd("d", 2, ReceberData)
          Case 7 'domingo
            ReceberData = DateAdd("d", 1, ReceberData)
        End Select
        dbPrevisaoRecebe.Recordset.AddNew
        dbPrevisaoRecebe.Recordset("codigoconta") = .Recordset("contas.codigoconta")
        dbPrevisaoRecebe.Recordset("conta") = .Recordset("contas.descri")
        dbPrevisaoRecebe.Recordset("codigoformapagamento") = .Recordset("codigoFormaDePg")
        dbPrevisaoRecebe.Recordset("descri") = .Recordset("formadepagamento.descri")
        dbPrevisaoRecebe.Recordset("dataentrada") = txtData.Value
        dbPrevisaoRecebe.Recordset("dataprevista") = ReceberData
        dbPrevisaoRecebe.Recordset("valorbruto") = .Recordset("valorbruto")
        dbPrevisaoRecebe.Recordset("valorliquidoPrevisto") = .Recordset("valor")
        dbPrevisaoRecebe.Recordset("valordesconto") = .Recordset("valordesconto")
        dbPrevisaoRecebe.Recordset("valortarifa") = .Recordset("valordescTarifa")
        dbPrevisaoRecebe.Recordset("valoroperacao") = .Recordset("valordescoper")
        dbPrevisaoRecebe.Recordset("operacoes") = .Recordset("operacoes")
        dbPrevisaoRecebe.Recordset.Update
      Else
        ReceberData = txtData.Value
        Select Case Weekday(ReceberData)
          Case 1 'sabado
            ReceberData = DateAdd("d", 2, ReceberData)
          Case 7 'domingo
            ReceberData = DateAdd("d", 1, ReceberData)
        End Select
        dbPrevisaoRecebe.Recordset.AddNew
        dbPrevisaoRecebe.Recordset("codigoconta") = .Recordset("contas.codigoconta")
        dbPrevisaoRecebe.Recordset("conta") = .Recordset("contas.descri")
        dbPrevisaoRecebe.Recordset("codigoformapagamento") = .Recordset("codigoFormaDePg")
        dbPrevisaoRecebe.Recordset("descri") = .Recordset("formadepagamento.descri")
        dbPrevisaoRecebe.Recordset("dataentrada") = txtData.Value
        dbPrevisaoRecebe.Recordset("dataprevista") = ReceberData
        dbPrevisaoRecebe.Recordset("valorbruto") = .Recordset("valorbruto")
        dbPrevisaoRecebe.Recordset("valorliquidoPrevisto") = .Recordset("valor")
        dbPrevisaoRecebe.Recordset("valordesconto") = .Recordset("valordesconto")
        dbPrevisaoRecebe.Recordset("valortarifa") = .Recordset("valordescTarifa")
        dbPrevisaoRecebe.Recordset("valoroperacao") = .Recordset("valordescoper")
        dbPrevisaoRecebe.Recordset("operacoes") = .Recordset("operacoes")
        dbPrevisaoRecebe.Recordset!confirmado = True
        dbPrevisaoRecebe.Recordset!datarecebida = txtData.Value
        dbPrevisaoRecebe.Recordset!dataconfirmada = Now
        dbPrevisaoRecebe.Recordset!difrecebido = 0
        dbPrevisaoRecebe.Recordset.Update
        With dbConcilia
          .Refresh
          .Recordset.AddNew
          .Recordset!codigoconta = dbPrevisaoRecebe.Recordset!codigoconta
          .Recordset!Data = ReceberData
          .Recordset!tipo = "Fechamento"
          .Recordset!Codigo = codigoFechamento
          .Recordset!Descri = QTemp.Recordset("formadepagamento.descri")
          .Recordset!NrDocumento = "777777777"
          .Recordset!Valor = QTemp.Recordset!Valor
          .Recordset.Update
        End With
        dbContas.Refresh
        dbContas.Refresh
        dbContas.Refresh
        dbContas.Recordset.Find "codigoconta=" & .Recordset("contas.codigoconta")
        If .Recordset.EOF = True Then
          MsgBox "Conta " & .Recordset("contas.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
        Else
          Tempvalor = .Recordset("valor")
          dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + Tempvalor
          dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
          dbContas.Recordset.Update
        End If
        TiraSaldo ReceberData
      End If
      .Recordset.MoveNext
    Loop
  End If
End With
With dbCheques
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset!CodigoSoma = "1"
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
With dbFechamento
  .Refresh
  .Recordset.Find "codigofechamento=" & codigoFechamento
  If .Recordset.EOF = True Then
    MsgBox "Erro na tabela de fechamento!"
  Else
    .Recordset("totalVendas") = CCur(lblTotalVendas.Caption)
    .Recordset("TotalDespesa") = CCur(lblTotalDespesas.Caption)
    .Recordset("Totalrecebimento") = CCur(lblTotalRecebido.Caption)
    .Recordset("diferenca") = CCur(lblDiferenca.Caption)
    .Recordset("confirmado") = True
    .Recordset.Update
  End If
End With

cboPosto.Enabled = True
cboResponsavel.Enabled = True
txtData.Enabled = True
cboTurno.Enabled = True
cmdAnterior.Enabled = False
cmdProximo.Enabled = False
cmdFinalizar.Enabled = False
cmdInlueBomba.Enabled = True
cmdCancelar.Caption = "&Sair"
cboPosto.SetFocus
For i = 0 To Tela.Count - 1
  Tela(i).Visible = False
Next i
intTela = 0
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double

Largura = 190
Dia = Now

frmImprime.Show
frmImprime.ScaleMode = vbMillimeters

Cabeca Largura, Dia

With dbBicoMovimento
  StrTemp = "Movimento de Bicos"
  frmImprime.CurrentX = 2
  frmImprime.Print StrTemp
  
  frmImprime.CurrentY = frmImprime.CurrentY + 1
  X1 = 0
  Y1 = frmImprime.CurrentY
  X2 = Largura
  
  StrTemp = "Bico"
  frmImprime.CurrentX = 2
  frmImprime.Print StrTemp;
  StrTemp = "Abertura"
  frmImprime.CurrentX = 55 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp;
  StrTemp = "Fechamento"
  frmImprime.CurrentX = 93 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp;
  StrTemp = "Retorno"
  frmImprime.CurrentX = 118 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp;
  StrTemp = "Vendas"
  frmImprime.CurrentX = 138 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp;
  StrTemp = "$ Venda"
  frmImprime.CurrentX = 158 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp;
  StrTemp = "Total"
  frmImprime.CurrentX = Largura - 2 - frmImprime.TextWidth(StrTemp)
  frmImprime.Print StrTemp
  
  
  frmImprime.CurrentY = frmImprime.CurrentY + 1
  Y2 = frmImprime.CurrentY
  frmImprime.Line (0, frmImprime.CurrentY)-(Largura, frmImprime.CurrentY)
  frmImprime.CurrentY = frmImprime.CurrentY + 1
  
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = .Recordset!bico
      frmImprime.CurrentX = 2
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!ValorInicial
      frmImprime.CurrentX = 55 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!Valorfinal
      frmImprime.CurrentX = 93 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!Retorno
      frmImprime.CurrentX = 118 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!vendas
      frmImprime.CurrentX = 138 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!precounitario
      frmImprime.CurrentX = 158 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp;
      StrTemp = .Recordset!ValorVendido
      frmImprime.CurrentX = Largura - 2 - frmImprime.TextWidth(StrTemp)
      frmImprime.Print StrTemp
      
      Y2 = frmImprime.CurrentY
      .Recordset.MoveNext
    Loop
  End If
  frmImprime.Line (X1, Y1)-(X2, Y1)
  frmImprime.Line (X1, Y1)-(X1, Y2)
  frmImprime.Line (X2, Y1)-(X2, Y2)
  frmImprime.Line (X1, Y2)-(X2, Y2)
End With


End Sub

Private Sub cmdInclueBico_Click()
Dim ValorUnitario As Currency, ValorInicial As Double, Valorfinal As Double
Dim Quantidade As Double, Mecanico As Double, Retorno As Double


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
Valorfinal = CDbl(txtBicoEncerra.Text)
If Valorfinal < dbBico.Recordset("ultimonumero") Then
  If Valorfinal < 3000 Then
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

If IsNumeric(txtMecanico.Text) = False Then
  MsgBox "Informe um numero de leitura mecânica correto!", vbCritical, "Erro!"
  txtMecanico.SetFocus
  Exit Sub
End If
Mecanico = CDbl(txtMecanico.Text)

If Mecanico < dbBico.Recordset("ultimomecanico") Then
  If Mecanico < 3000 Then
    If dbBico.Recordset("ultimomecanico") > 996000 Then
      Resposta = MsgBox("Este lançamento está acusando que a numeração do bico ultrapassou o número 999999, isto está correto?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
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
Tempvalor = (Mecanico - dbBico.Recordset("ultimomecanico")) - (Valorfinal - dbBico.Recordset("ultimonumero"))
If Tempvalor > 5 Or Tempvalor < -5 Then
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
  .Recordset.Find "codigoproduto=" & dbBico.Recordset("codigoproduto")
  If .Recordset.EOF = True Then
    MsgBox "Produto da bomba não encontrado!", vbCritical, "Erro!"
    Exit Sub
  End If
  ValorUnitario = dbBico.Recordset("precovenda")
End With

If Valorfinal < ValorInicial Then
  Quantidade = (999999 - ValorInicial) + Valorfinal - Retorno
Else
  Quantidade = Valorfinal - ValorInicial - Retorno
End If

With dbBicoMovimento
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Refresh
    .Recordset.Find "bico=" & cboBico.Text
    If .Recordset.EOF = True Then
      .Recordset.AddNew
    Else
      Resposta = MsgBox("Bico já lançado! Deseja alterar?", vbYesNo + vbDefaultButton2, "Bico lançado!")
      If Resposta = vbNo Then Exit Sub
    End If
  End If
  .Recordset("codigoFechamento") = codigoFechamento
  .Recordset("Data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("codigobico") = dbBico.Recordset("codigobico")
  .Recordset("bico") = dbBico.Recordset("bico")
  .Recordset("valorinicial") = ValorInicial
  .Recordset("valorfinal") = Valorfinal
  .Recordset("mecanicoInicial") = dbBico.Recordset("ultimomecanico")
  .Recordset("mecanicofinal") = Mecanico
  .Recordset("precocompra") = dbProdutos.Recordset!precocompra
  .Recordset("precounitario") = ValorUnitario
  .Recordset("vendas") = Quantidade
  .Recordset("retorno") = Retorno
  .Recordset("valorvendido") = Quantidade * ValorUnitario
  .Recordset("tanque") = dbBico.Recordset("tanque")
  .Recordset("codigoProduto") = dbProdutos.Recordset("Codigoproduto")
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With
DataGrid1.Refresh
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

DataPrevista = CDate(Format(dbClientes.Recordset("diapagamento"), "00") & "/" & txtData.Month & "/" & txtData.Year)
If DataPrevista < Date Then
  DataPrevista = DateAdd("m", 1, DataPrevista)
End If

With dbClientesNota
  .Recordset.AddNew
  .Recordset("codigofechamento") = codigoFechamento
  .Recordset("codigocliente") = dbClientes.Recordset("codigoCliente")
  .Recordset("nome") = dbClientes.Recordset("nome")
  .Recordset("datalanc") = Now
  .Recordset("dataprevista") = DataPrevista
  .Recordset("valorprevisto") = CCur(txtNotaValor.Text)
  .Recordset!Data = txtData.Value
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With
Totaliza
cboClientesNota.Text = ""
txtNotaValor.Text = ""
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
If cboConta.Text <> dbContas.Recordset("descri") Then
  MsgBox "Selecione uma conta válida!", vbCritical, "Erro!"
  cboConta.SetFocus
  Exit Sub
End If

With dbDespesasLanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = codigoFechamento
  .Recordset!Origem = "Fechamento"
  .Recordset("data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("codigoconta") = dbContas.Recordset("codigoconta")
  .Recordset("conta") = dbContas.Recordset("Descri")
  .Recordset("codigodespesa") = dbDespesas.Recordset("codigodespesa")
  .Recordset("descri") = dbDespesas.Recordset("descri")
  .Recordset("obs") = txtDespesaObs.Text
  .Recordset!compensado = True
  .Recordset("valor") = -CCur(txtDespesaValor.Text)
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With

Totaliza
cboDespesa.Text = ""
txtDespesaValor.Text = ""
txtDespesaObs.Text = ""
cboConta.Text = ""

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
  .Recordset("codigofechamento") = codigoFechamento
  .Recordset("codigoformadepg") = dbFormaDePg.Recordset("codigoPagamento")
  .Recordset("descri") = dbFormaDePg.Recordset("descri")
  .Recordset("valorbruto") = ValorBruto
  .Recordset("valordescoper") = Operacao
  .Recordset("valordesctarifa") = Tarifa
  .Recordset("valordesconto") = DescontoPorcento
  .Recordset("valor") = Liquido
  .Recordset("operacoes") = TotalOper
  .Recordset("data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With

Totaliza

cboRecebimento.Text = ""
txtValorRecebe.Text = ""
txtOperacoes.Text = ""
cboRecebimento.SetFocus

End Sub

Private Sub cmdIncluirTanque_Click()
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
If IsNumeric(txtReposicao.Text) = False Then
  txtReposicao.Text = "0"
End If
With dbTanquesMovimento
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Recordset.Find "tanque=" & dbTanques.Recordset("tanque")
    If .Recordset.EOF = True Then
      .Recordset.AddNew
    End If
  End If
  .Recordset("codigofechamento") = codigoFechamento
  .Recordset("codigoposto") = dbPosto.Recordset("codigoposto")
  .Recordset("tanque") = dbTanques.Recordset("tanque")
  .Recordset("data") = txtData.Value
  .Recordset("hora") = Now
  .Recordset("quantidade") = CDbl(txtRegua.Text)
  .Recordset("reposicao") = CDbl(txtReposicao.Text)
  .Recordset("estoqueantes") = dbTanques.Recordset("estoque")
  .Recordset("estoquedepois") = 0
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With
Totaliza
txtRegua.Text = ""
txtReposicao.Text = ""
cboTanque.SetFocus

End Sub

Private Sub cmdIncluirVendas_Click()
Dim CodigoProduto As Double, codigoPosto As Double
Dim Descricao As String, Qtd As Double, ValorUnitario As Currency
Dim ValorTotal As Currency, CodigoVendedor As Double
Dim Comissao As Double, ValorComissao As Currency
Dim CodProduto As Double

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
If cboProduto.Text <> dbProdutos2.Recordset("descri") Then
  MsgBox "Escolha um Produto a ser incluído na lista de Vendidos!", vbCritical, "Erro!"
  cboProduto.SetFocus
  Exit Sub
End If
If IsNumeric(txtProdutoQuantidade.Text) = False Then
  MsgBox "Informe uma quantidade correta!"
  txtProdutoQuantidade.SetFocus
  Exit Sub
End If

Qtd = CDbl(txtProdutoQuantidade.Text)
With dbProdutos2
  CodigoProduto = .Recordset("codigoproduto")
  Descricao = .Recordset("descri")
  ValorUnitario = .Recordset("precovenda")
  ValorTotal = ValorUnitario * Qtd
  Comissao = .Recordset("comissao")
  ValorComissao = ValorTotal * Comissao
  ValorComissao = ValorComissao + .Recordset!comissaovalor
  CodProduto = .Recordset("codigo")
End With
CodigoVendedor = dbResponsavel.Recordset("codigovendedor")
codigoPosto = dbPosto.Recordset("codigoposto")

With dbVendas
  .Recordset.AddNew
  .Recordset("codigoposto") = codigoPosto
  .Recordset("codigofechamento") = codigoFechamento
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
  .Recordset.Update
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
  .Refresh
  Do While .Recordset.State = adStateExecuting
    DoEvents
  Loop
End With
Totaliza
txtCodProduto.Text = ""
cboProduto.Text = ""
txtProdutoQuantidade.Text = ""
lblProdutoTotal.Caption = ""
txtCodProduto.SetFocus

End Sub

Private Sub cmdInlueBomba_Click()

If dbPosto.Recordset("nome") <> cboPosto.Text Then
  MsgBox "Posto inválido!", vbCritical, "Erro!"
  cboPosto.SetFocus
  Exit Sub
End If
If dbResponsavel.Recordset.EOF = True Then
  MsgBox "Responsável inválido!", vbCritical, "Erro!"
  cboResponsavel.SetFocus
  Exit Sub
End If
'If dbResponsavel.Recordset("nome") <> cboResponsavel.Text Then
'  MsgBox "Responsável inválido!", vbCritical, "Erro!"
'  cboResponsavel.SetFocus
'  Exit Sub
'End If
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
  .RecordSource = "select *from FechamentoDiario where codigoposto=" & dbPosto.Recordset("codigoPosto") & " and data=#" & DataInglesa(Trim(Str(txtData.Value))) & "# and cancelado=0 and confirmado=-1 and codigoturno=" & dbTurno.Recordset!codigoturno
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe esse fechamento!"
    Exit Sub
  End If
  codigoFechamento = .Recordset("codigofechamento")
End With
With dbResponsavel
  .RecordSource = "select *from vendedores where codigovendedor=" & dbFechamento.Recordset!codigoresponsavel
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    txtResponsavel.Text = .Recordset!nome
  Else
    txtResponsavel.Text = ""
  End If
End With
AbreFechamento codigoFechamento, dbPosto.Recordset("codigoposto")
Totaliza
intTela = 0
Totaliza
Tela(0).Visible = True
cmdProximo.Enabled = True
cmdInlueBomba.Enabled = False
cmdCancelar.Caption = "&Cancelar"
cboPosto.Enabled = False
cboResponsavel.Enabled = False
txtData.Enabled = False
cboTurno.Enabled = False
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdProximo_Click()
For i = 0 To Tela.Count - 1
  Tela(i).Visible = False
Next i
If intTela < 0 Then
  Exit Sub
Else
  intTela = intTela + 1
  cmdAnterior.Enabled = True
  Tela(intTela).Visible = True
  If intTela = Tela.Count - 1 Then
    cmdProximo.Enabled = False
    cmdFinalizar.Enabled = True
  End If
End If

If intTela = Tela.Count - 1 Then
  Screen.MousePointer = vbHourglass
  Totaliza
  Screen.MousePointer = vbDefault
End If
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

With dbCheques
  .Refresh
  .Recordset.Find "comp='" & MaskEdBox1(0).Text & "'"
  If .Recordset.EOF = False Then
    .Recordset.Find "banco='" & MaskEdBox1(1).Text & "'"
    If .Recordset.EOF = False Then
      .Recordset.Find "agencia='" & MaskEdBox1(2).Text & "'"
      If .Recordset.EOF = False Then
        .Recordset.Find "conta='" & MaskEdBox1(3).Text & "'"
        If .Recordset.EOF = False Then
          .Recordset.Find "chequeNr='" & MaskEdBox1(4).Text & "'"
          If .Recordset.EOF = False Then
            MsgBox "Cheque já cadastrado!"
            Exit Sub
          End If
        End If
      End If
    End If
  End If
  .Recordset.AddNew
  .Recordset!codigoFechamento = codigoFechamento
  .Recordset!cmc7 = CodBar
  .Recordset!comp = MaskEdBox1(0).Text
  .Recordset!banco = MaskEdBox1(1).Text
  .Recordset!agencia = MaskEdBox1(2).Text
  .Recordset!conta = MaskEdBox1(3).Text
  .Recordset!chequenr = MaskEdBox1(4).Text
  .Recordset!datalanc = Now
  .Recordset!Datacheque = MaskEdBox1(5).Text
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset!CodigoSoma = "2"
  .Recordset.Update
  .Refresh
  .Refresh
  .Refresh
End With

QCheques.Refresh
QCheques.Refresh
QCheques.Refresh

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
txtValor.Text = ""
MaskEdBox1(0).SetFocus

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
  MSComm1.PortOpen = True
End If
AbreFechamento 0, 0
codigoFechamento = 0
Totaliza
intTela = -1
For i = 0 To Tela.Count - 1
  Tela(i).Visible = False
Next i
txtData.Value = Date
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

Private Sub Timer1_Timer()

If MSComm1.InBufferCount > 0 Then
  Timer1.Enabled = False

  'recebeu o codigo de barras armazena na variavel o codigo de barras
  CodBar = ""
  CodBar = MSComm1.Input
  If Len(CodBar) > 1 Then
    Do While Asc(Mid(CodBar, Len(CodBar) - 1, 1)) <> 131
      DoEvents
      CodBar = CodBar & MSComm1.Input
    Loop
    CodBar = Mid(CodBar, 1, Len(CodBar) - 1)
    CodBar = Converte(Trim(CodBar))
    If Len(CodBar) = 33 Then
      'txtCodigo.Text = CodBar
      On Error Resume Next
      MaskEdBox1(0).Text = Mid(CodBar, 11, 3)
      MaskEdBox1(1).Text = Mid(CodBar, 2, 3)
      MaskEdBox1(2).Text = Mid(CodBar, 5, 4)
      MaskEdBox1(3).Text = Mid(CodBar, 26, 6) & "-" & Mid(CodBar, 32, 1)
      MaskEdBox1(4).Text = Mid(CodBar, 14, 6)
      
      MaskEdBox1(5).SetFocus
      
      With Data1
        .Recordset.FindFirst "comp='" & MaskEdBox1(0) & "' and banco='" & MaskEdBox1(1) & "' and agencia='" & MaskEdBox1(2) & "' and conta='" & MaskEdBox1(3) & "' and chequenr='" & MaskEdBox1(4) & "'"
        If .Recordset.NoMatch = False Then
          MaskEdBox1(5).Text = .Recordset("datacheque")
          txtValor.Text = Format(.Recordset("valor"), "Currency")
        End If
      End With
    End If
  End If
  Timer1.Enabled = True
End If

End Sub

Private Sub txtCodProduto_LostFocus()
With dbProdutos2
  If txtCodProduto.Text = "" Then Exit Sub
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "codigo='" & txtCodProduto.Text & "'"
  If .Recordset.EOF = False Then
    cboProduto.Text = .Recordset("descri")
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

Private Sub txtDespesaValor_LostFocus()
With txtDespesaValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtNotaValor_LostFocus()
If IsNumeric(txtNotaValor.Text) = False Then Exit Sub
txtNotaValor.Text = Format(txtNotaValor.Text, "Currency")
End Sub

Private Sub txtProdutoQuantidade_LostFocus()
Dim Valor As Currency
lblProdutoTotal.Caption = ""
With txtProdutoQuantidade
  If IsNumeric(.Text) = False Then Exit Sub
  If dbProdutos2.Recordset("descri") <> cboProduto.Text Then Exit Sub
  Valor = dbProdutos2.Recordset("precovenda")
  Valor = Valor * CDbl(.Text)
  lblProdutoTotal.Caption = Format(Valor, "currency")
End With
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
End With
End Sub

Private Sub txtValorRecebe_LostFocus()
With txtValorRecebe
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With

End Sub
