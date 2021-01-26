VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEntradaDeProdutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Produtos"
   ClientHeight    =   6960
   ClientLeft      =   240
   ClientTop       =   720
   ClientWidth     =   7950
   Icon            =   "FrmEntradaDeProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPedido 
      Caption         =   "Pedido de Compra"
      Height          =   375
      Left            =   1680
      TabIndex        =   38
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtNrNota 
      Height          =   315
      Left            =   3000
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtDiasParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5160
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin MSAdodcLib.Adodc dbStatus 
      Height          =   330
      Left            =   3720
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc dbDespesaLanc 
      Height          =   330
      Left            =   3720
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
      RecordSource    =   "select *from despesaslanc2"
      Caption         =   "dbDespesaLanc"
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
   Begin MSAdodcLib.Adodc dbMovimento 
      Height          =   330
      Left            =   3720
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
      RecordSource    =   "select *from ProdutosEntrada2"
      Caption         =   "dbMovimento"
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
      Left            =   3720
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc dbPosto 
      Height          =   330
      Left            =   3720
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
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   480
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc QNota 
      Height          =   330
      Left            =   480
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=0"
      Caption         =   "QNota"
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
   Begin MSAdodcLib.Adodc dbProdutosEntrada 
      Height          =   330
      Left            =   480
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select *from produtosentrada2"
      Caption         =   "dbProdutosEntrada"
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
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   6480
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbNotasCorpo 
      Height          =   330
      Left            =   480
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
      RecordSource    =   "select *from ProdutosNotasCorpo where codigoProdutoNota=0 order by codigo"
      Caption         =   "dbNotasCorpo"
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
   Begin VB.CommandButton cmdNovaNota 
      Caption         =   "Nova"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtEmissao 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   37680
   End
   Begin MSDataListLib.DataCombo cboFornecedor 
      Bindings        =   "FrmEntradaDeProdutos.frx":0442
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.TextBox txtCodFornecedor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin MSAdodcLib.Adodc dbNotas 
      Height          =   330
      Left            =   480
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
      RecordSource    =   "select *from ProdutosNotas"
      Caption         =   "dbNotas"
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
   Begin MSAdodcLib.Adodc dbFornecedores 
      Height          =   330
      Left            =   480
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
      RecordSource    =   "select *from fornecedores order by nome"
      Caption         =   "dbFornecedores"
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
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtValorUnitario 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtTanque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remover"
         Height          =   375
         Left            =   5760
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6720
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmEntradaDeProdutos.frx":045F
         Height          =   3015
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Codigo"
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
            DataField       =   "ValorUnitario"
            Caption         =   "Valor Unitario"
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
         BeginProperty Column03 
            DataField       =   "Quantidade"
            Caption         =   "Quantidade"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "total"
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
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2369,764
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1409,953
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtCodProduto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboProduto 
         Bindings        =   "FrmEntradaDeProdutos.frx":047A
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário:"
         Height          =   195
         Left            =   1080
         TabIndex        =   23
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Left            =   3720
         TabIndex        =   27
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "VTotal"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "QNota"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   5520
         TabIndex        =   34
         Top             =   4560
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2280
         TabIndex        =   25
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Qtd.:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComCtl2.DTPicker txtVencimento 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   37680
   End
   Begin MSDataListLib.DataCombo cboPosto 
      Bindings        =   "FrmEntradaDeProdutos.frx":0493
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Nota:"
      Height          =   195
      Left            =   3000
      TabIndex        =   37
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label15 
      Caption         =   "Dias entre parcelas:"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Posto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      Height          =   195
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "FrmEntradaDeProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoNota As Double

Private Sub cboFornecedor_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboFornecedor_LostFocus()
Me.KeyPreview = True
With dbFornecedores
  If cboFornecedor.Text = "" Then Exit Sub
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboFornecedor.Text & "'"
  If .Recordset.EOF = False Then
    txtCodFornecedor.Text = .Recordset!codigofornecedor
    cboFornecedor.Text = .Recordset!Nome
  End If
End With

End Sub

Private Sub cboProduto_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboProduto_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboProduto_LostFocus()
Me.KeyPreview = True
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    txtCodProduto.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdCancelar_Click()
Dim Resposta As Integer

Resposta = MsgBox("Deseja cancelar a nota atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbYes Then
  dbNotas.Recordset.Delete
  With dbNotasCorpo
    .Refresh
    Do While .Recordset.RecordCount <> 0
      .Recordset.Delete
      .Refresh
      .Refresh
    Loop
  End With
End If

txtCodFornecedor.Enabled = True
cboFornecedor.Enabled = True
txtEmissao.Enabled = True
txtVencimento.Enabled = True
txtNrNota.Enabled = True
cmdNovaNota.Enabled = True
Frame1.Visible = False
cmdConfirma.Enabled = False


End Sub

Private Sub cmdConfirma_Click()
Dim EstoqueAntigo As Double, PrecoCompraAntigo As Currency
Dim EstoqueNovo As Double, PrecoCompraNovo As Currency
Dim Variacao As Currency, Venda As Currency
Dim TempComissao As Currency, Aguardando As Boolean, Parcelas As Integer, DiasParcelas As Integer

If IsNumeric(txtParcelas.Text) = False Then
  MsgBox "Informe a quantidade de parcelas!"
  txtParcelas.SetFocus
  Exit Sub
End If
If CInt(txtParcelas.Text) < 1 Then
  MsgBox "O número de parcelas não pode ser menor que 1!"
  txtParcelas.SetFocus
  Exit Sub
End If
If CInt(txtParcelas.Text) > 1 Then
  If IsNumeric(txtDiasParcelas.Text) = False Then
    MsgBox "Informe quantos dias de intervalo entre as parcelas!"
    txtDiasParcelas.SetFocus
    Exit Sub
  End If
  If CInt(txtDiasParcelas.Text) < 1 Then
    MsgBox "Os dias entre parcelas deve ser maior que 0!"
    txtDiasParcelas.SetFocus
    Exit Sub
  End If
End If
With dbNotasCorpo
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Para confirmar uma nota deve existir pelo menos um produto lançado!"
    Exit Sub
  End If
  .Recordset.MoveFirst
  Aguardando = False
  If .Recordset!Tanque <> 0 Then
    Resposta = MsgBox("Esta nota deve aguardar uma confirmação para entrada no tanque?", vbYesNo + vbDefaultButton2)
    If Resposta = vbYes Then
      Aguardando = True
    Else
      Aguardando = False
    End If
  End If
  Do While .Recordset.EOF = False
    If .Recordset!Tanque <> 0 Then
      'Verifica se não vai ultrapassar o estoque máximo!
      dbTanque.Refresh
      If dbTanque.Recordset.RecordCount <> 0 Then
        dbTanque.Recordset.Sort = "codigoposto, tanque"
        dbTanque.Recordset.Find "codigoposto=" & dbPosto.Recordset!codigoPosto
        If dbTanque.Recordset.EOF = False Then
          dbTanque.Recordset.Find "tanque=" & .Recordset!Tanque
          If dbTanque.Recordset.EOF = False Then
            If dbTanque.Recordset!Estoque + .Recordset!Quantidade > dbTanque.Recordset!estoquefisico Then
              If Aguardando = False Then
                MsgBox "O Tanque " & .Recordset!Tanque & " ficará além da sua capacidade física! Corrija o lançamento!"
                Exit Sub
              End If
            End If
          End If
        End If
      End If
    End If
    .Recordset.MoveNext
  Loop
  
  .Recordset.MoveFirst
  
  Do While .Recordset.EOF = False
    If .Recordset!Tanque <> 0 Then
      'acrecenta no estoque do tanque
      If Aguardando = False Then
        dbTanque.Refresh
        If dbTanque.Recordset.RecordCount <> 0 Then
          dbTanque.Recordset.Sort = "codigoposto, tanque"
          dbTanque.Recordset.Find "codigoposto=" & dbPosto.Recordset!codigoPosto
          If dbTanque.Recordset.EOF = False Then
            dbTanque.Recordset.Find "tanque=" & .Recordset!Tanque
            If dbTanque.Recordset.EOF = False Then
              dbTanque.Recordset!Estoque = dbTanque.Recordset!Estoque + .Recordset!Quantidade
              dbTanque.Recordset.Update
            End If
          End If
        End If
      End If
    End If
    dbProdutos.Refresh
    dbProdutos.Refresh
    If dbProdutos.Recordset.RecordCount <> 0 Then
      'Registra na tabela de produtos
      dbProdutos.Recordset.Find "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.EOF = False Then
        PrecoCompraAntigo = dbProdutos.Recordset!precocompra
        EstoqueAntigo = dbProdutos.Recordset!Estoque
        PrecoCompraNovo = .Recordset!ValorUnitario
        EstoqueNovo = EstoqueAntigo + .Recordset!Quantidade
        Variacao = (EstoqueAntigo * PrecoCompraNovo) - (EstoqueAntigo * PrecoCompraAntigo)
        dbProdutos.Recordset!Estoque = EstoqueNovo
        dbProdutos.Recordset!precocompra = PrecoCompraNovo
        If IsNull(dbProdutos.Recordset!lucrominimo) = True Then
          dbProdutos.Recordset!lucrominimo = 0
        End If
        Venda = 0
        If dbProdutos.Recordset!lucrominimo <> 0 Then
          Venda = PrecoCompraNovo + (PrecoCompraNovo * (dbProdutos.Recordset!lucrominimo / 100))
          If dbProdutos.Recordset!Comissao <> 0 Then
            TempComissao = Venda / (1 - dbProdutos.Recordset!Comissao)
            Venda = TempComissao
          End If
          If dbProdutos.Recordset!comissaovalor <> 0 Then
            Venda = Venda + dbProdutos.Recordset!comissaovalor
          End If
        Else
          Venda = dbProdutos.Recordset!PrecoVenda
        End If
        If Venda <> dbProdutos.Recordset!PrecoVenda Then
TentaDeNovo:
          StrTemp = InputBox("O produto " & dbProdutos.Recordset!Codigo & " - " & dbProdutos.Recordset!Descri & " está sendo alterado o preço de venda de " & Format(dbProdutos.Recordset!PrecoVenda, "Currency") & " para " & Format(Venda, "Currency") & "!", "Alteração de preço!", Format(Venda, "Currency"))
          Do While IsNumeric(StrTemp) = False
            DoEvents
            StrTemp = InputBox("O produto " & dbProdutos.Recordset!Codigo & " - " & dbProdutos.Recordset!Descri & " está sendo alterado o preço de venda de " & Format(dbProdutos.Recordset!PrecoVenda, "Currency") & " para " & Format(Venda, "Currency") & "!", "Alteração de preço!", Format(Venda, "Currency"))
          Loop
          If Venda < (CCur(StrTemp) - 0.5) Or Venda > (CCur(StrTemp) + 0.5) Then
            Permissao = False
            frmPermissao.Show vbModal
            If Permissao = False Then
              GoTo TentaDeNovo
            Else
              Venda = CCur(StrTemp)
            End If
          Else
            Venda = CCur(StrTemp)
          End If
          
          dbProdutos.Recordset!PrecoVenda = Venda
        End If
        dbProdutos.Recordset!Variacao = dbProdutos.Recordset!Variacao + Variacao
        dbProdutos.Recordset.Update
      End If
    End If
    With dbMovimento
      .Recordset.AddNew
      .Recordset!Data = txtEmissao.Value
      .Recordset!CodigoProduto = dbNotasCorpo.Recordset!CodigoProduto
      .Recordset!Codigo = dbNotasCorpo.Recordset!Codigo
      .Recordset!Descri = dbNotasCorpo.Recordset!Descri
      .Recordset!PrecoAntigo = PrecoCompraAntigo
      .Recordset!PrecoNovo = PrecoCompraNovo
      .Recordset!VariaEstoque = Variacao
      .Recordset!Quantidade = dbNotasCorpo.Recordset!Quantidade
      .Recordset!valornota = dbNotasCorpo.Recordset!Total
      .Recordset!Tanque = dbNotasCorpo.Recordset!Tanque
      .Recordset!CodigoNota = CodigoNota
      .Recordset.Update
    End With
    
    With dbStatus
      .Refresh
      .Recordset("variacaoestoque") = .Recordset("variacaoestoque") + Variacao
      .Recordset.Update
      .Refresh
    End With
    .Recordset!Aguardando = Aguardando
    .Recordset.Update
    .Recordset.MoveNext
  Loop
  Parcelas = CInt(txtParcelas.Text)
  If Parcelas > 1 Then
    DiasParcelas = CInt(txtDiasParcelas.Text)
  End If
  For i = 1 To Parcelas
    With dbDespesaLanc
      .Recordset.AddNew
      .Recordset!CodigoFechamento = 0
      .Recordset!Origem = "Despesa"
      .Recordset!Data = txtEmissao.Value
      .Recordset!Hora = Now
      If Parcelas > 1 Then
        If i = 1 Then
          .Recordset!Vencimento = txtVencimento.Value
        Else
          .Recordset!Vencimento = DateAdd("d", (i - 1) * DiasParcelas, txtVencimento.Value)
        End If
      Else
        .Recordset!Vencimento = txtVencimento.Value
      End If
      .Recordset!CodigoDespesa = -1
      .Recordset!Descri = "Compra de Produto"
      .Recordset!obs = Left(cboFornecedor.Text, 25) & Left("-Nota Nr.: " & txtNrNota.Text, 25)
      .Recordset!Valor = -CCur(lblTotal.Caption) / Parcelas
      .Recordset!Produto = True
      .Recordset!fechamentodiario = True
      .Recordset.Update
      .Refresh
    End With
  Next i
End With
With dbNotas
  .Recordset!Confirmado = True
  .Recordset.Update
End With
txtCodFornecedor.Enabled = True
cboFornecedor.Enabled = True
txtEmissao.Enabled = True
txtVencimento.Enabled = True
txtNrNota.Enabled = True
cmdNovaNota.Enabled = True
cmdConfirma.Enabled = False
cboPosto.Enabled = True

txtCodFornecedor.Text = ""
cboFornecedor.Text = ""
txtNrNota.Text = ""
txtParcelas.Text = ""
txtDiasParcelas.Text = ""
Frame1.Visible = False

With dbNotasCorpo
  .RecordSource = "select *from produtosnotascorpo where codigoprodutoNota=0"
  .Refresh
  .Recordset.Sort = "codigo"
End With

cboPosto.SetFocus
End Sub

Private Sub cmdIncluir_Click()
Dim vUnitario As Currency, Quantidade As Double, Total As Currency
Dim Tanque As Double

If cboProduto.Text <> dbProdutos.Recordset!Descri Then
  MsgBox "Produto inválido!"
  txtCodProduto.SetFocus
  Exit Sub
End If
If IsNumeric(txtQtd.Text) = False Then
  MsgBox "Informe uma quantidade válida!"
  txtQtd.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If
Tanque = 0
If dbProdutos.Recordset!Combustivel = True Then
  If IsNumeric(txtTanque.Text) = False Then
    MsgBox "Tanque inválido!"
    txtTanque.SetFocus
    Exit Sub
  Else
    Tanque = CDbl(txtTanque.Text)
  End If
End If
Total = CCur(txtValor.Text)
Quantidade = CDbl(txtQtd.Text)
vUnitario = Total / Quantidade

'TempValor = vUnitario + dbProdutos.Recordset!comissaovalor + (dbProdutos.Recordset!precovenda * (dbProdutos.Recordset!Comissao / 100))
'If TempValor >= (dbProdutos.Recordset!precovenda / 2) Then
'  MsgBox "Margem de lucro abaixo de 50%! Custo= " & Format(TempValor, "Currency") & " / Venda= " & Format(dbProdutos.Recordset!precovenda, "Currency"), vbCritical
'End If

With dbNotasCorpo
  .Refresh
  .Recordset.AddNew
  .Recordset!codigoprodutonota = CodigoNota
  .Recordset!CodigoProduto = dbProdutos.Recordset!CodigoProduto
  .Recordset!Codigo = dbProdutos.Recordset!Codigo
  .Recordset!Descri = dbProdutos.Recordset!Descri
  .Recordset!ValorUnitario = vUnitario
  .Recordset!Quantidade = Quantidade
  .Recordset!Total = Total
  .Recordset!Tanque = Tanque
  .Recordset.Update
  StrTemp = .Recordset.Sort
  .Refresh
  .Recordset.Sort = StrTemp
End With
QNota.Refresh

txtCodProduto.Text = ""
cboProduto.Text = ""
txtQtd.Text = ""
txtTanque.Text = ""
txtValor.Text = ""
txtValorUnitario.Text = ""

txtCodProduto.SetFocus
End Sub

Private Sub cmdNovaNota_Click()
Call cboProduto_LostFocus

If DateDiff("d", Date, txtEmissao.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar nota com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtEmissao.Value) <= -15 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar nota com data anterior a 15 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) >= 90 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar vencimento com data futura acima de 90 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) <= -1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar nota já vencida!"
    Exit Sub
  End If
End If

If cboPosto.Text <> dbPosto.Recordset!Nome Then
  MsgBox "Posto inválido!"
  cboPosto.SetFocus
  Exit Sub
End If
Call cboFornecedor_LostFocus
If dbFornecedores.Recordset!Nome <> cboFornecedor.Text Then
  MsgBox "Fornecedor inválido!"
  cboFornecedor.SetFocus
  Exit Sub
End If

If txtVencimento.Value < Date Then
  MsgBox "Vencimento inválido!"
  txtVencimento.SetFocus
  Exit Sub
End If
If txtNrNota.Text = "" Then
  MsgBox "Número de Nota inválido!"
  txtNrNota.SetFocus
  Exit Sub
End If

With dbNotas
  .RecordSource = "select *from produtosnotas where codigofornecedor=" & dbFornecedores.Recordset!codigofornecedor & " and nrnota='" & txtNrNota.Text & "'"
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If .Recordset!Confirmado = True Then
      MsgBox "Nota já lançada.!"
      Exit Sub
    Else
      CodigoNota = .Recordset!CodigoEntrada
    End If
  Else
    .Recordset.AddNew
    .Recordset!codigofornecedor = dbFornecedores.Recordset!codigofornecedor
    .Recordset!fornecedor = dbFornecedores.Recordset!Nome
    .Recordset!nrnota = ""
    .Recordset!datalancada = Now
    .Recordset!Vencimento = txtVencimento.Value
    .Recordset!datanota = txtEmissao.Value
    .Recordset!Origem = "Entrada de Produtos"
    .Recordset!codigoPosto = dbPosto.Recordset!codigoPosto
    .Recordset.Update
    .Refresh
    .Refresh
    CodigoNota = .Recordset!CodigoEntrada
  End If
End With
With dbNotasCorpo
  .RecordSource = "select *from produtosnotascorpo where codigoprodutoNota=" & CodigoNota
  .Refresh
  .Recordset.Sort = "codigo"
End With
With QNota
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=" & CodigoNota
  .Refresh
End With

txtCodFornecedor.Enabled = False
cboFornecedor.Enabled = False
txtEmissao.Enabled = False
txtVencimento.Enabled = False
txtNrNota.Enabled = False
cmdNovaNota.Enabled = False
cmdConfirma.Enabled = True
cboPosto.Enabled = False

Frame1.Visible = True

txtCodProduto.SetFocus

End Sub

Private Sub cmdPedido_Click()
frmPedidoDeCompra.Show
frmPedidoDeCompra.SetFocus
End Sub

Private Sub cmdRemove_Click()
Dim Resposta As Integer
With dbNotasCorpo
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o produto atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  StrTemp = .Recordset.Sort
  .Refresh
  .Refresh
  .Refresh
  .Recordset.Sort = StrTemp
End With
QNota.Refresh
End Sub

Private Sub cmdSair_Click()
If cmdNovaNota.Enabled = False Then
  Call cmdCancelar_Click
End If
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If dbNotasCorpo.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  dbNotasCorpo.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  dbNotasCorpo.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()

With dbNotas
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbNotasCorpo
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbFornecedores
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbProdutosEntrada
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With QNota
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select sum(total) as VTotal from produtosnotascorpo where codigoprodutonota=0"
  .Refresh
End With
With dbProdutos
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbPosto
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbDespesaLanc
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbMovimento
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbTanque
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbStatus
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With

txtEmissao.Value = Date
txtVencimento.Value = Date
Select Case Usuarios.Grupo.ControleNotas
  Case 1 'Somente leitura
    cmdNovaNota.Enabled = False
    cmdIncluir.Enabled = False
    cmdRemove.Enabled = False
    cmdCancelar.Enabled = False
    cmdConfirma.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_Terminate()
If cmdNovaNota.Enabled = False Then
  Call cmdCancelar_Click
End If
End Sub

Private Sub txtCodFornecedor_GotFocus()
With txtCodFornecedor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodFornecedor_LostFocus()
With dbFornecedores
  If IsNumeric(txtCodFornecedor.Text) = False Then Exit Sub
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "Codigofornecedor=" & txtCodFornecedor.Text
  If .Recordset.EOF = False Then
    txtCodFornecedor.Text = .Recordset!codigofornecedor
    cboFornecedor.Text = .Recordset!Nome
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

Private Sub txtEmissao_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtEmissao_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtQtd_GotFocus()
With txtQtd
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_GotFocus()
txtValor.SelStart = 0
txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
If IsNumeric(txtValor.Text) = False Then Exit Sub
txtValor.Text = Format(txtValor.Text, "#,##0.000")
If IsNumeric(txtQtd.Text) = False Then Exit Sub
a = CCur(txtValor.Text) / CDbl(txtQtd)
txtValorUnitario.Text = Format(a, "#,##0.0000")
End Sub

Private Sub txtValorUnitario_GotFocus()
txtValorUnitario.SelStart = 0
txtValorUnitario.SelLength = Len(txtValorUnitario.Text)
End Sub

Private Sub txtValorUnitario_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValorUnitario_LostFocus()
With txtValorUnitario
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "#,##0.0000")
  If IsNumeric(txtQtd.Text) = False Then Exit Sub
  a = CDbl(txtQtd.Text) * CCur(.Text)
  txtValor.Text = Format(a, "#,##0.000")
End With
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
