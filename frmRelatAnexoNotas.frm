VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelatAnexoNotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexo para Notas de Cobrança"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTipoRelat 
      Height          =   315
      ItemData        =   "frmRelatAnexoNotas.frx":0000
      Left            =   120
      List            =   "frmRelatAnexoNotas.frx":0002
      TabIndex        =   17
      Text            =   "Anexo de Notas Selecionado"
      Top             =   6120
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc dbServicosTotal 
      Height          =   375
      Left            =   6000
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select *from Postos"
      Caption         =   "dbServicosTotal"
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
   Begin MSAdodcLib.Adodc dbServicos 
      Height          =   375
      Left            =   6000
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   $"frmRelatAnexoNotas.frx":0004
      Caption         =   "dbServicos"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5953
      _Version        =   393216
      TabOrientation  =   2
      TabHeight       =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cupons Fiscais"
      TabPicture(0)   =   "frmRelatAnexoNotas.frx":00D7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Produtos"
      TabPicture(1)   =   "frmRelatAnexoNotas.frx":00F3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "lblTotalProdutos"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Serviços"
      TabPicture(2)   =   "frmRelatAnexoNotas.frx":010F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "DataGrid4"
      Tab(2).ControlCount=   3
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmRelatAnexoNotas.frx":012B
         Height          =   2415
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
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
         Caption         =   "Cupons Fiscais"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Data"
            Caption         =   "Data"
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
         BeginProperty Column01 
            DataField       =   "Cupom"
            Caption         =   "Cupom"
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
         BeginProperty Column03 
            DataField       =   "ValorPrevisto"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "currency"
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
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1035,213
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmRelatAnexoNotas.frx":0141
         Height          =   2775
         Left            =   -74280
         TabIndex        =   10
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
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
         Caption         =   "Produtos"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "codigo"
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
            DataField       =   "descri"
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
            DataField       =   "quantidade"
            Caption         =   "Qtd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "valorunitariodif"
            Caption         =   "Valor Unitário"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "currency"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "total"
            Caption         =   "Valor Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "currency"
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
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2250,142
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1335,118
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmRelatAnexoNotas.frx":015A
         Height          =   2775
         Left            =   -74280
         TabIndex        =   13
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
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
         Caption         =   "Serviços"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "codigo"
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
            DataField       =   "descri"
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
            DataField       =   "quantidade"
            Caption         =   "Qtd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "valorunitariodif"
            Caption         =   "Valor Unitário"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "currency"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "total"
            Caption         =   "Valor Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "currency"
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
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2250,142
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1335,118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69120
         TabIndex        =   15
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   14
         Top             =   3000
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   12
         Top             =   3000
         Width           =   405
      End
      Begin VB.Label lblTotalProdutos 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69120
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3690
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3180
         TabIndex        =   8
         Top             =   2640
         Width           =   405
      End
   End
   Begin MSAdodcLib.Adodc dbProdutosTotal 
      Height          =   375
      Left            =   6000
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select *from Postos"
      Caption         =   "dbProdutosTotal"
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
      Height          =   375
      Left            =   6000
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   $"frmRelatAnexoNotas.frx":0173
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
   Begin MSAdodcLib.Adodc dbClientesCobranca 
      Height          =   375
      Left            =   2520
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select *from clientescobranca"
      Caption         =   "dbClientesCobranca"
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
   Begin MSAdodcLib.Adodc dbNotasTotal 
      Height          =   375
      Left            =   2520
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select sum(valorprevisto) as total from qclientesnota2produtos where servico=0 and codigosoma='1'"
      Caption         =   "dbNotasTotal"
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
   Begin MSAdodcLib.Adodc dbNotas 
      Height          =   375
      Left            =   2520
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select data, cupom, placa, valorprevisto from qclientesnota2produtos where servico=0 and codigosoma='1' order by cupom"
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
   Begin MSAdodcLib.Adodc dbPostos 
      Height          =   375
      Left            =   2520
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select *from Postos"
      Caption         =   "dbPostos"
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
      Height          =   375
      Left            =   2520
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      RecordSource    =   "Select *from clientes"
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
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6480
      Picture         =   "frmRelatAnexoNotas.frx":0246
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "Imprimir"
      Top             =   5880
      Width           =   735
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   3360
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   57475073
      CurrentDate     =   39331
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatAnexoNotas.frx":0CC8
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4048
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
      Caption         =   "Faturas"
      ColumnCount     =   7
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
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Cliente"
         Caption         =   "Cliente"
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
      BeginProperty Column06 
         DataField       =   "NrNota"
         Caption         =   "Nr. Nota"
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
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3509,858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1214,929
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   5040
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   57475073
      CurrentDate     =   39331
   End
   Begin VB.Label Label7 
      Caption         =   "Tipo de Relatório:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período de Faturas a ser impresso:"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   5880
      Width           =   2460
   End
End
Attribute VB_Name = "frmRelatAnexoNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImprimePorPlaca()
With dbNotas
  .ConnectionString = CaminhoADO
  .RecordSource = "Select data, cupom, placa, valorprevisto from qclientesnota2produtos where codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "' order by placa, data, cupom"
  .Refresh
  If .Recordset.EOF = False And .Recordset.BOF = False Then
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    ImprimeADOGrid DataGrid2, Printer, dbNotas, 3, , , 2, 2, , "Abastecimentos por Veículo", NomePosto, Chr(vbKeyReturn) & "Impresso em: " & Format(Date, "long Date")
    
    Printer.EndDoc
    
  Else
    MsgBox "Não existe cupom a ser impresso!"
  End If
End With



NaoImprime:

With dbNotas
  .ConnectionString = CaminhoADO
  .RecordSource = "Select data, cupom, placa, valorprevisto from qclientesnota2produtos where servico=0 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "' order by cupom"
  .Refresh
End With
End Sub

Private Sub ImprimePeriodo()
Dim Largura As Double, StrTemp As String
With dbClientesCobranca
  .RecordSource = "Select *from clientescobranca where datasoma between #" & DataInglesa(txtDataIni.Value) & " 00:00:01# and #" & DataInglesa(txtDataFim.Value) & " 23:59:59# order by datafechamento desc, cliente"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe fatura a ser impresso o anexo!"
    Exit Sub
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Largura = 190
  
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    DbClientes.Recordset.MoveFirst
    DbClientes.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
    If DbClientes.Recordset.EOF = False Then
      Cabeca Largura
      ImprimeADOGrid DataGrid2, Printer, dbNotas, 3, True, 2
      ImprimeADOGrid DataGrid3, Printer, dbProdutos, 4, True, 2
      Printer.NewPage
      If dbServicos.Recordset.RecordCount <> 0 Then
        Cabeca2 Largura
        ImprimeADOGrid DataGrid4, Printer, dbServicos, 4, True, 2
        Printer.NewPage
      End If
    End If
    .Recordset.MoveNext
  Loop
  Printer.EndDoc
  .RecordSource = "Select *from clientescobranca order by datasoma desc, cliente"
  .Refresh
End With

NaoImprime:

End Sub

Private Sub ImprimeSelecionado()
Dim Largura As Double, StrTemp As String
With dbClientesCobranca
  If .Recordset.EOF = True Then
    MsgBox "Não existe fatura a ser impresso o anexo!"
    Exit Sub
  End If
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Largura = 190
  
  
  DbClientes.Recordset.MoveFirst
  DbClientes.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
  If DbClientes.Recordset.EOF = False Then
    Cabeca Largura
    ImprimeADOGrid DataGrid2, Printer, dbNotas, 3, True, 2
    ImprimeADOGrid DataGrid3, Printer, dbProdutos, 4, True, 2
    Printer.NewPage
    If dbServicos.Recordset.RecordCount <> 0 Then
      Cabeca2 Largura
      ImprimeADOGrid DataGrid4, Printer, dbServicos, 4, True, 2
    End If
  End If
  Printer.EndDoc
End With

NaoImprime:

End Sub

Private Sub CompletaCupons()
Dim Ws As Workspace, db As Database, dbCupons As Recordset
Dim dbNotas As Recordset, Qtd As Double, ValorTotal As Currency
Dim CodigoProduto As Double

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)
Set dbNotas = db.OpenRecordset("select *from clientesnota2 where data>=#08/01/2007# and litros=0 and cupom<>0 order by cupom")

If dbNotas.RecordCount <> 0 Then
  dbNotas.MoveLast
  dbNotas.MoveFirst
  Do While dbNotas.EOF = False
    Set dbCupons = db.OpenRecordset("select *from cuponsfiscais where numerocupom='" & Format(dbNotas!Cupom, "000000") & "'")
    
    If dbCupons.RecordCount <> 0 Then
      dbCupons.MoveLast
      dbCupons.MoveFirst
      If dbCupons.RecordCount = 1 Then
        Qtd = dbCupons!QtdProduto
        ValorTotal = dbCupons!ValorTotal
        CodigoProduto = dbCupons!CodigoProduto
      Else
        Qtd = dbCupons!QtdProduto
        ValorTotal = dbCupons!ValorTotal
        CodigoProduto = dbCupons!CodigoProduto
        dbCupons.MoveNext
        Do While dbCupons.EOF = False
          Qtd = Qtd + dbCupons!QtdProduto
          ValorTotal = ValorTotal + dbCupons!ValorTotal
          CodigoProduto = dbCupons!CodigoProduto
          dbCupons.MoveNext
        Loop
      End If
      dbNotas.Edit
      dbNotas!Litros = Qtd
      dbNotas!Qtd = Qtd
      dbNotas!valorUnitario = ValorTotal / Qtd
      dbNotas!CodigoProduto = CodigoProduto
      dbNotas.Update
    End If
    dbNotas.MoveNext
  Loop
End If
MsgBox "Terminado"
End Sub

Private Sub Cabeca(ByVal Largura As Double)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Largura = 190

StrTemp = "DEMONSTRATIVO DE VENDAS REALIZADAS NO PERÍODO"
Printer.FontSize = 14
Printer.FontBold = True
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

Printer.FontSize = 10
StrTemp = "Emitente: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!Nome
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "Endereço: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!Endereco
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "C.N.P.J.: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!CNPJ
Printer.FontBold = False
Printer.Print StrTemp;

StrTemp = "I.E.: "
Printer.CurrentX = Printer.CurrentX + 15
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!ie
Printer.FontBold = False
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Adquirente: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = DbClientes.Recordset!Nome
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "Endereço: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!Endereco) = False Then
  StrTemp = DbClientes.Recordset!Endereco
Else
  StrTemp = ""
End If
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "C.N.P.J.: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!CNPJ) = False Then
  StrTemp = DbClientes.Recordset!CNPJ
Else
  StrTemp = String(25, "_")
End If
Printer.FontBold = False
Printer.Print StrTemp;

StrTemp = "I.E.: "
Printer.CurrentX = Printer.CurrentX + 15
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!ie) = False Then
  StrTemp = DbClientes.Recordset!ie
Else
  StrTemp = String(25, "_")
End If
Printer.FontBold = False
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5


End Sub

Private Sub Cabeca2(ByVal Largura As Double)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Largura = 190

StrTemp = "DEMONSTRATIVO DE SERVIÇOS REALIZADOS NO PERÍODO"
Printer.FontSize = 14
Printer.FontBold = True
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

Printer.FontSize = 10
StrTemp = "Emitente: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!Nome
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "Endereço: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!Endereco
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "C.N.P.J.: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!CNPJ
Printer.FontBold = False
Printer.Print StrTemp;

StrTemp = "I.E.: "
Printer.CurrentX = Printer.CurrentX + 15
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = dbPostos.Recordset!ie
Printer.FontBold = False
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5

StrTemp = "Adquirente: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
StrTemp = DbClientes.Recordset!Nome
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "Endereço: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!Endereco) = False Then
  StrTemp = DbClientes.Recordset!Endereco
Else
  StrTemp = ""
End If
Printer.FontBold = False
Printer.Print StrTemp

StrTemp = "C.N.P.J.: "
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!CNPJ) = False Then
  StrTemp = DbClientes.Recordset!CNPJ
Else
  StrTemp = String(25, "_")
End If
Printer.FontBold = False
Printer.Print StrTemp;

StrTemp = "I.E.: "
Printer.CurrentX = Printer.CurrentX + 15
Printer.FontBold = True
Printer.Print StrTemp;
If IsNull(DbClientes.Recordset!ie) = False Then
  StrTemp = DbClientes.Recordset!ie
Else
  StrTemp = String(25, "_")
End If
Printer.FontBold = False
Printer.Print StrTemp

Printer.DrawWidth = 3
Printer.CurrentY = Printer.CurrentY + 0.5
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.5


End Sub


Private Sub cmdImprime_Click()

Select Case cboTipoRelat
  Case "Anexo de Notas Selecionado"
    ImprimeSelecionado
  Case "Anexo por Período Faturado"
    ImprimePeriodo
  Case "Abastecimentos Por Placa"
    ImprimePorPlaca
End Select
End Sub

Private Sub dbClientesCobranca_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If dbClientesCobranca.Recordset.EOF = True Then Exit Sub
With dbNotas
  .ConnectionString = CaminhoADO
  .RecordSource = "Select data, cupom, placa, valorprevisto from qclientesnota2produtos where servico=0 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "' order by cupom"
  .Refresh
End With
With dbNotasTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total from qclientesnota2produtos where servico=0 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade, valorunitariodif, codigo, descri from qclientesnota2produtos where servico=0 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "' group by valorunitariodif, codigo, descri order by codigo"
  .Refresh
End With
With dbProdutosTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade from qclientesnota2produtos where servico=0 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalProdutos.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalProdutos.Caption = Format(0, "Currency")
  End If
End With
With dbServicos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade, valorunitariodif, codigo, descri from qclientesnota2produtos where servico=-1 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "' group by valorunitariodif, codigo, descri order by codigo"
  .Refresh
End With
With dbServicosTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(valorprevisto) as total, sum(qtd) as quantidade from qclientesnota2produtos where servico=-1 and codigosoma='" & dbClientesCobranca.Recordset!codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalProdutos.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalProdutos.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF5
    If Shift = 1 Then
      CompletaCupons
    End If
End Select
End Sub

Private Sub Form_Load()
txtDataIni.Value = Date
txtDataFim.Value = Date

With cboTipoRelat
  .Clear
  .AddItem "Anexo de Notas Selecionado"
  .AddItem "Anexo por Período Faturado"
  .AddItem "Abastecimentos Por Placa"
  .Text = "Anexo de Notas Selecionado"
End With
With dbClientesCobranca
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from clientescobranca order by datasoma desc, cliente"
  .Refresh
End With
With DbClientes
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbPostos
  .ConnectionString = CaminhoADO
  .Refresh
End With
End Sub
