VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadBombas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Bombas de Combustível"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmCadBombas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "frmCadBombas.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbProdutos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picButtons"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Adodc1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Numeração"
      TabPicture(1)   =   "frmCadBombas.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "cboBico"
      Tab(1).Control(5)=   "txtDataCaixa"
      Tab(1).Control(6)=   "cboTurno"
      Tab(1).Control(7)=   "txtNumeroInicial"
      Tab(1).Control(8)=   "cmdIncluir"
      Tab(1).Control(9)=   "cmdRemover"
      Tab(1).Control(10)=   "DataGrid2"
      Tab(1).Control(11)=   "dbBicos"
      Tab(1).Control(12)=   "dbTurnos"
      Tab(1).Control(13)=   "dbBicosEncerrantesNovo"
      Tab(1).Control(14)=   "dbFechamentos"
      Tab(1).ControlCount=   15
      Begin MSAdodcLib.Adodc dbFechamentos 
         Height          =   375
         Left            =   -72840
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "select datacaixa, horaini, fechado from fechamentodecaixa order by datacaixa, horaini"
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
      Begin MSAdodcLib.Adodc dbBicosEncerrantesNovo 
         Height          =   375
         Left            =   -72840
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "select *from BicosEncerrantesNovo order by datacaixa, horaini, bico"
         Caption         =   "dbBicosEncerrantesNovo"
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
         Height          =   375
         Left            =   -72840
         Top             =   2520
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "select *from turnos order by horaini"
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
      Begin MSAdodcLib.Adodc dbBicos 
         Height          =   375
         Left            =   -72840
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "select *from bicos order by bico"
         Caption         =   "dbBicos"
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
         Bindings        =   "frmCadBombas.frx":047A
         Height          =   3615
         Left            =   -74760
         TabIndex        =   27
         Top             =   1080
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   6376
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
            DataField       =   "DataCaixa"
            Caption         =   "DataCaixa"
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
         BeginProperty Column03 
            DataField       =   "Inicial"
            Caption         =   "Inicial"
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
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   "Remover"
         Height          =   375
         Left            =   -74760
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -69240
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNumeroInicial 
         Height          =   285
         Left            =   -70920
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo cboTurno 
         Bindings        =   "frmCadBombas.frx":049F
         Height          =   315
         Left            =   -73200
         TabIndex        =   21
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txtDataCaixa 
         Height          =   300
         Left            =   -74760
         TabIndex        =   19
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   39870
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   4860
         Width           =   6765
         _ExtentX        =   11933
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
         RecordSource    =   "SELECT Bicos.* FROM Bicos ORDER BY bico"
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
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   720
         TabIndex        =   30
         Top             =   2940
         Width           =   5415
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Tanque"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   9
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "ultimoNumero"
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Bico"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   615
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmCadBombas.frx":04B6
            DataField       =   "CodigoProduto"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   840
            TabIndex        =   5
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   "CodigoProduto"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Bindings        =   "frmCadBombas.frx":04D0
            DataField       =   "PrecoVenda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$ ""#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   2760
            TabIndex        =   11
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "$ Venda:"
            Height          =   195
            Index           =   5
            Left            =   2760
            TabIndex        =   10
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tanque:"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   8
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Último Número:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bomba:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.PictureBox picButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         ScaleHeight     =   450
         ScaleWidth      =   6765
         TabIndex        =   29
         Top             =   4500
         Width           =   6765
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   300
            Left            =   233
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Fechar"
            Height          =   300
            Left            =   5693
            TabIndex        =   17
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Gravar"
            Height          =   300
            Left            =   4598
            TabIndex        =   16
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Atuali&zar"
            Height          =   300
            Left            =   3503
            TabIndex        =   15
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remover"
            Height          =   300
            Left            =   2400
            TabIndex        =   14
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Adicionar"
            Height          =   300
            Left            =   1313
            TabIndex        =   13
            Top             =   0
            Width           =   975
         End
      End
      Begin MSAdodcLib.Adodc dbProdutos 
         Height          =   330
         Left            =   1440
         Top             =   1260
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         RecordSource    =   "select codigoproduto, descri from produtos where combustivel=-1 order by descri"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCadBombas.frx":04DB
         Height          =   2295
         Left            =   960
         TabIndex        =   1
         Top             =   540
         Width           =   4695
         _ExtentX        =   8281
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
         ColumnCount     =   4
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
            DataField       =   "ultimoNumero"
            Caption         =   "Último Número"
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
         BeginProperty Column03 
            DataField       =   "PrecoVenda"
            Caption         =   "Preço Venda"
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
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1649,764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1170,142
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboBico 
         Bindings        =   "frmCadBombas.frx":04F0
         Height          =   315
         Left            =   -71760
         TabIndex        =   23
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Bico"
         Text            =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Bico:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Inicial:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   -73200
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data do Caixa:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCadBombas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(0).SetFocus
  txtFields(1).Enabled = True
End Sub


Private Sub cmdDelete_Click()
  Dim Resposta As Integer
  
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo, "Excluir!")
  If Resposta = vbNo Then
    Exit Sub
  End If
  
  With Adodc1.Recordset
    If .EOF = False Then
      .Delete adAffectCurrent
    End If
    .Requery
  End With
  Frame1.Enabled = False
End Sub

Private Sub cmdEditar_Click()
Frame1.Enabled = True
txtFields(0).SetFocus
If txtFields(1).Text = 0 Then
  txtFields(1).Enabled = True
End If
End Sub

Private Sub cmdIncluir_Click()
Dim Caixa As String
With dbTurnos
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe turno cadastrado!"
    Exit Sub
  End If
  If cboTurno.Text = "" Then
    MsgBox "Selecione um turno primeiro!"
    cboTurno.SetFocus
    Exit Sub
  End If
  .Recordset.Find "descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = True Then
    MsgBox "Turno inválido!"
    cboTurno.SetFocus
    Exit Sub
  End If
  cboTurno.Text = .Recordset!Descri
End With
With dbBicos
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe bico cadastrado!"
    Exit Sub
  End If
  If IsNumeric(cboBico.Text) = False Then
    MsgBox "Bico incorreto!"
    cboBico.SetFocus
    Exit Sub
  End If
  .Recordset.Find "bico=" & cboBico.Text
  If .Recordset.EOF = True Then
    MsgBox "Bico incorreto!"
    cboBico.SetFocus
    Exit Sub
  End If
End With
If IsNumeric(txtNumeroInicial.Text) = False Then
  MsgBox "Número inválido!"
  txtNumeroInicial.SetFocus
  Exit Sub
End If
Caixa = Str(txtDataCaixa.Value) & " " & Str(dbTurnos.Recordset!HoraIni)
With dbFechamentos
  .Refresh
  .Recordset.Filter = "dia >=#" & Caixa & "#"
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    If .Recordset!Dia = Caixa Then
      If .Recordset!fechado = True Then
        MsgBox "Caixa finalizado!"
        Exit Sub
      End If
    Else
      If .Recordset!fechado = True Then
        MsgBox "Existe caixa superior finalizado!"
        Exit Sub
      End If
    End If
  End If
End With

With dbBicosEncerrantesNovo
  .Recordset.AddNew
  .Recordset!Bico = dbBicos.Recordset!Bico
  .Recordset!DataCaixa = txtDataCaixa.Value
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset!Turno = dbTurnos.Recordset!Descri
  .Recordset!HoraIni = dbTurnos.Recordset!HoraIni
  .Recordset!inicial = CDbl(txtNumeroInicial.Text)
  .Recordset!dataalterado = Now
  .Recordset!Usuario = Usuarios.Nome
  .Recordset.Update
End With
cboTurno.Text = ""
cboBico.Text = ""
txtNumeroInicial.Text = ""
txtDataCaixa.SetFocus
dbBicosEncerrantesNovo.Refresh
dbBicosEncerrantesNovo.Refresh
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  dbProdutos.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer

With dbBicosEncerrantesNovo
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um encerrante primeiro!"
    Exit Sub
  End If
  Caixa = Str(.Recordset!DataCaixa) & " " & Str(.Recordset!HoraIni)
End With

With dbFechamentos
  .Refresh
  .Recordset.Filter = "dia >=#" & Caixa & "#"
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    If .Recordset!Dia = Caixa Then
      If .Recordset!fechado = True Then
        MsgBox "Caixa finalizado!"
        Exit Sub
      End If
    Else
      If .Recordset!fechado = True Then
        MsgBox "Existe caixa superior finalizado!"
        Exit Sub
      End If
    End If
  End If
End With
Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbBicosEncerrantesNovo
  .Recordset.Delete adAffectCurrent
  .Refresh
End With
dbBicosEncerrantesNovo.Refresh
dbBicosEncerrantesNovo.Refresh
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  
  A = Adodc1.Recordset.AbsolutePosition
  Adodc1.Recordset.Update
  Adodc1.Recordset.AbsolutePosition = A
  
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Frame1.Enabled = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If txtFields(1).Text = "" Then Exit Sub
If txtFields(1).Text = 0 Then
  txtFields(1).Enabled = True
Else
  txtFields(1).Enabled = False
End If
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

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select

End Sub

Private Sub Form_Load()
With Adodc1
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbBicos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbBicosEncerrantesNovo
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbFechamentos
  .ConnectionString = CaminhoADO
  .RecordSource = "select datacaixa+horaini as Dia, fechado from fechamentodecaixa order by datacaixa, horaini"
  .Refresh
End With

Select Case Usuarios.Grupo.CadBomba
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
  Case 2 'Liberado
    
End Select
If Usuarios.Grupo.AdmEstatus = 2 Then
  cmdIncluir.Enabled = True
  cmdRemover.Enabled = True
Else
  cmdIncluir.Enabled = False
  cmdRemover.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then
  On Error Resume Next
  Select Case KeyAscii
    Case Asc(".")
      KeyAscii = 0
      SendKeys ","
  End Select
End If
End Sub
