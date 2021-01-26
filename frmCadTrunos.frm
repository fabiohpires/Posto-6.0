VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadTurnos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Casdastro de Pdvs / Turnos"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmCadTrunos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   5760
      TabIndex        =   34
      Top             =   4560
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pdvs"
      TabPicture(0)   =   "frmCadTrunos.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Picture1"
      Tab(0).Control(3)=   "dbPdvs"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Turnos"
      TabPicture(1)   =   "frmCadTrunos.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picButtons"
      Tab(1).Control(1)=   "DataGrid1"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Adodc1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Pdvs / Turnos"
      TabPicture(2)   =   "frmCadTrunos.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblLabels(5)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtHoraIni"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboTurno"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cboPdv"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DataGrid3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdIncluir"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdRemover"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "dbPdvsTurnos"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin MSAdodcLib.Adodc dbPdvsTurnos 
         Height          =   375
         Left            =   840
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "select *from PdvsTurnos order by DescriPdv, HoraIni"
         Caption         =   "dbPdvsTurnos"
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
      Begin VB.CommandButton cmdRemover 
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
         Left            =   6120
         TabIndex        =   32
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdIncluir 
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
         Left            =   5520
         TabIndex        =   31
         Top             =   1080
         Width           =   495
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmCadTrunos.frx":0496
         Height          =   2415
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   6375
         _ExtentX        =   11245
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
            DataField       =   "DescriPdv"
            Caption         =   "Pdv"
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
            DataField       =   "DescriTurno"
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
         BeginProperty Column02 
            DataField       =   "HoraIni"
            Caption         =   "Hora Inicial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2550,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2099,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989,858
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboPdv 
         Bindings        =   "frmCadTrunos.frx":04B1
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmCadTrunos.frx":04C6
         Height          =   1695
         Left            =   -74760
         TabIndex        =   0
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3465,071
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74880
         TabIndex        =   39
         Top             =   2160
         Width           =   6375
         Begin VB.CheckBox Check1 
            Caption         =   "Permite ficar fechado enquanto outro PDV está em operação"
            DataField       =   "Intermitente"
            DataSource      =   "dbPdvs"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox txtFields 
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
            DataSource      =   "dbPdvs"
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   15
            TabIndex        =   2
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Descri"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "dbPdvs"
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   4
            Top             =   480
            Width           =   3375
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            DataField       =   "HoraIni"
            DataSource      =   "dbPdvs"
            Height          =   300
            Left            =   5280
            TabIndex        =   6
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "hh:mm"
            PromptChar      =   " "
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Index           =   4
            Left            =   1800
            TabIndex        =   3
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Início:"
            Height          =   195
            Index           =   3
            Left            =   5280
            TabIndex        =   5
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74880
         ScaleHeight     =   330
         ScaleWidth      =   6465
         TabIndex        =   38
         Top             =   3480
         Width           =   6465
         Begin VB.CommandButton cmdEditaPDV 
            Caption         =   "&Editar"
            Height          =   300
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdGravaPdv 
            Caption         =   "&Gravar"
            Height          =   300
            Left            =   4365
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdAtualizaPdv 
            Caption         =   "Atuali&zar"
            Height          =   300
            Left            =   3270
            TabIndex        =   11
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRemovePdv 
            Caption         =   "&Remover"
            Height          =   300
            Left            =   2175
            TabIndex        =   10
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdAdicionaPdv 
            Caption         =   "&Adicionar"
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox picButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74880
         ScaleHeight     =   330
         ScaleWidth      =   6465
         TabIndex        =   37
         Top             =   3420
         Width           =   6465
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Adicionar"
            Height          =   300
            Left            =   1080
            TabIndex        =   21
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remover"
            Height          =   300
            Left            =   2175
            TabIndex        =   22
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Atuali&zar"
            Height          =   300
            Left            =   3270
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Gravar"
            Height          =   300
            Left            =   4365
            TabIndex        =   24
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   300
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCadTrunos.frx":04DB
         Height          =   1935
         Left            =   -74760
         TabIndex        =   13
         Top             =   420
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3413
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
            DataField       =   "HoraIni"
            Caption         =   "Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "HoraFim"
            Caption         =   "Fim"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3539,906
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   929,764
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74880
         TabIndex        =   36
         Top             =   2340
         Width           =   6375
         Begin VB.TextBox txtFields 
            DataField       =   "Descri"
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
            TabIndex        =   15
            Top             =   480
            Width           =   3855
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            DataField       =   "HoraIni"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   4080
            TabIndex        =   17
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "hh:mm"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            DataField       =   "HoraFim"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   5160
            TabIndex        =   19
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "hh:mm"
            PromptChar      =   " "
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fim:"
            Height          =   195
            Index           =   7
            Left            =   5160
            TabIndex        =   18
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Início:"
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   765
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -74880
         Top             =   3780
         Width           =   6465
         _ExtentX        =   11404
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
         RecordSource    =   "select *from turnos order by horaini"
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
      Begin MSAdodcLib.Adodc dbPdvs 
         Height          =   330
         Left            =   -74880
         Top             =   3840
         Width           =   6465
         _ExtentX        =   11404
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
         RecordSource    =   "select *from pdvs order by Descri"
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
      Begin MSDataListLib.DataCombo cboTurno 
         Bindings        =   "frmCadTrunos.frx":04F0
         Height          =   315
         Left            =   2280
         TabIndex        =   28
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox txtHoraIni 
         Height          =   300
         Left            =   4440
         TabIndex        =   30
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "hh:mm"
         PromptChar      =   " "
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   29
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "PDV:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Somente será necessário cadastrar este item no caso de PDV que interrompa seu funcionamento enquanto outro PDV está funcionando."
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmCadTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String



Private Sub Adodc1_Reposition()

End Sub

Private Sub Adodc1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(0).SetFocus
End Sub

Private Sub cmdAdicionaPdv_Click()
  dbPDVs.Recordset.AddNew
  cmdAdicionaPdv.Enabled = False
  cmdRemovePdv.Enabled = False
  cmdAtualizaPdv.Enabled = False
  Frame2.Enabled = True
  txtFields(2).SetFocus
End Sub

Private Sub cmdAtualizaPdv_Click()
dbPDVs.Refresh
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

Private Sub cmdEditaPDV_Click()
If dbPDVs.Recordset.RecordCount = 0 Then Exit Sub
Frame2.Enabled = True
txtFields(2).SetFocus
End Sub

Private Sub cmdEditar_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Frame1.Enabled = True
txtFields(0).SetFocus
End Sub

Private Sub cmdGravaPdv_Click()
  On Error Resume Next
  With dbPDVs
    A = .Recordset.AbsolutePosition
    .Recordset.Update
    .Recordset.AbsolutePosition = A
  End With
  
  cmdAdicionaPdv.Enabled = True
  cmdRemovePdv.Enabled = True
  cmdAtualizaPdv.Enabled = True
  Frame2.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
If cboPdv.Text = "" Then
  MsgBox "Selecione um Pdv!"
  cboPdv.SetFocus
  Exit Sub
End If
If cboTurno.Text = "" Then
  MsgBox "Selecione um Turno!"
  cboTurno.SetFocus
  Exit Sub
End If
If IsDate(txtHoraIni.Text) = False Then
  MsgBox "Selecione uma hora válida!"
  txtHoraIni.SetFocus
  Exit Sub
End If

With dbPDVs
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe Pdv Cadastrado!"
    Exit Sub
  End If
  .Recordset.MoveFirst
  .Recordset.Find "Descri='" & cboPdv.Text & "'"
  If .Recordset.EOF = True Then
    MsgBox "Pdv não encontrado!"
    cboPdv.SetFocus
    Exit Sub
  End If
End With

With Adodc1
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe Turno Cadastrado!"
    Exit Sub
  End If
  .Recordset.MoveFirst
  .Recordset.Find "Descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = True Then
    MsgBox "Turno não encontrado!"
    cboTurno.SetFocus
    Exit Sub
  End If
End With

With dbPdvsTurnos
  .Recordset.AddNew
  .Recordset!CodigoPdv = dbPDVs.Recordset!CodigoPdv
  .Recordset!descripdv = dbPDVs.Recordset!Descri
  .Recordset!CodigoTurno = Adodc1.Recordset!CodigoTurno
  .Recordset!descriturno = Adodc1.Recordset!Descri
  .Recordset!HoraIni = CDate(txtHoraIni.Text)
  .Recordset.Update
  .Refresh
  .Refresh
End With

cboPdv.Text = ""
cboTurno.Text = ""
txtHoraIni.Text = ""
cboPdv.SetFocus

End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdRemovePdv_Click()
Dim Resposta As Integer
  
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo, "Excluir!")
  If Resposta = vbNo Then
    Exit Sub
  End If
  
  With dbPDVs.Recordset
    If .EOF = False Then
      .Delete
      If .EOF = False Then
      .MoveNext
      Else
        If .BOF = False Then .MoveLast
      End If
    End If
  End With
  
  Frame2.Enabled = False
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja remover o registro atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbPdvsTurnos
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione um registro primeiro!"
    Exit Sub
  End If
  .Recordset.Delete adAffectCurrent
  .Refresh
  .Refresh
End With
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

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DataGrid1.Columns(ColIndex).DataField Then
  strOrdem = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = DataGrid1.Columns(ColIndex).DataField
End If
With Adodc1
  .RecordSource = "select *from turnos order by " & strOrdem
  .Refresh
End With
End Sub

Private Sub dbPdvs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
dbPDVs.Caption = "Registro: " & dbPDVs.Recordset.AbsolutePosition
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
SSTab1.Tab = 0
With Adodc1
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbPDVs
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbPdvsTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With
Select Case Usuarios.Grupo.CadTurnos
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


