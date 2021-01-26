VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEstatus2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Novo Estatus do Sistema"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12120
   Icon            =   "frmEstatus2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   16536
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Fechamentos"
      TabPicture(0)   =   "frmEstatus2.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "optFechaDia"
      Tab(0).Control(1)=   "optFechaMes"
      Tab(0).Control(2)=   "dbStatusDiario"
      Tab(0).Control(3)=   "DBGrid4"
      Tab(0).Control(4)=   "Label36"
      Tab(0).Control(5)=   "lblLucroAcumulado"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Estatus"
      TabPicture(1)   =   "frmEstatus2.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblZeradoEm"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblUltimo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblVerifica"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "adoContas"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdZerarTrimestral"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdValorEstoque"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "dbEstoque"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "dbProdutos"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dbStatus"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dbTemp"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "qEstoque"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "qLucroVendas"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdZerar"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmdImprimir"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdAtualiza"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdOk"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      Begin VB.OptionButton optFechaDia 
         Caption         =   "Fechamentos Diários"
         Height          =   255
         Left            =   -71640
         TabIndex        =   85
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optFechaMes 
         Caption         =   "Fechamentos Mensais"
         Height          =   255
         Left            =   -74520
         TabIndex        =   84
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Data dbStatusDiario 
         Caption         =   "dbStatusDiario"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -70800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from statusdiario"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   375
         Left            =   10920
         TabIndex        =   37
         Top             =   8880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAtualiza 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   8880
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   9840
         TabIndex        =   4
         Top             =   8880
         Width           =   855
      End
      Begin VB.CommandButton cmdZerar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fechamento Mensal"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   8880
         Width           =   1935
      End
      Begin VB.Data qLucroVendas 
         Caption         =   "qLucroVendas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"frmEstatus2.frx":047A
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data qEstoque 
         Caption         =   "qEstoque"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"frmEstatus2.frx":0585
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbTemp 
         Caption         =   "dbTemp"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select categoria, sum(estoque) as tEstoque , sum(estoque * precocompra) as total from produtos group by categoria"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbStatus 
         Caption         =   "dbStatus"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from status"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbProdutos 
         Caption         =   "dbProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtos"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbEstoque 
         Caption         =   "dbEstoque"
         Connect         =   "Access 2000;"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from QEstatusEstoque"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdValorEstoque 
         Caption         =   "Valor Estoque"
         Height          =   375
         Left            =   8400
         TabIndex        =   2
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CommandButton cmdZerarTrimestral 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fechamento Trimestral"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   8880
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc adoContas 
         Height          =   330
         Left            =   8280
         Top             =   3120
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
         RecordSource    =   "Select descri, saldo from Contas order by descri"
         Caption         =   "adoContas"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   8415
         Left            =   5760
         TabIndex        =   38
         Top             =   120
         Width           =   6255
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmEstatus2.frx":063B
            Height          =   1455
            Left            =   1920
            TabIndex        =   39
            Top             =   2040
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2566
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
            Caption         =   "Contas"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "descri"
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
            BeginProperty Column01 
               DataField       =   "saldo"
               Caption         =   "Saldo"
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
                  ColumnWidth     =   2115,213
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1305,071
               EndProperty
            EndProperty
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "frmEstatus2.frx":0653
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmEstatus2.frx":066A
            TabIndex        =   40
            Top             =   240
            Width           =   6015
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Capital Inicial:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   7680
            Width           =   1275
         End
         Begin VB.Label lblCapitalInicial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   7920
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Pagar:"
            Height          =   195
            Left            =   2400
            TabIndex        =   78
            Top             =   7440
            Width           =   1635
         End
         Begin VB.Label lblContasAPagar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   77
            Top             =   7440
            Width           =   1815
         End
         Begin VB.Label lblTotal2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   76
            Top             =   7035
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   75
            Top             =   7035
            Width           =   1155
         End
         Begin VB.Label lblNotasAReceber 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   74
            Top             =   5280
            Width           =   1815
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Notas Futuras:"
            Height          =   195
            Left            =   2400
            TabIndex        =   73
            Top             =   5280
            Width           =   1635
         End
         Begin VB.Label lblPrevisaoRecebe 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   72
            Top             =   3750
            Width           =   1815
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cartões Pendentes:"
            Height          =   195
            Left            =   2040
            TabIndex        =   71
            Top             =   3750
            Width           =   1995
         End
         Begin VB.Label lblEstatus 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   70
            Top             =   7920
            Width           =   1815
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Capital Atual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   69
            Top             =   7920
            Width           =   1635
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Notas em Cobrança:"
            Height          =   195
            Left            =   2400
            TabIndex        =   68
            Top             =   5535
            Width           =   1635
         End
         Begin VB.Label lblNotasCobra 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   67
            Top             =   5535
            Width           =   1815
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldos em Contas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2445
            TabIndex        =   66
            Top             =   3480
            Width           =   1590
         End
         Begin VB.Label lblCaixa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   65
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques Pré:"
            Height          =   195
            Left            =   2040
            TabIndex        =   64
            Top             =   4260
            Width           =   1995
         End
         Begin VB.Label lblChequesPre 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   63
            Top             =   4260
            Width           =   1815
         End
         Begin VB.Label lblChequesDevolvidos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   62
            Top             =   4770
            Width           =   1815
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques Devolvidos no Banco:"
            Height          =   195
            Left            =   1785
            TabIndex        =   61
            Top             =   4770
            Width           =   2250
         End
         Begin VB.Label lblChequesCobranca 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   60
            Top             =   5025
            Width           =   1815
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques em Cobrança:"
            Height          =   195
            Left            =   2040
            TabIndex        =   59
            Top             =   5025
            Width           =   1995
         End
         Begin VB.Label lblComissoes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   58
            Top             =   7680
            Width           =   1815
         End
         Begin VB.Label lblComissoesDescri 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comissões a Pagar:"
            Height          =   195
            Left            =   2400
            TabIndex        =   57
            Top             =   7680
            Width           =   1635
         End
         Begin VB.Label lblValorPendente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   56
            Top             =   6270
            Width           =   1815
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagamentos Antecipados:"
            Height          =   195
            Left            =   1800
            TabIndex        =   55
            Top             =   6270
            Width           =   2235
         End
         Begin VB.Label lblCustodia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   54
            Top             =   4515
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques em Custódia:"
            Height          =   195
            Left            =   2040
            TabIndex        =   53
            Top             =   4515
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa Confirmado - Distribuido:"
            Height          =   195
            Left            =   1800
            TabIndex        =   52
            Top             =   6525
            Width           =   2235
         End
         Begin VB.Label lblAcumulaCaixa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   51
            Top             =   6525
            Width           =   1815
         End
         Begin VB.Label lblVales 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   50
            Top             =   6780
            Width           =   1815
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vales / Diferença de Caixa:"
            Height          =   195
            Left            =   1800
            TabIndex        =   49
            Top             =   6780
            Width           =   2235
         End
         Begin VB.Label lblTotalEstoque 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4575
            TabIndex        =   48
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estoque Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   47
            Top             =   1800
            Width           =   1530
         End
         Begin VB.Label lblChequesEmFechamento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   46
            Top             =   4005
            Width           =   1815
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques dep. em Fechamento:"
            Height          =   195
            Left            =   1440
            TabIndex        =   45
            Top             =   4005
            Width           =   2595
         End
         Begin VB.Label lblContasAReceber 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   44
            Top             =   6015
            Width           =   1815
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber:"
            Height          =   195
            Left            =   2400
            TabIndex        =   43
            Top             =   6015
            Width           =   1635
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber Sub-Locação:"
            Height          =   195
            Left            =   600
            TabIndex        =   42
            Top             =   5760
            Width           =   3435
         End
         Begin VB.Label lblAluguel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   4200
            TabIndex        =   41
            Top             =   5760
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   7455
         Left            =   480
         TabIndex        =   6
         Top             =   120
         Width           =   5175
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmEstatus2.frx":1571
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmEstatus2.frx":158C
            TabIndex        =   7
            Top             =   240
            Width           =   4935
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frmEstatus2.frx":22DB
            Height          =   1335
            Left            =   120
            OleObjectBlob   =   "frmEstatus2.frx":22F6
            TabIndex        =   8
            Top             =   2160
            Width           =   4935
         End
         Begin VB.Label lblQuantidadeVendida 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   95
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblTotalVendido 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   94
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblArredondamento 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   93
            Top             =   6720
            Width           =   1575
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Arredondamento de Encerrantes:"
            Height          =   195
            Left            =   720
            TabIndex        =   92
            Top             =   6720
            Width           =   2505
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Locação Restituição:"
            Height          =   195
            Left            =   480
            TabIndex        =   91
            Top             =   6240
            Width           =   2790
         End
         Begin VB.Label lblSubLocacaoRestituicao 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   90
            Top             =   6240
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Despesas:"
            Height          =   195
            Left            =   2520
            TabIndex        =   36
            Top             =   5520
            Width           =   750
         End
         Begin VB.Label lblDespesa 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   35
            Top             =   5520
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lucro Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   34
            Top             =   7080
            Width           =   1110
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   33
            Top             =   7080
            Width           =   1575
         End
         Begin VB.Label lblTaxasPg 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   32
            Top             =   4800
            Width           =   1575
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Taxas em Forma de Pg:"
            Height          =   195
            Left            =   1590
            TabIndex        =   31
            Top             =   4800
            Width           =   1680
         End
         Begin VB.Label lblDifCaixa 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   30
            Top             =   4560
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença de Caixa:"
            Height          =   195
            Left            =   1875
            TabIndex        =   29
            Top             =   4560
            Width           =   1395
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Dif. Recebimentos:"
            Height          =   195
            Left            =   1920
            TabIndex        =   28
            Top             =   5040
            Width           =   1350
         End
         Begin VB.Label lblDifRecebido 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   27
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Label lblDifClientes 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   26
            Top             =   5280
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Dif. Rec. Clientes Cobrança:"
            Height          =   195
            Left            =   1995
            TabIndex        =   25
            Top             =   5280
            Width           =   1275
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Juros de Cheques:"
            Height          =   195
            Left            =   1530
            TabIndex        =   24
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label lblJuros 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   23
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Acerto de Estoque:"
            Height          =   195
            Left            =   1905
            TabIndex        =   22
            Top             =   4320
            Width           =   1365
         End
         Begin VB.Label lblAcertoEstoque 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """ ""#.##0,00;("" ""#.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label lblEstacionamento 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   20
            Top             =   5760
            Width           =   1575
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Estacionamento:"
            Height          =   195
            Left            =   2085
            TabIndex        =   19
            Top             =   5760
            Width           =   1185
         End
         Begin VB.Label lblPrecoDiferenciado 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   18
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Preço Diferenciado de Clientes:"
            Height          =   195
            Left            =   930
            TabIndex        =   17
            Top             =   3840
            Width           =   2340
         End
         Begin VB.Label lblTotalLucro 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3495
            TabIndex        =   16
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber:"
            Height          =   195
            Left            =   765
            TabIndex        =   15
            Top             =   6480
            Width           =   2505
         End
         Begin VB.Label lblOutrosFaturamentos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   14
            Top             =   6480
            Width           =   1575
         End
         Begin VB.Label lblSubLocacao 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   3360
            TabIndex        =   13
            Top             =   6000
            Width           =   1575
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber Sub-Locação:"
            Height          =   195
            Left            =   480
            TabIndex        =   12
            Top             =   6000
            Width           =   2790
         End
         Begin VB.Label lblVariacao 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Diferença:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   3480
            Width           =   1500
         End
         Begin VB.Label lblDiferenca 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   9
            Top             =   3480
            Width           =   1335
         End
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "frmEstatus2.frx":2EA9
         Height          =   2295
         Left            =   -74520
         OleObjectBlob   =   "frmEstatus2.frx":2EC6
         TabIndex        =   83
         Top             =   600
         Width           =   8295
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lucro Acumulado: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71040
         TabIndex        =   87
         Top             =   3000
         Width           =   1635
      End
      Begin VB.Label lblLucroAcumulado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69360
         TabIndex        =   86
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   82
         Top             =   8280
         Width           =   900
      End
      Begin VB.Label lblVerifica 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   81
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label lblUltimo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O último caixa finalizado foi:"
         Height          =   615
         Left            =   480
         TabIndex        =   80
         Top             =   7680
         Width           =   5175
      End
      Begin VB.Label lblZeradoEm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   79
         Top             =   8280
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmEstatus2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataCaixa As Date, CodigoTurno As Double, Turno As String
Dim ExibeStatusMensal As Boolean

Private Sub AtualizaFechamentos()
Dim Total As Currency
With dbStatusDiario
  If ExibeStatusMensal = True Then
    .RecordSource = "select *from statusdiario where ExibeStatusMensal=-1 and fechamentotrimestral=0 order by datacaixa"
  Else
    .RecordSource = "select *from statusdiario where ExibeStatusMensal=0 and fechamentotrimestral=0 order by datacaixa"
  End If
  Total = 0
  On Error Resume Next
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Total = Total + .Recordset!lucroacumulado
      .Recordset.MoveNext
      If Err.Number <> 0 Then Exit Do
    Loop
  End If
  On Error GoTo 0
  lblLucroAcumulado.Caption = Format(Total, "Currency")
End With
End Sub

Private Sub Atualizar()
Dim Total As Currency, TempValor As Currency, Estoque As Double, ValorEstoque As Currency
Dim Total2 As Currency, Lucro As Currency, Diferenca As Currency, ValorDiferenca As Currency
Dim StrBloqueia As String, TotalVendido As Currency, QuantidadeVendida As Double

Screen.MousePointer = vbHourglass

Lucro = 0
With qLucroVendas
  .RecordSource = "select categoria, sum(qtdcomprado) as comprado, sum(valorcomprado) as VComprado, sum(DifEstoque) as DiferencaEstoque, sum(ValorDifEstoque) as Varia, sum(lucroMedio) as Lucro, sum(Acumulativo) as Vendido, sum(totalvendido) as total from produtos group by categoria"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    TotalVendido = 0
    QuantidadeVendida = 0
    Do While .Recordset.EOF = False
      If IsNull(.Recordset!varia) = False Then
        Total = Total + .Recordset!varia
        Diferenca = Diferenca + .Recordset!diferencaestoque
        ValorDiferenca = ValorDiferenca + .Recordset!varia
      End If
      
      If IsNull(.Recordset!Total) = False Then TotalVendido = TotalVendido + .Recordset!Total
      If IsNull(.Recordset!Vendido) = False Then QuantidadeVendida = QuantidadeVendida + .Recordset!Vendido
      
      If IsNull(.Recordset!Lucro) = False Then
        Total = Total + .Recordset!Lucro
        Lucro = Lucro + .Recordset!Lucro
      End If
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

lblDiferenca.Caption = Format(Diferenca, "#,##0.00")
lblVariacao.Caption = Format(ValorDiferenca, "Currency")
lblTotalVendido.Caption = Format(TotalVendido, "Currency")
lblQuantidadeVendida.Caption = Format(QuantidadeVendida, "#,##0.0")

lblTotalLucro.Caption = Format(Lucro, "Currency")

With dbStatus
  .RecordSource = "select *from status"
  .Refresh
  If IsNumeric(.Recordset!Juros) = True Then
    lblJuros.Caption = Format(.Recordset!Juros, "currency")
  Else
    lblJuros.Caption = Format(0, "currency")
  End If
  If IsNumeric(.Recordset!clientediferenciado) = True Then
    lblPrecoDiferenciado.Caption = Format(.Recordset!clientediferenciado, "Currency")
  Else
    lblPrecoDiferenciado.Caption = Format(0, "Currency")
  End If
  If IsNumeric(.Recordset!Arredondamento) = True Then
    lblArredondamento.Caption = Format(-.Recordset!Arredondamento, "Currency")
  Else
    lblArredondamento.Caption = Format(0, "Currency")
  End If
  
End With

Total = Total + CCur(lblJuros.Caption)
Total = Total + CCur(lblPrecoDiferenciado.Caption)
Total = Total + CCur(lblArredondamento.Caption)

With dbStatus
  .RecordSource = "select *from status"
  .Refresh
  If IsNumeric(.Recordset!Estacionamento) = True Then
    lblEstacionamento.Caption = Format(.Recordset!Estacionamento, "currency")
  Else
    lblEstacionamento.Caption = Format(0, "currency")
  End If
End With

Total = Total + CCur(lblEstacionamento.Caption)

On Error Resume Next
With dbTemp
  .RecordSource = "select sum(clientescobrancacomposicao.valor) as total from qclientescobrancacomposicao where origem = 'Aluguel' and fechaaluguel=0 and reembolso=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblSubLocacao.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblSubLocacao.Caption = Format(0, "Currency")
  End If
End With
Total = Total + CCur(lblSubLocacao.Caption)

With dbTemp
  .RecordSource = "select sum(clientescobrancacomposicao.valor) as total from qclientescobrancacomposicao where origem = 'Aluguel' and fechaaluguel=0 and reembolso=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblSubLocacaoRestituicao.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblSubLocacaoRestituicao.Caption = Format(0, "Currency")
  End If
End With

Total = Total + CCur(lblSubLocacaoRestituicao.Caption)
On Error GoTo 0

With dbTemp
  .RecordSource = "select sum(valor) as total from clientescobranca where origem = 'Outros' and fechaaluguel=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblOutrosFaturamentos.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblOutrosFaturamentos.Caption = Format(0, "Currency")
  End If
End With

Total = Total + CCur(lblOutrosFaturamentos.Caption)


With dbEstoque
  .RecordSource = "select sum(valor) as total from despesaslanc2 where fechamento=0 and fechamentodiario=-1 and produto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    Total = Total + .Recordset!Total
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblDespesa.Caption = Format(TempValor, "Currency")
  
  .RecordSource = "select sum(juros) as total from clientescobranca where fechames=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  .RecordSource = "select *from status"
  .Refresh
  If IsNull(.Recordset!difcheques) = False Then
    TempValor = TempValor + .Recordset!difcheques
  Else
    TempValor = 0
  End If
  Total = Total + TempValor
  lblDifClientes.Caption = Format(TempValor, "Currency")
  
End With

'Acerto de Estoque
If IsNumeric(dbStatus.Recordset!acertoestoque) = True Then
  lblAcertoEstoque.Caption = Format(dbStatus.Recordset!acertoestoque, "Currency")
  TempValor = CCur(lblAcertoEstoque.Caption)
  Total = Total + TempValor
Else
  lblAcertoEstoque.Caption = Format(0, "Currency")
End If

'taxas de formas de recebimento
With dbTemp
  .RecordSource = "select  sum(valor - valorbruto) as total from formadepagamentorecebido2 where fechames=0 and fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblTaxasPg.Caption = Format(TempValor, "Currency")
  Total = Total + TempValor
  
  .RecordSource = "select sum(diferenca) as total from cartoes where confirmado=-1 and fechadiferenca=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblDifRecebido.Caption = Format(TempValor, "Currency")
  Total = Total + TempValor

End With



'Diferença de caixa
With dbTemp
  .RecordSource = "select sum(diferenca) as total from fechamentodecaixa where fechado=-1 and distribuido=-1 and fechames=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblDifCaixa.Caption = Format(TempValor, "Currency")
  Total = Total + TempValor
End With

lblTotal.Caption = Format(Total, "Currency")

'Aqui começa a parte 2
With dbTemp
  With qEstoque
    .RecordSource = "Select categoria, sum(qtdcomprado) as comprado, sum(valorcomprado) as VComprado, sum(estoque) as tEstoque , sum(valorestoque) as total from produtos group by categoria"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      TempValor = 0
      Do While .Recordset.EOF = False
        If IsNull(.Recordset!Total) = False Then
          TempValor = TempValor + .Recordset!Total
        End If
        .Recordset.MoveNext
      Loop
      .Recordset.MoveFirst
    End If
  End With
  lblTotalEstoque.Caption = Format(TempValor, "Currency")
  
  Total2 = Total2 + TempValor
  

  'Saldo em caixa
  With adoContas
    TempValor = 0
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      Do While .Recordset.BOF = False
        TempValor = TempValor + .Recordset!Saldo
        .Recordset.MovePrevious
      Loop
      .Recordset.MoveFirst
    End If
    lblCaixa.Caption = Format(TempValor, "Currency")
    Total2 = Total2 + TempValor
  End With
  
  'Cartões Pendentes
  .RecordSource = "select sum (valorliquido) as Liquido from cartoes where cartoes.confirmado=0"
  .Refresh
  If IsNull(.Recordset("liquido")) = False Then
    TempValor = .Recordset("liquido")
  Else
    TempValor = 0
  End If
  lblPrevisaoRecebe.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Cheques em Fechamento
  .RecordSource = "select sum(valor) as total from cheques where compensado=-1 and custodia=0 and fechamentodiario=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = -.Recordset!Total
  Else
    TempValor = 0
  End If
  .RecordSource = "select sum(valor) as total from cheques where custodia=-1 and fechamentodiario=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  lblChequesEmFechamento.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Cheques pre
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and custodia=0 and cobrando=0 and devolvido=0 and protesto=0 and fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblChequesPre.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Cheques em custódia
  .RecordSource = "select sum(valor) as total from compensapendente where conciliado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblCustodia.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Cheques devolvidos
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0 and fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblChequesDevolvidos.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Cheques cobrando
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblChequesCobranca.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Notas de clientes2
  .RecordSource = "select sum (valorprevisto) as Liquido from clientesnota2 where confirmado=0 and fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset("liquido")) = False Then
    TempValor = .Recordset("liquido")
  Else
    TempValor = 0
  End If
  lblNotasAReceber.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  .RecordSource = " select sum (valor) as total from clientescobranca where pago=0 and origem='Fiado'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblNotasCobra.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  .RecordSource = " select sum (valor) as total from clientescobranca where pago=0 and origem='Aluguel'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblAluguel.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  .RecordSource = " select sum (valor) as total from clientescobranca where pago=0 and origem='Outros'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  lblContasAReceber.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  .RecordSource = "select sum (valortotal) as total from Pedidos where Recebido=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  
  lblValorPendente.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  .RecordSource = "select sum (valor) as total from vales where fechado=-1 and cobrado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  
'  .RecordSource = "select sum (vales) as total from vendedorespagamento where pago=-1 and confirmadonocaixa=0"
'  .Refresh
'  If IsNull(.Recordset!Total) = False Then
'    TempValor = TempValor - .Recordset!Total
'  End If
  
  lblVales.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  'Valor a distribuido sem confirmação da primeira parte
  .RecordSource = "select sum(totalcombustivel+totalprodutos) as total from fechamentodecaixa where fechado=-1 and distribuido=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = .Recordset!Total
  Else
    TempValor = 0
  End If
  
  'Valor distribuido sem caixa confirmado
  .RecordSource = "select sum(totalcombustivel+totalprodutos) as total from fechamentodecaixa where fechado=0 and distribuido=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  
  'Pagamentos já distribuido
  .RecordSource = "select sum(totalrecebido-totaldespesas) as total from fechamentodecaixa where distribuido=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  
  'pagamentos de funcionários já distribuidos
   .RecordSource = "select sum(saldoapagar) as total from vendedorespagamento where pago=-1 and confirmadonocaixa=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  
  'Juros já distribuido
  .RecordSource = "select  sum(valordesconto) as total from qformadepgrecebidofechamento2 where formadepagamentorecebido2.fechames=0 and distribuido=0 and fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  'Vales já distribuido
  .RecordSource = "select  fechamentodecaixa.*, vales.* from fechamentodecaixa, vales where fechamentodecaixa.codigofechamento=vales.codigocaixa and fechamentodecaixa.distribuido=0 and vales.fechado=-1 and vales.cobrado=0"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      TempValor = TempValor - .Recordset!Valor
      .Recordset.MoveNext
    Loop
  End If
  
  'Vales já distribuido
  .RecordSource = "select  fechamentodecaixa.*, vales.* from fechamentodecaixa, vales where fechamentodecaixa.codigofechamento=vales.codigocaixa and fechamentodecaixa.distribuido=0 and vales.fechado=-1 and vales.cobrado=-1"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      TempValor = TempValor - .Recordset!Valor
      .Recordset.MoveNext
    Loop
  End If
  On Error Resume Next
  'Comissões já pagas
  If ComissaoAcumulativa = True Then
    .RecordSource = "select sum(valorcomissao) as total from  (Venda2 left JOIN Fechamentodecaixa ON Venda2.CodigoFechamento = Fechamentodecaixa.CodigoFechamento) where fechado=-1 and distribuido=0 and pago=-1 and comissaoacumulativa=-1"
    .Refresh
    If IsNull(.Recordset!Total) = False Then
      TempValor = TempValor + .Recordset!Total
    End If
  End If
  
  'Clientes de Nota
  .RecordSource = "select sum(valorprevisto) as total from clientesnota2 where confirmado=-1 and fechamentodiario=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  
  'Comissões já pagas
  .RecordSource = "select sum(valorcomissao) as total from qvendas where pago=-1 and fechado=-1 and distribuido=0 and fechames=0 and comissaoacumulativa=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = TempValor - .Recordset!Total
  End If
  
  
  lblAcumulaCaixa.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  
  lblTotal2.Caption = Format(Total2, "Currency")
  
  
  
  .RecordSource = "select sum(valor) as total, sum(valorpago) as pago from despesaslanc2 where despesaslanc2.origem='Despesa' and compensado=0"
  .Refresh
  If IsNull(.Recordset("total")) = False Then
    TempValor = .Recordset("total") - .Recordset!Pago
  Else
    TempValor = 0
  End If
  lblContasAPagar.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
'  If ComissaoAcumulativa = True Then
    .RecordSource = "select sum(valorcomissao) as total from venda2 where venda2.pago=0 and venda2.fechamentodiario=-1"
'  Else
'    .RecordSource = "select sum(valorcomissao) as total from qvendas where venda2.pago=-1 and venda2.fechamentodiario=-1 and fechames=0"
'  End If
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    TempValor = -.Recordset!Total
  Else
    TempValor = 0
  End If
  
  lblComissoes.Caption = Format(TempValor, "Currency")
  Total2 = Total2 + TempValor
  
  lblEstatus.Caption = Format(Total2, "Currency")
  
  .RecordSource = "Select *from fechamentodecaixa where fechado=-1 order by datacaixa desc,horaini desc"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    lblUltimo.Caption = "Não existe caixa finalizado"
  Else
    lblUltimo.Caption = "O último caixa finalizado foi: " & .Recordset!DataCaixa & " - " & .Recordset!Turno
    DataCaixa = .Recordset!DataCaixa
    CodigoTurno = .Recordset!CodigoTurno
    Turno = .Recordset!Turno
  End If
  
  .RecordSource = "Select *from fechamentodecaixa where fechado=0 order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    lblUltimo.Caption = lblUltimo.Caption & " / Não existe caixa digitado sem finalizar"
  Else
    lblUltimo.Caption = lblUltimo.Caption & " / O último caixa digitado sem finalizar é: " & .Recordset!DataCaixa & " - " & .Recordset!Turno
  End If
  
  dbStatus.Connect = Conectar
  dbStatus.DatabaseName = Caminho
  dbStatus.Refresh
  On Error Resume Next
  If IsNull(dbStatus.Recordset!CapitalInicial) = False Then
    TempValor = CCur(lblEstatus.Caption) - (CCur(lblTotal.Caption) + dbStatus.Recordset!CapitalInicial)
  Else
    TempValor = CCur(lblTotal.Caption) - CCur(lblEstatus.Caption)
  End If
  lblVerifica.Caption = Format(TempValor, "Currency")
  On Error GoTo 0
  With dbStatus
    .RecordSource = "select *from status"
    .Refresh
    On Error Resume Next
    If IsNull(.Recordset!CapitalInicial) = True Then
      .Recordset.Edit
      .Recordset!CapitalInicial = CCur(lblEstatus.Caption)
      .Recordset.Update
    End If
    On Error GoTo 0
  End With
  
  On Error Resume Next
  .RecordSource = "select *from statusdiario"
  .Refresh
  If Err.Number = 0 Then
    On Error GoTo 0
    If .Recordset.RecordCount <> 0 Then
      .Recordset.FindFirst "dataCaixa=#" & DataInglesa(DataCaixa) & "# and exibestatusmensal=0"
      If .Recordset.NoMatch = True Then
        .Recordset.AddNew
      Else
        .Recordset.Edit
      End If
    Else
      .Recordset.AddNew
    End If
    .Recordset!DataLanc = Date
    .Recordset!DataCaixa = DataCaixa
    .Recordset!Turno = Turno
    .Recordset!CodigoTurno = CodigoTurno
    .Recordset!capitaldodia = CCur(lblEstatus.Caption)
    .Recordset!lucroacumulado = CCur(lblTotal.Caption)
    On Error Resume Next
    Diferenca = .Recordset!capitaldodia - (dbStatus.Recordset!CapitalInicial + .Recordset!lucroacumulado)
    On Error GoTo 0
    .Recordset!Diferenca = Diferenca
    .Recordset!ExibeStatusMensal = False
    .Recordset.Update
  End If
  On Error GoTo 0
End With

lblVerifica.Caption = Format(Diferenca, "Currency")

On Error Resume Next
dbStatusDiario.Refresh
On Error GoTo 0

StrBloqueia = ReadINI("Estatus", "Erro", "", App.Path & "\Posto.ini")
If StrBloqueia = "" Then
  Open App.Path & "\Posto.ini" For Append As #1
  Print #1, ""
  Print #1, ";1-Bloqueia ou 2-Libera "
  Print #1, "[Estatus]"
  Print #1, "Erro=1"
  Print #1, ""
  Print #1, ""
  Close #1
  StrBloqueia = "1"
End If

With mdiPosto.StatusBar1
  .Panels(6).Text = "Dif. Status: " & lblVerifica.Caption
  .Refresh
  If Diferenca > 5 Or Diferenca < -5 Then
    If Usuarios.Nome <> "Usuário Master" Then
      MsgBox "Sistema com erro no status! Avisar o suporte técnico!"
    End If
    If Usuarios.Nome <> "" Then
      If Usuarios.Grupo.AdmEstatus <> 2 Then
        StrBloqueia = ReadINI("Estatus", "Erro", "", App.Path & "\Posto.ini")
        If StrBloqueia = "" Then
          Open App.Path & "\Posto.ini" For Append As #1
          Print #1, ""
          Print #1, ";1-Bloqueia ou 2-Libera "
          Print #1, "[Estatus]"
          Print #1, "Erro=1"
          Print #1, ""
          Print #1, ""
          Close #1
          StrBloqueia = "1"
        End If
        If StrBloqueia = "1" Then
          End
        End If
      End If
    End If
  End If
End With

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdAtualiza_Click()
Atualizar
End Sub

Private Sub cmdImprimir_Click()
Dim Largura As Double, StrTemp As String
Dim X1 As Double, X2 As Double
Dim Y1 As Double, Y2 As Double
Dim A1 As Double, A2 As Double
Dim B1 As Double, B2 As Double

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0


Printer.ScaleMode = vbMillimeters

Largura = 190
StrTemp = "Estatus do Sistema de Posto de Combustível"
Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = Format(Now, "Long Date")
Printer.FontSize = 8
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
X1 = 0
Y1 = Printer.CurrentY
Printer.Line (0, Y1)-(Largura, Y1)
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = True
StrTemp = "Lucro de Vendas"
Printer.CurrentX = 45 - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
B1 = Printer.CurrentY
Printer.Line (2, Printer.CurrentY)-(88, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Descrição"
Printer.CurrentX = 3
Printer.Print StrTemp;

StrTemp = "Quantidade"
Printer.CurrentX = 44 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Vendido"
Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Lucro"
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (2, Printer.CurrentY)-(88, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = False

With qLucroVendas
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = ""
      If IsNull(.Recordset!Categoria) = False Then
        StrTemp = .Recordset!Categoria
      End If
      Printer.CurrentX = 3
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Vendido, "#,###.0")
      Printer.CurrentX = 44 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Total, "#,###.0")
      Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Lucro, "Currency")
      Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    Printer.CurrentY = Printer.CurrentY + 1
    B2 = Printer.CurrentY
    Printer.Line (2, B2)-(88, B2)
    Printer.Line (2, B1)-(2, B2)
    Printer.Line (26, B1)-(26, B2)
    Printer.Line (45, B1)-(45, B2)
    Printer.Line (65, B1)-(65, B2)
    Printer.Line (88, B1)-(88, B2)
    Printer.CurrentY = Printer.CurrentY + 1
    
    StrTemp = "Lucro Total: " & Format(lblTotalLucro.Caption, "Currency")
    Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "Variação de Estoque"
    Printer.CurrentX = 45 - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    B1 = Printer.CurrentY
    Printer.Line (2, Printer.CurrentY)-(88, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    StrTemp = "Descrição"
    Printer.CurrentX = 3
    Printer.Print StrTemp;
    
    StrTemp = "Qtd"
    Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = "Valor"
    Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (2, Printer.CurrentY)-(88, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    Printer.FontBold = False

    
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = ""
      If IsNull(.Recordset!Categoria) = False Then
        StrTemp = .Recordset!Categoria
      End If
      Printer.CurrentX = 3
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!diferencaestoque, "#,##0.00")
      Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!varia, "Currency")
      Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    Printer.CurrentY = Printer.CurrentY + 1
    B2 = Printer.CurrentY
    Printer.Line (2, B2)-(88, B2)
    Printer.Line (2, B1)-(2, B2)
    Printer.Line (44, B1)-(44, B2)
    Printer.Line (65, B1)-(65, B2)
    Printer.Line (88, B1)-(88, B2)
    Printer.CurrentY = Printer.CurrentY + 1
  End If
  StrTemp = Format(lblDiferenca.Caption, "#,##0.00")
  Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(lblVariacao.Caption, "Currency")
  Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
End With

StrTemp = "Preço Diferenciado de Clientes:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblPrecoDiferenciado.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Juros de Cheques:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblJuros.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Acerto de Estoque:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblAcertoEstoque.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Diferença de Caixa:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblDifCaixa.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Taxas em Forma de Pg.:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblTaxasPg.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Dif. Recebimentos:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblDifRecebido.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Dif. Rec. Clientes: Cobrança"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblDifClientes.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Despesas:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblDespesa.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Estacionamento:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblEstacionamento.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Contas a Receber Sub-Locação:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblSubLocacao.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Sub-Locação Restituição:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblSubLocacaoRestituicao.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Contas a Receber:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblOutrosFaturamentos.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Arredondamento de Encerrantes:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblArredondamento.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.FontBold = True
StrTemp = "Lucro Total:"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblTotal.Caption
Printer.CurrentX = 87 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.CurrentY = Printer.CurrentY + 1
Y2 = Printer.CurrentY

Printer.CurrentY = Y1
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = True
StrTemp = "Compra"
Printer.CurrentX = ((Largura - 92) / 2) + 92 - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp;

Printer.FontBold = True
StrTemp = "Estoque"
Printer.CurrentX = ((Largura - 152) / 2) + 152 - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
B1 = Printer.CurrentY
Printer.Line (92, Printer.CurrentY)-(Largura - 2, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Descrição"
Printer.CurrentX = 93
Printer.Print StrTemp;

StrTemp = "Compr."
Printer.CurrentX = 136 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Compr."
Printer.CurrentX = 152 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Qtd."
Printer.CurrentX = 167 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total"
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (92, Printer.CurrentY)-(Largura - 2, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = False

With qEstoque
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = ""
      If IsNull(.Recordset!Categoria) = False Then
        StrTemp = .Recordset!Categoria
      End If
      Printer.CurrentX = 93
      Printer.Print StrTemp;
      
      StrTemp = ""
      If IsNull(.Recordset!Comprado) = False Then
        StrTemp = Format(.Recordset!Comprado, "#,###")
      End If
      Printer.CurrentX = 136 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = ""
      If IsNull(.Recordset!vcomprado) = False Then
        StrTemp = Format(.Recordset!vcomprado, "#,###.00")
      End If
      Printer.CurrentX = 152 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!tEstoque, "#,###.0")
      Printer.CurrentX = 167 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!Total, "Currency")
      Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    Printer.CurrentY = Printer.CurrentY + 1
    B2 = Printer.CurrentY
    Printer.Line (92, B2)-(Largura - 2, B2)
    Printer.Line (92, B1)-(92, B2)
    Printer.Line (125, B1)-(125, B2)
    Printer.Line (137, B1)-(137, B2)
    Printer.Line (153, B1)-(153, B2)
    Printer.Line (168, B1)-(168, B2)
    Printer.Line (Largura - 2, B1)-(Largura - 2, B2)
    Printer.CurrentY = Printer.CurrentY + 1
   End If
End With

StrTemp = lblTotalEstoque.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Saldos em Contas"
Printer.CurrentX = 93
Printer.Print StrTemp

With adoContas
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = .Recordset!Descri
      Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      StrTemp = Format(.Recordset!Saldo, "Currency")
      Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With

Printer.FontBold = True
StrTemp = "Total em Contas:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblCaixa.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.Print ""

StrTemp = "Cartões Pendentes:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblPrevisaoRecebe.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Cheques dep. em Fechamento:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblChequesEmFechamento.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Cheques Pré:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblChequesPre.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Cheques em Custódia:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblCustodia.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Cheques Devolvidos no Banco:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblChequesDevolvidos.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Cheques em Cobrança:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblChequesCobranca.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Notas Futuras:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblNotasAReceber.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Notas em Cobrança:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblNotasCobra.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Contas a Receber Sub-Locação:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblAluguel.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Contas a Receber:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblContasAReceber.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Pagamentos Antecipados:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblValorPendente.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Caixa Confirmado - Distribuido:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblAcumulaCaixa.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Vales / Diferença de Caixa:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblVales.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.FontBold = True
StrTemp = "Total:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblTotal2.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.Print ""

StrTemp = "Contas a Pagar:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblContasAPagar.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = lblComissoesDescri.Caption
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblComissoes.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.FontBold = True
StrTemp = "Capital Atual:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblEstatus.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.CurrentY = Printer.CurrentY + 1
If Printer.CurrentY > Y2 Then
  Y2 = Printer.CurrentY
End If

Printer.Line (0.5, Y1)-(0.5, Y2)
Printer.Line (90, Y1)-(90, Y2)
Printer.Line (Largura, Y1)-(Largura, Y2)
Printer.Line (0.5, Y2)-(Largura, Y2)

Printer.CurrentY = Printer.CurrentY + 1

If Turno = "" Then
  StrTemp = "Não existe caixa finalizado"
Else
  StrTemp = "O último caixa finalizado foi: " & DataCaixa & " - " & Turno
End If
Printer.FontBold = True
Printer.CurrentX = 0
Printer.Print StrTemp;


Printer.FontBold = True
StrTemp = "Diferença de Estatus:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblVerifica.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False

Printer.FontBold = True
StrTemp = "Capital Inicial:"
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;
StrTemp = lblCapitalInicial.Caption
Printer.CurrentX = Largura - 3 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp
Printer.FontBold = False



Resposta = MsgBox("Deseja imprimir a lista de fechamentos anteriores?", vbYesNo)
If Resposta = vbYes Then
  optFechaMes.Value = True
  Call optFechaMes_Click
  ImprimeGrid DBGrid4, Printer, dbStatusDiario, 4
End If

Printer.EndDoc

Exit Sub
NaoImprime:
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdValorEstoque_Click()
Dim db As New ADODB.Connection

db.Open CaminhoADO
db.Execute "update produtos set valorestoque=(estoque*precocompra)-variacao"
db.Execute "update produtos set precomedio=precocompra"

'Db.Execute "update produtos set estoque=estoque-difestoque"
'Db.Execute "update produtos set valorestoque=valorestoque-valordifestoque"
'Db.Execute "update produtos set difestoque=0"
'Db.Execute "update produtos set valordifestoque=0"

db.Close

Call cmdAtualiza_Click

End Sub

Private Sub cmdZerar_Click()
Dim Resposta As Integer, ZeradoEm As String, DataFechamento As Date
Dim CapitalInicial As Currency

With dbTemp
  .RecordSource = "select *from fechamentodecaixa where fechames=0 and fechado=-1 and distribuido=0"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    MsgBox "Existe caixa finalizado a primeira parte mas não finalizado a segunda parte! Não será zerado o Status!"
    Exit Sub
  End If
  .RecordSource = "select *from fechamentodecaixa where fechames=0 and fechado=0 and distribuido=-1"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    MsgBox "Existe caixa finalizado a segunda parte mas não finalizado a primeira parte! Não será zerado o Status!"
    Exit Sub
  End If
End With
Resposta = MsgBox("Deseja zerar o status?", vbYesNo + vbDefaultButton2)
If Resposta = vbYes Then
  Screen.MousePointer = vbHourglass
  
  
  With dbTemp
    On Error Resume Next
    .RecordSource = "select *from statusdiario where exibestatusmensal=-1"
    .Refresh
    If Err.Number = 0 Then
      On Error GoTo 0
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "dataCaixa=#" & DataInglesa(DataCaixa) & "# and exibestatusmensal=-1"
        If .Recordset.NoMatch = True Then
          .Recordset.AddNew
        Else
          If Usuarios.Nome <> "Usuário Master" Then
            MsgBox "Não foi possível zerar pois já existe um registro de fechamento!"
            Exit Sub
          Else
            Resposta = MsgBox("Já foi zerado o mês! Deseja zerar novamente?", vbYesNo)
            If Resposta = vbNo Then Exit Sub
          End If
        End If
      Else
        .Recordset.AddNew
      End If
      If .Recordset.EditMode = dbEditInProgress Then
        .Recordset!DataLanc = Date
        .Recordset!DataCaixa = DataCaixa
        .Recordset!Turno = Turno
        .Recordset!CodigoTurno = CodigoTurno
        .Recordset!capitaldodia = CCur(lblEstatus.Caption)
        .Recordset!lucroacumulado = CCur(lblTotal.Caption)
        .Recordset!Diferenca = .Recordset!capitaldodia - (dbStatus.Recordset!CapitalInicial + .Recordset!lucroacumulado)
        .Recordset!ExibeStatusMensal = True
        .Recordset!fechamentoMensal = True
        .Recordset!DataFechamento = DataCaixa
        .Recordset.Update
      End If
      With dbStatus
        .Connect = Conectar
        .DatabaseName = Caminho
        .RecordSource = "select *from status"
        .Refresh
        .Recordset.Edit
        .Recordset!CapitalInicial = CCur(lblEstatus.Caption)
        .Recordset.Update
      End With
    End If
    .RecordSource = "select *from statusdiario where fechamentomensal=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechamentoMensal = True
        .Recordset!DataFechamento = DataCaixa
        .Recordset.MoveNext
      Loop
    End If
    On Error GoTo 0
    
    .RecordSource = "Select *from fechamentodecaixa where fechado=-1 order by datacaixa desc,horaini desc"
    .Refresh
    If .Recordset.RecordCount = 0 Then
      ZeradoEm = "Fechamento em " & Format(Date, "long date") & " sem caixa finalizado"
      DataFechamento = Date
    Else
      ZeradoEm = "Último Fechamento até o caixa: " & Chr(13) & Format(.Recordset!DataCaixa, "short date") & " - " & .Recordset!Turno
      DataFechamento = .Recordset!DataCaixa
    End If
  End With
  With dbTemp
    .RecordSource = "select *from produtos"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!LucroVenda = 0
        .Recordset!Variacao = 0
        .Recordset!acumulativo = 0
        .Recordset!TotalVendido = 0
        .Recordset!qtdcomprado = 0
        .Recordset!valorcomprado = 0
        .Recordset!DifEstoque = 0
        .Recordset!valordifestoque = 0
        .Recordset!LucroMedio = 0
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from status"
    .Refresh
    CapitalInicial = .Recordset!CapitalInicial
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Edit
      For i = 0 To .Recordset.Fields.Count - 1
        .Recordset(i) = 0
      Next i
      .Recordset!difcheques = 0
      lblZeradoEm.Caption = ZeradoEm
      .Recordset!ZeradoEm = ZeradoEm
      .Recordset!CapitalInicial = CapitalInicial
      .Recordset.Update
    End If
    .RecordSource = "select *from fechamentodecaixa where fechames=0 and fechado=-1 and distribuido=-1"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechames = True
        .Recordset!DataFechamento = DataFechamento
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from formadepagamentorecebido2 where fechames=0 and fechamentodiario=-1"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechames = True
        .Recordset!DataFechamento = DataFechamento
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from cartoes where fechadiferenca=0 and confirmado=-1"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechadiferenca = True
        .Recordset!DataFechamento = DataFechamento
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from cartoes where fechataxa=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!DataFechamento = DataFechamento
        .Recordset!fechataxa = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from clientescobranca where fechames=0 and pago=-1"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechames = True
        .Recordset!DataFechames = DataFechamento
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from clientescobranca where origem = 'Aluguel' and fechaaluguel=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechaaluguel = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from clientescobranca where origem = 'Outros' and fechaaluguel=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!fechaaluguel = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    
    .RecordSource = "select *from despesaslanc2 where fechamento=0 and origem='Pg Funcionários'"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!DataFechamento = DataFechamento
        .Recordset!Fechamento = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from despesaslanc2 where FechamentoDiario=-1 and origem<>'Pg Funcionários' and codigofechamento=0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!DataFechamento = DataFechamento
        .Recordset!Fechamento = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
    .RecordSource = "select *from despesaslanc2 where FechamentoDiario=-1 and origem<>'Pg Funcionários' and codigofechamento<>0"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!DataFechamento = DataFechamento
        .Recordset!Fechamento = True
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
  End With
  Call cmdAtualiza_Click
  Screen.MousePointer = vbDefault
End If
End Sub

Private Sub cmdZerarTrimestral_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja zerar o trimestre do status?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbTemp
  .RecordSource = "select *from statusdiario where fechamentotrimestral=0"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Edit
      .Recordset!fechamentotrimestral = -1
      .Recordset!datatrimestre = DataCaixa
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  On Error GoTo 0
End With
End Sub

Private Sub Form_Load()
ExibeStatusMensal = True
On Error Resume Next
If Usuarios.Nome = "Usuário Master" Then
  cmdValorEstoque.Visible = True
Else
  cmdValorEstoque.Visible = False
End If
'If ComissaoAcumulativa = True Then
  lblComissoesDescri.Caption = "Comissões a Pagar:"
'Else
'  lblComissoesDescri.Caption = "Comissões Pagas:"
'End If
With adoContas
  .ConnectionString = CaminhoADO
  .Refresh
End With

With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from status order by zeradoem"
  'On Error Resume Next
  .Refresh
  If IsNull(.Recordset!ZeradoEm) = False Then
    lblZeradoEm.Caption = .Recordset!ZeradoEm
  Else
    lblZeradoEm.Caption = "Fechamento ainda não efetuado!"
  End If
  On Error Resume Next
  If IsNull(.Recordset!CapitalInicial) = False Then
    lblCapitalInicial.Caption = Format(.Recordset!CapitalInicial, "Currency")
  End If
  On Error GoTo 0
End With

With dbEstoque
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbTemp
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qEstoque
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qLucroVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbStatusDiario
  .Connect = Conectar
  .DatabaseName = Caminho
  On Error Resume Next
  .Refresh
  On Error GoTo 0
End With

Atualizar
AtualizaFechamentos


Select Case Usuarios.Grupo.AdmEstatus
  Case 1 'Somente leitura
    cmdZerar.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub optFechaDia_Click()
ExibeStatusMensal = False
AtualizaFechamentos
End Sub

Private Sub optFechaMes_Click()
ExibeStatusMensal = True
AtualizaFechamentos
End Sub

