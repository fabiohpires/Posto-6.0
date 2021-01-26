VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadFormaDePg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Forma de Pagamento"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmCadFormaDePg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5850
      Width           =   8310
      _ExtentX        =   14658
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
      RecordSource    =   "select *from formaDePagamento order by descri"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmCadFormaDePg.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picButtons"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Adodc2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbContas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Recebimentos"
      TabPicture(1)   =   "frmCadFormaDePg.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "lblTotalBruto"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "lblTotalLiquido"
      Tab(1).Control(6)=   "DataGrid2"
      Tab(1).Control(7)=   "QPagamentoRecebidoTotal"
      Tab(1).Control(8)=   "QPagamentoRecebido"
      Tab(1).Control(9)=   "txtDataFim"
      Tab(1).Control(10)=   "txtDataIni"
      Tab(1).Control(11)=   "cmdExibir"
      Tab(1).Control(12)=   "cmdImprime"
      Tab(1).ControlCount=   13
      Begin MSAdodcLib.Adodc dbContas 
         Height          =   330
         Left            =   2760
         Top             =   1680
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   2760
         Top             =   2040
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   "select *from contas order by descri"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCadFormaDePg.frx":047A
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
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
         ColumnCount     =   5
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
            DataField       =   "Reembolso"
            Caption         =   "Dias"
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
            DataField       =   "DescontoPorcento"
            Caption         =   "Desconto %"
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
            DataField       =   "DescontoValor"
            Caption         =   "Desconto $"
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
            DataField       =   "DescontoPorOperacao"
            Caption         =   "P/ Oper."
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
               ColumnWidth     =   3284,788
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   555,024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   975,118
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdImprime 
         Height          =   495
         Left            =   -70200
         Picture         =   "frmCadFormaDePg.frx":048F
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "Imprimir"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdExibir 
         Caption         =   "Exibir"
         Height          =   375
         Left            =   -71640
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox picButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         ScaleHeight     =   330
         ScaleWidth      =   7815
         TabIndex        =   33
         Top             =   5100
         Width           =   7815
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Adicionar"
            Height          =   300
            Left            =   1785
            TabIndex        =   26
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Remover"
            Height          =   300
            Left            =   2880
            TabIndex        =   27
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Atuali&zar"
            Height          =   300
            Left            =   3975
            TabIndex        =   28
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Gravar"
            Height          =   300
            Left            =   5070
            TabIndex        =   29
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   300
            Left            =   6165
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   300
            Left            =   705
            TabIndex        =   25
            Top             =   0
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker txtDataIni 
         Height          =   300
         Left            =   -74760
         TabIndex        =   35
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72941569
         CurrentDate     =   38191
      End
      Begin MSComCtl2.DTPicker txtDataFim 
         Height          =   300
         Left            =   -73080
         TabIndex        =   37
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72941569
         CurrentDate     =   38191
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   2175
         Left            =   150
         TabIndex        =   32
         Top             =   2880
         Width           =   7695
         Begin MSMask.MaskEdBox txtDataCorte 
            DataField       =   "dataCorte"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   5640
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "diasCorte"
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
            Index           =   3
            Left            =   7080
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chkDataCorte 
            Caption         =   "Data de Corte"
            DataField       =   "Corte"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   5640
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmCadFormaDePg.frx":0F11
            DataField       =   "PlanoDeConta"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1560
            TabIndex        =   23
            Top             =   1680
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   "CodigoConta"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmCadFormaDePg.frx":0F28
            DataField       =   "CodigoConta"
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   3360
            TabIndex        =   19
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Descri"
            BoundColumn     =   "CodigoConta"
            Text            =   ""
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "CodigoNoPosto"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Não Acumular"
            DataField       =   "NaoAcumula"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   5520
            TabIndex        =   24
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "grupo"
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
            Index           =   5
            Left            =   2640
            TabIndex        =   17
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "Reembolso"
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
            Index           =   2
            Left            =   3240
            TabIndex        =   4
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Descri"
            DataSource      =   "Adodc1"
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Mês"
            DataField       =   "Mes"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   3840
            TabIndex        =   5
            Top             =   480
            Width           =   615
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            DataField       =   "DescontoPorcento"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   4560
            TabIndex        =   7
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            DataField       =   "DescontoValor"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            DataField       =   "DescontoPorOperacao"
            DataSource      =   "Adodc1"
            Height          =   300
            Left            =   1320
            TabIndex        =   15
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   " "
         End
         Begin VB.Label lblDias 
            Caption         =   "Dias"
            Height          =   255
            Left            =   7080
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plano de Conta:"
            Height          =   195
            Index           =   3
            Left            =   1560
            TabIndex        =   22
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código no Posto:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   16
            Top             =   840
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Desconto %:"
            Height          =   195
            Index           =   7
            Left            =   4560
            TabIndex        =   6
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Desconto $:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Dias:"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   3
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Conta destino:"
            Height          =   195
            Index           =   0
            Left            =   3360
            TabIndex        =   18
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Desc. p/ Oper. $:"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   14
            Top             =   840
            Width           =   1245
         End
      End
      Begin MSAdodcLib.Adodc QPagamentoRecebido 
         Height          =   375
         Left            =   -73080
         Top             =   4200
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
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
         RecordSource    =   "select *from QFormadePgRecebidoFechamento2 where codigoformadepg=0"
         Caption         =   "QPagamentoRecebido"
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
      Begin MSAdodcLib.Adodc QPagamentoRecebidoTotal 
         Height          =   375
         Left            =   -73080
         Top             =   4560
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
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
         RecordSource    =   "select *from QFormadePgRecebidoFechamento2 where codigopagamento=0"
         Caption         =   "QPagamentoRecebidoTotal"
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
         Bindings        =   "frmCadFormaDePg.frx":0F3D
         Height          =   3975
         Left            =   -74880
         TabIndex        =   44
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7011
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
            DataField       =   "DataCaixa"
            Caption         =   "Data Caixa"
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
         BeginProperty Column02 
            DataField       =   "Data"
            Caption         =   "Data"
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
            DataField       =   "Operacoes"
            Caption         =   "Operacoes"
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
            DataField       =   "ValorBruto"
            Caption         =   "Valor Bruto"
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
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalLiquido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -68640
         TabIndex        =   43
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Total Liquido:"
         Height          =   255
         Left            =   -69720
         TabIndex        =   42
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label lblTotalBruto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -71280
         TabIndex        =   41
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Total Bruto:"
         Height          =   255
         Left            =   -72240
         TabIndex        =   40
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "a"
         Height          =   255
         Left            =   -73320
         TabIndex        =   36
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Período do bordero:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCadFormaDePg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdemPg As String, StrOrdemRecebido As String

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1
End Sub

Private Sub chkDataCorte_Click()
If chkDataCorte.Value = vbChecked Then
  txtDataCorte.Visible = True
  lblDias.Visible = True
  txtFields(3).Visible = True
Else
  txtDataCorte.Visible = False
  lblDias.Visible = False
  txtFields(3).Visible = False
End If
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(0).SetFocus
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
txtFields(0).SetFocus
End Sub

Private Sub cmdExibir_Click()
Call DataGrid2_HeadClick(0)
End Sub

Private Sub cmdImprime_Click()
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

ImprimeADOGrid DataGrid2, Printer, QPagamentoRecebido, 4, , , , 5, , "Cartões Recebidos", Adodc1.Recordset!Descri, Format(Date, "long date")

Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Adodc2.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    .Recordset.Update
    Codigo = .Recordset!CodigoPagamento
    Grupo = .Recordset!Grupo
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
On Error Resume Next
With QPagamentoRecebido
  .ConnectionString = CaminhoADO
  If .RecordSource = "select *from QFormadePgRecebidoFechamento2 where codigoformadepg=" & Adodc1.Recordset!CodigoPagamento & " and fechamentodiario=-1  and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# order by " & DBGrid1.Columns(ColIndex).DataField & ", datacaixa, horaini, data" Then
    .RecordSource = "select *from QFormadePgRecebidoFechamento2 where codigoformadepg=" & Adodc1.Recordset!CodigoPagamento & " and fechamentodiario=-1  and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# order by " & DBGrid1.Columns(ColIndex).DataField & " desc, datacaixa, horaini, data"
  Else
    .RecordSource = "select *from QFormadePgRecebidoFechamento2 where codigoformadepg=" & Adodc1.Recordset!CodigoPagamento & " and fechamentodiario=-1  and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# order by " & DBGrid1.Columns(ColIndex).DataField & ", datacaixa, horaini, data"
  End If
  .Refresh
End With
With QPagamentoRecebidoTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(valorbruto) as bruto, sum(valor) as liquido from QFormadePgRecebidoFechamento2 where codigoformadepg=" & Adodc1.Recordset!CodigoPagamento & " and fechamentodiario=-1  and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# "
  .Refresh
  If IsNull(.Recordset!Bruto) = False Then
    lblTotalBruto.Caption = Format(.Recordset!Bruto, "Currency")
  Else
    lblTotalBruto.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!Liquido) = False Then
    lblTotalLiquido.Caption = Format(.Recordset!Liquido, "Currency")
  Else
    lblTotalLiquido.Caption = Format(0, "Currency")
  End If
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
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With Adodc1
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from formaDePagamento order by descri"
  .Refresh
End With
With Adodc2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With QPagamentoRecebido
  .ConnectionString = CaminhoADO
  .Refresh
End With
With QPagamentoRecebidoTotal
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
Select Case Usuarios.Grupo.CadFormaDePg
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

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataFim_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtFields_Change(Index As Integer)
Select Case Index
  Case 1, 3, 4
    On Error Resume Next
    Select Case KeyAscii
      Case Asc(".")
        KeyAscii = 0
        SendKeys ","
    End Select
End Select
End Sub
