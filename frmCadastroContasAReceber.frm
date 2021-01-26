VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadastroContasAReceber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incluir Contas a Receber"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11580
   Icon            =   "frmCadastroContasAReceber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo cboDescri 
      Bindings        =   "frmCadastroContasAReceber.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.CommandButton cmdRemoverComposicao 
      Caption         =   "Remover"
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdIncluirResumo 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkReembolso 
      Caption         =   "Reembolso"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CheckBox chkAluguel 
      Caption         =   "Aluguel"
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   3120
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   7095
      Begin MSAdodcLib.Adodc dbClientes 
         Height          =   330
         Left            =   360
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientes order by nome"
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
      Begin MSAdodcLib.Adodc dbClientesCobranca 
         Height          =   330
         Left            =   360
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientescobranca order by datafechamento"
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
      Begin MSAdodcLib.Adodc dbClientesTipo 
         Height          =   330
         Left            =   360
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from clientesTipo order by tipocliente"
         Caption         =   "dbClientesTipo"
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
      Begin MSAdodcLib.Adodc qClientesCobranca 
         Height          =   330
         Left            =   360
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select sum(valor) as total from clientescobranca"
         Caption         =   "qClientesCobranca"
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
      Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
         Height          =   330
         Left            =   3120
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from bloqueiafechamento"
         Caption         =   "dbBloqueiaFechamento"
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
      Begin MSAdodcLib.Adodc dbCompoisicao 
         Height          =   330
         Left            =   3120
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from ClientesCobrancaComposicao where codigocobranca=0"
         Caption         =   "dbCompoisicao"
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
      Begin MSAdodcLib.Adodc dbCompoisicaoTipo 
         Height          =   330
         Left            =   3120
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from Composicaotipo order by descri"
         Caption         =   "dbCompoisicaoTipo"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadastroContasAReceber.frx":0462
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "DataFechamento"
         Caption         =   "Vencimento"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Origem"
         Caption         =   "Origem"
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
         DataField       =   "Obs"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2129,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4589,858
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   10320
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtDescri 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtNrNota 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Top             =   4440
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker txtVencimento 
      Height          =   300
      Left            =   8760
      TabIndex        =   7
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39582
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin MSDataListLib.DataCombo cboCliente 
      Bindings        =   "frmCadastroContasAReceber.frx":0483
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboClientesTipo 
      Bindings        =   "frmCadastroContasAReceber.frx":049C
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "TipoCliente"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCadastroContasAReceber.frx":04B9
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   6975
      _ExtentX        =   12303
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
         DataField       =   "Reembolso"
         Caption         =   "Reemb."
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
      BeginProperty Column02 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   3750,236
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qClientesCobranca"
      Height          =   255
      Left            =   9360
      TabIndex        =   27
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   8640
      TabIndex        =   26
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "Tipo de Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Nr. Nota:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Valor:"
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vencimento:"
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label4 
      Caption         =   "Código:"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCadastroContasAReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCliente_LostFocus()
With dbClientes
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboCliente.Text = "" Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & cboCliente.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!CodigoCliente
    If .Recordset!tipocliente = "Aluguel" Then
      chkAluguel.Value = vbChecked
    Else
      chkAluguel.Value = vbUnchecked
    End If
  End If
End With
End Sub

Private Sub cboClientesTipo_LostFocus()
With dbClientesTipo
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboClientesTipo.Text = "" Then
    With dbClientes
      .RecordSource = "select *from clientes order by nome"
      .Refresh
    End With
  Else
    .Recordset.Find "tipocliente='" & cboClientesTipo.Text & "'"
    If .Recordset.EOF = False Then
      With dbClientes
        .RecordSource = "select *from clientes where tipocliente='" & cboClientesTipo.Text & "' order by nome"
        .Refresh
      End With
    Else
      With dbClientes
        .RecordSource = "select *from clientes order by nome"
        .Refresh
      End With
    End If
  End If
End With
End Sub

Private Sub cmdAtualizar_Click()
cboClientesTipo.Text = ""
Call cboClientesTipo_LostFocus
txtCodigo.Text = ""
cboCliente.Text = ""
txtValor.Text = ""
txtNrNota.Text = ""
txtDescri.Text = ""
dbClientes.Refresh
dbClientesCobranca.Refresh
qClientesCobranca.Refresh
dbClientesTipo.Refresh

End Sub

Private Sub cmdIncluir_Click()
Dim Origem As String

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtVencimento.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If DateDiff("d", Date, txtVencimento.Value) >= 30 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar despesa com data futura acima de 30 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) <= -15 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar despesa com data anterior a 15 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) >= 120 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar vencimento com data acima de 90 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtVencimento.Value) <= -1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode lançar despesa já vencida!"
    Exit Sub
  End If
End If


If chkAluguel.Value = vbChecked Then
  Origem = "Aluguel"
Else
  Origem = "Outros"
End If
With dbClientes
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Escolha um cliente primeiro"
    Exit Sub
  End If
  If cboCliente.Text <> .Recordset!Nome Then
    MsgBox "Cliente não localizado!"
    Exit Sub
  End If
End With


With dbClientesCobranca
  .Recordset.AddNew
  .Recordset!datasoma = Now
  .Recordset!NrNota = txtNrNota.Text
  .Recordset!DataFechamento = txtVencimento.Value
  .Recordset!CodigoCliente = txtCodigo.Text
  .Recordset!Cliente = cboCliente.Text
  .Recordset!Valor = 0
  .Recordset!Obs = txtDescri.Text
  .Recordset!tipocliente = dbClientes.Recordset!tipocliente
  .Recordset!Origem = Origem
  .Recordset.Update
  .Refresh
End With

Call cmdAtualizar_Click

cboCliente.SetFocus

End Sub

Private Sub cmdIncluirResumo_Click()
Dim CodigoCobranca As Double

If cboDescri.Text = "" Then
  MsgBox "Informe uma descrição!"
  txtDescriResumo.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor incorreto!"
  Exit Sub
End If
With dbClientesCobranca
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma cobrança primeiro!"
    Exit Sub
  End If
  If .Recordset!Origem = "Fiado" Then
    MsgBox "Esta cobrança é de clientes de fiado. Não pode ser alterada!"
    Exit Sub
  End If
  If .Recordset!fechaaluguel = -1 Then
    MsgBox "Já pertence a fechamento anterior!"
    Exit Sub
  End If
  CodigoCobranca = .Recordset!CodigoCobranca
End With
If chkReembolso.Value = vbChecked Then
  reembolso = True
Else
  reembolso = False
End If
With dbCompoisicao
  .Recordset.AddNew
  .Recordset!CodigoCobranca = CodigoCobranca
  .Recordset!Descri = cboDescri.Text
  .Recordset!reembolso = reembolso
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset.Update
  .Refresh
  .Refresh
End With
With dbClientesCobranca
  .Recordset!Valor = .Recordset!Valor + txtValor.Text
  .Recordset.Update
End With
With dbClientes
  .Refresh
  Call cboCliente_LostFocus
  .Recordset!Saldo = .Recordset!Saldo + CCur(txtValor.Text)
  .Recordset!TotalBoleto = .Recordset!TotalBoleto + CCur(txtValor.Text)
  .Recordset.Update
End With
cboDescri.Text = ""
chkReembolso.Value = vbUnchecked
txtValor.Text = ""
cboDescri.SetFocus
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer, StrTemp As String

With dbClientesCobranca
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe cobrança para ser removida!"
    Exit Sub
  End If
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma cobrança primeiro!"
    Exit Sub
  End If
  If .Recordset!Pago = True Then
    MsgBox "A cobrança atual já está paga!"
    Exit Sub
  End If
  If .Recordset!fechaaluguel = True Then
    MsgBox "A cobrança atual já foi contabilizada no fechamento!"
    Exit Sub
  End If
  If .Recordset!Valor <> 0 Then
    MsgBox "A cobrança precisa estar com valor R$ 0,00 para poder ser removida!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover a cobrança atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  
  .Recordset.Delete
  .Refresh
  .Refresh
End With
End Sub

Private Sub cmdRemoverComposicao_Click()
Dim Resposta As Integer, TempValor As Currency
Resposta = MsgBox("Deseja remover o item selecionado?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

With dbClientesCobranca
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma cobrança primeiro!"
    Exit Sub
  End If
  If .Recordset!Origem = "Fiado" Then
    MsgBox "Esta cobrança é de clientes de fiado. Não pode ser alterada!"
    Exit Sub
  End If
  If .Recordset!fechaaluguel = -1 Then
    MsgBox "Já pertence a fechamento anterior!"
    Exit Sub
  End If
End With
With dbCompoisicao
  TempValor = .Recordset!Valor
  .Recordset.Delete adAffectCurrent
  .Refresh
  .Refresh
End With
With dbClientesCobranca
  With dbClientes
    StrTemp = .RecordSource
    .RecordSource = "select *from clientes"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Find "codigocliente=" & dbClientesCobranca.Recordset!CodigoCliente
      If .Recordset.EOF = True Then
        MsgBox "Erro na tabela de clientes, não pode ser removido.!"
        Exit Sub
      End If
    End If
    .Recordset!Saldo = .Recordset!Saldo - TempValor
    .Recordset!TotalBoleto = .Recordset!TotalBoleto - TempValor
    .Recordset.Update
    .RecordSource = StrTemp
    .Refresh
  End With
  .Recordset!Valor = .Recordset!Valor - TempValor
  .Recordset.Update
End With

End Sub

Private Sub dbClientesCobranca_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim CodigoCobranca As Double, Total As Double
Total = 0
With dbClientesCobranca
  If .Recordset.EOF = True Or .Recordset.BOF = True Or IsNull(.Recordset!CodigoCobranca) = True Then
    CodigoCobranca = 0
  Else
    CodigoCobranca = .Recordset!CodigoCobranca
  End If
End With
With dbCompoisicao
  .RecordSource = "select *from clientescobrancacomposicao where codigocobranca=" & CodigoCobranca & " order by descri"
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      Total = Total + .Recordset!Valor
      .Recordset.MoveNext
    Loop
    .Recordset.MoveFirst
  End If
End With
With dbClientesCobranca
  If .Recordset.EOF = False Or .Recordset.BOF = False Or IsNull(.Recordset!CodigoCobranca) = False Then
    If .Recordset.RecordCount <> 0 Then
      If .Recordset!Valor <> Total Then
        .Recordset!Valor = Total
        .Recordset.Update
      End If
    End If
  End If
End With
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
With dbClientes
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbClientesCobranca
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from clientescobranca where origem<>'Fiado' and pago=0 order by datafechamento"
  .Refresh
End With
With qClientesCobranca
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(valor) as total from clientescobranca where origem<>'Fiado' and pago=0"
  .Refresh
End With
With dbClientesTipo
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCompoisicaoTipo
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew "Descri", "Locativo"
    .Recordset.AddNew "Descri", "IPTU"
    .Recordset.AddNew "Descri", "Água"
    .Recordset.AddNew "Descri", "Luz"
    .Recordset.AddNew "Descri", "Condomínio"
    .Recordset.AddNew "Descri", "Boleto"
    .Recordset.AddNew "Descri", "Multa"
    .Recordset.AddNew "Descri", "Juros"
  End If
End With

End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbClientes
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "codigocliente=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    cboCliente.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
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
