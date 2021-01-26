VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDespesasParaContabilidade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Despesas para Contabilidade"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmDespesasParaContabilidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmDespesasParaContabilidade.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "Imprimir"
      Top             =   7320
      Width           =   735
   End
   Begin VB.ComboBox cboConfirmado 
      Height          =   315
      ItemData        =   "frmDespesasParaContabilidade.frx":0EC4
      Left            =   5640
      List            =   "frmDespesasParaContabilidade.frx":0ED1
      TabIndex        =   12
      Text            =   "Totas"
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboAutoriazado 
      Height          =   315
      ItemData        =   "frmDespesasParaContabilidade.frx":0F0C
      Left            =   3240
      List            =   "frmDespesasParaContabilidade.frx":0F1C
      TabIndex        =   11
      Text            =   "Não Autorizadas"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdSubtrair 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdSomar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbSoma 
      Height          =   330
      Left            =   4680
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select sum(valor) as total from despesaslanc2 where autorizacao=0 and fechamento=0"
      Caption         =   "dbSoma"
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
      Left            =   4560
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 order by data"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDespesasParaContabilidade.frx":0F59
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Não Enviadas"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Vencimento"
         Caption         =   "Venc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Data"
         Caption         =   "Lanc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "ValorPago"
         Caption         =   "Pago"
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
      BeginProperty Column06 
         DataField       =   "Autorizacao"
         Caption         =   "Autorizada"
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
      BeginProperty Column07 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1019,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dbSoma2 
      Height          =   330
      Left            =   2280
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select sum(valor) as total from despesaslanc2 where codigoenviar='Enviando'"
      Caption         =   "dbSoma2"
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
   Begin MSAdodcLib.Adodc dbDespesaLanc2 
      Height          =   330
      Left            =   2280
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select *from DespesasLanc2 where codigoenviar='Enviando'  order by data"
      Caption         =   "dbDespesaLanc2"
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
      Bindings        =   "frmDespesasParaContabilidade.frx":0F75
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Enviando"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Vencimento"
         Caption         =   "Venc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Data"
         Caption         =   "Lanc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "ValorPago"
         Caption         =   "Pago"
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
      BeginProperty Column06 
         DataField       =   "Autorizacao"
         Caption         =   "Autorizada"
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
      BeginProperty Column07 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1019,906
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   38286
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   38286
   End
   Begin VB.CheckBox chkFechadas 
      Caption         =   "Despesas já fechadas de meses anteriores"
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   8040
      TabIndex        =   19
      Top             =   7320
      Width           =   405
   End
   Begin VB.Label Label7 
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
         SubFormatType   =   0
      EndProperty
      DataSource      =   "dbSoma2"
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   8520
      TabIndex        =   17
      Top             =   3840
      Width           =   405
   End
   Begin VB.Label Label5 
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
         SubFormatType   =   0
      EndProperty
      DataSource      =   "dbSoma"
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Confirmada Descrição:"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Autorização:"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Período de lançamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmDespesasParaContabilidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigoSoma As String, ColunaDeQuebra As Integer

Private Sub Filtrar()
Dim StrTemp As String, StrTemp2 As String
Dim StrDescri As String, strFechamento As String
Dim A As Double, B As Double


If chkFechadas.Value = vbChecked Then
  strFechamento = ""
Else
  strFechamento = " and fechamento=0"
End If

Select Case cboConfirmado
  Case "Não Confirmada Descrição"
    StrDescri = " and usuario='Não'"
  Case "Confirmada Descrição"
    StrDescri = " and usuario='Sim'"
  Case "Todas"
    StrDescri = ""
End Select

StrTemp = "select *from despesaslanc2 where codigoenviar='1' and paracontabilidade=0 and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and fechamentodiario=-1"
StrTemp2 = "select sum(valor) as total from despesaslanc2 where codigoenviar='1' and paracontabilidade=0 and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and fechamentodiario=-1"

Select Case cboAutoriazado.Text
  Case "Não Autorizadas"
    StrTemp = StrTemp & " and autorizacao=0" & strFechamento & StrDescri & " and produto=0 order by data, hora"
    StrTemp2 = StrTemp2 & " and autorizacao=0" & strFechamento & StrDescri & " and produto=0"
  Case "Autorizadas"
    StrTemp = StrTemp & " and autorizacao=-1" & strFechamento & StrDescri & " and produto=0 order by data, hora"
    StrTemp2 = StrTemp2 & " and autorizacao=-1" & strFechamento & StrDescri & " and produto=0"
  Case "Todas"
    StrTemp = StrTemp & " and produto=0" & strFechamento & StrDescri & " order by data, hora"
    StrTemp2 = StrTemp2 & " and produto=0" & strFechamento & StrDescri
  Case "Despesas Bancárias"
    StrTemp = StrTemp & " and produto=0 and origem='Conciliação'" & strFechamento & StrDescri & " order by data, hora"
    StrTemp2 = StrTemp2 & " and produto=0 and origem='Conciliação'" & strFechamento & StrDescri
End Select

With dbDespesalanc
  On Error Resume Next
  A = .Recordset.AbsolutePosition
  On Error GoTo 0
  .ConnectionString = CaminhoADO
  .RecordSource = StrTemp
  .Refresh
  .Refresh
  On Error Resume Next
  .Recordset.AbsolutePosition = A
  On Error GoTo 0
End With
With dbSoma
  .ConnectionString = CaminhoADO
  .RecordSource = StrTemp2
  .Refresh
  .Refresh
End With

With dbDespesaLanc2
  On Error Resume Next
  B = .Recordset.AbsolutePosition
  On Error GoTo 0
  .ConnectionString = CaminhoADO
  .Refresh
  .Refresh
  On Error Resume Next
  .Recordset.AbsolutePosition = B
  On Error GoTo 0
End With
With dbSoma2
  .ConnectionString = CaminhoADO
  .Refresh
  .Refresh
End With
End Sub

Private Sub cboAutoriazado_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboAutoriazado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboAutoriazado_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cboConfirmado_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConfirmado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboConfirmado_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmdEnviar_Click()
Dim DataEnviado As Date, Resposta As Integer
Dim db As New ADODB.Connection


codigoSoma = GeraCodigo()
DataEnviado = Now

Resposta = MsgBox("Deseja enviar agora os documentos selecionados?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

db.Open CaminhoADO

db.Execute "update despesaslanc2 set ParaContabilidade=-1, DataContabilidade=#" & DataInglesa(DataEnviado) & " " & Time & "#, CodigoEnviar='" & codigoSoma & "' where CodigoEnviar='Enviando'"
db.Close
dbDespesaLanc2.Refresh
End Sub

Private Sub cmdExibir_Click()
Filtrar
End Sub

Private Sub cmdImprime_Click()
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

ImprimeADOGrid DataGrid2, Printer, dbDespesaLanc2, 4, True, , ColunaDeQuebra, , , "Protocolo de Envio de Despesas para Contabilidade", NomePosto, "Impresso em: " & Format(Date, "Long date") & " - " & Format(Time, "short time") & Chr(13) & "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")


Printer.EndDoc

NaoImprime:

End Sub

Private Sub cmdSomar_Click()
With dbDespesalanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma despesa primeiro!"
    Exit Sub
  End If
  .Recordset!codigoenviar = codigoSoma
  .Recordset.Update
End With
Call cmdExibir_Click
DataGrid1.SetFocus
End Sub

Private Sub cmdSubtrair_Click()
With dbDespesaLanc2
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma despesa primeiro!"
    Exit Sub
  End If
  .Recordset!codigoenviar = "1"
  .Recordset.Update
End With
Call cmdExibir_Click
DataGrid2.SetFocus
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
dbDespesalanc.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
ColunaDeQuebra = ColIndex
dbDespesaLanc2.Recordset.Sort = DataGrid2.Columns(ColIndex).DataField
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
  Case Asc("+")
    KeyAscii = 0
    Call cmdSomar_Click
  Case Asc("-")
    KeyAscii = 0
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub Form_Load()
ColunaDeQuebra = -1
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date

codigoSoma = "Enviando"
Filtrar
End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyPress(KeyAscii As Integer)
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
