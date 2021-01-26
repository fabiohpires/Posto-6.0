VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatContasAReceber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Contas a Receber"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
   Icon            =   "frmRelatContasAReceber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   10680
      Picture         =   "frmRelatContasAReceber.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cboOrigem 
      Height          =   315
      ItemData        =   "frmRelatContasAReceber.frx":0EC4
      Left            =   5280
      List            =   "frmRelatContasAReceber.frx":0ED4
      TabIndex        =   7
      Text            =   "Todos"
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cboPagamento 
      Height          =   315
      ItemData        =   "frmRelatContasAReceber.frx":0EF7
      Left            =   3240
      List            =   "frmRelatContasAReceber.frx":0F04
      TabIndex        =   5
      Text            =   "Todos"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   1680
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
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
      Begin MSAdodcLib.Adodc dbCompoisicao 
         Height          =   330
         Left            =   360
         Top             =   1800
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
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatContasAReceber.frx":0F21
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5741
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
      ColumnCount     =   8
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3149,858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3780,284
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboCliente 
      Bindings        =   "frmRelatContasAReceber.frx":0F42
      Height          =   315
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboClientesTipo 
      Bindings        =   "frmRelatContasAReceber.frx":0F5B
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "TipoCliente"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   37620
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   37620
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmRelatContasAReceber.frx":0F78
      Height          =   1455
      Left            =   120
      TabIndex        =   20
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
   Begin VB.Label Label5 
      Caption         =   "Origem:"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pagamento:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label4 
      Caption         =   "Código:"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label28 
      Caption         =   "Tipo de Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
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
      TabIndex        =   17
      Top             =   4800
      Width           =   2055
   End
End
Attribute VB_Name = "frmRelatContasAReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub cboCliente_LostFocus()
With dbClientes
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboCliente.Text = "" Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & cboCliente.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!CodigoCliente
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

Private Sub cmdExibir_Click()
Dim StrTemp As String, StrTotal As String

StrTemp = "select *from clientescobranca where datafechamento between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
StrTotal = "select sum(valor) as total from clientescobranca where datafechamento between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"

Select Case cboPagamento.Text
  Case "Pagos"
    StrTemp = StrTemp & " and pago=-1"
    StrTotal = StrTotal & " and pago=-1"
  Case "Não Pagos"
    StrTemp = StrTemp & " and pago=0"
    StrTotal = StrTotal & " and pago=0"
End Select

If cboClientesTipo.Text <> "" Then
  StrTemp = StrTemp & " and tipocliente='" & cboClientesTipo.Text & "'"
  StrTotal = StrTotal & " and tipocliente='" & cboClientesTipo.Text & "'"
End If

If cboCliente.Text <> "" Then
  If dbClientes.Recordset.EOF = False And dbClientes.Recordset.BOF = False Then
    If dbClientes.Recordset!Nome = cboCliente.Text Then
      StrTemp = StrTemp & " and codigocliente=" & dbClientes.Recordset!CodigoCliente
      StrTotal = StrTotal & " and codigocliente=" & dbClientes.Recordset!CodigoCliente
    End If
  End If
End If

If cboOrigem.Text <> "" And cboOrigem.Text <> "Todos" Then
  StrTemp = StrTemp & " and origem='" & cboOrigem.Text & "'"
  StrTotal = StrTotal & " and origem='" & cboOrigem.Text & "'"
End If

With dbClientesCobranca
  .ConnectionString = CaminhoADO
  .RecordSource = StrTemp & strOrdem
  .Refresh
End With
With qClientesCobranca
  .ConnectionString = CaminhoADO
  .RecordSource = StrTotal
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, StrTemp2 As String
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = Chr(13)
StrTemp = StrTemp & "Período:" & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataIni.Value, "short date")

If cboPagamento.Text <> "" Then
  StrTemp = StrTemp & "      Pagamento: " & cboPagamento.Text
End If
If cboOrigem.Text <> "" Then
  StrTemp = StrTemp & "      Origem: " & cboOrigem.Text
End If
If cboClientesTipo.Text <> "" Then
  StrTemp = StrTemp & Chr(13) & "Tipo Cliente: " & cboOrigem.Text
End If
If cboCliente.Text <> "" Then
  StrTemp = StrTemp & "     Codigo: " & txtCodigo.Text & "   Cliente: " & cboCliente.Text
End If

StrTemp2 = "Impresso em: " & Format(Date, "long date")

ImprimeADOGrid DataGrid1, Printer, dbClientesCobranca, 3, True, , , , , "Relatório de Contas a Receber" & Chr(13) & NomePosto, StrTemp, StrTemp2

Printer.EndDoc

NaoImprime:

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DataGrid1.Columns(ColIndex) Then
  strOrdem = " order by " & DataGrid1.Columns(ColIndex) & " desc"
Else
  strOrdem = " order by " & DataGrid1.Columns(ColIndex)
End If
End Sub

Private Sub dbClientesCobranca_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim CodigoCobranca As Double
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
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date

strOrdem = " order by DataFechamento"

With dbClientes
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbClientesCobranca
  .ConnectionString = CaminhoADO
  .Refresh
End With
With qClientesCobranca
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbClientesTipo
  .ConnectionString = CaminhoADO
  .Refresh
End With

Call cmdExibir_Click

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

