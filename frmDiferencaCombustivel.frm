VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDiferencaCombustivel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diferença de Combustível"
   ClientHeight    =   3030
   ClientLeft      =   8025
   ClientTop       =   8115
   ClientWidth     =   6870
   Icon            =   "frmDiferencaCombustivel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6870
   Begin VB.CommandButton cmdPrimeira 
      Caption         =   ">>|"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   ">>"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdProcima 
      Caption         =   "<<"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdUltima 
      Caption         =   "|<<"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin MSAdodcLib.Adodc dbFechamentos 
      Height          =   375
      Left            =   3360
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select datacaixa, horaini, turno, fechado from fechamentodecaixa order by datacaixa desc, horaini desc"
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
   Begin MSAdodcLib.Adodc dbDifCombustivel 
      Height          =   375
      Left            =   3360
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select *from diferencacombustivel where codigofechamento=0"
      Caption         =   "dbDifCombustivel"
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
      Bindings        =   "frmDiferencaCombustivel.frx":0442
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
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
         DataField       =   "TanqueNr"
         Caption         =   "Tq."
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
         DataField       =   "Estoque"
         Caption         =   "Sistema"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Vendido"
         Caption         =   "Vendido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tanque"
         Caption         =   "Posto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Diferenca"
         Caption         =   "Diferença"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,###"
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
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "fechado"
      BeginProperty DataFormat 
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
      DataSource      =   "dbFechamentos"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Finalizado:"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "turno"
      DataSource      =   "dbFechamentos"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "datacaixa"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "dbFechamentos"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Data do Caixa:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmDiferencaCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnterior_Click()
If dbFechamentos.Recordset.BOF = True Then Exit Sub
dbFechamentos.Recordset.MovePrevious
End Sub

Private Sub cmdPrimeira_Click()
If dbFechamentos.Recordset.BOF = True Then Exit Sub
dbFechamentos.Recordset.MoveFirst
End Sub

Private Sub cmdProcima_Click()
If dbFechamentos.Recordset.EOF = True Then Exit Sub
dbFechamentos.Recordset.MoveNext
End Sub

Private Sub cmdUltima_Click()
If dbFechamentos.Recordset.EOF = True Then Exit Sub
dbFechamentos.Recordset.MoveLast
End Sub

Private Sub dbFechamentos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim CodigoFechamento As Double
CodigoFechamento = 0
With dbFechamentos
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.BOF = False And .Recordset.EOF = False Then
      CodigoFechamento = .Recordset!CodigoFechamento
    End If
  End If
End With

With dbDifCombustivel
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from diferencacombustivel where codigofechamento=" & CodigoFechamento
  .Refresh
End With
End Sub

Private Sub Form_Load()
With dbDifCombustivel
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from diferencacombustivel where codigofechamento=0"
  .Refresh
End With
With dbFechamentos
  .ConnectionString = CaminhoADO
  .RecordSource = "select CodigoFechamento, DataCaixa, horaini, turno, fechado from fechamentodecaixa order by datacaixa desc, horaini desc"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.Find "fechado=-1"
    If .Recordset.EOF = True Then
      .Recordset.MoveFirst
    End If
  End If
End With
End Sub

