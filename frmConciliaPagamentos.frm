VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConciliaPagamentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliação de Pagamentos"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmConciliaPagamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37648
   End
   Begin MSAdodcLib.Adodc dbPendencias 
      Height          =   330
      Left            =   2280
      Top             =   2880
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from compensapendente where conciliado=0"
      Caption         =   "dbPendencias"
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
      Bindings        =   "frmConciliaPagamentos.frx":0442
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5530
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
         DataField       =   "Data"
         Caption         =   "Data"
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
      BeginProperty Column02 
         DataField       =   "NrDoc"
         Caption         =   "NrDocumento"
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
         DataField       =   "Descri"
         Caption         =   "Descri"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1649,764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmConciliaPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Totaliza()
Dim Total As Currency
With dbPendencias
  .Refresh
  lblTotal.Caption = ""
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    Total = Total + .Recordset!Valor
    .Recordset.MoveNext
  Loop
  lblTotal.Caption = Format(Total, "Currency")
  .Recordset.MoveFirst
End With
End Sub
Private Sub cmdConfirmar_Click()
If dbPendencias.Recordset.RecordCount = 0 Then Exit Sub
If dbPendencias.Recordset.EOF = True Then
  MsgBox "Selecione um registro primeiro!"
  Exit Sub
End If
With frmConciliacao
  With .dbConcilia
    .Recordset.AddNew
    .Recordset!codigoconta = frmConciliacao.dbContas.Recordset!codigoconta
    .Recordset!Data = txtData.Value
    .Recordset!tipo = "Pagamento"
    .Recordset!Codigo = dbPendencias.Recordset!CodigoDespesaLanc
    .Recordset!Descri = dbPendencias.Recordset!Descri
    .Recordset!NrDocumento = dbPendencias.Recordset!NrDoc & " "
    .Recordset!Valor = dbPendencias.Recordset!Valor
    .Recordset.Update
  End With
'  With .dbContas
'    .Recordset!Saldo = .Recordset!Saldo + dbPendencias.Recordset!Valor
'    .Recordset.Update
'  End With
  .TiraSaldo txtData.Value
End With

With dbPendencias
  .Recordset!datacompensado = txtData.Value
  .Recordset!databaixado = Now
  .Recordset!conciliado = True
  .Recordset.Update
  .Refresh
  .Refresh
  .Refresh
End With
Totaliza
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbPendencias_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
txtData.Value = dbPendencias.Recordset!Data
End Sub

Private Sub Form_Activate()
Totaliza
End Sub

Private Sub Form_Load()
With dbPendencias
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
txtData.Value = Date
Totaliza
End Sub
