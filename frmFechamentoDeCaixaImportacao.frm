VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFechamentoDeCaixaImportacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "frmFechamentoDeCaixaImportacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFechamentoDeCaixaImportacao.frx":0442
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
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
         DataField       =   "datacaixa"
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
         DataField       =   "turno"
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
         DataField       =   "planodeconta"
         Caption         =   "Plano de Conta"
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
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1544,882
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc dbCaixas 
      Height          =   330
      Left            =   3000
      Top             =   2400
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbCaixas"
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
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   128
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc dbCaixas2 
      Height          =   330
      Left            =   3000
      Top             =   2760
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbCaixas2"
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
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   265
      FullHeight      =   33
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmFechamentoDeCaixaImportacao.frx":0459
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
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
         DataField       =   "datacaixa"
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
         DataField       =   "turno"
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
         DataField       =   "planodeconta"
         Caption         =   "Plano de Conta"
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
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1544,882
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Caixa Final:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Caixa Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFechamentoDeCaixaImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigoPosto As String
Dim StrConexao As String

Private Sub cmdImportar_Click()
With dbCaixas
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  If .Recordset.BOF = True Then Exit Sub
  
  If dbCaixas.Recordset.AbsolutePosition < dbCaixas2.Recordset.AbsolutePosition Then
    MsgBox "O caixa inicial deve ser menor que o caixa final!"
    DataGrid1.SetFocus
    Exit Sub
  End If
  
  Do While .Recordset.EOF = False
    With frmFechamentoDeCaixaNovo
      .dbPdvs.Recordset.MoveFirst
      .dbPdvs.Recordset.Find "codigo='" & dbCaixas.Recordset!planodeconta & "'"
      If .dbPdvs.Recordset.EOF = False Then
        .cmdCancelar_Click
        .cboPdvs.Text = .dbPdvs.Recordset!Descri
        .cboPdvs.Refresh
        .txtData.Value = dbCaixas.Recordset!DataCaixa
        .txtData.Refresh
        .cboTurno.Text = dbCaixas.Recordset!Turno
        .cboTurno.Refresh
        .AbreCaixa
        If .dbFechamentos.Recordset!fechado = False Then
          .Importar
        End If
      End If
    End With
    If .Recordset.AbsolutePosition = dbCaixas2.Recordset.AbsolutePosition Then Exit Do
    .Recordset.MovePrevious
  Loop
End With
'Shell "notepad " & App.Path & "\NotasBloqueadas.txt"
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
Dim db As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset

db.Open CaminhoADO
dbConfig.CursorLocation = adUseClient
dbConfig.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly

StrConexao = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig!ftp
codigoPosto = dbConfig!Porta

dbConfig.Close
db.Close

On Error Resume Next
db.Open StrConexao
db.Execute "update caixas set planodeconta='2100000000' where planodeconta is null"
db.Close

With dbCaixas
  .ConnectionString = StrConexao
  .RecordSource = "select datacaixa, turno, planodeconta from caixas where linhaexportada like '000%' and codigoposto='" & codigoPosto & "' order by datacaixa desc, turno desc"
  .Refresh
End With
With dbCaixas2
  .ConnectionString = StrConexao
  .RecordSource = dbCaixas.RecordSource
  .Refresh
End With

End Sub
