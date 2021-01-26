VERSION 5.00
Object = "{CEFAFC26-8C37-11D1-9618-D4DB04C10000}#1.0#0"; "YoconTray.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmImportadorAutomatico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação automática"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6600
      Top             =   120
   End
   Begin YOCONTRAYLib.YoconTray YoconTray1 
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdVefivicar 
      Caption         =   "Verificar"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtHoras 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "3"
      Top             =   4200
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   4920
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmImportadorAutomatico.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         DataField       =   "NomePosto"
         Caption         =   "Nome"
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
         DataField       =   "LocalDB"
         Caption         =   "Local do banco-de-dados"
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
            ColumnWidth     =   2145,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4334,74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdProcurar 
      Caption         =   "Procurar"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtCaminho 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc dbImportar 
      Height          =   375
      Left            =   3720
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Connect         =   $"frmImportadorAutomatico.frx":0019
      OLEDBString     =   $"frmImportadorAutomatico.frx":00B1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from importar order by nomeposto"
      Caption         =   "dbImportar"
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
   Begin VB.Label Label4 
      Caption         =   "Horas:"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Importar a cada"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Local do Banco-de-Dados:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "Exibir"
      Begin VB.Menu mnuConfigurar 
         Caption         =   "Configurar"
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "Fechar o Programa"
      End
   End
End
Attribute VB_Name = "frmImportadorAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim dbIntegrador As New ADODB.Recordset, IconeAtual As Integer

Private Sub cmdIncluir_Click()
With dbImportar
  .Recordset.AddNew
  .Recordset!Nomeposto = txtNome.Text
  .Recordset!localdb = txtCaminho.Text
  
End With
End Sub

Private Sub cmdProcurar_Click()
On Error GoTo TrataErro
With CommonDialog1
  .Filter = "Banco de Dados Access|*.mdb"
  .ShowOpen
  txtCaminho.Text = .filename
End With

TrataErro:

End Sub

Private Sub cmdVefivicar_Click()
Dim Horas As Integer
Horas = 0

IconeAtual = 102
Timer2.Enabled = True
Verifica Horas, dbImportar
Timer2.Enabled = False
IconeAtual = 101
YoconTray1.ModifyIcon 1, IconeAtual, "Importação Automática"
End Sub

Private Sub Form_Load()
IconeAtual = 101
With YoconTray1
  If .AddIcon(1, LoadResPicture(IconeAtual, vbResIcon), "Importação Automática") = False Then
    .AddIcon 1, LoadResPicture(IconeAtual, vbResIcon), "Importação Automática"
  End If
End With

Configura.ChequesNoCaixa = CInt(ReadINI("cheques", "Cheques", 0, App.Path & "\Posto.ini"))
Configura.NotaNoCaixa = CInt(ReadINI("Notas no Caixa", "Nocaixa", 0, App.Path & "\Posto.ini"))
Configura.NotaBloqueia = CInt(ReadINI("Notas no Caixa", "Bloqueia", 0, App.Path & "\Posto.ini"))

strMdb = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="

CaminhoImporta = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Importacao.mdb"

txtHoras.Text = GetSetting(App.EXEName, "Configura", "Horas", "3")

With dbImportar
  .ConnectionString = CaminhoImporta
  .Refresh
End With
Me.Hide
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Terminate()
On Error Resume Next
YoconTray1.DeleteIcon 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
YoconTray1.DeleteIcon 1
End Sub

Private Sub mnuConfigurar_Click()
Me.Show
Me.SetFocus
End Sub

Private Sub mnuFechar_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja sair do programa!", vbYesNo)
If Resposta = vbYes Then
  YoconTray1.DeleteIcon 1
  End
End If
End Sub

Private Sub Timer1_Timer()
Dim Horas As Integer
If IsNumeric(txtHoras.Text) = False Then
  Horas = -3
Else
  Horas = -CInt(txtHoras.Text)
End If
IconeAtual = 102
Timer2.Enabled = True
Verifica Horas, dbImportar
Timer2.Enabled = False
IconeAtual = 101
YoconTray1.ModifyIcon 1, LoadResPicture(IconeAtual, vbResIcon), "Importação Automática"


End Sub

Private Sub Timer2_Timer()
Select Case IconeAtual
  Case 102
    IconeAtual = 103
  Case 103
    IconeAtual = 104
  Case 104
    IconeAtual = 105
  Case 105
    IconeAtual = 106
  Case 106
    IconeAtual = 102
End Select

With YoconTray1
  .ModifyIcon 1, LoadResPicture(IconeAtual, vbResIcon), "Verificando Importação"
End With
End Sub

Private Sub txtHoras_Change()
SaveSetting App.EXEName, "Configura", "Horas", txtHoras.Text
End Sub

Private Sub YoconTray1_LeftClick(ByVal ID As Integer)
PopupMenu mnuExibir
End Sub

Private Sub YoconTray1_LeftDoubleClick(ByVal ID As Integer)
Me.Show
Me.SetFocus
End Sub

Private Sub YoconTray1_RightClick(ByVal ID As Integer)
PopupMenu mnuExibir
End Sub
