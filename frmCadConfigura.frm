VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCadConfigura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurações"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmCadConfigura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      DataField       =   "ClientesNotaPlano"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   2040
      TabIndex        =   35
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      DataField       =   "MicroCreditoPlano"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   2040
      TabIndex        =   34
      Top             =   4320
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      DataField       =   "ServidorPista"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   2040
      TabIndex        =   29
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      DataField       =   "LocalCustodia"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   2040
      TabIndex        =   27
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox txtPorta 
      Alignment       =   1  'Right Justify
      DataField       =   "Porta"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   6360
      TabIndex        =   25
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtFTP 
      DataField       =   "ftp"
      DataSource      =   "dbConfig2"
      Height          =   285
      Left            =   2040
      TabIndex        =   23
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtCPMF 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   31
      Top             =   4680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dbConfig2 
      Height          =   375
      Left            =   2640
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select *from config"
      Caption         =   "dbConfig2"
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
      Caption         =   " Modem "
      Height          =   2295
      Left            =   3600
      TabIndex        =   33
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cboCom2 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   "Sem"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboBaud2 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":0442
         Left            =   1920
         List            =   "frmCadConfigura.frx":046D
         TabIndex        =   13
         Text            =   "9600"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboParidade2 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":04C5
         Left            =   1920
         List            =   "frmCadConfigura.frx":04D8
         TabIndex        =   15
         Text            =   "n"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cboDataBit2 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":04EB
         Left            =   1920
         List            =   "frmCadConfigura.frx":04FE
         TabIndex        =   17
         Text            =   "8"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cboStopBit2 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":0511
         Left            =   1920
         List            =   "frmCadConfigura.frx":051E
         TabIndex        =   19
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Porta de Comunicação:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Velocidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Paridade:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Data bit:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "bit de parada:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Leitora de Código-de-Barras "
      Height          =   2295
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cboStopBit 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":052D
         Left            =   1920
         List            =   "frmCadConfigura.frx":053A
         TabIndex        =   9
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox cboDataBit 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":0549
         Left            =   1920
         List            =   "frmCadConfigura.frx":055C
         TabIndex        =   7
         Text            =   "8"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cboParidade 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":056F
         Left            =   1920
         List            =   "frmCadConfigura.frx":0582
         TabIndex        =   5
         Text            =   "n"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cboBaud 
         Height          =   315
         ItemData        =   "frmCadConfigura.frx":0595
         Left            =   1920
         List            =   "frmCadConfigura.frx":05C0
         TabIndex        =   3
         Text            =   "9600"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboCom 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Text            =   "Sem"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "bit de parada:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Data bit:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1680
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Paridade:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Velocidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Porta de Comunicação:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta Notas a Cobra:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta Microcrédito:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Servidor Microsffer Pista:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Local para custódia:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Codigo Posto:"
      Height          =   195
      Left            =   5280
      TabIndex        =   24
      Top             =   2880
      Width           =   990
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Servidor para importação:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPMF:"
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "frmCadConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancela_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If IsNumeric(txtCPMF.Text) = False Then
  MsgBox "Taxa de CPMF inválida!"
  txtCPMF.SetFocus
  Exit Sub
End If
SaveSetting App.EXEName, "Base", "CPMF", txtCPMF.Text
SaveSetting App.EXEName, "Base", "COM", cboCom.Text
SaveSetting App.EXEName, "Base", "Baud", cboBaud.Text
SaveSetting App.EXEName, "Base", "Paridade", cboParidade.Text
SaveSetting App.EXEName, "Base", "DataBit", cboDataBit.Text
SaveSetting App.EXEName, "Base", "StopBit", cboStopBit.Text

SaveSetting App.EXEName, "Base", "COM2", cboCom2.Text
SaveSetting App.EXEName, "Base", "Baud2", cboBaud2.Text
SaveSetting App.EXEName, "Base", "Paridade2", cboParidade2.Text
SaveSetting App.EXEName, "Base", "DataBit2", cboDataBit2.Text
SaveSetting App.EXEName, "Base", "StopBit2", cboStopBit2.Text
dbConfig2.Recordset.Update
CPMF = CDbl(txtCPMF) / 100
Unload Me
End Sub





Private Sub Form_Load()

txtCPMF.Text = GetSetting(App.EXEName, "Base", "CPMF", "0,38")
With dbConfig2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With cboCom
  .Clear
  .AddItem "Sem"
  For i = 1 To 16
    .AddItem Trim("COM" & Trim(Str(i)))
  Next i
  .Text = GetSetting(App.EXEName, "Base", "COM", "Sem")
End With
cboBaud.Text = GetSetting(App.EXEName, "Base", "Baud", "9600")
cboParidade.Text = GetSetting(App.EXEName, "Base", "Paridade", "n")
cboDataBit.Text = GetSetting(App.EXEName, "Base", "DataBit", "8")
cboStopBit.Text = GetSetting(App.EXEName, "Base", "StopBit", "1")
With cboCom2
  .Clear
  .AddItem "Sem"
  For i = 1 To 16
    .AddItem Trim("COM" & Trim(Str(i)))
  Next i
  .Text = GetSetting(App.EXEName, "Base", "COM2", "Sem")
End With
cboBaud2.Text = GetSetting(App.EXEName, "Base", "Baud2", "9600")
cboParidade2.Text = GetSetting(App.EXEName, "Base", "Paridade2", "n")
cboDataBit2.Text = GetSetting(App.EXEName, "Base", "DataBit2", "8")
cboStopBit2.Text = GetSetting(App.EXEName, "Base", "StopBit2", "1")
Select Case Usuarios.Grupo.CadConfiguracao
  Case 1 'Somente leitura
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtCPMF.Enabled = False
  Case 2 'Liberado
    
End Select
End Sub

Private Sub txtCPMF_GotFocus()
With txtCPMF
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCPMF_LostFocus()
With txtCPMF
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = True Then
    .Text = Format(.Text, "#0,00")
  End If
End With
End Sub
