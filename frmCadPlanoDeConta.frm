VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadPlanoDeConta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plano de Conta"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   Icon            =   "frmCadPlanoDeConta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   7095
      Begin VB.TextBox txtFields 
         DataField       =   "COD_ITEM"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   120
         MaxLength       =   50
         TabIndex        =   23
         ToolTipText     =   "Código do item relacionado a esta conta contábil"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "COD_CTA_REF"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   21
         ToolTipText     =   "CNPJ do estabelecimento, no caso da conta informada no campo COD_CTA ser específica de um estabelecimento"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "COD_CTA_REF"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Código da conta correlacionada no Plano de Contas Referenciado, publicado pela RFB"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "NIVEL"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   50
         TabIndex        =   17
         ToolTipText     =   "Nível da conta analítica/grupo de contas"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "IND_CTA"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmCadPlanoDeConta.frx":0442
         Left            =   3120
         List            =   "frmCadPlanoDeConta.frx":044C
         TabIndex        =   16
         ToolTipText     =   "Indicador do tipo de conta"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "COD_NAT_CC"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmCadPlanoDeConta.frx":048B
         Left            =   120
         List            =   "frmCadPlanoDeConta.frx":04A1
         TabIndex        =   14
         ToolTipText     =   "Código da natureza da conta/grupo de contas"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         DataField       =   "NOME_CTA"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Nome da conta analítica/grupo de contas. "
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "COD_CTA"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Código da conta analítica/grupo de contas. "
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código do item:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ do Estabelecimento:"
         Height          =   195
         Index           =   6
         Left            =   3720
         TabIndex        =   22
         Top             =   1440
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da conta correlacionada:"
         Height          =   195
         Index           =   5
         Left            =   1320
         TabIndex        =   20
         Top             =   1440
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nível:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Conta:"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   15
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da conta/grupo de contas:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7410
      TabIndex        =   1
      Top             =   6045
      Width           =   7410
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2408
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6375
      Width           =   7410
      _ExtentX        =   13070
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
      RecordSource    =   "select *from PlanosDeConta order by cod_cta"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadPlanoDeConta.frx":0535
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
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
         DataField       =   "COD_CTA"
         Caption         =   "Código"
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
         DataField       =   "NOME_CTA"
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
      BeginProperty Column02 
         DataField       =   "DataAltera"
         Caption         =   "Data Altera"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   1649,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadPlanoDeConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Adodc1_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  Adodc1.Recordset!DataAltera = Now
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

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    A = .Recordset.AbsolutePosition
    Adodc1.Recordset!DataAltera = Now
    .Recordset.Update
    .Recordset.AbsolutePosition = A
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

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select

End Sub

Private Sub Form_Load()
With Adodc1
  .ConnectionString = CaminhoADO
  .Refresh
End With


Select Case Usuarios.Grupo.CadConta
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


