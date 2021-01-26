VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadPosto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Posto de Combustível"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmCadPosto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadPosto.frx":0442
      Height          =   1335
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2355
      _Version        =   393216
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Nome"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   6105,26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5730
      Width           =   6975
      _ExtentX        =   12303
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
      RecordSource    =   "select *from postos order by nome"
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   36
      Top             =   1440
      Width           =   6735
      Begin VB.CheckBox Check2 
         Caption         =   "A comissão é acumulativa"
         DataField       =   "ComissaoAcumulativa"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Bairro"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   12
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CCM"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000/0000-00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   120
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Municipio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000/0000-00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   10
         Left            =   120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Estado"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000"".""000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   9
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "DataJucesp"
         DataSource      =   "Adodc1"
         Height          =   300
         Left            =   3720
         TabIndex        =   27
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72941569
         CurrentDate     =   39409
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Jucesp"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000"".""000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   8
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "O tanque é medido na abertura do caixa"
         DataField       =   "MedeTanqueAntes"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "IE"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000"".""000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   7
         Left            =   4680
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Contato"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fax"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(000)#000-0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CNPJ"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"".""000"".""000/0000-00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(000)#000-0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Endereco"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Complemento"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Index           =   13
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "CCM:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   390
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   11
         Left            =   2160
         TabIndex        =   16
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Município:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Data Jucesp:"
         Height          =   195
         Index           =   9
         Left            =   3720
         TabIndex        =   26
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Jucesp:"
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   24
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   10
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "IE:"
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   12
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   20
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   18
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   465
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
      ScaleWidth      =   6975
      TabIndex        =   35
      Top             =   5400
      Width           =   6975
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   31
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   33
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   34
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCadPosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(0).SetFocus
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
    .Recordset.Update
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
Select Case Usuarios.Grupo.CadPostos
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub




