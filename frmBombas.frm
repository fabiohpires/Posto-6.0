VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadBombas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Bombas de Combustível"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3840
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from postos order by nome"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
      Caption         =   "Adodc2"
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
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6765
      TabIndex        =   14
      Top             =   2460
      Width           =   6765
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2408
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Prazo"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "PrecoVenda"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Tanque"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "ultimoNumero"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmBombas.frx":0000
         DataField       =   "CodigoProduto"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoProduto"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Bico"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmBombas.frx":0033
         DataField       =   "CodigoPosto"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   "CodigoPosto"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prazo:"
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   21
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$ Venda:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   18
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Último Número:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bomba:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2790
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select *from bicos"
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
End
Attribute VB_Name = "frmcadBombas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
MsgBox ErrorNumber & " - " & Description
fCancelDisplay = True
End Sub

Private Sub Adodc2_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
MsgBox ErrorNumber & " - " & Description
fCancelDisplay = True
End Sub

Private Sub Adodc3_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
MsgBox ErrorNumber & " - " & Description
fCancelDisplay = True
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
End Sub

Private Sub Adodc1_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If adReason = adRsnUpdate Then
  If adStatus = adStatusOK Then
    If QuerGravar = False Then
      Adodc1.Recordset.CancelUpdate
    End If
  End If
End If

End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
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
    a = .Recordset.AbsolutePosition
    .Recordset.Save
    Do While .Recordset.State = adStateExecuting
      DoEvents
    Loop
    For i = 0 To 1
      .Refresh
    Next i
    Do While .Recordset.State = adStateExecuting
      DoEvents
    Loop
    .Recordset.AbsolutePosition = a
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
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select

End Sub

Private Sub Form_Load()
With Adodc1
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With Adodc2
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from produtos where combustivel=-1 order by descri"
  .Refresh
End With
With Adodc3
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from postos order by nome"
  .Refresh
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

