VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadTanque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Tanques"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmCadTanque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2145
      Width           =   6930
      _ExtentX        =   12224
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
      RecordSource    =   "select *from Tanques order by Tanque"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3720
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
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
      Left            =   3720
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      ScaleWidth      =   6930
      TabIndex        =   11
      Top             =   1815
      Width           =   6930
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2408
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmCadTanque.frx":0442
         DataField       =   "CodigoProduto"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoProduto"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txtEstoque 
         Alignment       =   1  'Right Justify
         DataField       =   "Estoque"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "EstoqueFisico"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Tanque"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Estoque:"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade Máxima:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanque:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmCadTanque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_Reposition()
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
End Sub

Private Sub Adodc1_Validate(Action As Integer, Save As Integer)
If Save = True Then
  If QuerGravar = False Then
    Adodc1.Recordset.CancelUpdate
  End If
End If
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  Adodc1.Recordset!codigoPosto = Adodc2.Recordset!codigoPosto
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(2).SetFocus
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
txtFields(2).SetFocus
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Adodc2.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
  With Adodc1
    A = .Recordset.AbsolutePosition
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
With Adodc2
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from postos order by nome"
  .Refresh
End With
With Adodc3
  .ConnectionString = CaminhoADO
  .Refresh
End With
If Usuarios.Nome = "Usuário Master" Then
  txtEstoque.Enabled = True
End If
Select Case Usuarios.Grupo.CadTanques
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


