VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadProdutosCategoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   Icon            =   "frmCadProdutosCategoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbProdutosCategoria 
      Height          =   330
      Left            =   2280
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "select *from ProdutosCategoria order by categoria"
      Caption         =   "dbProdutosCategoria"
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
   Begin MSAdodcLib.Adodc dbProdutosSubCategoria 
      Height          =   330
      Left            =   2280
      Top             =   960
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "select *from ProdutosSubCategoria order by descri"
      Caption         =   "dbProdutosCategoria"
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
   Begin MSDataListLib.DataCombo cboSubCategoria 
      Bindings        =   "frmCadProdutosCategoria.frx":0442
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboCategoria 
      Bindings        =   "frmCadProdutosCategoria.frx":0467
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Categoria"
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Sub Categoria:"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "frmCadProdutosCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCategoria As String, codSubCategoria As Double, Precodigo As String

Private Sub cmdCancelar_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
With dbProdutosCategoria
  .Refresh
  If .Recordset.EOF = True Then
    MsgBox "Não existe categoria cadastrada!"
    Exit Sub
  End If
  If cboCategoria.Text <> "" Then
    .Recordset.Find "categoria='" & cboCategoria.Text & "'"
    If .Recordset.EOF = False Then
      strCategoria = .Recordset!Categoria
    End If
  End If
End With
With dbProdutosSubCategoria
  .Refresh
  If .Recordset.EOF = True Then
    MsgBox "Não existe sub categoria cadastrada!"
    Exit Sub
  End If
  If cboSubCategoria.Text <> "" Then
    .Recordset.Find "descri='" & cboSubCategoria.Text & "'"
    If .Recordset.EOF = False Then
      codSubCategoria = .Recordset!codigosubcategoria
      Precodigo = .Recordset!Precodigo
    End If
  End If
End With
Me.Hide
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
strCategoria = ""
codSubCategoria = 0
With dbProdutosCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutosSubCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With
End Sub
