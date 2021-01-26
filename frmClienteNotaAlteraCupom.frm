VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmClientesNotasAlteraCupom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altera Cupom de Cliente"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "frmClienteNotaAlteraCupom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select codigocliente, nome from clientes order by nome"
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin MSDBCtls.DBCombo cboCliente 
      Bindings        =   "frmClienteNotaAlteraCupom.frx":0442
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label11 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmClientesNotasAlteraCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoCliente As Double, Nome As String

Private Sub cboCliente_LostFocus()
With dbClientes
  .Refresh
  .Recordset.FindFirst "nome='" & cboCliente.Text & "'"
  If .Recordset.NoMatch = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cmdCancelar_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
CodigoCliente = dbClientes.Recordset!CodigoCliente
Nome = dbClientes.Recordset!Nome
Me.Hide
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
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select codigocliente, nome from clientes order by nome"
  .Refresh
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbClientes
  .Refresh
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.FindFirst "codigocliente=" & txtCodigo.Text
  If .Recordset.NoMatch = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub
