VERSION 5.00
Begin VB.Form frmPermissao2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorização!"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmPermissao2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsuario 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   285
      TabIndex        =   5
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   510
   End
End
Attribute VB_Name = "frmPermissao2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ws As Workspace, db As Database
Dim Senha As Recordset

Private Sub Command1_Click()
If txtUsuario.Text = "fabio" And txtSenha.Text = "oepos21" Then
  Permissao = True
Else
  Senha.FindFirst "usuario='" & txtUsuario.Text & "'"
  Permissao = False
  If Senha.NoMatch = False Then
    If txtSenha.Text = Criptografa(Senha!Senha, 225) Then
      Permissao = True
    End If
  End If
End If
Unload Me
End Sub

Private Sub Command2_Click()
Permissao = False
Unload Me
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
Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho)
Set Senha = db.OpenRecordset("select *from usuarios order by usuario")
Permissao = False
End Sub

