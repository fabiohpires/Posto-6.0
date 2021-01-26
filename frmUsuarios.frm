VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuários"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbGrupos 
      Caption         =   "dbGrupos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from UsuariosGrupos order by descri"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dbUsuarios 
      Caption         =   "dbUsuarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Usuarios"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtNomeNovo 
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   2640
      TabIndex        =   19
      Top             =   240
      Width           =   3375
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin MSDBCtls.DBCombo cboGrupo 
         Bindings        =   "frmUsuarios.frx":0442
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin VB.TextBox txtConfirma 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Grupo:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Confirma:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   585
      End
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmUsuarios.frx":0459
      Height          =   3570
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6297
      _Version        =   393216
      ListField       =   "Nome"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   2640
      TabIndex        =   12
      Top             =   3120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuários:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoGrupo As String

Private Sub Limpar()
txtNome.Text = ""
txtUsuario.Text = ""
txtSenha.Text = ""
txtConfirma.Text = ""
cboGrupo.Text = ""
End Sub

Private Sub Ler()
With dbUsuarios
  If .Recordset.EOF = True Then
    CodigoGrupo = -1
    txtNome.Text = ""
    txtUsuario.Text = ""
    txtSenha.Text = ""
    txtConfirma.Text = ""
    cboGrupo.Text = ""
  Else
    dbGrupos.Refresh
    If dbGrupos.Recordset.RecordCount <> 0 Then
      On Error Resume Next
      CodigoGrupo = Criptografa(.Recordset!CodigoGrupo, 223)
      dbGrupos.Recordset.FindFirst "codigogrupo=" & CodigoGrupo
      If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description
        Exit Sub
      End If
      If dbGrupos.Recordset.NoMatch = False Then
        CodigoGrupo = dbGrupos.Recordset!CodigoGrupo
        cboGrupo.Text = dbGrupos.Recordset!Descri
      Else
        CodigoGrupo = ""
        cboGrupo.Text = ""
      End If
    Else
      CodigoGrupo = ""
      cboGrupo.Text = ""
    End If
    txtNome.Text = .Recordset!Nome
    txtUsuario.Text = .Recordset!Usuario
    txtSenha.Text = Criptografa(.Recordset!Senha, 225)
    txtConfirma.Text = Criptografa(.Recordset!Confirma, 224)
  End If
End With
End Sub

Private Sub Gravar()
If txtSenha.Text <> txtConfirma.Text Then
  MsgBox "Senha não confere"
  txtSenha.SetFocus
  Exit Sub
End If
If txtUsuario.Text = "" Then
  MsgBox "Indique um nome de usuário!"
  txtUsuario.SetFocus
  Exit Sub
End If
If txtSenha.Text = "" Then
  MsgBox "Indique uma senha!"
  txtSenha.SetFocus
  Exit Sub
End If
With dbUsuarios
  If .Recordset.EOF = True Then
    MsgBox "Erro na tabela de Usuários"
    Exit Sub
  Else
    dbGrupos.Refresh
    If dbGrupos.Recordset.RecordCount <> 0 Then
      dbGrupos.Recordset.FindFirst "descri='" & cboGrupo.Text & "'"
      If dbGrupos.Recordset.NoMatch = False Then
        CodigoGrupo = dbGrupos.Recordset!CodigoGrupo
        cboGrupo.Text = dbGrupos.Recordset!Descri
      Else
        CodigoGrupo = ""
        cboGrupo.Text = ""
      End If
    Else
      CodigoGrupo = ""
      cboGrupo.Text = ""
    End If
    .Recordset.Edit
    .Recordset!CodigoGrupo = Criptografa(CodigoGrupo, 223)
    .Recordset!Nome = txtNome.Text
    .Recordset!Usuario = txtUsuario.Text
    .Recordset!Senha = Criptografa(txtSenha.Text, 225)
    .Recordset!Confirma = Criptografa(txtConfirma.Text, 224)
    .Recordset.Update
  End If
End With
End Sub

Private Sub cmdGravar_Click()
Gravar
End Sub

Private Sub cmdNovo_Click()
With dbUsuarios
  If txtNomeNovo.Text = "" Then
    MsgBox "Indique um Nome para o Usuário!"
    txtNomeNovo.SetFocus
    Exit Sub
  End If
  Limpar
  .Recordset.AddNew
  .Recordset!Nome = txtNomeNovo.Text
  .Recordset.Update
  .Refresh
  DBList1.ReFill
  .Recordset.FindFirst "nome='" & txtNomeNovo.Text & "'"
  Ler
End With
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbUsuarios
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja excluir o usuario atual?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
  DBList1.ReFill
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBList1_Click()
With dbUsuarios
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.FindFirst "nome='" & DBList1.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  Limpar
  Ler
End With
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
With dbGrupos
  .Connect = Conectar
  .DatabaseName = CaminhoUsuarios
  .Refresh
End With
With dbUsuarios
  .Connect = Conectar
  .DatabaseName = CaminhoUsuarios
  .Refresh
End With
Select Case Usuarios.Grupo.AdmUsuarios
  Case 1 'Somente leitura
    cmdGravar.Enabled = False
    cmdRemover.Enabled = False
    cmdNovo.Enabled = False
    Frame1.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub
