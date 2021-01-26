VERSION 5.00
Begin VB.Form frmUsuarioAlteraSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Altera Senha"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "frmUsuarioAlteraSenha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAtual 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtNova 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtConfirma 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha Atual:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nova Senha:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Confirma:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   660
   End
End
Attribute VB_Name = "frmUsuarioAlteraSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Confirma As Boolean

Private Sub cmdCancela_Click()
Confirma = False
Me.Hide
End Sub

Private Sub cmdOk_Click()
If txtConfirma.Text <> txtNova.Text Then
  MsgBox "A senha nova não pode ser confirmada!", vbCritical, "Erro!"
  txtNova.SetFocus
  Exit Sub
End If
Confirma = True
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

Private Sub txtAtual_GotFocus()
With txtAtual
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtConfirma_GotFocus()
With txtConfirma
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNova_GotFocus()
With txtNova
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

