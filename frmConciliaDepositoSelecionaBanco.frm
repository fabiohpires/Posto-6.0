VERSION 5.00
Begin VB.Form frmConciliaDepositoSelecionaBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleciona Lay-out"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   Icon            =   "frmConciliaDepositoSelecionaBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox lstBanco 
      Height          =   1425
      ItemData        =   "frmConciliaDepositoSelecionaBanco.frx":0442
      Left            =   120
      List            =   "frmConciliaDepositoSelecionaBanco.frx":0444
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Lay-out do banco:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmConciliaDepositoSelecionaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Banco As String

Private Sub cmdOk_Click()
For i = 0 To lstBanco.ListCount - 1
  If lstBanco.Selected(i) = True Then
    Banco = lstBanco.List(i)
    Exit For
  End If
Next i
Me.Hide
End Sub
