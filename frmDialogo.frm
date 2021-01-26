VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDialogo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferência de arquivo!"
   ClientHeight    =   1785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmDialogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation Animation1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1508
      _Version        =   393216
      FullWidth       =   385
      FullHeight      =   57
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
   End
End
Attribute VB_Name = "frmDialogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
frmCadChequeCliente.Desconectar = True
End Sub

Private Sub Form_Load()
Dim Filme As String
Filme = RetornaDiretorio(Caminho) & "FILECOPY.AVI"
Animation1.Open Filme
Animation1.Play
End Sub

