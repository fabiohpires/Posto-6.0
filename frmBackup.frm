VERSION 5.00
Object = "{E88121A0-9FA9-11CF-9D9F-00AA003A3AA3}#1.0#0"; "ZlibTool.ocx"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Back Up"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ZLIBTOOLLib.ZlibTool ZlibTool1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdRestaurar 
      Caption         =   "Restaurar"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Back Up"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Local para Backup:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancela_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Dir(Caminho) = "" Then
  MsgBox "Arquivo para backup não encontrado!"
  Exit Sub
End If

On Error GoTo TrataErro
cmdCancela.Enabled = False
cmdOk.Enabled = False
cmdRestaurar.Enabled = False
With ZlibTool1
  .InputFile = Caminho
  .OutputFile = Dir1.Path & "\" & NomePosto & ".bak"
  .Visible = True
  .Compress
End With
Unload Me

Exit Sub

TrataErro:
  MsgBox Err.Number & " - " & Err.Description
  cmdCancela.Enabled = True
  cmdOk.Enabled = True
  cmdRestaurar.Enabled = True

End Sub

Private Sub cmdRestaurar_Click()
Dim Resposta As Integer
Resposta = MsgBox("Esta operação irá cancelar qualquer lançamento feito após o backup que está sendo restaurado! Tem certeza que deseja continuar?", vbYesNo)
If Resposta = vbNo Then Exit Sub
Arquivo = ""
If Right(Dir1.Path, 1) <> "\" Then
  If Dir(Dir1.Path & "\" & NomePosto & ".bak") = "" Then
    MsgBox "Arquivo de backup não encontrado!"
    Exit Sub
  End If
  Arquivo = Dir1.Path & "\" & NomePosto & ".bak"
Else
  If Dir(Dir1.Path & NomePosto & ".bak") = "" Then
    MsgBox "Arquivo de backup não encontrado!"
    Exit Sub
  End If
  Arquivo = Dir1.Path & NomePosto & ".bak"
End If
On Error GoTo TrataErro
If Arquivo = "" Then Exit Sub
cmdCancela.Enabled = False
cmdOk.Enabled = False
cmdRestaurar.Enabled = False

With ZlibTool1
  .InputFile = Arquivo
  .OutputFile = Caminho
  .Visible = True
  .Decompress
End With
Unload Me
Exit Sub
TrataErro:
  MsgBox Err.Number & " - " & Err.Description
  cmdCancela.Enabled = True
  cmdOk.Enabled = True
  cmdRestaurar.Enabled = True
  
End Sub

Private Sub Drive1_Change()
On Error GoTo TrataErro
Dir1.Path = Drive1.Drive
Exit Sub
TrataErro:
  MsgBox Err.Number & " - " & Err.Description
End Sub

