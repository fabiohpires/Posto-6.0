VERSION 5.00
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form frmConfigura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurações"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "frmConfigura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNomePosto 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox txtCodigoPosto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtDestino 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
   End
   Begin VB.TextBox txtOrigem 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Gravar Configurações"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox chkIniciarComWindows 
      Caption         =   "Iniciar com o Windows"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtMsde 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Text            =   "SQLOLEDB.1"
      Top             =   1680
      Width           =   2055
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   4800
      Top             =   1440
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin VB.Label Label2 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Posto:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Local para salvar as informações:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Servidor do banco de dados:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Driver de MSDE:"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkIniciarComWindows_Click()
If chkIniciarComWindows = vbChecked Then
  Ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador")
  If Ret = "" Then
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador", App.Path & "\Integrador.exe"
  End If
Else
  DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador"
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
If txtCodigoPosto.Text = "" Then
  MsgBox "Informe o código do posto!"
  txtCodigoPosto.SetFocus
  Exit Sub
End If
If txtNomePosto.Text = "" Then
  MsgBox "Informe o nome do posto!"
  txtNomePosto.SetFocus
  Exit Sub
End If
If txtOrigem.Text = "" Then
  MsgBox "Caminho inválido!"
  txtOrigem.SetFocus
  Exit Sub
End If
If txtDestino.Text = "" Then
  MsgBox "Destino inválido!"
  txtDestino.SetFocus
  Exit Sub
End If
'If Mid(txtDestino.Text, Len(txtDestino.Text)) <> "\" Then
'  txtDestino.Text = txtDestino.Text & "\"
'End If
If txtMsde.Text = "" Then
  MsgBox "MSDE inválido!"
  txtMsde.SetFocus
  Exit Sub
End If

If frmIntegrador.GravarConfiguracoes(txtOrigem.Text, txtDestino.Text, frmIntegrador.txtDataCaixa.Value, "01", txtMsde.Text, txtCodigoPosto.Text, txtNomePosto.Text) = True Then
  Unload Me
End If

End Sub

Private Sub Form_Load()
txtOrigem.Text = frmIntegrador.Caminho
txtDestino.Text = frmIntegrador.Destino
txtMsde.Text = frmIntegrador.strMSDE
txtCodigoPosto.Text = frmIntegrador.CodigoPosto
txtNomePosto.Text = frmIntegrador.NomePosto
End Sub
