VERSION 5.00
Begin VB.Form frmCadChequeCliente2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes de Cheque"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmCadChequeCliente2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Adodc1 
      Caption         =   "Adodc1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from chequesClientes order by nome"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone2"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Codigo"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Endereco"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CEP"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CIC"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   120
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "RG"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Origem"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   7
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   17
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Origem2"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   8
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   18
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CNPJ"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   9
         Left            =   120
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "IE"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   10
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Celular:"
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Bairro:"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nome:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Endereco:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "CEP:"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "CIC:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "RG:"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Emissão:"
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "CNPJ:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "IE:"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   21
         Top             =   2640
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCadChequeCliente2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Adodc1.Recordset.Cancel
Unload Me
End Sub

Private Sub cmdOk_Click()
Adodc1.Recordset.Update
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
Frame1.Enabled = True
With Adodc1
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  .Recordset.AddNew
End With

End Sub
