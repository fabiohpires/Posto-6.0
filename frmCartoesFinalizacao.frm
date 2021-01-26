VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCartoesFinalizacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finalizações de Cartão"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmCartoesFinalizacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data qTotal 
      Caption         =   "qTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Meus documentos\Projeto For Windows\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "QFormadePgRecebidoFechamento2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmCartoesFinalizacao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Imprimir"
      Top             =   2760
      Width           =   735
   End
   Begin VB.Data QFormaDePgRecebido 
      Caption         =   "QFormaDePgRecebido"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Meus documentos\Projeto For Windows\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "QFormadePgRecebidoFechamento2"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCartoesFinalizacao.frx":0EC4
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frmCartoesFinalizacao.frx":0EE5
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label lblTotalLiquido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblTotalBruto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmCartoesFinalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprime_Click()
Dim StrTemp As String


On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0
If QFormaDePgRecebido.Recordset.EOF = True Then
  StrTemp = "Cartoes: " & frmCartoes.qCartoes.Recordset!Descri
Else
  StrTemp = "Cartoes: " & frmCartoes.qCartoes.Recordset!Descri & "  - Passados em:  " & Format(QFormaDePgRecebido.Recordset!Data, "Short Date")
End If
ImprimeGrid DBGrid1, Printer, QFormaDePgRecebido, 4, , , , 5, , "Cartoes Passados", NomePosto, StrTemp

Printer.EndDoc

NaoImprime:
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_DblClick()
'AtualizaCartoesNoDia frmCartoes.qPendentes.Recordset!CodigoCartao
With QFormaDePgRecebido
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

Private Sub Form_Load()
With QFormaDePgRecebido
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

