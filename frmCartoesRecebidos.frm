VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCartoesRecebidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartoes Recebidos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmCartoesRecebidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   3518
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Data dbCartoes 
      Caption         =   "dbCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cartoes"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ConciliaNova"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCartoesRecebidos.frx":0442
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "frmCartoesRecebidos.frx":045F
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmCartoesRecebidos.frx":136A
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmCartoesRecebidos.frx":1382
      TabIndex        =   1
      Top             =   3120
      Width           =   8175
   End
End
Attribute VB_Name = "frmCartoesRecebidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoConta As Double

Public Sub ExibeCartoes()
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from concilianova where codigoconta=" & CodigoConta & " and tipo='Cartão' order by data"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
  End If
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbConciliaNova_Reposition()
On Error Resume Next
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cartoes where codigosoma='" & dbConciliaNova.Recordset!nrdocumento & "' order by dataprevista"
  .Refresh
End With
End Sub

Private Sub Form_Load()
ExibeCartoes
End Sub
