VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFechamentoConfirmaEntrada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirma Entrada de Combustível Pendente"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "frmFechamentoConfirmaEntrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Data dbTanques 
      Caption         =   "dbTanques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Tanques order by tanque"
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data dbProdutosNotasCorpo 
      Caption         =   "dbProdutosNotasCorpo"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProdutosNotasCorpo"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data qProdutosNotas 
      Caption         =   "qProdutosNotas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from QProdutosNotas order by datanota, nrnota"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmFechamentoConfirmaEntrada.frx":0442
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "frmFechamentoConfirmaEntrada.frx":045F
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frmFechamentoConfirmaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoCaixa As Double
Dim strOrdem As String

Private Sub cmdConfirmar_Click()
Dim db As New ADODB.Connection
Dim dbDifComb As New ADODB.Recordset
Dim dbPostos As New ADODB.Recordset

Dim Resposta  As Integer
Resposta = MsgBox("Deseja confirmar a entrada no tanque que está selecionada?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub

db.Open CaminhoADO
dbPostos.CursorLocation = adUseClient
dbPostos.Open "select *from postos", db, adOpenForwardOnly, adLockReadOnly

dbDifComb.CursorLocation = adUseClient
dbDifComb.Open "select *from diferencacombustivel where codigofechamento=" & CodigoCaixa, db, adOpenKeyset, adLockOptimistic


With dbTanques
  .Recordset.FindFirst "tanque=" & qProdutosNotas.Recordset!Tanque
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de tanques! Tanque não encontrado"
    Exit Sub
  End If
End With
With dbProdutosNotasCorpo
  .Recordset.FindFirst "codigoprodutonotacorpo=" & qProdutosNotas.Recordset!codigoprodutonotacorpo
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de notas! Corpo da nota não encontrado!"
    Exit Sub
  End If
End With
With dbTanques
  If MedetanqueAntes = True Then
    If .Recordset!Estoque + qProdutosNotas.Recordset!Quantidade > .Recordset!estoquefisico + 1000 Then
      MsgBox "A capacidade do tanque é menor que o estoque + entrada atual!"
      Exit Sub
    End If
  Else
    dbDifComb.MoveFirst
    dbDifComb.Find "tanquenr=" & qProdutosNotas.Recordset!Tanque
    If .Recordset!Estoque - dbDifComb!Vendido + qProdutosNotas.Recordset!Quantidade > .Recordset!estoquefisico + 1000 Then
      MsgBox "A capacidade do tanque é menor que o estoque + entrada atual!"
      Exit Sub
    End If
  End If
  .Recordset.Edit
  .Recordset!Estoque = .Recordset!Estoque + qProdutosNotas.Recordset!Quantidade
  .Recordset.Update
End With
With dbProdutosNotasCorpo
  .Recordset.Edit
  .Recordset!Aguardando = 0
  .Recordset!CodigoCaixa = CodigoCaixa
  .Recordset.Update
End With
qProdutosNotas.Refresh

dbDifComb.Close
dbPostos.Close
db.Close

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", NrNota" Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc, NrNota"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", NrNota"
End If

With qProdutosNotas
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from QProdutosNotas where aguardando=-1" & strOrdem
  .Refresh
End With

End Sub

Private Sub Form_Load()
strOrdem = " order by DataNota, NrNota"
With dbTanques
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With qProdutosNotas
  .DatabaseName = Caminho
  .Connect = Conectar
  .RecordSource = "select *from QProdutosNotas where aguardando=-1" & strOrdem
  .Refresh
End With
With dbProdutosNotasCorpo
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
End Sub
