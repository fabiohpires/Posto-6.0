VERSION 5.00
Begin VB.Form frmFechamentoDeCaixaTransfere 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferência"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmFechamentoDeCaixaTransfere.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDestino 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cboOrigem 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cboProduto 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtQtd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransfere 
      Caption         =   "Transferir"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Quantidade:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Tanque Destino:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Tanque Origem:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFechamentoDeCaixaTransfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim dbProdutos As New ADODB.Recordset
Dim dbOrigem As New ADODB.Recordset
Dim dbDestino As New ADODB.Recordset
Dim dbTanques As New ADODB.Recordset

Private Sub FiltraTanque()
If dbProdutos.EOF = True Or dbProdutos.BOF = True Then Exit Sub

cboOrigem.Clear
cboDestino.Clear

dbTanques.Filter = "codigoproduto=" & dbProdutos!CodigoProduto
If dbTanques.EOF = True Then Exit Sub
Do While dbTanques.EOF = False
    cboOrigem.AddItem dbTanques!Tanque
    cboDestino.AddItem dbTanques!Tanque
    dbTanques.MoveNext
Loop

End Sub

Private Sub cboProduto_LostFocus()
If dbProdutos.RecordCount = 0 Then Exit Sub
dbProdutos.MoveFirst
If cboProduto.Text = "" Then Exit Sub
dbProdutos.Find "descri='" & cboProduto.Text & "'"
If dbProdutos.EOF = False Then
    txtCodigo.Text = dbProdutos!Codigo
End If
FiltraTanque
End Sub

Private Sub cmdCancelar_Click()
dbProdutos.Close
dbTanques.Close
db.Close
Unload Me
End Sub

Private Sub cmdTransfere_Click()
Dim CodigoProduto As Double, TanqueOrigem As Integer, TanqueDestino As Integer
Dim Qtd As Double

cmdTransfere.Enabled = False
cmdCancelar.Enabled = False
If txtCodigo.Text = "" Then
    MsgBox "Escolha um produto!"
    txtCodigo.SetFocus
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    Exit Sub
End If
If dbProdutos.EOF = True Then
    dbProdutos.MoveFirst
    Call txtCodigo_LostFocus
    If dbProdutos.EOF = True Then
        MsgBox "Código do Produto não encontrado!"
        txtCodigo.SetFocus
        cmdTransfere.Enabled = True
        cmdCancelar.Enabled = True
        Exit Sub
    End If
End If

If IsNumeric(cboOrigem.Text) = False Then
    MsgBox "Tanque Inválido!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboOrigem.SetFocus
    Exit Sub
End If
If IsNumeric(cboDestino.Text) = False Then
    MsgBox "Tanque Inválido!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboDestino.SetFocus
    Exit Sub
End If

TanqueOrigem = CInt(cboOrigem.Text)
TanqueDestino = CInt(cboDestino.Text)

If dbTanques.RecordCount = 0 Then
    MsgBox "Erro na tabela de tanques!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    Exit Sub
End If

If TanqueOrigem = TanqueDestino Then
    MsgBox "Tanque origem igual ao tanque destino!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboOrigem.SetFocus
    Exit Sub
End If

If IsNumeric(txtQtd.Text) = False Then
    MsgBox "Quantidade inválida!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    txtQtd.SetFocus
    Exit Sub
End If

Qtd = CDbl(txtQtd.Text)

dbTanques.MoveFirst
dbTanques.Find "tanque=" & TanqueOrigem
If dbTanques.EOF = True Then
    MsgBox "Tanque origem inválido!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboOrigem.SetFocus
    Exit Sub
End If
CodigoProduto = dbTanques!CodigoProduto

dbTanques.MoveFirst
dbTanques.Find "tanque=" & TanqueDestino
If dbTanques.EOF = True Then
    MsgBox "Tanque destino inválido!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboDestino.SetFocus
    Exit Sub
End If

If dbTanques!CodigoProduto <> CodigoProduto Then
    MsgBox "Produto do Tanque origem é diferente do tanque destino!"
    cmdTransfere.Enabled = True
    cmdCancelar.Enabled = True
    cboOrigem.SetFocus
    Exit Sub
End If

dbTanques.MoveFirst
dbTanques.Find "tanque=" & TanqueOrigem
If dbTanques.EOF = False Then
    dbTanques!Estoque = dbTanques!Estoque - Qtd
    dbTanques.Update
End If

dbTanques.MoveFirst
dbTanques.Find "tanque=" & TanqueDestino
If dbTanques.EOF = False Then
    dbTanques!Estoque = dbTanques!Estoque + Qtd
    dbTanques.Update
End If

dbProdutos.Close
dbTanques.Close
db.Close

Unload Me

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
db.Open CaminhoADO
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "select codigoproduto, codigo, descri from produtos where combustivel=-1 order by descri", db, adOpenKeyset, adLockOptimistic

With cboProduto
    .Clear
    Do While dbProdutos.EOF = False
        .AddItem dbProdutos!Descri
        dbProdutos.MoveNext
    Loop
End With

dbTanques.CursorLocation = adUseClient
dbTanques.Open "select *from tanques order by tanque", db, adOpenKeyset, adLockOptimistic


End Sub

Private Sub txtCodigo_LostFocus()
If dbProdutos.RecordCount = 0 Then Exit Sub
dbProdutos.MoveFirst
If txtCodigo.Text = "" Then Exit Sub
dbProdutos.Find "codigo='" & txtCodigo.Text & "'"
If dbProdutos.EOF = False Then
    cboProduto.Text = dbProdutos!Descri
End If
FiltraTanque

End Sub

Private Sub txtQtd_GotFocus()
txtQtd.SelStart = 0
txtQtd.SelLength = Len(txtQtd.Text)
End Sub
