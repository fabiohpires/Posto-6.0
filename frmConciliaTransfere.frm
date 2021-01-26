VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmConciliaTransfere 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferência entre contas"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   Icon            =   "frmConciliaTransfere.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ConciliaNova"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbMovimentacao 
      Caption         =   "dbMovimentacao"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from movimentacao"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbRecebe 
      Caption         =   "dbRecebe"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from PrevisaoRecebimentos"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbDestino 
      Caption         =   "dbDestino"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbOrigem 
      Caption         =   "dbOrigem"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from CompensaPendente"
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo cboDestino 
      Bindings        =   "frmConciliaTransfere.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo cboOrigem 
      Bindings        =   "frmConciliaTransfere.frx":045A
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37654
   End
   Begin VB.TextBox txtNrDoc 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Documento:"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conta a Creditar:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta a Debitar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmConciliaTransfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDestino_LostFocus()
With dbDestino
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboDestino.Text & "'"
  If .Recordset.NoMatch = False Then
    cboDestino.Text = .Recordset!Descri
  End If
End With

End Sub

Private Sub cboOrigem_LostFocus()
With dbOrigem
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboOrigem.Text & "'"
  If .Recordset.NoMatch = False Then
    cboOrigem.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
'If DateDiff("d", Date, txtData.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode Transferir com data futura!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtData.Value) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode Transferir com data anterior a 10 dias!"
'    Exit Sub
'  End If
'End If

If dbOrigem.Recordset.RecordCount = 0 Then
  MsgBox "Não existe conta cadastrada!"
  Exit Sub
End If
If dbDestino.Recordset.RecordCount = 0 Then
  MsgBox "Não existe conta cadastrada!"
  Exit Sub
End If
If cboOrigem.Text <> dbOrigem.Recordset!Descri Then
  MsgBox "Conta a ser debitada inválida!"
  cboOrigem.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um valor Correto!"
  txtValor.SetFocus
  Exit Sub
End If
With dbConcilia
  .Recordset.AddNew
  .Recordset!CodigoConta = dbOrigem.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  If dbOrigem.Recordset!temcpmf = False Then
    .Recordset!compensado = True
    .Recordset!Data = txtData.Value
  Else
    .Recordset!compensado = False
  End If
  .Recordset!Tipo = "Transferência"
  .Recordset!Codigo = 999999994
  .Recordset!Descri = "Débito para " & dbDestino.Recordset!Descri
  .Recordset!NrDocumento = txtNrDoc.Text & " "
  .Recordset!Valor = -CCur(txtValor.Text)
  .Recordset.Update
  .Refresh
  
  .Recordset.AddNew
  .Recordset!CodigoConta = dbDestino.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  If dbDestino.Recordset!temcpmf = False Then
    .Recordset!compensado = True
    .Recordset!Data = txtData.Value
  Else
    .Recordset!compensado = False
  End If
  .Recordset!Tipo = "Transferência"
  .Recordset!Codigo = 999999994
  .Recordset!Descri = "Crédito de " & dbOrigem.Recordset!Descri
  .Recordset!NrDocumento = txtNrDoc.Text & " "
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset.Update
  .Refresh
End With
With dbOrigem
  .Refresh
  .Recordset.FindFirst "descri='" & cboOrigem.Text & "'"
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo - CCur(txtValor.Text)
  .Recordset.Update
End With
With dbDestino
  CodigoConta = .Recordset!CodigoConta
  .Refresh
  .Recordset.FindFirst "codigoconta=" & CodigoConta
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo + CCur(txtValor.Text)
  .Recordset.Update
End With

With dbMovimentacao
  .Recordset.AddNew
  .Recordset!Data = Now
  .Recordset!Tipo = "Transferência"
  .Recordset!CodigoConta = dbOrigem.Recordset!CodigoConta
  .Recordset!Conta = dbOrigem.Recordset!Descri
  .Recordset!Descri = Left("Transf. Débito para " & cboDestino.Text, 50)
  .Recordset!Valor = CCur(txtValor.Text)
  .Recordset!Saldo = dbOrigem.Recordset!Saldo
  .Recordset.Update
  .Refresh
  .Refresh
End With
txtValor.Text = ""
txtNrDoc.Text = ""
cboOrigem.SetFocus

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
With dbPendencias
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbOrigem
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbDestino
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbRecebe
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbMovimentacao
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbConcilia
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
txtData.Value = Date
Select Case Usuarios.Grupo.BancoTransfere
  Case 1 'Somente leitura
    cmdOK.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub
