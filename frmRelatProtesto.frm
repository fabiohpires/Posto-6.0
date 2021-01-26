VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmRelatProtesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprime Protesto"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   53
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5880
      TabIndex        =   52
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimirAcima 
      Caption         =   "Imprimir Acima"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimirTudo 
      Caption         =   "Imprimir Tudo"
      Height          =   375
      Left            =   1440
      TabIndex        =   46
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Data dbPostos 
      Caption         =   "dbPostos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "J:\rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Postos"
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "J:\rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ChequesClientes order by nome"
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdPreencher 
      Caption         =   "Preencher"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo cboNome 
      Bindings        =   "frmRelatProtesto.frx":0000
      Height          =   315
      Left            =   2160
      TabIndex        =   15
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   3000
      TabIndex        =   51
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtRg 
      Height          =   285
      Left            =   600
      TabIndex        =   49
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtTelefone2 
      Height          =   285
      Left            =   6360
      TabIndex        =   44
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtEstado2 
      Height          =   285
      Left            =   4560
      TabIndex        =   42
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtCidade2 
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox txtCep2 
      Height          =   285
      Left            =   6360
      TabIndex        =   38
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtEndereco2 
      Height          =   285
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   6135
   End
   Begin VB.TextBox txtDocDevedor2 
      Height          =   285
      Left            =   6000
      TabIndex        =   34
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtDevedor2 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   3360
      Width           =   5775
   End
   Begin VB.TextBox txtTelefone1 
      Height          =   285
      Left            =   6360
      TabIndex        =   30
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtEstado1 
      Height          =   285
      Left            =   4560
      TabIndex        =   28
      Text            =   "SP"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtCidade1 
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Text            =   "São Paulo"
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtCep1 
      Height          =   285
      Left            =   6360
      TabIndex        =   24
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtEndereco1 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   6135
   End
   Begin VB.TextBox txtDocDevedor1 
      Height          =   285
      Left            =   6000
      TabIndex        =   20
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtDevedor1 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   5775
   End
   Begin VB.TextBox txtSaldo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6720
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtVencimento 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtEmissao 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtNTitulo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtTdoc 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "CH"
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7920
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7920
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label25 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Dados de quem vai entregar no cartório:"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label22 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   2400
      TabIndex        =   50
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "R.G.:"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   6360
      TabIndex        =   43
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   4560
      TabIndex        =   41
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   6360
      TabIndex        =   37
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Doc do Devedor:"
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Devedor 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   6360
      TabIndex        =   29
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   6360
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Doc do Devedor:"
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Devedor 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Saldo:"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Valor:"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Vencimento:"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Emissão:"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nº do Título:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "T.Doc:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatProtesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboNome_LostFocus()
With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboNome.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboNome.Text & "'"
  If .Recordset.NoMatch = False Then
    cboNome.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!codigochequecliente
  End If
End With
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then GoTo naoImprime
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  
  Printer.ScaleMode = vbMillimeters
  
  StrTemp = txtRg.Text
  Printer.CurrentX = 8
  Printer.CurrentY = 113
  Printer.Print StrTemp
  
  StrTemp = txtNome.Text
  Printer.CurrentX = 10
  Printer.CurrentY = 119
  Printer.Print StrTemp
  
  Printer.EndDoc
naoImprime:

End Sub

Private Sub cmdImprimirAcima_Click()
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then GoTo naoImprime
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  
  Printer.ScaleMode = vbMillimeters
  
  With dbPostos
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 0
    Printer.CurrentY = 28
    Printer.Print StrTemp
    
    StrTemp = .Recordset!telefone
    Printer.CurrentX = 147
    Printer.CurrentY = 28
    Printer.Print StrTemp
    
    StrTemp = .Recordset!Endereco
    Printer.CurrentX = 0
    Printer.CurrentY = 36
    Printer.Print StrTemp
    
    StrTemp = .Recordset!complemento
    Printer.CurrentX = 139
    Printer.CurrentY = 36
    Printer.Print StrTemp
    
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 0
    Printer.CurrentY = 44
    Printer.Print StrTemp
  End With
  
  StrTemp = txtTdoc.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtNTitulo.Text
  Printer.CurrentX = 11
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtEmissao.Text
  Printer.CurrentX = 41
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtVencimento.Text
  Printer.CurrentX = 71
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtValor.Text
  Printer.CurrentX = 101
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtSaldo.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtDevedor1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 61
  Printer.Print StrTemp
  
  StrTemp = txtDocDevedor1.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 61
  Printer.Print StrTemp
  
  StrTemp = txtEndereco1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 69
  Printer.Print StrTemp
  
  StrTemp = txtCep1.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 69
  Printer.Print StrTemp
  
  StrTemp = txtCidade1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtEstado1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtCidade1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtEstado1.Text
  Printer.CurrentX = 82
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtTelefone1.Text
  Printer.CurrentX = 112
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtDevedor2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 85
  Printer.Print StrTemp
  
  StrTemp = txtDocDevedor2.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 85
  Printer.Print StrTemp
  
  StrTemp = txtEndereco2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 94
  Printer.Print StrTemp
  
  StrTemp = txtCep2.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 94
  Printer.Print StrTemp
  
  StrTemp = txtCidade2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = txtEstado2.Text
  Printer.CurrentX = 82
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = txtTelefone2.Text
  Printer.CurrentX = 112
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = "CNPJ: " & dbPostos.Recordset!cnpj
  Printer.CurrentX = 0
  Printer.CurrentY = 140
  Printer.Print StrTemp
  
  Printer.EndDoc
naoImprime:

End Sub

Private Sub cmdImprimirTudo_Click()
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then GoTo naoImprime
  On Error GoTo 0
  
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  
  Printer.ScaleMode = vbMillimeters
  
  With dbPostos
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 0
    Printer.CurrentY = 28
    Printer.Print StrTemp
    
    StrTemp = .Recordset!telefone
    Printer.CurrentX = 147
    Printer.CurrentY = 28
    Printer.Print StrTemp
    
    StrTemp = .Recordset!Endereco
    Printer.CurrentX = 0
    Printer.CurrentY = 36
    Printer.Print StrTemp
    
    StrTemp = .Recordset!complemento
    Printer.CurrentX = 139
    Printer.CurrentY = 36
    Printer.Print StrTemp
    
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 0
    Printer.CurrentY = 44
    Printer.Print StrTemp
  End With
  
  StrTemp = txtTdoc.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtNTitulo.Text
  Printer.CurrentX = 11
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtEmissao.Text
  Printer.CurrentX = 41
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtVencimento.Text
  Printer.CurrentX = 71
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtValor.Text
  Printer.CurrentX = 101
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtSaldo.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 52
  Printer.Print StrTemp
  
  StrTemp = txtDevedor1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 61
  Printer.Print StrTemp
  
  StrTemp = txtDocDevedor1.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 61
  Printer.Print StrTemp
  
  StrTemp = txtEndereco1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 69
  Printer.Print StrTemp
  
  StrTemp = txtCep1.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 69
  Printer.Print StrTemp
  
  StrTemp = txtCidade1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtEstado1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtCidade1.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtEstado1.Text
  Printer.CurrentX = 82
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtTelefone1.Text
  Printer.CurrentX = 112
  Printer.CurrentY = 77
  Printer.Print StrTemp
  
  StrTemp = txtDevedor2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 85
  Printer.Print StrTemp
  
  StrTemp = txtDocDevedor2.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 85
  Printer.Print StrTemp
  
  StrTemp = txtEndereco2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 94
  Printer.Print StrTemp
  
  StrTemp = txtCep2.Text
  Printer.CurrentX = 140
  Printer.CurrentY = 94
  Printer.Print StrTemp
  
  StrTemp = txtCidade2.Text
  Printer.CurrentX = 0
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = txtEstado2.Text
  Printer.CurrentX = 82
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = txtTelefone2.Text
  Printer.CurrentX = 112
  Printer.CurrentY = 102
  Printer.Print StrTemp
  
  StrTemp = txtRg.Text
  Printer.CurrentX = 8
  Printer.CurrentY = 113
  Printer.Print StrTemp
  
  StrTemp = txtNome.Text
  Printer.CurrentX = 10
  Printer.CurrentY = 119
  Printer.Print StrTemp
  
  StrTemp = "CNPJ: " & dbPostos.Recordset!cnpj
  Printer.CurrentX = 0
  Printer.CurrentY = 140
  Printer.Print StrTemp
  
  Printer.EndDoc
naoImprime:

End Sub

Private Sub cmdPreencher_Click()
With dbClientes
  If .Recordset.EOF = True Then Exit Sub
  If cboNome.Text <> .Recordset!Nome Then
    MsgBox "Cliente não encontrado!"
    Exit Sub
  End If
  
  txtDevedor1.Text = .Recordset!Nome
  If IsNull(.Recordset!cic) = False Then txtDocDevedor1.Text = .Recordset!cic
  If IsNull(.Recordset!cnpj) = False Then txtDocDevedor1.Text = .Recordset!cnpj
  txtCep1.Text = .Recordset!CEP
  txtEndereco1.Text = .Recordset!Endereco
  txtTelefone1.Text = .Recordset!telefone
End With
End Sub

Private Sub cmdSair_Click()
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
txtRg.Text = GetSetting(App.EXEName, "Protesto", "RG", "")
txtNome.Text = GetSetting(App.EXEName, "Protesto", "Nome", "")
With dbClientes
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
With dbPostos
  .DatabaseName = Caminho
  .Connect = Conectar
  .Refresh
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting App.EXEName, "Protesto", "RG", txtRg.Text
SaveSetting App.EXEName, "Protesto", "Nome", txtNome.Text
End Sub

Private Sub Form_Terminate()
SaveSetting App.EXEName, "Protesto", "RG", txtRg.Text
SaveSetting App.EXEName, "Protesto", "Nome", txtNome.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.EXEName, "Protesto", "RG", txtRg.Text
SaveSetting App.EXEName, "Protesto", "Nome", txtNome.Text
End Sub

Private Sub txtCep1_GotFocus()
With txtCep1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCep2_GotFocus()
With txtCep2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCidade1_GotFocus()
With txtCidade1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCidade2_GotFocus()
With txtCidade2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If txtCodigo.Text = "" Then Exit Sub
  .Recordset.FindFirst "codigochequecliente=" & txtCodigo.Text
  If .Recordset.NoMatch = False Then
    cboNome.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!codigochequecliente
  End If
End With
End Sub

Private Sub txtDevedor1_GotFocus()
With txtDevedor1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDevedor2_GotFocus()
With txtDevedor2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDocDevedor1_GotFocus()
With txtDocDevedor1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDocDevedor2_GotFocus()
With txtDocDevedor2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEmissao_GotFocus()
With txtEmissao
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEmissao_LostFocus()
With txtEmissao
  If .Text = "" Then Exit Sub
  If IsDate(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Short Date")
End With
End Sub

Private Sub txtEndereco1_GotFocus()
With txtEndereco1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEndereco2_GotFocus()
With txtEndereco2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEstado1_GotFocus()
With txtEstado1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtEstado2_GotFocus()
With txtEstado2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNome_GotFocus()
With txtRg
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNTitulo_GotFocus()
With txtNTitulo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRg_GotFocus()
With txtRg
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtSaldo_GotFocus()
With txtSaldo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtSaldo_LostFocus()
With txtSaldo
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtTdoc_GotFocus()
With txtTdoc
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTelefone1_GotFocus()
With txtTelefone1
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTelefone2_GotFocus()
With txtTelefone2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtVencimento_GotFocus()
With txtVencimento
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVencimento_LostFocus()
With txtVencimento
  If .Text = "" Then Exit Sub
  If IsDate(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Short Date")
End With
End Sub
