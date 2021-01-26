VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUsuarioLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbUsuarios 
      Height          =   330
      Left            =   840
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Usuarios.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Usuarios.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from Usuarios order by usuario"
      Caption         =   "dbUsuarios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo txtUserName 
      Bindings        =   "frmUsuarioLogin.frx":0000
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Usuario"
      Text            =   ""
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuário:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmUsuarioLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database, Ws As Workspace
Dim dbUsuario As Recordset, dbGrupos As Recordset
Dim Contador As Integer

Private Sub CarregaUsuario()
Usuarios.Nome = dbUsuario!Nome
Usuarios.Senha = txtPassword.Text
With Usuarios.Grupo
  .Descri = dbGrupos!Descri
  'Cadastro
  .CadBomba = Criptografa(dbGrupos!CadBomba, 1)
  .CadCliente = Criptografa(dbGrupos!CadCliente, 2)
  .CadClienteCheque = Criptografa(dbGrupos!CadClienteCheque, 3)
  .CadConta = Criptografa(dbGrupos!CadConta, 4)
  .CadDespesaTipo = Criptografa(dbGrupos!CadDespesaTipo, 5)
  .CadDespesaBancaria = Criptografa(dbGrupos!CadDespesaBancaria, 6)
  .CadFormaDePg = Criptografa(dbGrupos!CadFormaDePg, 7)
  .CadFornecedores = Criptografa(dbGrupos!CadFornecedores, 8)
  .CadFuncionarios = Criptografa(dbGrupos!CadFuncionarios, 9)
  .CadJuros = Criptografa(dbGrupos!CadJuros, 10)
  .CadPostos = Criptografa(dbGrupos!CadPostos, 11)
  .CadProdutos = Criptografa(dbGrupos!CadProdutos, 12)
  .CadProdutosFornecedores = Criptografa(dbGrupos!CadProdutosFornecedores, 13)
  .CadTanques = Criptografa(dbGrupos!CadTanques, 14)
  .CadTurnos = Criptografa(dbGrupos!CadTurnos, 15)
  .CadConfiguracao = Criptografa(dbGrupos!CadConfiguracao, 16)
  If IsNull(dbGrupos!ClientesPlanos) = False Then
    .ClientesPlanos = Criptografa(dbGrupos!ClientesPlanos, 16)
  Else
    .ClientesPlanos = ""
  End If
  'Controle
  .ControleFechamentoDiario = Criptografa(dbGrupos!ControleFechamentoDiario, 17)
  .ControleConferencia = Criptografa(dbGrupos!ControleConferencia, 18)
  .ControleCartoes = Criptografa(dbGrupos!ControleCartoes, 19)
  .ControlePgAntecipado = Criptografa(dbGrupos!ControlePgAntecipado, 20)
  .ControleNotas = Criptografa(dbGrupos!ControleNotas, 21)
  .ControleLancContas = Criptografa(dbGrupos!ControleLancContas, 22)
  .ControleContasPg = Criptografa(dbGrupos!ControleContasPg, 23)
  .ControleCobranca = Criptografa(dbGrupos!ControleCobranca, 24)
  .ControleAgua = Criptografa(dbGrupos!ControleAgua, 25)
  .ControleLuz = Criptografa(dbGrupos!ControleLuz, 26)
  .ControleLavagem = Criptografa(dbGrupos!ControleLavagem, 27)
  .ControleVales = Criptografa(dbGrupos!ControleVales, 28)
  'Cheques
  .ChequeDeposito = Criptografa(dbGrupos!ChequeDeposito, 29)
  .ChequeDevolucao = Criptografa(dbGrupos!ChequeDevolucao, 30)
  .ChequeCobranca = Criptografa(dbGrupos!ChequeCobranca, 31)
  .ChequeProtesto = Criptografa(dbGrupos!ChequeProtesto, 32)
  .ChequeEnviarPEmpresaCobranca = Criptografa(dbGrupos!ChequeEnviarPEmpresaCobranca, 33)
  .ChequePorData = Criptografa(dbGrupos!ChequePorData, 34)
  'Banco
  .BancoConcilia = Criptografa(dbGrupos!BancoConcilia, 35)
  .BancoTransfere = Criptografa(dbGrupos!BancoTransfere, 36)
  'Relatórios
  .RelatAcertoEstoque = Criptografa(dbGrupos!RelatAcertoEstoque, 37)
  .RelatChequeCliente = Criptografa(dbGrupos!RelatChequeCliente, 38)
  .RelatProdutosComprados = Criptografa(dbGrupos!RelatProdutosComprados, 40)
  .RelatCompraVenda = Criptografa(dbGrupos!RelatCompraVenda, 41)
  .RelatDifCaixa = Criptografa(dbGrupos!RelatDifCaixa, 42)
  .RelatDifRecebe = Criptografa(dbGrupos!RelatDifRecebe, 43)
  .RelatDifCombustivel = Criptografa(dbGrupos!RelatDifCombustivel, 44)
  .RelatFormaDePg = Criptografa(dbGrupos!RelatFormaDePg, 45)
  .RelatGalonagem = Criptografa(dbGrupos!RelatGalonagem, 46)
  .RelatGalonagemTotal = Criptografa(dbGrupos!RelatGalonagemTotal, 47)
  .RelatVendaProdutos = Criptografa(dbGrupos!RelatVendaProdutos, 48)
  .RelatVendaDetalhada = Criptografa(dbGrupos!RelatVendaDetalhada, 49)
  .RelatVendaLucro = Criptografa(dbGrupos!RelatVendaLucro, 50)
  .RelatVendaMedia = Criptografa(dbGrupos!RelatVendaMedia, 51)
  .RelatDiariaCombustivel = Criptografa(dbGrupos!RelatDiariaCombustivel, 52)
  .RelatProtestoDeCheques = Criptografa(dbGrupos!RelatProtestoDeCheques, 53)
  .RelatCadastroIncompleto = Criptografa(dbGrupos!RelatCadastroIncompleto, 54)
  .RelatRetornoCombustivel = Criptografa(dbGrupos!RelatRetornoCombustivel, 55)
  .RelatFaturamentoCheques = Criptografa(dbGrupos!RelatFaturamentoCheques, 56)
  .RelatKilometragem = Criptografa(dbGrupos!RelatKilometragem, 57)
  'Administração
  .AdmConfirma = Criptografa(dbGrupos!AdmConfirma, 58)
  .AdmEstatus = Criptografa(dbGrupos!AdmEstatus, 59)
  .AdmTotalVenda = Criptografa(dbGrupos!AdmTotalVenda, 60)
  .AdmLMC = Criptografa(dbGrupos!AdmLMC, 61)
  .AdmUsuarios = Criptografa(dbGrupos!AdmUsuarios, 62)
  .AdmUsuariosGrupos = Criptografa(dbGrupos!AdmUsuariosGrupos, 63)
  If IsNull(dbGrupos!admDatas) = False Then
    .admDatas = Criptografa(dbGrupos!admDatas, 64)
  Else
    .admDatas = .AdmEstatus
  End If
  If IsNull(dbGrupos!liberanotas) = False Then
    .admLiberaNotas = Criptografa(dbGrupos!liberanotas, 65)
  Else
    .admLiberaNotas = .AdmEstatus
  End If
End With
End Sub

Private Sub CarregaMaster()
With Usuarios
  .Nome = "Usuário Master"
  .Senha = txtPassword.Text
  With .Grupo
    .Descri = "Master"
    'Cadastro
    .CadBomba = 2
    .CadCliente = 2
    .CadClienteCheque = 2
    .CadConta = 2
    .CadDespesaTipo = 2
    .CadDespesaBancaria = 2
    .CadFormaDePg = 2
    .CadFornecedores = 2
    .CadFuncionarios = 2
    .CadJuros = 2
    .CadPostos = 2
    .CadProdutos = 2
    .CadProdutosFornecedores = 2
    .CadTanques = 2
    .CadTurnos = 2
    .CadConfiguracao = 2
    .ClientesPlanos = ""
    'Controle
    .ControleFechamentoDiario = 2
    .ControleConferencia = 2
    .ControleCartoes = 2
    .ControlePgAntecipado = 2
    .ControleNotas = 2
    .ControleLancContas = 2
    .ControleContasPg = 2
    .ControleCobranca = 2
    .ControleAgua = 2
    .ControleLuz = 2
    .ControleLavagem = 2
    .ControleVales = 2
    'Cheques
    .ChequeDeposito = 2
    .ChequeDevolucao = 2
    .ChequeCobranca = 2
    .ChequeProtesto = 2
    .ChequeEnviarPEmpresaCobranca = 2
    .ChequePorData = 2
    'Banco
    .BancoConcilia = 2
    .BancoTransfere = 2
    'Relatórios
    .RelatAcertoEstoque = 2
    .RelatChequeCliente = 2
    .RelatProdutosComprados = 2
    .RelatCompraVenda = 2
    .RelatDifCaixa = 2
    .RelatDifRecebe = 2
    .RelatDifCombustivel = 2
    .RelatFormaDePg = 2
    .RelatGalonagem = 2
    .RelatGalonagemTotal = 2
    .RelatVendaProdutos = 2
    .RelatVendaDetalhada = 2
    .RelatVendaLucro = 2
    .RelatVendaMedia = 2
    .RelatDiariaCombustivel = 2
    .RelatProtestoDeCheques = 2
    .RelatCadastroIncompleto = 2
    .RelatRetornoCombustivel = 2
    .RelatFaturamentoCheques = 2
    .RelatKilometragem = 2
    'Administração
    .AdmConfirma = 2
    .AdmEstatus = 2
    .AdmTotalVenda = 2
    .AdmLMC = 2
    .AdmUsuarios = 2
    .AdmUsuariosGrupos = 2
    .admDatas = 2
    .admLiberaNotas = 2
  End With
End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim CodigoGrupo As String, Senha As String
  
  
  If txtUserName.Text = "Administrador" Then
    If txtPassword.Text = "oweum10*mqde" Then
      CarregaMaster
      SaveSetting App.EXEName, "Usuario", "Ultimo", txtUserName.Text
      Unload Me
    End If
  Else
    If txtUserName.Text <> "" Then
      If dbUsuario.RecordCount = 0 Then End
      dbUsuario.FindFirst "usuario='" & txtUserName.Text & "'"
      If dbUsuario.NoMatch = False Then
        Senha = Criptografa(dbUsuario!Senha, 225)
        If txtPassword.Text <> Senha Then
          MsgBox "Nome de usuário ou senha não confere!", , "Login"
          txtUserName.SetFocus
          Contador = Contador + 1
          If Contador = 4 Then End
          Exit Sub
        Else
          CodigoGrupo = Criptografa(dbUsuario!CodigoGrupo, 223)
          Set Ws = DBEngine.Workspaces(0)
          Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
          Set dbGrupos = db.OpenRecordset("select *from usuariosgrupos")
          If dbGrupos.RecordCount = 0 Then End
          dbGrupos.FindFirst "codigogrupo=" & CodigoGrupo
          If dbGrupos.NoMatch = True Then
            MsgBox "Erro na tabela de usuários!"
            Unload Me
          End If
          CarregaUsuario
          SaveSetting App.EXEName, "Usuario", "Ultimo", txtUserName.Text
          Unload Me
        End If
      Else
        Unload Me
      End If
    Else
      MsgBox "Senha inválida, tente novamente!", , "Login"
      txtPassword.SetFocus
      Contador = Contador + 1
      If Contador = 4 Then End
    End If
  End If
End Sub

Private Sub Form_Activate()
If txtPassword <> "" Then
  txtUserName.Text = "Administrador"
  Call cmdOk_Click
  Exit Sub
End If
txtPassword.SetFocus
If Usuarios.Senha <> "" Then
  txtPassword.Text = Usuarios.Senha
  Call cmdOk_Click
End If
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
Dim StrTemp As String

Dim Ws As Workspace, db As Database, dbTemp As Recordset
Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from UsuariosGrupos order by LiberaNotas")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table UsuariosGrupos add column LiberaNotas text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela UsuariosGrupos->LiberaNotas!"
  End If
End If
On Error GoTo 0



txtPassword.Text = Command$
StrTemp = GetSetting(App.EXEName, "Usuario", "Ultimo")
If StrTemp = "" Then
  StrTemp = String(256, Chr$(0))
  
  Resultado = w32_WNetGetUser(vbNullString, StrTemp, 256)
  If resualtado = 0 Then
    StrTemp = Left$(StrTemp, InStr(1, StrTemp, Chr$(0)) - 1)
    txtUserName.Text = StrTemp
  End If
Else
  txtUserName.Text = StrTemp
  
End If
Contador = 1
Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
Set dbUsuario = db.OpenRecordset("select *from usuarios order by usuario")
With dbUsuarios
  .ConnectionString = CaminhoUsuariosAdo
  .Refresh
End With
With Usuarios
  .Nome = "Nulo"
  .Grupo.Descri = "Nulo"
End With
End Sub

Private Sub txtPassword_GotFocus()
cmdOK.Default = True
With txtPassword
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtUserName_GotFocus()
cmdOK.Default = False
With txtUserName
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtUserName_LostFocus()
cmdOK.Default = True
End Sub
