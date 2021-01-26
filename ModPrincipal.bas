Attribute VB_Name = "ModPrincipal"

Public Selecionando As Boolean, Provedor As String
Public Caminho As String, Conectar As String, CPMF As Double, Diretorio As String
Public CaminhoADO As String, CaminhoUsuariosAdo As String
Public dbUsuario As String
Public CaminhoUsuarios As String
Public Permissao As Boolean
Public NomePosto As String
Public ComissaoAcumulativa As Boolean
Public Usuarios As Usuario
Public Configura As ConfigIni

Public Type DadosCheque
  COMP As String
  Banco As String
  Agencia As String
  Conta As String
  Cheque As String
End Type

Public Type ConfigIni
  NotaNoCaixa As Integer
  NotaBloqueia As Integer
  ChequesNoCaixa As Integer
  PrecoDiferente As Integer
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Type Permite
  Descri As String
  'Cadastro
  CadBomba As Double
  CadCliente As Double
  CadClienteCheque As Double
  CadConta As Double
  CadDespesaTipo As Double
  CadDespesaBancaria As Double
  CadFormaDePg As Double
  CadFornecedores As Double
  CadFuncionarios As Double
  CadJuros As Double
  CadPostos As Double
  CadProdutos As Double
  CadProdutosFornecedores As Double
  CadTanques As Double
  CadTurnos As Double
  CadConfiguracao As Double
  ClientesPlanos As String
  'Controle
  ControleFechamentoDiario As Double
  ControleConferencia As Double
  ControleCartoes As Double
  ControlePgAntecipado As Double
  ControleNotas As Double
  ControleLancContas As Double
  ControleContasPg As Double
  ControleCobranca As Double
  ControleAgua As Double
  ControleLuz As Double
  ControleLavagem As Double
  ControleVales As Double
  'Cheques
  ChequeDeposito As Double
  ChequeDevolucao As Double
  ChequeCobranca As Double
  ChequeProtesto As Double
  ChequeEnviarPEmpresaCobranca As Double
  ChequePorData As Double
  'Banco
  BancoConcilia As Double
  BancoTransfere As Double
  'Relatórios
  RelatAcertoEstoque As Double
  RelatChequeCliente As Double
  RelatEntrada As Double
  RelatProdutosComprados As Double
  RelatCompraVenda As Double
  RelatDifCaixa As Double
  RelatDifRecebe As Double
  RelatDifCombustivel As Double
  RelatFormaDePg As Double
  RelatGalonagem As Double
  RelatGalonagemTotal As Double
  RelatVendaProdutos As Double
  RelatVendaDetalhada As Double
  RelatVendaLucro As Double
  RelatVendaMedia As Double
  RelatDiariaCombustivel As Double
  RelatProtestoDeCheques As Double
  RelatCadastroIncompleto As Double
  RelatRetornoCombustivel As Double
  RelatFaturamentoCheques As Double
  RelatKilometragem As Double
  'Administração
  AdmConfirma As Double
  AdmEstatus As Double
  AdmTotalVenda As Double
  AdmLMC As Double
  AdmUsuarios As Double
  AdmUsuariosGrupos As Double
  admDatas As Double
  admLiberaNotas As Double
End Type

Public Type Usuario
  Nome As String
  Grupo As Permite
  Senha As String
End Type

Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function w32_WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpszLocalName As String, ByVal lpszUserName As String, lpcchBuffer As Long) As Long

Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long

Public Function AtualizaCodigoFuncionario(ByVal CodigoNovo As Double, ByVal idFuncionario As Double, ByVal CodigoAntigo As Double) As Boolean
Dim db As New ADODB.Connection
Dim dbVendedores As New ADODB.Recordset


db.Open CaminhoADO
dbVendedores.CursorLocation = adUseClient
dbVendedores.Open "select *from vendedores where codigo=" & CodigoNovo, db
If dbVendedores.RecordCount >= 1 Then
  MsgBox ("Já existe outro funcionário com esse código!")
  AtualizaCodigoFuncionario = False
Else
  db.Execute "update venda2 set codigovendedor=" & CodigoNovo & " where codigopagamento=" & idFuncionario
  AtualizaCodigoFuncionario = True
End If

db.Close
End Function

Public Sub AtualizaSequenciaCaixa()
Dim db As New ADODB.Connection, dbCaixas As New ADODB.Recordset
db.Open CaminhoADO
dbCaixas.CursorLocation = adUseClient
dbCaixas.Open "Select *from fechamentodecaixa order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
On Error Resume Next
If dbCaixas.RecordCount <> 0 Then
  dbCaixas.Find "sequencia=null"
  If Err.Number <> 0 Then Exit Sub
  If dbCaixas.EOF = False Then
    Do While dbCaixas.EOF = False
      dbCaixas!Sequencia = dbCaixas.AbsolutePosition
      dbCaixas.Update
      dbCaixas.MoveNext
    Loop
  End If
  dbCaixas.MoveFirst
  Do While dbCaixas.EOF = False
    If dbCaixas!Sequencia <> dbCaixas.AbsolutePosition Then
      dbCaixas!Sequencia = dbCaixas.AbsolutePosition
      dbCaixas.Update
    End If
    dbCaixas.MoveNext
  Loop
End If

End Sub

Public Function AtualizaSistema(ByVal Internet As Boolean, ByVal CaminhoVersao As String) As Boolean
Dim Versao As Double, Revisao As Double, Compilacao As Double, StrTemp As String
Dim A As Integer, B As Integer
Dim Atualizar As Boolean
If CaminhoVersao = "" Then Exit Function
If Internet = False Then
  On Error GoTo TrataErro
  If Dir(CaminhoVersao) <> "" Then
    StrTemp = ReadINI("Versao", "Versao", "", CaminhoVersao)
    A = InStr(1, StrTemp, ".")
    Versao = CDbl(Mid(StrTemp, 1, A - 1)) * 10000000
    B = A + 1
    A = InStr(B, StrTemp, ".")
    Versao = Versao + CDbl(Mid(StrTemp, B, A - B)) * 10000
    B = A + 1
    Versao = Versao + CDbl(Mid(StrTemp, B))
    Revisao = (App.Major * 10000000) + (App.Minor * 10000) + (App.Revision)
    If Versao > Revisao Then
      MsgBox "Existe atualização!"
      StrTemp = ReadINI("Versao", "Caminho", "", CaminhoVersao)
      If Dir(StrTemp) <> "" Then
        Shell StrTemp, vbNormalFocus
        End
      End If
    Else
      AtualizaSistema = True
    End If
  Else
    AtualizaSistema = True
  End If
Else
  On Error Resume Next
  Load frmAtualizacao
  AtualizaSistema = True
End If
TrataErro:
End Function

Public Function CalculaJurosBoleto(ByVal Vencimento As Date, ByVal DataPagamento As Date, ByVal ValorBoleto As Currency) As Currency
Dim db As New ADODB.Connection
Dim dbJuros As New ADODB.Recordset
Dim Dias As Double, TempValor As Currency, Juros As Double
Dim DiaDaSemana As Double, FimDeSemana As Integer

If DataPagamento <= Vencimento Then
  CalculaJurosBoleto = ValorBoleto
  Exit Function
End If

db.Open CaminhoADO
dbJuros.CursorLocation = adUseClient
Dias = DateDiff("d", Vencimento, DataPagamento)
dbJuros.Open "Select *from jurosboleto where inicio<=" & Dias & " and final>=" & Dias & " order by inicio", db, adOpenKeyset, adLockOptimistic

JurosValor = 0
txJuros = 0
FimDeSemana = 0
DiaDaSemana = Weekday(Vencimento)
Select Case DiaDaSemana
  Case vbSunday 'domingo
    FimDeSemana = 1
  Case vbSaturday 'sabado
    FimDeSemana = 2
End Select
If Dias - FimDeSemana > 1 Then
  If dbJuros.RecordCount <> 0 Then
    If IsNull(dbJuros!JurosValor) = False Then
      JurosValor = dbJuros!JurosValor
    End If
    If dbJuros!Juros > 0 Then
      txJuros = (dbJuros!Juros * ValorBoleto) * Dias
    End If
  End If
  
End If
JurosValor = JurosValor + txJuros
dbJuros.Close
db.Close

TempValor = ValorBoleto + JurosValor
CalculaJurosBoleto = TempValor

End Function

Public Function AlteraCodigoProduto(ByVal CodigoProduto As Double, ByVal Codigo As Double, ByVal CodigoNovo As Double) As Boolean
Dim db As New ADODB.Connection
AlteraCodigoProduto = False
On Error GoTo TrataErro
db.Open CaminhoADO
db.Execute "update clientesnota2 set codigoproduto=" & CodigoNovo & " where codigoproduto=" & Codigo
db.Execute "update clientesnotatemp set codigo=" & CodigoNovo & " where codigo=" & Codigo
db.Execute "update clientesprodutos set codproduto=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update notascorpo set codigoproduto='" & CodigoNovo & "' where codigoproduto='" & Codigo & "'"
db.Execute "update pedidos set codproduto=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update produtosAcerto set codigoproduto=" & CodigoNovo & " where codproduto=" & CodigoProduto
db.Execute "update produtosalteradetalhe set codigo=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update produtosentrada2 set codigo=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update produtoshistorico set codigo=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update produtosnotascorpo set codigo=" & CodigoNovo & " where codigoproduto=" & CodigoProduto
db.Execute "update venda2 set codproduto=" & CodigoNovo & " where codigoproduto=" & CodigoProduto


AlteraCodigoProduto = True
db.Close
Exit Function

TrataErro:
MsgBox Err.Number & " - " & Err.Description
db.Close

End Function

Public Function SelecionaDiretorio(ByVal Formulario As Form) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = Formulario.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    SelecionaDiretorio = sPath
End Function


Public Sub Main()
  frmSplash.Show
  frmSplash.Refresh
  
  Provedor = "Microsoft.Jet.OLEDB.4.0"
  'Provedor = "SQLOLEDB.1"
  
  Selecionando = False
  frmSelecionaPosto.Show vbModal
  
  If Dir(App.Path & "\Posto.ini") = "" Then
    CriaPostoINI
  End If
  
  Configura.ChequesNoCaixa = CInt(ReadINI("cheques", "Cheques", 0, App.Path & "\Posto.ini"))
  Configura.NotaNoCaixa = CInt(ReadINI("Notas no Caixa", "Nocaixa", 0, App.Path & "\Posto.ini"))
  Configura.NotaBloqueia = CInt(ReadINI("Notas no Caixa", "Bloqueia", 0, App.Path & "\Posto.ini"))
  Configura.PrecoDiferente = CInt(ReadINI("Tabela", "Preco", 22, App.Path & "\Posto.ini"))
  If Configura.PrecoDiferente = 22 Then
    WriteINI "Tabela", "Preco", 0, App.Path & "\Posto.ini"
    Configura.PrecoDiferente = 0
  End If
  
  If Caminho = "" Then End
  If RetornaDiretorio(CaminhoUsuarios) = "" Then
    CaminhoUsuarios = App.Path & "\" & CaminhoUsuarios
  End If
  
  If Dir(CaminhoUsuarios) = "" Then
    CaminhoUsuarios = RetornaDiretorio(Caminho) & "Usuarios.mdb"
  End If
  
  StrTemp = GetSetting(App.EXEName, "Base", "CPMF", "0,38")
  If IsNumeric(StrTemp) = True Then
    CPMF = CDbl(StrTemp) / 100
  End If
  
  dbUsuario = "Administrador"
  Conectar = "Access"
  
  If Provedor = "SQLOLEDB.1" Then
    CaminhoADO = "Provider=" & Provedor & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Maria Vitoria;Data Source=temvale17"
  Else
    CaminhoADO = "Provider=" & Provedor & ";Data Source=" & Caminho & ";Persist Security Info=False"
  End If
  
  AtualizaADO
  
  On Error GoTo TrataErro
  
  Load mdiPosto
  mdiPosto.Show
  mdiPosto.Caption = "Posto de Combustível - " & NomePosto & " - " & Caminho
  
  VerificaAcesso
  
  frmUsuarioLogin.Show vbModal
  
  If Usuarios.Nome = "Nulo" Then End
  
  With mdiPosto
    If Usuarios.Nome = "Usuário Master" Then
      .mnuCadComissoes.Visible = True
    Else
      .mnuCadComissoes.Visible = False
    End If
  End With
  
  VerificaAcesso
  Selecionando = True
  
  Load frmEstatus2
  Unload frmEstatus2
  
  Unload frmSplash
  
  Exit Sub
  
TrataErro:
  
  
  
  End
End Sub

Public Function QuerGravar() As Boolean
  Dim Resposta As Integer
  Resposta = MsgBox("Deseja gravar as alterações!", vbYesNo, "Quer gravar?")
  If Resposta = vbYes Then
    QuerGravar = True
  Else
    QuerGravar = False
  End If
End Function

Public Function Converte(ByVal Texto As String) As String
  StrTemp = Texto
  StrTemp2 = ""
  On Error Resume Next
  StrTemp2 = Mid(StrTemp, 2, Len(StrTemp) - 2)
  Converte = StrTemp2
End Function

Public Function GeraCodigo() As String
Dim StrTemp As String, Dia As String
Dia = Str(Now)
For i = Len(Dia) To 1 Step -1
  If IsNumeric(Mid(Dia, i, 1)) = True Then
    StrTemp = StrTemp & Mid(Dia, i, 1)
  End If
Next i
GeraCodigo = StrTemp
End Function

Public Sub CompensaCustodia(ByVal CodigoPrevisao As Double)
Dim Ws As Workspace, db As Database
Dim dbCheques As Recordset, DbClientes As Recordset

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)

db.Execute "update cheques set compensado=-1 where codigoprevisaorecebe=" & CodigoPrevisao

Set dbCheques = db.OpenRecordset("select *from cheques where codigoprevisaorecebe=" & CodigoPrevisao)
Set DbClientes = db.OpenRecordset("select *from chequesclientes")
If dbCheques.RecordCount <> 0 Then
  dbCheques.MoveLast
  dbCheques.MoveFirst
  Do While dbCheques.EOF = False
    DbClientes.FindFirst "codigochequecliente=" & dbCheques!CodigoCliente
    If DbClientes.NoMatch = False Then
      DbClientes.Edit
      DbClientes!Depositados = DbClientes!Depositados + 1
      DbClientes!valordepositado = DbClientes!valordepositado + dbCheques!Valor
      If IsNull(DbClientes!saldopendente) = True Then DbClientes!saldopendente = 0
      DbClientes!saldopendente = DbClientes!saldopendente - dbCheques!Valor
      If DbClientes!saldopendente < 0 Then DbClientes!saldopendente = 0
      DbClientes.Update
    End If
    dbCheques.MoveNext
  Loop
End If
End Sub


Public Function AtualizaADOSPED(Optional BancoDeDados As String = "")
Dim strSql As String
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

If BancoDeDados = "" Then
  BancoDeDados = CaminhoADO
End If

db.Open BancoDeDados


CriaCampo db, "Produtos", "TipoItem", "Text", "30"
CriaCampo db, "Produtos", "CodigoNCM", "Text", "10"
CriaCampo db, "Produtos", "CodigoGenero", "Text", "30"
CriaCampo db, "Produtos", "CodigoSEFAZ", "Text", "30"
CriaCampo db, "Produtos", "CSOSN", "Text", "30"
CriaCampo db, "Produtos", "ReducaoICMS", "double"
CriaCampo db, "Produtos", "BC_ICMS_ST", "double"
CriaCampo db, "Produtos", "IPIEntrada", "Text", "30"
CriaCampo db, "Produtos", "IPISaida", "Text", "30"
CriaCampo db, "Produtos", "PISEntrada", "Text", "30"
CriaCampo db, "Produtos", "PISSaida", "Text", "30"
CriaCampo db, "Produtos", "NatRecPIS", "Text", "30"
CriaCampo db, "Produtos", "COFINSEntrada", "Text", "30"
CriaCampo db, "Produtos", "COFINSSaida", "Text", "30"
CriaCampo db, "Produtos", "COFINSIsento", "Text", "30"
CriaCampo db, "Produtos", "OBS", "memo"
CriaCampo db, "Produtos", "ContaContabil", "Text", "25"



'Cria a nova tabela de Planos De Conta.
On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosDeConta order by ProdutoResumido", db
If Err.Number = 0 Then
  dbTemp.Close
  On Error GoTo 0
  On Error Resume Next
  db.Execute "drop table PlanosDeConta"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao excluir a tabela 'PlanosDeConta Antiga'"
  End If
  
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table PlanosDeConta (CodigoPlanoDeConta counter, DataAltera datetime, COD_NAT_CC Text(30), IND_CTA Text(30), NIVEL double, COD_CTA Text(60), NOME_CTA Text(60), COD_CTA_REF Text(60), CNPJ_EST Text(20), COD_ITEM Text(60))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'PlanosDeConta Nova'"
  End If
    
End If


'****************************************************************************************
'****************************************************************************************
'Definir se é nota de serviço, produto ou somente para o sistema(sem lançamento fiscal)
'****************************************************************************************
'****************************************************************************************
CriaCampo db, "ProdutosNotas", "TipoNota", "Text", "40"

CriaCampo db, "ProdutosNotas", "IndMov", "Text", "40"
CriaCampo db, "ProdutosNotas", "IndOperacao", "Text", "60"
CriaCampo db, "ProdutosNotas", "IndEmitente", "Text", "40"
CriaCampo db, "ProdutosNotas", "IndParticipante", "Text", "60"
CriaCampo db, "ProdutosNotas", "CodSituacao", "Text", "40"
CriaCampo db, "ProdutosNotas", "Serie", "Text", "20"
CriaCampo db, "ProdutosNotas", "SubSerie", "Text", "20"
CriaCampo db, "ProdutosNotas", "ChaveNFSE", "Text", "60"
CriaCampo db, "ProdutosNotas", "Desconto", "currency"
CriaCampo db, "ProdutosNotas", "BaseCalcPis", "currency"
CriaCampo db, "ProdutosNotas", "ValorPis", "currency"
CriaCampo db, "ProdutosNotas", "BaseCalcCOFINS", "currency"
CriaCampo db, "ProdutosNotas", "ValorCOFINS", "currency"
CriaCampo db, "ProdutosNotas", "ValorPISRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorCOFINSRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorISS", "currency"
CriaCampo db, "ProdutosNotas", "BaseCalculoISS", "currency"
CriaCampo db, "ProdutosNotas", "AliquotaISS", "double"
CriaCampo db, "ProdutosNotas", "Talao", "Text", "10"
CriaCampo db, "ProdutosNotas", "CNPJLocalServico", "Text", "20"
CriaCampo db, "ProdutosNotas", "CodServico", "Text", "20"
CriaCampo db, "ProdutosNotas", "ModeloLivro", "Text", "20"
CriaCampo db, "ProdutosNotas", "Especie", "Text", "20"
CriaCampo db, "ProdutosNotas", "LP", "Text", "20"
CriaCampo db, "ProdutosNotas", "CC", "Text", "20"
CriaCampo db, "ProdutosNotas", "ValorMateriais", "currency"
CriaCampo db, "ProdutosNotas", "ValorSubEmpreitada", "currency"
CriaCampo db, "ProdutosNotas", "Municipio", "double"
CriaCampo db, "ProdutosNotas", "ValorISSRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorISSIsento", "currency"
CriaCampo db, "ProdutosNotas", "ValorRemessa", "currency"
CriaCampo db, "ProdutosNotas", "ValorIRRFRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorCSLLRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorINSSRetido", "currency"
CriaCampo db, "ProdutosNotas", "ValorIsentoPISCofins", "currency"
CriaCampo db, "produtosnotas", "CFPS", "double"
CriaCampo db, "produtosnotas", "CodObra", "text", "30"
CriaCampo db, "produtosnotas", "OBSLivre", "text", "30"
CriaCampo db, "produtosnotas", "GeraCredito", "bit"
CriaCampo db, "produtosnotas", "AliquotaSimplesNacional", "double"
CriaCampo db, "produtosnotas", "JuntaDarf", "bit"
CriaCampo db, "produtosnotas", "CodDarf", "Text", "20"

CriaCampo db, "produtosnotas", "CodInfo", "Text", "20"
CriaCampo db, "produtosnotas", "TextoComplementar", "memo"

'REGISTRO A111: PROCESSO REFERENCIADO
CriaCampo db, "produtosnotas", "NumProcesso", "Text", "20"
CriaCampo db, "produtosnotas", "IndProcesso", "Text", "50"

'REGISTRO A170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO
CriaCampo db, "ProdutosNotasCorpo", "ValorDesconto", "currency"
CriaCampo db, "ProdutosNotasCorpo", "NatBCCredito", "currency"
CriaCampo db, "ProdutosNotasCorpo", "IndOrigemCred", "Text", "50"
CriaCampo db, "ProdutosNotasCorpo", "CST_PIS", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "NatRecPIS", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "ValorBaseCalcPIS", "currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaPIS", "double"
CriaCampo db, "ProdutosNotasCorpo", "ValorPIS", "currency"
CriaCampo db, "ProdutosNotasCorpo", "CST_COFINS", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "NatRecCOFINS", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "ValorBcCOFINS", "currency"
CriaCampo db, "ProdutosNotasCorpo", "NatRecCOFINS", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaCOFINS", "double"
CriaCampo db, "ProdutosNotasCorpo", "ValorCOFINS", "currency"


'REGISTRO C111: PROCESSO REFERENCIADO
CriaCampo db, "ProdutosNotas", "NumProcesso", "Text", "20"
CriaCampo db, "ProdutosNotas", "IndProcesso", "Text", "50"
CriaCampo db, "ProdutosNotas", "CodItem", "Text", "60"

'REGISTRO C170: COMPLEMENTO DO DOCUMENTO - ITENS DO DOCUMENTO (CÓDIGOS 01, 1B, 04 e 55)
CriaCampo db, "ProdutosNotasCorpo", "ValorDesconto", "currency"
CriaCampo db, "ProdutosNotasCorpo", "MovimentacaoFisica", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "CstIcms", "Text", "50"
CriaCampo db, "ProdutosNotasCorpo", "CFOP", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "CodNatOperacao", "Text", "10"
CriaCampo db, "ProdutosNotasCorpo", "BaseCalcICMS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaICMS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "ValorICMS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "BaseCalcICMSst", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaST", "Double"
CriaCampo db, "ProdutosNotasCorpo", "ValorICMSst", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "IndApuracao", "Text", "20"
CriaCampo db, "ProdutosNotasCorpo", "CST_IPI", "Text", "2"
CriaCampo db, "ProdutosNotasCorpo", "CodEnquadraIPI", "Text", "3"
CriaCampo db, "ProdutosNotasCorpo", "ValorBcIPI", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaIPI", "Double"
CriaCampo db, "ProdutosNotasCorpo", "ValorIPI", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "CST_PIS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "ValorBcPIS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AiquotaPIS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "QuantBcPIS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaPISQuant", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "ValorPIS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "CST_COFINS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "ValorBcCOFINS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaCOFINS", "Double"
CriaCampo db, "ProdutosNotasCorpo", "QuantBcCOFINS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "AliquotaCOFINSQuant", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "ValorCOFINS", "Currency"
CriaCampo db, "ProdutosNotasCorpo", "CodCTA", "Text", "60"
CriaCampo db, "ProdutosNotasCorpo", "NatReceitaPIS", "Text", "3"
CriaCampo db, "ProdutosNotasCorpo", "NatReceitaCOFINS", "Text", "20"

'REGISTRO C470: ITENS DO DOCUMENTO FISCAL EMITIDO POR ECF (CÓDIGO 02 e 2D).
CriaCampo db, "ProdutosNotasCorpo", "QuantCancelada", "Double"

'REGISTRO C510: ITENS DO DOCUMENTO NOTA FISCAL/CONTA ENERGIA ELÉTRICA (CÓDIGO 06),
'NOTA FISCAL/CONTA DE FORNECIMENTO D'ÁGUA CANALIZADA (CÓDIGO 29) E
'NOTA FISCAL/CONTA DE FORNECIMENTO DE GÁS (CÓDIGO 28).
CriaCampo db, "ProdutosNotasCorpo", "CodClassificacao", "Text", "4"
CriaCampo db, "ProdutosNotasCorpo", "IndReceita", "Text", "30"
CriaCampo db, "ProdutosNotasCorpo", "CodParticipante", "Text", "60"




db.Close


End Function

Public Function CriaCampo(ByVal db As ADODB.Connection, ByVal Tabela As String, ByVal NomeCampo As String, ByVal Tipo As String, Optional Tamanho As String = 0, Optional BancoDeDados As String = "") As Boolean
Dim TempTipo As String

Dim dbTemp As New ADODB.Recordset

CriaCampo = False

If BancoDeDados = "" Then
  BancoDeDados = "Provider=" & Provedor & ";Data Source=" & CaminhoUsuarios & ";Persist Security Info=False"
End If
If UCase(Tipo) = "TEXT" Then
    If Tamanho = "0" Then
        Exit Function
    Else
        TempTipo = Tipo & "(" & Tamanho & ")"
    End If
Else
    TempTipo = Tipo
End If

'db.Open BancoDeDados

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select " & NomeCampo & " from " & Tabela, db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table " & Tabela & " add column " & NomeCampo & " " & TempTipo
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela '" & Tabela & "->" & NomeCampo & "'"
    dbTemp.Close
    Exit Function
  End If
End If

dbTemp.Close
'db.Close

CriaCampo = True

End Function



Public Function AtualizaUsuarios(Optional BancoDeDados As String = "")
Dim strSql As String
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

If BancoDeDados = "" Then
  BancoDeDados = "Provider=" & Provedor & ";Data Source=" & CaminhoUsuarios & ";Persist Security Info=False"
End If

db.Open BancoDeDados

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from UsuariosGrupos order by ClientesPlanos", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "ALTER TABLE UsuariosGrupos Add column ClientesPlanos Text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'UsuariosGrupos->ClientesPlanos'"
  End If
End If
dbTemp.Close

db.Close


End Function

Public Function AtualizaADO(Optional BancoDeDados As String = "")
Dim strSql As String
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

If BancoDeDados = "" Then
  BancoDeDados = CaminhoADO
End If

db.Open BancoDeDados

AtualizaDbQuery

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Cartoes order by Obs", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  AtualizaAdo2 BancoDeDados
End If
dbTemp.Close

db.Close

End Function

Public Function AtualizaAdo2(Optional BancoDeDados As String = "")
Dim strSql As String
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

If BancoDeDados = "" Then
  BancoDeDados = CaminhoADO
End If

db.Open BancoDeDados

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Cartoes order by Obs", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Cartoes add column Obs memo"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cartoes->Obs'"
  End If
End If
dbTemp.Close

db.Close

db.Close


End Function


Public Function AtualizaAdo002(Optional BancoDeDados As String = "")
Dim strSql As String
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

If BancoDeDados = "" Then
  BancoDeDados = CaminhoADO
End If

db.Open BancoDeDados

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Pdvs order by Intermitente", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Pdvs add column Intermitente bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Pdvs->Intermitente'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Pdvs order by HoraIni", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Pdvs add column HoraIni datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Pdvs->HoraIni'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PdvsTurnos", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table PdvsTurnos (codigoPdv double, CodigoTurno double, DescriPdv Text(50), DescriTurno text(50), HoraIni datetime)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'PdvsTurnos'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ImportacaoErros", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ImportacaoErros (CodigoImportacao counter, codigoFechamento double, DataImportado datetime, Tipo Text(30), Descri Text(255), bico double, CodigoNoPosto double, CodigoFuncionario double, CodigoProduto double, ValorPosto Currency, ValorSistema Currency, CodigoClienteNoPosto double, CodigoClienteSistema double, LimiteNadata currency, ValorBloqueado currency, StatusCliente Text(30))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ImportacaoErros'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FormaDePagamentoTotalizado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table FormaDePagamentoTotalizado (Descri Text(60), Taxa double, ValorBruto currency, ValorLiquido Currency, Custo Currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'FormaDePagamentoTotalizado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ImportacaoErros order by Funcionario", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ImportacaoErros add column Funcionario double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ImportacaoErros->Funcionario'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ImportacaoErros order by Qtd", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ImportacaoErros add column Qtd double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ImportacaoErros->Qtd'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Status order by Arredondamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Status add column Arredondamento double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Status->Arredondamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotas order by Gravado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotas add column Gravado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotas->Gravado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by UnCaixa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column UnCaixa double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->UnCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from clientes order by Nota", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column Nota Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->Nota'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Notas order by Nota", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Notas add column Nota Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Notas->Nota'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by Arredondamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column Arredondamento Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->Arredondamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Cheques order by datalista", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Cheques add column DataLista datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cheques->DataLista'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotasCorpo order by CodBarras", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotasCorpo add column CodBarras text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotasCorpo->CodBarras'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Comissoes", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table Comissoes (id counter, CodigoProduto double, Codigo text(15), bico integer, CodigoFuncionario double, funcionario integer, Nome text(50), qtd double, VlUnitario currency, VlTotal currency, VlVendaC currency, VlTotalC currency, VlComissao Currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Comissoes'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by Comissao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column Comissao currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->Comissao'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Comissoes order by CodigoFechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Comissoes add column CodigoFechamento double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Comissoes->CodigoFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Venda2 order by Combustivel", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Venda2 add column Combustivel bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Vendas2->Combustivel'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Venda2 order by bico", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Venda2 add column Bico integer"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Vendas2->Bico'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column PlanoDeConta text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ResumoDia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ResumoDia (data datetime, combustivel currency, Produtos currency, total currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Comissoes'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesPlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesPlanoDeConta (id counter, CodigoPlano text(10), Descri text(100))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesPlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Config order by ClientesNotaPlano", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Config add column ClientesNotaPlano text(200)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->ClientesNotaPlano'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Config order by MicroCreditoPlano", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Config add column MicroCreditoPlano text(200)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->MicroCreditoPlano'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by NotaConferida2", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column NotaConferida2 bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->NotaConferida2'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesNota2 order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesNota2 add column PlanoDeConta Text(10)"
  db.Execute "update clientesnota2 set planodeconta='3500000000'"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by ComissaoAcumulativa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column ComissaoAcumulativa bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->ComissaoAcumulativa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by UltimaVenda", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column UltimaVenda DateTime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->ComissaoAcumulativa'"
  End If
End If
dbTemp.Close

db.Close

db.Close


End Function


Public Function AtualizaADO001()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim DbDao As Database, Ws As Workspace

Set Ws = DBEngine.Workspaces(0)
Set DbDao = Ws.OpenDatabase(Caminho, , , Conectar)

On Error Resume Next
DbDao.CreateQueryDef "qClientesCobrancaComposicao", "select ClientesCobrancaComposicao.*, ClientesCobranca.* from ClientesCobrancaComposicao, ClientesCobranca where ClientesCobrancaComposicao.codigocobranca=ClientesCobranca.codigocobranca"
On Error GoTo 0

db.Open CaminhoADO

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by Sequencia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column Sequencia double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->Sequencia'"
  End If
End If
dbTemp.Close


On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from qFormaDePgContasNovo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set Ws = DBEngine.Workspaces(0)
  Set DaoDB = Ws.OpenDatabase(Caminho, , , Conectar)
  DaoDB.CreateQueryDef "qFormaDePgContasNovo", "SELECT Contas.*, FormaDePagamento.* FROM FormaDePagamento, contas where Contas.CodigoConta = FormaDePagamento.CodigoConta"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'qFormaDePgContasNovo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Protestos", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table Protestos (CodigoProtesto counter, DataLanc datetime, DataDocumento DateTime, TipoDocumento Text(30), Status Text(20), Descri Text(100), Valor currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'Protestos'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Protestos order by CodigoClienteNota", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Protestos add column CodigoClienteNota double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Protestos->CodigoClienteNota'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Protestos order by CodigoClienteCheque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Protestos add column CodigoClienteCheque double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Protestos->CodigoClienteCheque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Protestos order by Nome", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Protestos add column Nome Text(100)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Protestos->Nome'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Protestos order by CPF_CNPJ", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Protestos add column CPF_CNPJ Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Protestos->CPF_CNPJ'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobrancaTemp", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesCobrancaTemp (CodigoTemp counter, CodigoCliente double, Nome Text(100), Tipo Text(20), Vencimento datetime, ValorPrevisto currency, Confirmado bit)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ClientesCobrancaTemp'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesNotaTemp", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesNotaTemp (CodigoTemp counter, CodigoClientesCobra double, CodigoClientesNota2 double, CodigoCliente double, Nome Text(100), Cupom Text(20), Data datetime, CodigoProduto double, Codigo double, Descri Text(50), Valor currency, Confirmado bit)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ClientesNotaTemp'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by ValorPrevisto", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column ValorPrevisto currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->ValorPrevisto'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobrancaTemp order by DataFechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobrancaTemp add column DataFechamento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobrancaTemp->DataFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobrancaTemp order by Praso", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobrancaTemp add column Praso double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobrancaTemp->Praso'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by DataFechaMes", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column DataFechaMes datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->DataFechaMes'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by FechaAluguel", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column FechaAluguel bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->FechaAluguel'"
  Else
    db.Execute "update clientescobranca set fechaaluguel=fechames"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by QtdComprado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column QtdComprado double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->QtdComprado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ValorComprado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ValorComprado Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ValorComprado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by TipoCliente", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column TipoCliente Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->TipoCliente'"
  End If
  db.Execute "update Clientes set TipoCliente='Fiado'"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by TipoCliente", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column TipoCliente Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->TipoCliente'"
  End If
  db.Execute "update ClientesCobranca set TipoCliente='Fiado'"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesTipo", db, adOpenKeyset
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesTipo (TipoCliente Text(20))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesTipo'"
  End If
  db.Execute "insert into ClientesTipo (TipoCliente) values ('Fiado')"
  db.Execute "insert into ClientesTipo (TipoCliente) values ('Nota Avulsa')"
  db.Execute "insert into ClientesTipo (TipoCliente) values ('Aluguel')"
  db.Execute "insert into ClientesTipo (TipoCliente) values ('Outros')"
End If
dbTemp.Requery
If dbTemp.RecordCount <> 0 Then
  dbTemp.MoveLast
  dbTemp.MoveFirst
  dbTemp.Find "tipocliente='Fiado'"
  If dbTemp.EOF = True Then
    db.Execute "insert into ClientesTipo (TipoCliente) values ('Fiado')"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by Obs", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column Obs Text(255)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->Obs'"
  End If
  db.Execute "update ClientesCobranca set TipoCliente='Fiado'"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobranca order by Origem", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesCobranca add column Origem Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->Origem'"
  End If
  db.Execute "update ClientesCobranca set Origem='Fiado'"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ValorEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ValorEstoque Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ValorEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by DifEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column DifEstoque double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->DifEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ValorDifEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ValorDifEstoque Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ValorDifEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by LucroMedio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column LucroMedio Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->LucroMedio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by PrecoMedio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column PrecoMedio Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->PrecoMedio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotasCorpo order by CodigoCaixa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotasCorpo add column CodigoCaixa double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotasCorpo->CodigoCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from diferencacombustivel order by ValorDiferenca", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table diferencacombustivel add column ValorDiferenca Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DiferencaCombustivel->ValorDiferenca'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from contas order by CustodiaLote", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table contas add column CustodiaLote Double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->CustodiaLote'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from cheques order by Autorizar", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Cheques add column Autorizar bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cheques->Autorizar'"
  End If
  db.Execute "update cheques set autorizar=0"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from cheques order by Autorizado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Cheques add column Autorizado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cheques->Autorizado'"
  End If
  db.Execute "update cheques set autorizado=0"
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from fechamentodecaixa order by NotaConferida", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentodeCaixa add column NotaConferida bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->NotaConferida'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by LucroMedio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column LucroMedio currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->LucroMedio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by PrecoMedio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column PrecoMedio currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->PrecoMedio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by CodigoEmpresa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column CodigoEmpresa double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->CodigoEmpresa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by CodigoFilial", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column CodigoFilial text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->CodigoFilial'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from StatusDiario", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table StatusDiario (CodigoStatus counter, DataLanc datetime, DataCaixa DateTime, CodigoTurno double, Turno Text(30), CapitalDoDia currency, LucroAcumulado currency, LucroDoDia currency, Diferenca Currency, ExibeStatusMensal bit, FechamentoMensal bit, DataFechamento Datetime, FechamentoTrimestral bit, DataTrimestre datetime)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'StatusDiario'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Status order by CapitalInicial", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Status add column CapitalInicial currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Status->CapitalInicial'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from lmc order by VendasNoDia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table lmc add column VendasNoDia currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'LMC->VendasNoDia'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from lmc order by AcumuladoNoMes", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table lmc add column AcumuladoNoMes currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'LMC->AcumuladoNoMes'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table VendasLeituraX (Data datetime, Categoria Text(50), SistemaQtd double, SistemaValor Currency, LeituraXQtd double, LeituraXValor Currency, DiferencaQtd double, DiferencaValor currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'VendasLeituraX'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by Aliquota", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column Aliquota Text(6)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->Aliquota'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by Combustivel", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column Combustivel bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->Combustivel'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by ReducaoZ", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column ReducaoZ Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->ReducaoZ'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by PrecoDiferenciado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column PrecoDiferenciado Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->PrecoDiferenciado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by DifReducaoZ", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column DifReducaoZ Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->DifReducaoZ'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by Departamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column Departamento Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->Departamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesCobrancaComposicao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesCobrancaComposicao (CodigoCobranca double, Descri Text(150), Reembolso bit ,Valor currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ClientesCobrancaComposicao'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ComposicaoTipo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ComposicaoTipo (Descri Text(150), Reembolso bit)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ComposicaoTipo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by QtdComprado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column QtdComprado double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->QtdComprado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ValorComprado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ValorComprado Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ValorComprado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by NaoBloqueia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column NaoBloqueia bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->NaoBloqueia'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesNota2 order by Autorizar", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesNota2 add column Autorizar bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->Autorizar'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesNota2 order by Autorizado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesNota2 add column Autorizado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->Autorizado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesNota2 order by Motivo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesNota2 add column Motivo text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->Motivo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosEntrada2 order by CodigoNota", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosEntrada2 add column CodigoNota double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosEntrada2->CodigoNota'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from despesaslanc2 order by datafechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table despesaslanc2 add column DataFechamento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'despesaslanc2->DataFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from cartoes order by datafechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table cartoes add column DataFechamento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'cartoes->DataFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from formadepagamentorecebido2 order by datafechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table formadepagamentorecebido2 add column DataFechamento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'formadepagamentorecebido2->DataFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from fechamentodecaixa order by datafechamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table fechamentodecaixa add column DataFechamento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'fechamentodecaixa->DataFechamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from cheques order by Usuario", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table cheques add column Usuario text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cheques->Usuario'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Status order by estacionamento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Status add column Estacionamento Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Status->Estacionamento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from produtosaltera order by horaini", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table produtosaltera add column HoraIni datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'produtosaltera->HoraIni'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from produtosnotascorpo order by Aguardando", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table produtosnotascorpo add column Aguardando bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->ParaContabilidade'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DespesasLanc2 order by ParaContabilidade", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DespesasLanc2 add column ParaContabilidade bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->ParaContabilidade'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DespesasLanc2 order by DataContabilidade", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DespesasLanc2 add column DataContabilidade datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->DataContabilidade'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DespesasLanc2 order by CodigoEnviar", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DespesasLanc2 add column CodigoEnviar Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->CodigoEnviar'"
  End If
  On Error GoTo 0
  On Error Resume Next
  db.Execute "update despesaslanc2 set codigoenviar='1' where paracontabilidade=0"
End If
dbTemp.Close




On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from CFOP order by Descri", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table CFOP add column Descri Text(250)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'CFOP->Descri'"
  End If
Else
  If dbTemp.Fields("Descri").DefinedSize = 20 Then
    dbTemp.Close
    db.Execute "alter table CFOP add column descri2 text(50)"
    db.Execute "update CFOP set descri2=descri"
    db.Execute "alter table CFOP drop column descri"
    db.Execute "alter table CFOP add Column Descri text(250)"
    db.Execute "update cfop CFOP descri=descri2"
    db.Execute "alter table CFOP drop column descri2"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from CFOP order by Aplicacao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table cfop add column Aplicacao memo"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'CFOP->Aplicacao'"
  End If
End If
dbTemp.Close


On Error GoTo 0
dbTemp.Open "select *from Despesaslanc2", db
If dbTemp.Fields("obs").DefinedSize = 50 Then
  dbTemp.Close
  db.Execute "alter table Despesaslanc2 add column obs2 text(50)"
  db.Execute "update despesaslanc2 set obs2=obs"
  db.Execute "alter table despesaslanc2 drop column obs"
  db.Execute "alter table Despesaslanc2 add column obs text(250)"
  db.Execute "update despesaslanc2 set obs=obs2"
  db.Execute "alter table despesaslanc2 drop column obs2"
End If
On Error GoTo 0
On Error Resume Next
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from clientesnota2 order by ClienteAntigo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientesnota2 add column ClienteAntigo double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->ClienteAntigo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from clientesnota2 order by UsuarioTroca", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientesnota2 add column UsuarioTroca text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->UsuarioTroca'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Despesatipo order by Mensal", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column Mensal bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendedoresPagamento->Mensal'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Despesatipo order by UltimoVencimento", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column UltimoVencimento DateTime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendedoresPagamento->UltimoVencimento'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Despesatipo order by Obrigatorio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column Obrigatorio text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendedoresPagamento->ConfirmadoNoCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DespesatipoObrigatorio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table DespesatipoObrigatorio (Descri text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'DespesatipoObrigatorio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by Plano", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column Plano text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->Plano'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by CodResumido", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column CodResumido text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->CodResumido'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by Interna", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column Interna bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->Interna'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Contas order by Tipo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Contas add column Tipo Text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Contas->Tipo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Despesatipo order by CodigoHistorico", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column CodigoHistorico Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Despesatipo->CodigoHistorico'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Despesatipo order by CodigoPlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column CodigoPlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Despesatipo->CodigoPlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ContasDespesas order by CodigoHistorico", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ContasDespesas add column CodigoHistorico Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ContasDespesas->CodigoHistorico'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ContasDespesas order by CodigoPlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ContasDespesas add column CodigoPlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ContasDespesas->CodigoPlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosCaixa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table PlanosCaixa (Produto text(10), ProdutoHistorico Text(10), Cliente text(10), ClienteHistorico Text(10), Despesa text(10), DespesaHistorico Text(10), Cheques text(10), ChequeHistorico Text(10))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PlanosCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosHistorico", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table PlanosHistorico (CodigoLancamento counter, Data datetime, Origem text(30), Debito text(10), Credito Text(10), Valor currency, CodigoHistoricoPadrao Text(10), Descri Text(200))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PlanosCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ChequesClientes order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ChequesClientes add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ChequesClientes->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Fornecedores order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Fornecedores add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Fornecedores->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FormaDePagamento order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FormaDePagamento add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Vendedores order by PlanoDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Vendedores add column PlanoDeConta Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Vendedores->PlanoDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosDeConta", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table PlanosDeConta (ProdutoPlano text(50), ProdutoResumido text(10), ClientePlano text(50), ClienteResumido text(10), DespesaPlano text(50), DespesaResumido text(10), CartaoPlano text(50), CartaoResumido text(10), ChequePlano text(50), ChequeResumido text(10), FornecedoresPlano text(50), FornecedoresResumido text(10), FuncionariosPlano text(50), FuncionariosResumido text(10))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PlanosDeConta'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosDeConta order by CaixaPlano", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table PlanosDeConta add column CaixaPlano Text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'PlanosDeConta->CaixaPlano'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PlanosDeConta order by CaixaResumido", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table PlanosDeConta add column CaixaResumido Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'PlanosDeConta->CaixaResumido'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Config order by LocalCustodia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Config add column LocalCustodia Text(255)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->LocalCustodia'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FormaDePagamento order by Corte", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FormaDePagamento add column Corte bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->Corte'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FormaDePagamento order by dataCorte", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FormaDePagamento add column dataCorte datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->dataCorte'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FormaDePagamento order by diasCorte", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FormaDePagamento add column diasCorte integer"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->diasCorte'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from lmc order by Obs", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table LMC add column OBS Memo"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'LMC->OBS'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from despesaslanc2 order by pagarcomo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DespesasLanc2 add column PagarComo text(100)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->PagarComo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by Sequencia", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column Sequencia double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->Sequencia'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from qProdutosVendaCaixa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  
  Set Ws = DBEngine.Workspaces(0)
  Set DbDao = Ws.OpenDatabase(Caminho, , , Conectar)
  
  On Error Resume Next
  DbDao.CreateQueryDef "qProdutosVendaCaixa", "select produtos.*, Fechamentodecaixa.*, venda2.*  from produtos, Fechamentodecaixa, venda2 where produtos.codigoproduto=venda2.codigoproduto and fechamentodecaixa.codigofechamento=venda2.codigofechamento"
  On Error GoTo 0
  
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'qProdutosVendaCaixa'"
  End If
End If
On Error Resume Next
dbTemp.Close
On Error GoTo 0

AtualizaSequenciaCaixa

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by DuracaoEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column DuracaoEstoque double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->DuracaoEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by EstoqueIdeal", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column EstoqueIdeal double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->EstoqueIdeal'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by DuracaoEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column DuracaoEstoque double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->DuracaoEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by EstoqueIdeal", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column EstoqueIdeal double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->EstoqueIdeal'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by Sugerido", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column Sugerido double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->Sugerido'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by LucroMinimo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column LucroMinimo double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->LucroMinimo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by PrecoSugerido", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column PrecoSugerido double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->PrecoSugerido'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Config order by ServidorPista", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Config add column ServidorPista text(255)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->ServidorPista'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesHistorico", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesHistorico (CodigoCliente double, Ativando bit, Dia datetime, Descri text(255), Usuario text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PlanosCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixaPista", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table FechamentoDeCaixaPista (CodigoFechamento double, CodigoConta text(15), valor currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PlanosCaixa'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosGrupoIF", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ProdutosGrupoIF (codigo counter, CodigoGrupo double, Descri text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ProdutosGrupoIF'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CodigoGrupoIF", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CodigoGrupoIF double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CodigoGrupoIF'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasLeituraX order by CodigoGrupoIF", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasLeituraX add column CodigoGrupoIF double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasLeituraX->CodigoGrupoIF'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosClassFisc", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ProdutosClassFisc (codigo counter, CodigoClass text(15), Descri text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ProdutosClassFisc'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosOrigem", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ProdutosOrigem (codigo counter, CodigoOrigem integer, Descri text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ProdutosOrigem'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosTributacao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ProdutosTributacao (codigo counter, CodigoTributacao integer, Descri text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ProdutosTributacao'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CodigoClasFisc", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CodigoClasFisc double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CodigoClasFisc'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CodigoOrigem", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CodigoOrigem double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CodigoOrigem'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CodigoTributacao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CodigoTributacao double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CodigoTributacao'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotas order by codigoTurno", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotas add column codigoTurno double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotas->codigoTurno'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosEntrada2 order by codigoTurno", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosEntrada2 add column codigoTurno double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosEntrada2->codigoTurno'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosEstoque", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "CREATE TABLE ProdutosEstoque(ID counter," & _
                                  "CodigoProduto double," & _
                                  "Codigo double," & _
                                  "Tanque text(3)," & _
                                  "DataCaixa DateTime," & _
                                  "codigoTurno double," & _
                                  "Turno text(50)," & _
                                  "horaini datetime," & _
                                  "Combustivel bit," & _
                                  "Abertura double," & _
                                  "Entrada double," & _
                                  "Saida double," & _
                                  "Acerto double," & _
                                  "Diferenca double," & _
                                  "Disponivel double," & _
                                  "DataAlterado datetime," & _
                                  "Usuario text(50))"
   
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'ProdutosEstoque'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicosEncerrantesNovo", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "CREATE TABLE BicosEncerrantesNovo(ID counter," & _
                                  "Bico double," & _
                                  "DataCaixa datetime," & _
                                  "CodigoTurno double," & _
                                  "Turno Text(20)," & _
                                  "HoraIni DateTime," & _
                                  "Inicial double," & _
                                  "DataAlterado datetime," & _
                                  "Usuario text(50))"
   
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'BicosEncerrantesNovo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotas order by FormaDePg", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotas add column FormaDePg Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotas->FormaDePg'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosEntrada2 order by FormaDePg", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosEntrada2 add column FormaDePg Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosEntrada2->ProdutosEntrada2'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DespesasLanc2 order by FormaDePg", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DespesasLanc2 add column FormaDePg Text(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DespesasLanc2->ProdutosEntrada2'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by Numero", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column Numero Text(60)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->Numero'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by CodMunicipio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column CodMunicipio Text(7)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->CodMunicipio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by Municipio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column Municipio Text(60)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->Municipio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by Municipio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column Municipio Text(60)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->Municipio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by Apurado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column Apurado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->Apurado'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from DiferencaCombustivel order by Apurado", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table DiferencaCombustivel add column Apurado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'DiferencaCombustivel->Apurado'"
  End If
End If
dbTemp.Close

dbTemp.CursorLocation = adUseClient
dbTemp.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly
If dbTemp.RecordCount <> 0 Then
  strSql = dbTemp!ftp
  dbTemp.Close
  db.Close
  
  On Error Resume Next
  db.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & strSql
  If Err.Number = 0 Then
    On Error GoTo 0
    On Error Resume Next
    dbTemp.Open "select planodeconta, datacaixa from caixas where datacaixa='" & Date & "' order by PlanoDeConta", db, adOpenForwardOnly, adLockReadOnly
    If Err.Number <> 0 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description
      On Error GoTo 0
      On Error Resume Next
      db.Execute "ALTER TABLE Caixas Add PlanoDeConta nVarChar(20)"
      If Err.Number <> 0 Then
        MsgBox "Erro " & Err.Number & " - " & Err.Description
      Else
        On Error GoTo 0
        db.Execute "update caixas set PlanoDeConta='2100000000'"
      End If
    Else
      dbTemp.Close
    End If
    On Error Resume Next
    
    db.Close
  End If
Else
  dbTemp.Close
  db.Close
End If
db.Open BancoDeDados

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by codEAN", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column codEAN text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->codEAN'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CFOP", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CFOP text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CFOP'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by CST", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column CST text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->CST'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by Origem", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column Origem text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->Origem'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by IPI", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column IPI text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->IPI'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by PIS", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column PIS text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->PIS'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by COFINS", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column COFINS text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->COFINS'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ISS", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ISS text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ISS'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by AliquotaICMS", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column AliquotaICMS text(7)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->AliquotaICMS'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Notas order by Eletronica", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Notas add column Eletronica Bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Notas->Eletronica'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Notas order by Gerada", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Notas add column Gerada Bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Notas->Gerada'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Postos order by Bairro", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Postos add column Bairro Text(40)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Postos->Bairro'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Postos order by CEP", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Postos add column CEP Text(40)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Postos->CEP'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Notas order by CodMunicipio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Notas add column CodMunicipio Text(7)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Notas->CodMunicipio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from NotasCorpo order by CFOP", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table NotasCorpo add column CFOP Text(4)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'NotasCorpo->CFOP'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from NotasCorpo order by Origem", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table NotasCorpo add column Origem Text(1)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'NotasCorpo->Origem'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from NotasCorpo order by ValorIcms", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table NotasCorpo add column ValorIcms currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'NotasCorpo->ValorIcms'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from NotasCorpo order by ReducaoIcms", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table NotasCorpo add column ReducaoIcms double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'NotasCorpo->ReducaoIcms'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by ReducaoIcms", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column ReducaoIcms double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->ReducaoIcms'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Estados", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "CREATE TABLE Estados(ID counter," & _
                                  "Codigo Text(2)," & _
                                  "Nome Text(30)," & _
                                  "Sigla text(2))"
   
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'BicosEncerrantesNovo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Municipios", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "CREATE TABLE Municipios(ID counter," & _
                                  "CodigoEstado Text(2)," & _
                                  "Codigo Text(7)," & _
                                  "Nome Text(60))"
   
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'BicosEncerrantesNovo'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Clientes order by CodMunicipio", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column CodMunicipio Text(7)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->CodMunicipio'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by Comissao", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column Comissao double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->Comissao'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from VendasTemp order by ComissaoValor", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendasTemp add column ComissaoValor double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendasTemp->ComissaoValor'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Postos order by ComissaoAcumulativa", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Postos add column ComissaoAcumulativa bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Postos->ComissaoAcumulativa'"
  Else
    db.Execute "update postos set comissaoacumulativa=" & ComissaoAcumulativa
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from Produtos order by LMC", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column LMC bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->LMC'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosNotasCorpo order by LMC", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosNotasCorpo add column LMC bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosNotasCorpo->LMC'"
  Else
    db.Execute "update ProdutosNotasCorpo set lmc=-1 where tanque<>0"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from PDVs", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "CREATE TABLE PDVs(codigoPDV counter," & _
                                  "Codigo Text(20)," & _
                                  "Descri Text(50))"
   
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao Criar a tabela 'PDVs'"
  Else
    dbTemp.Open "select *from PDVs", db, adOpenKeyset, adLockOptimistic
    dbTemp.AddNew
    dbTemp!Codigo = "2100000000"
    dbTemp!Descri = "Pista"
    dbTemp.Update
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from FechamentoDeCaixa order by CodigoPdv", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column CodigoPdv double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->CodigoPdv'"
  Else
    db.Execute "update fechamentodecaixa set codigopdv=1"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by DeOutroCaixaQtd", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column DeOutroCaixaQtd double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->DeOutroCaixaQtd'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by DeOutroCaixaValor", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column DeOutroCaixaValor Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->DeOutroCaixaValor'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by DesteCaixaQtd", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column DesteCaixaQtd double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->DesteCaixaQtd'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes order by DesteCaixaValor", db
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table BicoEncerrantes add column DesteCaixaValor Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'BicoEncerrantes->DesteCaixaValor'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from BicoEncerrantes", db
If Err.Number = 0 Then
  If VarType(dbTemp!Vendas) = vbLong Then
    dbTemp.Close
    On Error GoTo 0
    On Error Resume Next
    db.Execute "alter table BicoEncerrantes add column Vendas2 double"
    If Err.Number = 0 Then
      On Error GoTo 0
      On Error Resume Next
      db.Execute "update bicoencerrantes set vendas2=encerrante-abertura-retorno"
      If Err.Number = 0 Then
        On Error GoTo 0
        On Error Resume Next
        db.Execute "alter table BicoEncerrantes drop column Vendas"
        If Err.Number = 0 Then
          On Error GoTo 0
          On Error Resume Next
          db.Execute "alter table BicoEncerrantes add column Vendas double"
          db.Execute "update bicoencerrantes set vendas=vendas2"
        End If
      End If
    End If
  End If
  If VarType(dbTemp!Retorno) = vbLong Then
    dbTemp.Close
    On Error GoTo 0
    On Error Resume Next
    db.Execute "alter table BicoEncerrantes add column Retorno2 double"
    If Err.Number = 0 Then
      On Error GoTo 0
      On Error Resume Next
      db.Execute "update bicoencerrantes set retorno2=retorno"
      If Err.Number = 0 Then
        On Error GoTo 0
        On Error Resume Next
        db.Execute "alter table BicoEncerrantes drop column retorno"
        If Err.Number = 0 Then
          On Error GoTo 0
          On Error Resume Next
          db.Execute "alter table BicoEncerrantes add column Retorno double"
          db.Execute "update bicoencerrantes set retorno=retorno2"
        End If
      End If
    End If
  End If
End If
dbTemp.Close


DbDao.Close
Ws.Close
db.Close

End Function

Public Function AtualizaDb()
Dim db As Database, Ws As Workspace
Dim dbTemp As Recordset

If Dir(Caminho) = "" Then Exit Function

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from VendedoresPagamento order by ConfirmadoNoCaixa")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendedoresPagamento add column ConfirmadoNoCaixa bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendedoresPagamento->ConfirmadoNoCaixa'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Produtos order by PermiteNoCaixa")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Produtos add column PermiteNoCaixa bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->PermiteNoCaixa'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from vendedores order by codigoDespesa")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Vendedores add column codigoDespesa double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Vendedores->codigoDespesa'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from qClientesNota2Produtos")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.CreateQueryDef "qClientesNota2Produtos", "select clientesnota2.*, produtos.* from clientesnota2, produtos where clientesnota2.codigoproduto=produtos.codigo"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a consulta 'qClientesNota2Produtos'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ClientesNota2 order by valorunitario")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientesNota2 add column ValorUnitario Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->ValorUnitario'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ClientesNota2 order by qtd")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientesNota2 add column Qtd double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNota2->Qtd'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from status order by ClienteDiferenciado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table status add column ClienteDiferenciado currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'status->ClienteDiferenciado'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientesnota2 order by LucroDif")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table clientesnota2 add column LucroDif currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesNota2->LucroDif'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientesnota2 order by ValorUnitarioDif")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table clientesnota2 add column ValorUnitarioDif currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesNota2->ValorUnitarioDif'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientesnota2 order by ValorTotalDif")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table clientesnota2 add column ValorTotalDif currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesNota2->ValorTotalDif'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

frmSplash.lblWarning.Caption = "Atualizando Banco de Dados! Aguarde..."
frmSplash.Refresh

Set dbTemp = db.OpenRecordset("select *from VendedoresPagamento order by FechadoAte")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table VendedoresPagamento add column FechadoAte datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'VendedoresPagamento->FechadoAte'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Venda2 order by IDPagamento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Venda2 add column IDPagamento double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Venda2->IDPagamento'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from BicoMovimento")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table BicoMovimento"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'BicoMovimento'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Cheques2")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Cheques2"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Cheques2'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from ClientesNota")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table ClientesNota"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'ClientesNota'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Comissao")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Comissao"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Comissao'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Compensa")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Compensa"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Compensa'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from ContasAPagar")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table ContasAPagar"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'ContasAPagar'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from DespesasLanc")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table DespesasLanc"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'DespesasLanc'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from FechamentoDiario")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table FechamentoDiario"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'FechamentoDiario'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from FormaDePagamentoRecebido")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table FormaDePagamentoRecebido"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'FormaDePagamentoRecebido'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from ProdutosEntrada")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table ProdutosEntrada"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'ProdutosEntrada'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Regua")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Regua"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Regua'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Salario")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Salario"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Salario'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from TanquesMovimento")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table TanquesMovimento"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'TanquesMovimento'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from Venda")
If Err.Number = 0 Then
  On Error GoTo 0
  On Error Resume Next
  Set dbTemp = Nothing
  db.Execute "drop table Venda"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao remover a tabela 'Venda'"
  End If
End If
Set dbTemp = Nothing



On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from VendedoresPagamento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table VendedoresPagamento (codigoPagamento counter, DataCriado datetime, codigoVendedor double, Codigo double, Mes integer, Ano integer, Salario bit, Adiantamento bit, ValorBase currency, Vales Currency, Comissoes Currency, VR currency, VT currency, SaldoAPagar currency, Pago bit, CodigoCaixa double, DataCaixa datetime, Turno text(30), UsuarioCriou Text(50), UsuarioConfirmou Text(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'VendedoresPagamento'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from VendedoresPagamento order by Funcionario")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table VendedoresPagamento add column Funcionario Text(50)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'VendedoresPagamento->Funcionario'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from vales order by codigopagamento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Vales add column CodigoPagamento double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Vales->CodigoPagamento'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ProdutosNotas order by PgAntecipado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table ProdutosNotas add column PgAntecipado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ProdutosNotas->PgAntecipado'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from DespesasLanc2 order by PgAntecipado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table DespesasLanc2 add column PgAntecipado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'DespesasLanc2->PgAntecipado'"
  End If
End If
Set dbTemp = Nothing
Set db = Nothing
Set Ws = Nothing

End Function

Public Function AtualizaDbAntigo()
Dim db As Database, Ws As Workspace
Dim dbTemp As Recordset


Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from produtos order by Unidade")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table produtos add column Unidade Text(5)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->Unidade'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from notas order by codigoboleto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table notas add column CodigoBoleto double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Notas->CodigoBoleto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from config order by porta")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table config add column Porta Text(5)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->Porta'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from CuponsFiscais")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create table CuponsFiscais (CodigoCliente double, DataCupom datetime, HoraCupom Datetime, NumeroCupom Text(20), Placa text(15), Km double, Carro Text(25), QtdProduto double, Valortotal currency, CodigoProduto double, Tributo Text(5))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'CuponsFiscais'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next


Set dbTemp = db.OpenRecordset("select *from produtos order by DescriAbreviada")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table produtos add column DescriAbreviada text(15)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Produtos->DescriAbreviada'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientescobranca order by protestador")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientescobranca add column Protestador bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->Protestador'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from formadepagamento order by CodigoNoPosto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table formadepagamento add column CodigoNoPosto Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->CodigoNoPosto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from despesatipo order by CodigoNoPosto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Despesatipo add column CodigoNoPosto Text(10)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FormaDePagamento->CodigoNoPosto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientes order by UltimoAbastecimento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column UltimoAbastecimento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->UltimoAbastecimento'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from produtoshistorico order by EstoqueFinal")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table produtoshistorico add column EstoqueFinal double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosHistorico->EstoqueFinal'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ClientesNota2 order by CodigoProduto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ClientesNota2 add column CodigoProduto double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesNotas->CodigoProduto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Clientes order by CodigoNoPosto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column CodigoNoPosto double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->CodigoNoPosto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Fechamentodecaixa order by FinalizadoPor")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column FinalizadoPor Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->FinalizadoPor'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Fechamentodecaixa order by DistribuidoPor")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table FechamentoDeCaixa add column DistribuidoPor Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'FechamentoDeCaixa->DistribuidoPor'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Clientes order by UltimoAbastecimento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table Clientes add column UltimoAbastecimento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->UltimoAbastecimento'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from QChequesCx")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "drop table QChequesCx"
  On Error GoTo 0
  On Error Resume Next
  db.CreateQueryDef "QChequesCx", "SELECT cheques.*, fechamentodecaixa.* FROM cheques, fechamentodecaixa WHERE (((cheques.CodigoFechamento)=[fechamentodecaixa].[codigofechamento]));"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'QChequesCx'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Cheques order by Juros")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table cheques add column Juros currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Cheques->Juros'"
  End If
  db.Execute "update cheques set juros=valor-valornabomba"
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ChequesClientesCobraHistorico")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create table ChequesClientesCobraHistorico (codigohistorico counter, CodigoCliente double, LancadoEm datetime, Usuario text(30), Contato Text(50), Obs Text(255))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ChequesClientesCobraHistorico'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ChequesClientes order by DataCadastro")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table chequesclientes add column DataCadastro datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ChequesClientes->DataCadastro'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ChequesClientes order by DataDesativado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table chequesclientes add column DataDesativado datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ChequesClientes->DataDesativado'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ChequesClientes order by SaldoPendente")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table chequesclientes add column SaldoPendente currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ChequesClientes->SaldoPendente'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from produtosHistorico")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create table ProdutosHistorico (codigohistorico counter, LancadoEm datetime, DataAlteracao Datetime, CodigoProduto double, Codigo double, DescriProduto Text(50), DescriOperacao Text(100), PrecoCompra currency, PrecoVenda Currency, EstoqueAnterior Double, Quantidade double)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ProdutosHistorico'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from produtosaltera order by Alterado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table ProdutosAltera add column Alterado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ProdutosAltera->Alterado'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from config")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table config (ftp text(254))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Config'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from config order by ftp")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table config add column ftp text(254)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Config->FTP'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientes order by PodeOleo")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column PodeOleo bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->PodeOleo'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientes order by PodeCombustivel")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column PodeCombustivel bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->PodeCombustivel'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientes order by PodeLavagem")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column PodeLavagem bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'Clientes->PodeLavagem'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientescobranca order by protestado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientescobranca add column Protestado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesCobranca->Protestado'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("Select *from clientes order by protestado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column Protestado bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela Clientes->Protestado"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("Select *from clientescobranca order by Dataprotestado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientescobranca add column DataProtestado date"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela Clientes->DataProtestado"
  End If
End If

Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from clientes order by limitar")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column Limitar bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & "-" & Err.Description & " ao alterar a tabela Clientes->Limitar!"
  End If
End If

Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from clientes order by limite")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "alter table clientes add column Limite currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & "-" & Err.Description & " ao alterar a tabela Clientes->Limite!"
  End If
End If

On Error GoTo 0
On Error Resume Next
Set dbTemp = db.OpenRecordset("select *from clientesnota2 order by ValorTotalDif")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table clientesnota2 add column ValorTotalDif currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesNota2->ValorTotalDif'"
  End If
End If
Set dbTemp = Nothing

On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from status order by ClienteDiferenciado")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table status add column ClienteDiferenciado currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Status->ClienteDiferenciado'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ClientesProdutos")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "create table ClientesProdutos (CodigoCliente double, codigoproduto double, codProduto double, Descri Text(50), Preco currency, ValorASomar currency, Porcento double, Validade datetime, CodigoTurno double, Turno Text(20), Grupo Double)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesProdutos'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from ClientesProdutos order by HoraIni")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table ClientesProdutos add column HoraIni datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesProdutos->HoraIni'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Clientes order by FormaDePagamento")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Clientes add column FormaDePagamento Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Clientes->FormaDePagamento'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by LinhasCorpo")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table configNota add column LinhasCorpo double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->LinhasCorpo'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Notas order by Servico")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Notas add column Servico Text(100)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Notas->Servico'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Notas order by ServicoISS")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Notas add column ServicoISS Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Notas->ServicoISS'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Notas order by ServicoTotal")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Notas add column ServicoTotal Currency "
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Notas->ServicoTotal'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Produtos order by Servico")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Produtos add column Servico bit"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Produtos->Servico'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from Produtos order by DescriServico")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table Produtos add column DescriServico Text(100)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'Produtos->DescriServico'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoX")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoX double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoX'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoY")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoY double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoY'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoISSX")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoISSX double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoISSX'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoISSY")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoISSY double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoISSY'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoTotalX")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoTotalX double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoTotalX'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from confignota order by PrestacaoServicoTotalY")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table confignota add column PrestacaoServicoTotalY double"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ConfigNota->PrestacaoServicoTotalY'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from chequesclientes order by CadastradoPor")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table ChequesClientes add column CadastradoPor Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ChequesClientes->CadastradoPor'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from cheques order by UsuarioLanc")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table cheques add column UsuarioLanc Text(30)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'cheques->UsuarioLanc'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from DespesaTipoSubGrupo")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create table DespesaTipoSubGrupo (CodigoDespesaTipoSub counter, CodigoDespesaTipo double, Descri Text(100))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'DespesaTipoSubGrupo'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from JurosBoleto")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create table JurosBoleto (CodigoJuros counter, Inicio double, Final double, Juros double, JurosValor Currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'JurosBoleto'"
  End If
End If
Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next

Set dbTemp = db.OpenRecordset("select *from clientescobranca order by jurosdevido")
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Alter table clientescobranca add column JurosDevido currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesCobranca->JurosDevido'"
  End If
End If
Set dbTemp = Nothing
Set db = Nothing
Set Ws = Nothing

End Function

Public Function AtualizaDbQuery()
Dim db As Database, Ws As Workspace
Dim dbTemp As Recordset


Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho, , , Conectar)

On Error GoTo 0
On Error Resume Next
db.Execute "drop table qVendas"
On Error GoTo 0
On Error Resume Next

db.CreateQueryDef "qVendas", "SELECT Vendedores.*, Venda2.*, Fechamentodecaixa.* FROM (Venda2 INNER JOIN Fechamentodecaixa ON Venda2.CodigoFechamento = Fechamentodecaixa.CodigoFechamento) INNER JOIN Vendedores ON Venda2.CodigoVendedor = Vendedores.Codigo;"
If Err.Number <> 0 Then
  MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'QChequesCx'"
End If

Set dbTemp = Nothing
On Error GoTo 0
On Error Resume Next


Set dbTemp = Nothing
Set db = Nothing
Set Ws = Nothing

End Function


Private Sub CriaPostoINI()
A = FreeFile()
Open App.Path & "\Posto.ini" For Output As #A
Print #A, "; 1 - permite lançar nototas de cliente so com cliente - nrcupom - valor"
Print #A, "; 0 - tem que lançar tudo"
Print #A, "[Notas no Caixa]"
Print #A, "Nocaixa=0"
Print #A, "; 1 - aceita qualquer cheque"
Print #A, "; 0 - tem que estar cadastrado e ativo"
Print #A, "[cheques]"
Print #A, "Cheques=0"
Close #A
End Sub

Public Sub VerificaAcesso()
With Usuarios.Grupo
  'Cadastro
  If .CadBomba = 0 Then
    mdiPosto.mnuCadBicos.Enabled = False
  Else
    mdiPosto.mnuCadBicos.Enabled = True
  End If
  If .CadCliente = 0 Then
    mdiPosto.mnuCadClientes.Enabled = False
  Else
    mdiPosto.mnuCadClientes.Enabled = True
  End If
  If .CadClienteCheque = 0 Then
    mdiPosto.mnuCadClienteCheque.Enabled = False
  Else
    mdiPosto.mnuCadClienteCheque.Enabled = True
  End If
  If .CadConta = 0 Then
    mdiPosto.mnuCadContas.Enabled = False
  Else
    mdiPosto.mnuCadContas.Enabled = True
  End If
  If .CadDespesaTipo = 0 Then
    mdiPosto.mnuCadDespesasTipo.Enabled = False
  Else
    mdiPosto.mnuCadDespesasTipo.Enabled = True
  End If
  If .CadDespesaBancaria = 0 Then
    mdiPosto.mnuCadDespesaBanco.Enabled = False
  Else
    mdiPosto.mnuCadDespesaBanco.Enabled = True
  End If
  If .CadFormaDePg = 0 Then
    mdiPosto.mnuCadFormaDePg.Enabled = False
  Else
    mdiPosto.mnuCadFormaDePg.Enabled = True
  End If
  If .CadFornecedores = 0 Then
    mdiPosto.mnuCadFornecedor.Enabled = False
  Else
    mdiPosto.mnuCadFornecedor.Enabled = True
  End If
  If .CadFuncionarios = 0 Then
    mdiPosto.mnuCadVendedores.Enabled = False
  Else
    mdiPosto.mnuCadVendedores.Enabled = True
  End If
  If .CadJuros = 0 Then
    mdiPosto.mnuCadJuros.Enabled = False
  Else
    mdiPosto.mnuCadJuros.Enabled = True
  End If
  If .CadPostos = 0 Then
    mdiPosto.mnuCadPostos.Enabled = False
  Else
    mdiPosto.mnuCadPostos.Enabled = True
  End If
  If .CadProdutos = 0 Then
    mdiPosto.mnuCadProdutos.Enabled = False
  Else
    mdiPosto.mnuCadProdutos.Enabled = True
  End If
  If .CadProdutosFornecedores = 0 Then
    mdiPosto.mnuCadProdutosFornecedores.Enabled = False
  Else
    mdiPosto.mnuCadProdutosFornecedores.Enabled = True
  End If
  If .CadTanques = 0 Then
    mdiPosto.mnuCadTanque.Enabled = False
  Else
    mdiPosto.mnuCadTanque.Enabled = True
  End If
  If .CadTurnos = 0 Then
    mdiPosto.mnuCadTurnos.Enabled = False
  Else
    mdiPosto.mnuCadTurnos.Enabled = True
  End If
  If .CadConfiguracao = 0 Then
    mdiPosto.mnuCadConfigura.Enabled = False
  Else
    mdiPosto.mnuCadConfigura.Enabled = True
  End If
  'Controle
  If .ControleFechamentoDiario = 0 Then
    mdiPosto.mnuControleFechaDia.Enabled = False
  Else
    mdiPosto.mnuControleFechaDia.Enabled = True
  End If
  If .ControleConferencia = 0 Then
    mdiPosto.mnuConferencia.Enabled = False
  Else
    mdiPosto.mnuConferencia.Enabled = True
  End If
  If .ControleCartoes = 0 Then
    mdiPosto.mnuControleRecebimentos.Enabled = False
  Else
    mdiPosto.mnuControleRecebimentos.Enabled = True
  End If
  If .ControlePgAntecipado = 0 Then
    mdiPosto.mnuControlePgAntecipado.Enabled = False
  Else
    mdiPosto.mnuControlePgAntecipado.Enabled = True
  End If
  If .ControleNotas = 0 Then
    mdiPosto.mnuControleNotas.Enabled = False
  Else
    mdiPosto.mnuControleNotas.Enabled = True
  End If
  If .ControleLancContas = 0 Then
    mdiPosto.mnuControleDespesaLanc = False
  Else
    mdiPosto.mnuControleDespesaLanc = True
  End If
  If .ControleContasPg = 0 Then
    mdiPosto.mnuControlePgDespesa.Enabled = False
  Else
    mdiPosto.mnuControlePgDespesa.Enabled = True
  End If
  If .ControleCobranca = 0 Then
    mdiPosto.mnuControleCobra.Enabled = False
  Else
    mdiPosto.mnuControleCobra.Enabled = True
  End If
  If .ControleCobranca = 0 Then
    mdiPosto.mnuControleFaturaCliente.Enabled = False
  Else
    mdiPosto.mnuControleFaturaCliente.Enabled = True
  End If
  If .CadProdutos = 0 Then
    mdiPosto.mnuControlePrecos.Enabled = False
  Else
    mdiPosto.mnuControlePrecos.Enabled = True
  End If
  If .ControleAgua = 0 Then
    mdiPosto.mnuControleAgua.Enabled = False
  Else
    mdiPosto.mnuControleAgua.Enabled = True
  End If
  If .ControleLuz = 0 Then
    mdiPosto.mnuControleLuz.Enabled = False
  Else
    mdiPosto.mnuControleLuz.Enabled = True
  End If
  If .ControleLavagem = 0 Then
    mdiPosto.mnuControleLavagem.Enabled = False
  Else
    mdiPosto.mnuControleLavagem.Enabled = True
  End If
  If .ControleVales = 0 Then
    mdiPosto.mnuControleVales.Enabled = False
  Else
    mdiPosto.mnuControleVales.Enabled = True
  End If
  If .ControleFechamentoDiario = 0 Then
    mdiPosto.mnuControleVendasLeituraX.Enabled = False
  Else
    mdiPosto.mnuControleVendasLeituraX.Enabled = True
  End If
  'Cheques
  If .ChequeDeposito = 0 Then
    mdiPosto.mnuBancoDepositaCheque.Enabled = False
  Else
    mdiPosto.mnuBancoDepositaCheque.Enabled = True
  End If
  If .ChequeDevolucao = 0 Then
    mdiPosto.mnuBancoDevolucao.Enabled = False
  Else
    mdiPosto.mnuBancoDevolucao.Enabled = True
  End If
  If .ChequeCobranca = 0 Then
    mdiPosto.mnuBancoCobraCheque.Enabled = False
  Else
    mdiPosto.mnuBancoCobraCheque.Enabled = True
  End If
  If .ChequeProtesto = 0 Then
    mdiPosto.mnuChequeProtesto.Enabled = False
  Else
    mdiPosto.mnuChequeProtesto.Enabled = True
  End If
  If .ChequeEnviarPEmpresaCobranca = 0 Then
    mdiPosto.mnuChequeEmpresaDeCobranca.Enabled = False
  Else
    mdiPosto.mnuChequeEmpresaDeCobranca.Enabled = True
  End If
  If .ChequeEnviarPEmpresaCobranca = 0 Then
    mdiPosto.mnuChequesResgatados.Enabled = False
  Else
    mdiPosto.mnuChequesResgatados.Enabled = True
  End If
  If .ChequePorData = 0 Then
    mdiPosto.mnuBancoChequesData.Enabled = False
  Else
    mdiPosto.mnuBancoChequesData.Enabled = True
  End If
  'Banco
  If .BancoConcilia = 0 Then
    mdiPosto.mnuControleConcilia.Enabled = False
  Else
    mdiPosto.mnuControleConcilia.Enabled = True
  End If
  If .BancoTransfere = 0 Then
    mdiPosto.mnuBancoTransfere.Enabled = False
  Else
    mdiPosto.mnuBancoTransfere.Enabled = True
  End If
  'Relatórios
  If .RelatAcertoEstoque = 0 Then
    mdiPosto.mnuRelatAcertoEstoque.Enabled = False
  Else
    mdiPosto.mnuRelatAcertoEstoque.Enabled = True
  End If
  If .RelatChequeCliente = 0 Then
    mdiPosto.mnuRelatChequeCliente.Enabled = False
  Else
    mdiPosto.mnuRelatChequeCliente.Enabled = True
  End If
  If .RelatProdutosComprados = 0 Then
    mdiPosto.mnuRelatCompras.Enabled = False
  Else
    mdiPosto.mnuRelatCompras.Enabled = True
  End If
  If .RelatProdutosComprados = 0 Then
    mdiPosto.mnuRelatExtratoProdutos.Enabled = False
  Else
    mdiPosto.mnuRelatExtratoProdutos.Enabled = True
  End If
  If .RelatCompraVenda = 0 Then
    mdiPosto.mnuRelatComprasProd.Enabled = False
  Else
    mdiPosto.mnuRelatComprasProd.Enabled = True
  End If
  If .RelatDifCaixa = 0 Then
    mdiPosto.mnuRelatDiferencaCaixa.Enabled = False
  Else
    mdiPosto.mnuRelatDiferencaCaixa.Enabled = True
  End If
  If .RelatDifRecebe = 0 Then
    mdiPosto.mnuRelatDifRecebimentos.Enabled = False
  Else
    mdiPosto.mnuRelatDifRecebimentos.Enabled = True
  End If
  If .RelatDifCombustivel = 0 Then
    mdiPosto.mnuRelatDifComb.Enabled = False
  Else
    mdiPosto.mnuRelatDifComb.Enabled = True
  End If
  If .RelatFormaDePg = 0 Then
    mdiPosto.mnuRelatFormaDePg.Enabled = False
    mdiPosto.mnuRelatFormaDePgBordero.Enabled = False
  Else
    mdiPosto.mnuRelatFormaDePg.Enabled = True
    mdiPosto.mnuRelatFormaDePgBordero.Enabled = True
  End If
  If .RelatGalonagem = 0 Then
    mdiPosto.mnuRelatGalonagem.Enabled = False
  Else
    mdiPosto.mnuRelatGalonagem.Enabled = True
  End If
  If .RelatGalonagemTotal = 0 Then
    mdiPosto.mnuRelatGalonagemTotal.Enabled = False
  Else
    mdiPosto.mnuRelatGalonagemTotal.Enabled = True
  End If
  If .RelatVendaProdutos = 0 Then
    mdiPosto.mnuRelatVendas.Enabled = False
  Else
    mdiPosto.mnuRelatVendas.Enabled = True
  End If
  If .RelatVendaDetalhada = 0 Then
    mdiPosto.mnuRelatVendaDetalhada.Enabled = False
  Else
    mdiPosto.mnuRelatVendaDetalhada.Enabled = True
  End If
  If .RelatVendaLucro = 0 Then
    mdiPosto.mnuRelatVendaComissoes.Enabled = False
  Else
    mdiPosto.mnuRelatVendaComissoes.Enabled = True
  End If
  If .RelatVendaMedia = 0 Then
    mdiPosto.mnuRelatMediaVenda.Enabled = False
  Else
    mdiPosto.mnuRelatMediaVenda.Enabled = True
  End If
  If .RelatDiariaCombustivel = 0 Then
    mdiPosto.mnuRelatVendaDiariaCombustivel.Enabled = False
  Else
    mdiPosto.mnuRelatVendaDiariaCombustivel.Enabled = True
  End If
  If .RelatProtestoDeCheques = 0 Then
    mdiPosto.mnuRelatProtesto.Enabled = False
  Else
    mdiPosto.mnuRelatProtesto.Enabled = True
  End If
  If .RelatCadastroIncompleto = 0 Then
    mdiPosto.mnuRelatCadIncompleto.Enabled = False
  Else
    mdiPosto.mnuRelatCadIncompleto.Enabled = True
  End If
  If .RelatRetornoCombustivel = 0 Then
    mdiPosto.mnuRelatRetorno.Enabled = False
  Else
    mdiPosto.mnuRelatRetorno.Enabled = True
  End If
  If .RelatFaturamentoCheques = 0 Then
    mdiPosto.mnuRelatFaturaCheque.Enabled = False
  Else
    mdiPosto.mnuRelatFaturaCheque.Enabled = True
  End If
  If .RelatKilometragem = 0 Then
    mdiPosto.mnuRelatKilometragem.Enabled = False
  Else
    mdiPosto.mnuRelatKilometragem.Enabled = True
  End If
  If .RelatKilometragem = 0 Then
    mdiPosto.mnuRelatEstacionamento.Enabled = False
  Else
    mdiPosto.mnuRelatEstacionamento.Enabled = True
  End If
  'Administração
  If .AdmConfirma = 0 Then
    mdiPosto.mnuAdmConfirmaDespesa.Enabled = False
  Else
    mdiPosto.mnuAdmConfirmaDespesa.Enabled = True
  End If
  If .AdmEstatus = 0 Then
    mdiPosto.mnuAdmEstatus.Enabled = False
  Else
    mdiPosto.mnuAdmEstatus.Enabled = True
  End If
  If .AdmEstatus = 0 Then
    mdiPosto.mnuAdmBloqueiaFinaliza.Enabled = False
  Else
    mdiPosto.mnuAdmBloqueiaFinaliza.Enabled = True
  End If
  If .AdmEstatus = 0 Then
    mdiPosto.mnuRelatFaturaCliente.Enabled = False
  Else
    mdiPosto.mnuRelatFaturaCliente.Enabled = True
  End If
  If .AdmTotalVenda = 0 Then
    mdiPosto.mnuAdmTotalVenda.Enabled = False
  Else
    mdiPosto.mnuAdmTotalVenda.Enabled = True
  End If
  If .AdmLMC = 0 Then
    mdiPosto.mnuAdmLMC.Enabled = False
  Else
    mdiPosto.mnuAdmLMC.Enabled = True
  End If
  If .AdmUsuarios = 0 Then
    mdiPosto.mnuAdmUsuarios.Enabled = False
  Else
    mdiPosto.mnuAdmUsuarios.Enabled = True
  End If
  If .AdmUsuariosGrupos = 0 Then
    mdiPosto.mnuAdmGrupos.Enabled = False
  Else
    mdiPosto.mnuAdmGrupos.Enabled = True
  End If
End With
End Sub

Public Sub ImprimeADOGrid(ByVal Grade1 As DataGrid, ByVal ImprimeEm As Printer, ByVal Dados As Adodc, Optional ColunaTotal As Integer = -1, Optional Linhas As Boolean = True, Optional LinhaLargura As Integer = 1, Optional ColunaDeQuebra As Integer = -1, Optional ColunaTotal2 As Integer = -1, Optional ColunaTotal3 As Integer = -1, Optional Titulo1 As String = "", Optional Titulo2 As String = "", Optional Titulo3 As String = "", Optional ColunaTotal4 As Integer = -1, Optional ColunaTotal5 As Integer = -1, Optional ColunaTotal6 As Integer = -1, Optional ColunaTotal7 As Integer = -1, Optional ColunaTotal8 As Integer = -1)
Dim InicioX As Double, InicioY As Double, LarguraMaxima As Double
Dim Total As Double, StrTemp As String, Grade As DataGrid, StrQuebra As String
Dim Total2 As Double, Total3 As Double, Total4 As Double, Total5 As Double, Total6 As Double
Dim Total7 As Double, Total8 As Double
Dim fimTotal As Double, fimTotal2 As Double, fimTotal3 As Double
Dim fimTotal4 As Double, fimTotal5 As Double, fimTotal6 As Double
Dim fimTotal7 As Double, fimTotal8 As Double

Set Grade = Grade1
ImprimeEm.ScaleMode = vbTwips
With ImprimeEm.Font
  .Name = Grade.Font.Name
  .Size = Grade.Font.Size
  .Bold = Grade.Font.Bold
  .Italic = Grade.Font.Italic
End With
If Dados.Recordset.RecordCount = 0 Then Exit Sub

With Grade
  ImprimeEm.DrawWidth = LinhaLargura
  For i = 0 To .Columns.Count - 1
    If .Columns(i).Width > 200 Then
      LarguraMaxima = LarguraMaxima + .Columns(i).Width
    End If
  Next i
  If .Caption <> "" Then
    Titulo3 = .Caption & Titulo3
  End If
  
  If Titulo1 <> "" Then
    ImprimeEm.FontSize = 14
    ImprimeEm.FontBold = True
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo1) / 2)
    ImprimeEm.Print Titulo1
    ImprimeEm.Font = Grade.Font
  End If
  If Titulo2 <> "" Then
    ImprimeEm.FontSize = 10
    ImprimeEm.FontBold = True
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo2) / 2)
    ImprimeEm.Print Titulo2
  End If
  If Titulo3 <> "" Then
    ImprimeEm.FontSize = 10
    ImprimeEm.FontBold = False
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo3) / 2)
    ImprimeEm.Print Titulo3
  End If
  ImprimeEm.Font.Bold = Grade.Font.Bold
  ImprimeEm.Font.Size = Grade.Font.Size
  
  InicioY = ImprimeEm.CurrentY
  If Linhas = True Then
    ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  End If
  For i = 0 To .Columns.Count - 1
    If .Columns(i).Width > 200 Then
      If .Columns(i).Alignment = 1 Then
        ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
      Else
        ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
      End If
      ImprimeEm.Print Grade1.Columns(i).Caption;
      InicioX = InicioX + Grade1.Columns(i).Width
    End If
  Next i
  
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
  ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  
  InicioX = 0
  Dados.Recordset.MoveFirst
  If ColunaDeQuebra >= 0 Then
    StrQuebra = .Columns(ColunaDeQuebra).Text
  End If
  TamMaximo = ImprimeEm.ScaleHeight - (Grade.RowHeight * 5)
  Do While Dados.Recordset.EOF = False
    If ImprimeEm.CurrentY >= TamMaximo Then
      InicioX = 0
      If Linhas = True Then
        For i = 0 To .Columns.Count - 1
          If .Columns(i).Width > 200 Then
            ImprimeEm.CurrentX = InicioX
            ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
            InicioX = InicioX + .Columns(i).Width
          End If
        Next i
        ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
      Else
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      End If
      
      If ColunaTotal >= 0 Then
        InicioX = 0
        A = ColunaTotal
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total
        Else
          StrTemp = Format(Total, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal2 >= 0 Then
        InicioX = 0
        A = ColunaTotal2
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total2
        Else
          StrTemp = Format(Total2, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal2 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal3 >= 0 Then
        InicioX = 0
        A = ColunaTotal3
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total3
        Else
          StrTemp = Format(Total3, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal3 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal4 >= 0 Then
        InicioX = 0
        A = ColunaTotal4
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total4
        Else
          StrTemp = Format(Total4, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal4 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal5 >= 0 Then
        InicioX = 0
        A = ColunaTotal5
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total5
        Else
          StrTemp = Format(Total5, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal5 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal6 >= 0 Then
        InicioX = 0
        A = ColunaTotal6
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total6
        Else
          StrTemp = Format(Total6, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal6 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal7 >= 0 Then
        InicioX = 0
        A = ColunaTotal7
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total7
        Else
          StrTemp = Format(Total7, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal7 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal8 >= 0 Then
        InicioX = 0
        A = ColunaTotal8
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Total8
        Else
          StrTemp = Format(Total8, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal8 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        ImprimeEm.Print StrTemp;
      End If
      
      
      ImprimeEm.Print ""
      ImprimeEm.CurrentX = 0
      ImprimeEm.Print "Página " & ImprimeEm.Page
      
      InicioX = 0
      ImprimeEm.NewPage
      
      If Titulo1 <> "" Then
        ImprimeEm.FontSize = 14
        ImprimeEm.FontBold = True
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo1) / 2)
        ImprimeEm.Print Titulo1
      End If
      If Titulo2 <> "" Then
        ImprimeEm.FontSize = 10
        ImprimeEm.FontBold = True
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo2) / 2)
        ImprimeEm.Print Titulo2
      End If
      If Titulo3 <> "" Then
        ImprimeEm.FontSize = 10
        ImprimeEm.FontBold = False
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo3) / 2)
        ImprimeEm.Print Titulo3
      End If
      ImprimeEm.Font.Bold = Grade.Font.Bold
      ImprimeEm.Font.Size = Grade.Font.Size
      
      ImprimeEm.CurrentX = 0
      InicioY = ImprimeEm.CurrentY
      If Linhas = True Then
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      End If
      For i = 0 To .Columns.Count - 1
        If .Columns(i).Width > 200 Then
          If .Columns(i).Alignment = 1 Then
            ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
          Else
            ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
          End If
          ImprimeEm.Print .Columns(i).Caption;
          InicioX = InicioX + .Columns(i).Width
        End If
      Next i
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
      ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      InicioX = 0
    End If
    
    If ColunaDeQuebra >= 0 Then
      If StrQuebra <> .Columns(ColunaDeQuebra).Text Then
        InicioX = 0
        If Linhas = True Then
          For i = 0 To .Columns.Count - 1
            If .Columns(i).Width > 200 Then
              ImprimeEm.CurrentX = InicioX
              ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
              InicioX = InicioX + .Columns(i).Width
            End If
          Next i
          ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
        Else
          ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
          ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        End If
        
        If ColunaTotal >= 0 Then
          InicioX = 0
          A = ColunaTotal
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total
          Else
            StrTemp = Format(Total, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
            ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal2 >= 0 Then
          InicioX = 0
          A = ColunaTotal2
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total2
          Else
            StrTemp = Format(Total2, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal2 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal3 >= 0 Then
          InicioX = 0
          A = ColunaTotal3
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total3
          Else
            StrTemp = Format(Total3, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal3 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal4 >= 0 Then
          InicioX = 0
          A = ColunaTotal4
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total4
          Else
            StrTemp = Format(Total4, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal4 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
            ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal5 >= 0 Then
          InicioX = 0
          A = ColunaTotal5
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total5
          Else
            StrTemp = Format(Total5, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal5 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal6 >= 0 Then
          InicioX = 0
          A = ColunaTotal6
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total6
          Else
            StrTemp = Format(Total6, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal6 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal7 >= 0 Then
          InicioX = 0
          A = ColunaTotal7
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total7
          Else
            StrTemp = Format(Total7, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal7 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal8 >= 0 Then
          InicioX = 0
          A = ColunaTotal8
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Total8
          Else
            StrTemp = Format(Total8, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal8 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          ImprimeEm.Print StrTemp;
        End If
        
        InicioX = 0
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        InicioY = ImprimeEm.CurrentY
        ImprimeEm.CurrentX = 0
        
        
        InicioY = ImprimeEm.CurrentY
        If Linhas = True Then
          ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
          ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
          ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        End If
        For i = 0 To .Columns.Count - 1
          If .Columns(i).Width > 200 Then
            If .Columns(i).Alignment = 1 Then
              ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
            Else
              ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
            End If
            ImprimeEm.Print .Columns(i).Caption;
            InicioX = InicioX + .Columns(i).Width
          End If
        Next i
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        InicioX = 0
        If IsNumeric(Total) = True Then
          Total = 0
        Else
          Total = ""
        End If
        If IsNumeric(Total2) = True Then
          Total2 = 0
        Else
          Total2 = ""
        End If
        If IsNumeric(Total3) = True Then
          Total3 = 0
        Else
          Total3 = ""
        End If
        If IsNumeric(Total4) = True Then
          Total4 = 0
        Else
          Total4 = ""
        End If
        If IsNumeric(Total5) = True Then
          Total5 = 0
        Else
          Total5 = ""
        End If
        If IsNumeric(Total6) = True Then
          Total6 = 0
        Else
          Total6 = ""
        End If
        If IsNumeric(Total7) = True Then
          Total7 = 0
        Else
          Total7 = ""
        End If
        If IsNumeric(Total8) = True Then
          Total8 = 0
        Else
          Total8 = ""
        End If
        
      End If
    End If
    
    For i = 0 To .Columns.Count - 1
      If .Columns(i).Width > 200 Then
        If .Columns(i).Alignment = 1 Then
          ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Text & " ")
        Else
          ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
        End If
        ImprimeEm.Print .Columns(i).Text;
        InicioX = InicioX + .Columns(i).Width
      End If
    Next i
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
    If Linhas = True Then
      ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
    End If
    InicioX = 0
    If ColunaTotal >= 0 Then
      If IsNumeric(.Columns(ColunaTotal).Text) = True Then
        Total = Total + CDbl(.Columns(ColunaTotal).Text)
        fimTotal = fimTotal + CDbl(.Columns(ColunaTotal).Text)
      Else
        If Len(.Columns(ColunaTotal).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal).Text, Len(.Columns(ColunaTotal).Text) - 2, 1) = ":" Then
            Total = SomaHora(Total, .Columns(ColunaTotal).Text)
            fimTotal = SomaHora(fimTotal, .Columns(ColunaTotal).Text)
          Else
            If .Columns(ColunaTotal).Text <> "" Then
              Total = Total + 1
              fimTotal = fimTotal + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal).Text <> "" Then
            Total = Total + 1
            fimTotal = fimTotal + 1
          End If
        End If
      End If
    End If
    If ColunaTotal2 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal2).Text) = True Then
        Total2 = Total2 + CDbl(.Columns(ColunaTotal2).Text)
        fimTotal2 = fimTotal2 + CDbl(.Columns(ColunaTotal2).Text)
      Else
        If Len(.Columns(ColunaTotal2).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal2).Text, Len(.Columns(ColunaTotal2).Text) - 2, 1) = ":" Then
            Total2 = SomaHora(Total2, .Columns(ColunaTotal2).Text)
            fimTotal2 = SomaHora(fimTotal2, .Columns(ColunaTotal2).Text)
          Else
            If .Columns(ColunaTotal2).Text <> "" Then
              Total2 = Total2 + 1
              fimTotal2 = fimTotal2 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal2).Text <> "" Then
            Total2 = Total2 + 1
            fimTotal2 = fimTotal2 + 1
          End If
        End If
      End If
    End If
    If ColunaTotal3 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal3).Text) = True Then
        Total3 = Total3 + CDbl(Trim(.Columns(ColunaTotal3).Text))
        fimTotal3 = fimTotal3 + CDbl(.Columns(ColunaTotal3).Text)
      Else
        If Len(.Columns(ColunaTotal3).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal3).Text, Len(.Columns(ColunaTotal3).Text) - 2, 1) = ":" Then
            Total3 = SomaHora(Total3, .Columns(ColunaTotal3).Text)
            fimTotal3 = SomaHora(fimTotal3, .Columns(ColunaTotal3).Text)
          Else
            If .Columns(ColunaTotal3).Text <> "" Then
              Total3 = Total3 + 1
              fimTotal3 = fimTotal3 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal3).Text <> "" Then
            Total3 = Total3 + 1
            fimTotal3 = fimTotal3 + 1
          End If
        End If
      End If
    End If
    
    If ColunaTotal4 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal4).Text) = True Then
        Total4 = Total4 + CDbl(.Columns(ColunaTotal4).Text)
        fimTotal4 = fimTotal4 + CDbl(.Columns(ColunaTotal4).Text)
      Else
        If Len(.Columns(ColunaTotal4).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal4).Text, Len(.Columns(ColunaTotal4).Text) - 2, 1) = ":" Then
            Total4 = SomaHora(Total4, .Columns(ColunaTotal4).Text)
            fimTotal4 = SomaHora(fimTotal4, .Columns(ColunaTotal4).Text)
          Else
            If .Columns(ColunaTotal4).Text <> "" Then
              Total4 = Total4 + 1
              fimTotal4 = fimTotal4 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal4).Text <> "" Then
            Total4 = Total4 + 1
            fimTotal4 = fimTotal4 + 1
          End If
        End If
      End If
    End If
    If ColunaTotal5 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal5).Text) = True Then
        Total5 = Total5 + CDbl(.Columns(ColunaTotal5).Text)
        fimTotal5 = fimTotal5 + CDbl(.Columns(ColunaTotal5).Text)
      Else
        If Len(.Columns(ColunaTotal5).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal5).Text, Len(.Columns(ColunaTotal5).Text) - 2, 1) = ":" Then
            Total5 = SomaHora(Total5, .Columns(ColunaTotal5).Text)
            fimTotal5 = SomaHora(fimTotal5, .Columns(ColunaTotal5).Text)
          Else
            If .Columns(ColunaTotal5).Text <> "" Then
              Total5 = Total5 + 1
              fimTotal5 = fimTotal5 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal5).Text <> "" Then
            Total5 = Total5 + 1
            fimTotal5 = fimTotal5 + 1
          End If
        End If
      End If
    End If
    If ColunaTotal6 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal6).Text) = True Then
        Total6 = Total6 + CDbl(.Columns(ColunaTotal6).Text)
        fimTotal6 = fimTotal6 + CDbl(.Columns(ColunaTotal6).Text)
      Else
        If Len(.Columns(ColunaTotal6).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal6).Text, Len(.Columns(ColunaTotal6).Text) - 2, 1) = ":" Then
            Total6 = SomaHora(Total6, .Columns(ColunaTotal6).Text)
            fimTotal6 = SomaHora(fimTotal6, .Columns(ColunaTotal6).Text)
          Else
            If .Columns(ColunaTotal6).Text <> "" Then
              Total6 = Total6 + 1
              fimTotal6 = fimTotal6 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal6).Text <> "" Then
            Total6 = Total6 + 1
            fimTotal6 = fimTotal6 + 1
          End If
        End If
      End If
    End If
    If ColunaTotal7 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal7).Text) = True Then
        Total7 = Total7 + CDbl(.Columns(ColunaTotal7).Text)
        fimTotal7 = fimTotal7 + CDbl(.Columns(ColunaTotal7).Text)
      Else
        If Len(.Columns(ColunaTotal7).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal7).Text, Len(.Columns(ColunaTotal7).Text) - 2, 1) = ":" Then
            Total7 = SomaHora(Total7, .Columns(ColunaTotal7).Text)
            fimTotal7 = SomaHora(fimTotal7, .Columns(ColunaTotal7).Text)
          Else
            If .Columns(ColunaTotal7).Text <> "" Then
              Total7 = Total7 + 1
              fimTotal7 = fimTotal7 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal7).Text <> "" Then
            Total7 = Total7 + 1
            fimTotal7 = fimTotal7 + 1
          End If
        End If
      End If
    End If
    If ColunaTotal8 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal8).Text) = True Then
        Total8 = Total8 + CDbl(.Columns(ColunaTotal8).Text)
        fimTotal8 = fimTotal8 + CDbl(.Columns(ColunaTotal8).Text)
      Else
        If Len(.Columns(ColunaTotal8).Text) >= 3 Then
          If Mid(.Columns(ColunaTotal8).Text, Len(.Columns(ColunaTotal8).Text) - 2, 1) = ":" Then
            Total8 = SomaHora(Total8, .Columns(ColunaTotal8).Text)
            fimTotal8 = SomaHora(fimTotal8, .Columns(ColunaTotal8).Text)
          Else
            If .Columns(ColunaTotal8).Text <> "" Then
              Total8 = Total8 + 1
              fimTotal8 = fimTotal8 + 1
            End If
          End If
        Else
          If .Columns(ColunaTotal8).Text <> "" Then
            Total8 = Total8 + 1
            fimTotal8 = fimTotal8 + 1
          End If
        End If
      End If
    End If
    
    If ColunaDeQuebra >= 0 Then
      StrQuebra = .Columns(ColunaDeQuebra).Text
    End If
    Dados.Recordset.MoveNext
  Loop
  InicioX = 0
  If Linhas = True Then
    For i = 0 To .Columns.Count - 1
      If .Columns(i).Width > 200 Then
        ImprimeEm.CurrentX = InicioX
        ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
        InicioX = InicioX + .Columns(i).Width
      End If
    Next i
    ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
  Else
    ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  End If
  
  If ColunaDeQuebra <> -1 Then
    If ColunaTotal >= 0 Then
      InicioX = 0
      A = ColunaTotal
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total
      Else
        StrTemp = Format(Total, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal2 >= 0 Then
      InicioX = 0
      A = ColunaTotal2
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total2
      Else
        StrTemp = Format(Total2, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal2 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal3 >= 0 Then
      InicioX = 0
      A = ColunaTotal3
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total3
      Else
        StrTemp = Format(Total3, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal3 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal4 >= 0 Then
      InicioX = 0
      A = ColunaTotal4
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total4
      Else
        StrTemp = Format(Total4, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal4 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
        ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal5 >= 0 Then
      InicioX = 0
      A = ColunaTotal5
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total5
      Else
        StrTemp = Format(Total5, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal5 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal6 >= 0 Then
      InicioX = 0
      A = ColunaTotal6
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total6
      Else
        StrTemp = Format(Total6, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal6 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal7 >= 0 Then
      InicioX = 0
      A = ColunaTotal7
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total7
      Else
        StrTemp = Format(Total7, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal7 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    
    If ColunaTotal8 >= 0 Then
      InicioX = 0
      A = ColunaTotal8
      If .Columns(A).NumberFormat = "" Then
        StrTemp = Total8
      Else
        StrTemp = Format(Total8, .Columns(A).NumberFormat)
      End If
      For A = 0 To ColunaTotal8 - 1
        InicioX = InicioX + .Columns(A).Width
      Next A
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
      ImprimeEm.Print StrTemp;
    End If
    ImprimeEm.Print ""
  End If
  
    
  If ColunaTotal >= 0 Then
    InicioX = 0
    A = ColunaTotal
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal
    Else
      StrTemp = Format(fimTotal, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal2 >= 0 Then
    InicioX = 0
    A = ColunaTotal2
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal2
    Else
      StrTemp = Format(fimTotal2, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal2 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal3 >= 0 Then
    InicioX = 0
    A = ColunaTotal3
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal3
    Else
      StrTemp = Format(fimTotal3, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal3 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  
  If ColunaTotal4 >= 0 Then
    InicioX = 0
    A = ColunaTotal4
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal4
    Else
      StrTemp = Format(fimTotal4, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal4 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal5 >= 0 Then
    InicioX = 0
    A = ColunaTotal5
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal5
    Else
      StrTemp = Format(fimTotal5, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal5 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal6 >= 0 Then
    InicioX = 0
    A = ColunaTotal6
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal6
    Else
      StrTemp = Format(fimTotal6, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal6 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal7 >= 0 Then
    InicioX = 0
    A = ColunaTotal7
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal7
    Else
      StrTemp = Format(fimTotal7, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal7 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal8 >= 0 Then
    InicioX = 0
    A = ColunaTotal8
    If .Columns(A).NumberFormat = "" Then
      StrTemp = fimTotal8
    Else
      StrTemp = Format(fimTotal8, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal8 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    ImprimeEm.Print StrTemp;
  End If
  
  
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
End With
End Sub

Public Sub ImprimeADOGrid2(ByVal Grade1 As DataGrid, ByVal ImprimeEm As Printer, ByVal Dados As Adodc, Optional ColunaTotal As Integer = -1, Optional Linhas As Boolean = True, Optional LinhaLargura As Integer = 1, Optional ColunaDeQuebra As Integer = -1, Optional ColunaTotal2 As Integer = -1, Optional ColunaTotal3 As Integer = -1, Optional Titulo1 As String = "", Optional Titulo2 As String = "", Optional Titulo3 As String = "")
Dim InicioX As Double, InicioY As Double, LarguraMaxima As Double
Dim Total As Double, StrTemp As String, Grade As DataGrid, StrQuebra As String
Dim Total2 As Double, Total3 As Double
Set Grade = Grade1
ImprimeEm.ScaleMode = vbTwips
With ImprimeEm.Font
  .Name = Grade.Font.Name
  .Size = Grade.Font.Size
  .Bold = Grade.Font.Bold
  .Italic = Grade.Font.Italic
End With
If Dados.Recordset.RecordCount = 0 Then Exit Sub
With Grade
  ImprimeEm.DrawWidth = LinhaLargura
  For i = 0 To .Columns.Count - 1
    If .Columns(i).Width > 200 Then
      LarguraMaxima = LarguraMaxima + .Columns(i).Width
    End If
  Next i
  If .Caption <> "" Then
    Titulo3 = .Caption & Titulo3
  End If
  
  If Titulo1 <> "" Then
    ImprimeEm.FontSize = 14
    ImprimeEm.FontBold = True
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo1) / 2)
    ImprimeEm.Print Titulo1
    ImprimeEm.Font = Grade.Font
  End If
  If Titulo2 <> "" Then
    ImprimeEm.FontSize = 10
    ImprimeEm.FontBold = True
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo2) / 2)
    ImprimeEm.Print Titulo2
  End If
  If Titulo3 <> "" Then
    ImprimeEm.FontSize = 10
    ImprimeEm.FontBold = False
    ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo3) / 2)
    ImprimeEm.Print Titulo3
  End If
  ImprimeEm.Font.Bold = Grade.Font.Bold
  ImprimeEm.Font.Size = Grade.Font.Size
  
  InicioY = ImprimeEm.CurrentY
  If Linhas = True Then
    ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  End If
  For i = 0 To .Columns.Count - 1
    If .Columns(i).Width > 200 Then
      If .Columns(i).Alignment = 1 Then
        ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
      Else
        ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
      End If
      ImprimeEm.Print Grade1.Columns(i).Caption;
      InicioX = InicioX + Grade1.Columns(i).Width
    End If
  Next i
  
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
  ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  
  InicioX = 0
  Dados.Recordset.MoveFirst
  If ColunaDeQuebra >= 0 Then
    StrQuebra = .Columns(ColunaDeQuebra).Text
  End If
  TamMaximo = ImprimeEm.ScaleHeight - (Grade.RowHeight * 6)
  Do While Dados.Recordset.EOF = False
    If ImprimeEm.CurrentY >= TamMaximo Then
      InicioX = 0
      If Linhas = True Then
        For i = 0 To .Columns.Count - 1
          If .Columns(i).Width > 200 Then
            ImprimeEm.CurrentX = InicioX
            ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
            InicioX = InicioX + .Columns(i).Width
          End If
        Next i
        ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
      Else
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      End If
      
      If ColunaTotal >= 0 Then
        InicioX = 0
        A = ColunaTotal
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Str(Total)
        Else
          StrTemp = Format(Total, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        If .Columns(A).Alignment = 1 Then
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        Else
          ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
        End If
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal2 >= 0 Then
        InicioX = 0
        A = ColunaTotal2
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Str(Total2)
        Else
          StrTemp = Format(Total2, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal2 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        If .Columns(A).Alignment = 1 Then
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        Else
          ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
        End If
        ImprimeEm.Print StrTemp;
      End If
      
      If ColunaTotal3 >= 0 Then
        InicioX = 0
        A = ColunaTotal3
        If .Columns(A).NumberFormat = "" Then
          StrTemp = Str(Total3)
        Else
          StrTemp = Format(Total3, .Columns(A).NumberFormat)
        End If
        For A = 0 To ColunaTotal3 - 1
          InicioX = InicioX + .Columns(A).Width
        Next A
        If .Columns(A).Alignment = 1 Then
          ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
        Else
          ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
        End If
        ImprimeEm.Print StrTemp;
      End If
      
      ImprimeEm.Print ""
      ImprimeEm.CurrentX = 0
      ImprimeEm.Print "Página " & ImprimeEm.Page
      
      InicioX = 0
      ImprimeEm.NewPage
      
      If Titulo1 <> "" Then
        ImprimeEm.FontSize = 14
        ImprimeEm.FontBold = True
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo1) / 2)
        ImprimeEm.Print Titulo1
      End If
      If Titulo2 <> "" Then
        ImprimeEm.FontSize = 10
        ImprimeEm.FontBold = True
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo2) / 2)
        ImprimeEm.Print Titulo2
      End If
      If Titulo3 <> "" Then
        ImprimeEm.FontSize = 10
        ImprimeEm.FontBold = False
        ImprimeEm.CurrentX = (LarguraMaxima / 2) - (ImprimeEm.TextWidth(Titulo3) / 2)
        ImprimeEm.Print Titulo3
      End If
      ImprimeEm.Font.Bold = Grade.Font.Bold
      ImprimeEm.Font.Size = Grade.Font.Size
      
      ImprimeEm.CurrentX = 0
      InicioY = ImprimeEm.CurrentY
      If Linhas = True Then
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      End If
      For i = 0 To .Columns.Count - 1
        If .Columns(i).Width > 200 Then
          If .Columns(i).Alignment = 1 Then
            ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
          Else
            ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
          End If
          ImprimeEm.Print .Columns(i).Caption;
          InicioX = InicioX + .Columns(i).Width
        End If
      Next i
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
      ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
      InicioX = 0
    End If
    
    If ColunaDeQuebra >= 0 Then
      If StrQuebra <> .Columns(ColunaDeQuebra).Text Then
        InicioX = 0
        If Linhas = True Then
          For i = 0 To .Columns.Count - 1
            If .Columns(i).Width > 200 Then
              ImprimeEm.CurrentX = InicioX
              ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
              InicioX = InicioX + .Columns(i).Width
            End If
          Next i
          ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
        Else
          ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
          ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        End If
        
        If ColunaTotal >= 0 Then
          InicioX = 0
          A = ColunaTotal
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Str(Total)
          Else
            StrTemp = Format(Total, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          If .Columns(A).Alignment = 1 Then
            ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          Else
            ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
          End If
          Total = 0
          ImprimeEm.Print StrTemp;
        End If
        
        If ColunaTotal2 >= 0 Then
          InicioX = 0
          A = ColunaTotal2
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Str(Total2)
          Else
            StrTemp = Format(Total2, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal2 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          If .Columns(A).Alignment = 1 Then
            ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          Else
            ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
          End If
          ImprimeEm.Print StrTemp;
          Total2 = 0
        End If
        
        If ColunaTotal3 >= 0 Then
          InicioX = 0
          A = ColunaTotal3
          If .Columns(A).NumberFormat = "" Then
            StrTemp = Str(Total3)
          Else
            StrTemp = Format(Total3, .Columns(A).NumberFormat)
          End If
          For A = 0 To ColunaTotal3 - 1
            InicioX = InicioX + .Columns(A).Width
          Next A
          If .Columns(A).Alignment = 1 Then
            ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
          Else
            ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
          End If
          ImprimeEm.Print StrTemp;
          Total3 = 0
        End If
        
        InicioX = 0
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        InicioY = ImprimeEm.CurrentY
        ImprimeEm.CurrentX = 0
        
        
        InicioY = ImprimeEm.CurrentY
        If Linhas = True Then
          ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
          ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        End If
        For i = 0 To .Columns.Count - 1
          If .Columns(i).Width > 200 Then
            If .Columns(i).Alignment = 1 Then
              ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Caption & " ")
            Else
              ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
            End If
            ImprimeEm.Print .Columns(i).Caption;
            InicioX = InicioX + .Columns(i).Width
          End If
        Next i
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
        ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
        ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
        InicioX = 0
      End If
    End If
    
    For i = 0 To .Columns.Count - 1
      If .Columns(i).Width > 200 Then
        If .Columns(i).Alignment = 1 Then
          ImprimeEm.CurrentX = (InicioX + .Columns(i).Width) - ImprimeEm.TextWidth(.Columns(i).Text & " ")
        Else
          ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
        End If
        ImprimeEm.Print .Columns(i).Text;
        InicioX = InicioX + .Columns(i).Width
      End If
    Next i
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
    If Linhas = True Then
      ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
      ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
    End If
    InicioX = 0
    If ColunaTotal >= 0 Then
      If IsNumeric(.Columns(ColunaTotal).Text) = True Then
        Total = Total + CDbl(.Columns(ColunaTotal).Text)
      Else
        Total = Total + 1
      End If
    End If
    If ColunaTotal2 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal2).Text) = True Then
        Total2 = Total2 + CDbl(.Columns(ColunaTotal2).Text)
      Else
        Total2 = Total2 + 1
      End If
    End If
    If ColunaTotal3 >= 0 Then
      If IsNumeric(.Columns(ColunaTotal3).Text) = True Then
        Total3 = Total3 + CDbl(.Columns(ColunaTotal3).Text)
      Else
        Total3 = Total3 + 1
      End If
    End If
    
    If ColunaDeQuebra >= 0 Then
      StrQuebra = .Columns(ColunaDeQuebra).Text
    End If
    Dados.Recordset.MoveNext
  Loop
  InicioX = 0
  If Linhas = True Then
    For i = 0 To .Columns.Count - 1
      If .Columns(i).Width > 200 Then
        ImprimeEm.CurrentX = InicioX
        ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
        InicioX = InicioX + .Columns(i).Width
      End If
    Next i
    ImprimeEm.Line (InicioX, InicioY)-(InicioX, ImprimeEm.CurrentY)
  Else
    ImprimeEm.Line (0, ImprimeEm.CurrentY)-(LarguraMaxima, ImprimeEm.CurrentY)
    ImprimeEm.CurrentY = ImprimeEm.CurrentY + (.RowHeight / 20)
  End If
  
  If ColunaTotal >= 0 Then
    InicioX = 0
    A = ColunaTotal
    If .Columns(A).NumberFormat = "" Then
      StrTemp = Str(Total)
    Else
      StrTemp = Format(Total, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    If .Columns(A).Alignment = 1 Then
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    Else
      ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
    End If
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal2 >= 0 Then
    InicioX = 0
    A = ColunaTotal2
    If .Columns(A).NumberFormat = "" Then
      StrTemp = Str(Total2)
    Else
      StrTemp = Format(Total2, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal2 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    If .Columns(A).Alignment = 1 Then
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    Else
      ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
    End If
    ImprimeEm.Print StrTemp;
  End If
  If ColunaTotal3 >= 0 Then
    InicioX = 0
    A = ColunaTotal3
    If .Columns(A).NumberFormat = "" Then
      StrTemp = Str(Total3)
    Else
      StrTemp = Format(Total3, .Columns(A).NumberFormat)
    End If
    For A = 0 To ColunaTotal3 - 1
      InicioX = InicioX + .Columns(A).Width
    Next A
    If .Columns(A).Alignment = 1 Then
      ImprimeEm.CurrentX = (InicioX + .Columns(A).Width) - ImprimeEm.TextWidth(StrTemp & " ")
    Else
      ImprimeEm.CurrentX = InicioX + ImprimeEm.TextWidth("_")
    End If
    ImprimeEm.Print StrTemp;
  End If
  
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
  ImprimeEm.CurrentY = ImprimeEm.CurrentY + .RowHeight
End With
End Sub

Public Function Arredonda(ByVal Valor As Currency, Optional CentavosLimite As Currency = 0.1) As Currency
Dim TempValor As Currency
Dim TempValor2 As Currency
TempValor = Fix(Valor)
TempValor2 = Valor - TempValor
Do While TempValor2 >= CentavosLimite
  TempValor = TempValor + CentavosLimite
  TempValor2 = TempValor2 - CentavosLimite
Loop
If TempValor2 <> 0 Then
  If TempValor2 >= CentavosLimite / 2 Then
    TempValor = TempValor + CentavosLimite
  End If
End If
Arredonda = TempValor
End Function

Public Function FiltraProdutos(Optional strFiltro As String = "") As String
Load frmCadProdutosFiltro
frmCadProdutosFiltro.Filtro = strFiltro
frmCadProdutosFiltro.Show vbModal
strFiltro = frmCadProdutosFiltro.Filtro
Unload frmCadProdutosFiltro
FiltraProdutos = strFiltro
End Function

Public Function ConverteCMC7(ByVal CMC7 As String) As DadosCheque
Dim Cheque As DadosCheque
With Cheque
  If Len(CMC7) >= 3 Then
    If UCase(Mid(CMC7, 1, 2)) = "L," Or UCase(Mid(CMC7, 2, 1)) = "<" Then
      If Len(CMC7) >= 36 Then
        CMC7 = Mid(CMC7, 3)
        .COMP = Mid(CMC7, 11, 3)
        .Banco = Mid(CMC7, 2, 3)
        .Agencia = Mid(CMC7, 5, 4)
        .Conta = Mid(CMC7, 26, 6) & "-" & Mid(CMC7, 32, 1)
        .Cheque = Mid(CMC7, 14, 6)
      End If
    Else
      If Len(CMC7) >= 34 Then
        .COMP = Mid(CMC7, 11, 3)
        .Banco = Mid(CMC7, 2, 3)
        .Agencia = Mid(CMC7, 5, 4)
        .Conta = Mid(CMC7, 26, 6) & "-" & Mid(CMC7, 32, 1)
        .Cheque = Mid(CMC7, 14, 6)
      End If
    End If
  End If
End With
ConverteCMC7 = Cheque

End Function

Public Function DataDeCorte(ByRef DtIniCorte As Date, ByRef Dias As Double, ByRef DtPassagem As Date) As Date
Dim DtCorte As Date, Diferenca As Double

Diferenca = DateDiff("d", DtIniCorte, DtPassagem)
Diferenca = Diferenca - (Diferenca Mod Dias)
DtIniCorte = DateAdd("d", Diferenca, DtIniCorte)

If DtIniCorte < DtPassagem Then
  Do While DtIniCorte < DtPassagem
    DtIniCorte = DateAdd("d", Dias, DtIniCorte)
  Loop
Else
  Do While DtIniCorte >= DtPassagem
    DtIniCorte = DateAdd("d", -Dias, DtIniCorte)
  Loop
  DtIniCorte = DateAdd("d", Dias, DtIniCorte)
End If

DataDeCorte = DtIniCorte

End Function

Public Function NFPModelo1(ByVal DataIni As Date, ByVal DataFim As Date) As Boolean
Dim db As New ADODB.Connection, Destino As String, A As Integer, B As Integer
Dim dbPostos As New ADODB.Recordset, dbNotas As New ADODB.Recordset, dbNotasCorpo As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim StrLinha As String, StrTemp As String, IntA As Double
Dim Linhas20 As Double
Dim Linhas30 As Double
Dim Linhas40 As Double
Dim Linhas50 As Double
Dim Linhas60 As Double


NFPModelo1 = False

On Error Resume Next

Destino = ReadINI("Local", "Local", App.Path, App.Path & "\nfp.ini")
If Right(Destino, 1) <> "\" Then Destino = Destino & "\"
Destino = Destino & NomePosto & "-" & Format(DataIni, "YYYYmmdd") & "-" & Format(DataFim, "YYYYmmdd") & ".txt"
A = FreeFile()

Open Destino For Output As #A

db.Open CaminhoADO
dbPostos.Open "Select *from postos", db, adOpenKeyset, adLockOptimistic
dbProdutos.Open "Select codigo, aliquota from produtos order by codigo", db, adOpenKeyset, adLockOptimistic


StrLinha = "10|1,00|" & Format(RemoveString(dbPostos!CNPJ), "00000000000000") & "|" & Format(DataIni, "dd/mm/YYYY") & "|" & Format(DataFim, "dd/mm/YYYY")
Print #A, StrLinha

dbNotas.Open "Select *from notas where dataemissao between #" & DataInglesa(DataIni) & "# and #" & DataInglesa(DataFim) & "# order by notanr", db, adOpenKeyset, adLockOptimistic
Linhas20 = 0
Linhas30 = 0
Linhas40 = 0
Linhas50 = 0
Linhas60 = 0

If dbNotas.RecordCount <> 0 Then
  dbNotas.MoveLast
  dbNotas.MoveFirst
  Do While dbNotas.EOF = False
    dbNotasCorpo.Open "select *from notascorpo where codigonota=" & dbNotas!CodigoNota & " order by codigoproduto", db, adOpenKeyset, adLockOptimistic
    If dbNotasCorpo.RecordCount <> 0 Then
      Linhas20 = Linhas20 + 1
      StrLinha = "20|I||" & Left(dbNotas!NaturezaOP, 60)
      StrLinha = StrLinha & "|" & ReadINI("serie", "serie", "0", App.Path & "\nfp.ini")
      StrLinha = StrLinha & "|" & dbNotas!notanr
      StrLinha = StrLinha & "|" & Format(dbNotas!dataemissao, "dd/mm/YYYY") & " " & Format(dbNotas!horasaida, "HH:NN:SS") & "|" & Format(dbNotas!datasaida, "dd/mm/YYYY") & " " & Format(dbNotas!horasaida, "HH:NN:SS")
      'tipo de nota fiscal 1=saida 0=entrada
      If dbNotas!Entrada = True Then
        StrLinha = StrLinha & "|0|" & dbNotas!cfop
      Else
        StrLinha = StrLinha & "|1|" & dbNotas!cfop
      End If
      'ie do substituto tributario
      'Inscrição Municipal do Emitente se NF-e conjugada, com prestação de serviços sujeitos ao ISSQN
      StrLinha = StrLinha & "||"
      StrTemp = ""
      If IsNull(dbNotas!CNPJ) = False Then
        StrTemp = dbNotas!CNPJ
      End If
      StrTemp = RemoveString(StrTemp)
      If Len(StrTemp) = 11 Or Len(StrTemp) = 14 Then
        StrLinha = StrLinha & "|" & StrTemp
      Else
        StrLinha = StrLinha & "|"
      End If
      StrTemp = ""
      StrTemp = Left(dbNotas!Nome, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      IntA = InStr(1, dbNotas!Endereco, ",")
      
      If IntA <> 0 Then
        StrTemp = Mid(dbNotas!Endereco, 1, IntA - 1)
        StrLinha = StrLinha & "|" & StrTemp
        StrTemp = "" 'numero1
        StrTemp = Mid(dbNotas!Endereco, IntA + 1)
        StrLinha = StrLinha & "|" & StrTemp
      Else
        StrTemp = Left(dbNotas!Endereco, 60)
        StrLinha = StrLinha & "|" & StrTemp
        StrTemp = "0" 'numero
        StrLinha = StrLinha & "|" & StrTemp
      End If
      StrTemp = "" 'complemento
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!bairro, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!municipio, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!uf, 2)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = dbNotas!CEP
      StrTemp = Format(RemoveString(StrTemp), "00000000")
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "" 'País
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = dbNotas!fone
      StrTemp = RemoveString(StrTemp)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = dbNotas!ie
      StrTemp = RemoveString(StrTemp)
      StrLinha = StrLinha & "|" & StrTemp
      
      '***********************************************************************
      'Termina a cabeça da nota
      '***********************************************************************
      Print #A, StrLinha
      
      dbNotasCorpo.MoveLast
      dbNotasCorpo.MoveFirst
      B = 1
      Do While dbNotasCorpo.EOF = False
        If Format(dbNotasCorpo!Quantidade, "0.0000") = 0 Then GoTo ProcimoProduto
        If Format(dbNotasCorpo!ValorTotal, "0.00") = 0 Then GoTo ProcimoProduto
        Linhas30 = Linhas30 + 1
        StrLinha = "30"
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!CodigoProduto
        StrTemp = Left(StrTemp, 60)
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!descriproduto
        StrTemp = Left(StrTemp, 120)
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = "" 'codigo ncm
        StrTemp = Left(StrTemp, 8)
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!unidade
        StrTemp = Left(StrTemp, 6)
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!Quantidade
        StrTemp = Format(StrTemp, "0.0000")
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!valorUnitario
        StrTemp = Format(StrTemp, "0.0000")
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        StrTemp = dbNotasCorpo!ValorTotal
        StrTemp = Format(StrTemp, "0.00")
        StrLinha = StrLinha & "|" & StrTemp
        
        'Código da Situação Tributária:
        '1° Dígito: Origem da mercadoria
        '0 - Nacional
        '1 - Estrangeira - Importação direta
        '2 - Estrangeira - Adquirida no mercado interno
        '2° e 3° Dígitos: Tributação pelo ICMS
        '00 - Tributada integralmente;
        '10 - Tributada e com cobrança de; ICMS por substituição tributária;
        '20 - Com redução de base de cálculo;
        '30 - Isenta ou não tributada e com cobrança do ICMS por substituição tributária;
        '40 - Isenta;
        '41 - Não tributada;
        '50 - Suspensão;
        '51 - Diferimento;
        '60 - ICMS cobrado anteriormente por substituição tributária;
        '70 - Com redução de base de cálculo e cobrança de ICMS substituição tributária;
        '90 - Outras.
        StrTemp = "060"
        'StrTemp = dbNotasCorpo!subtributaria
        StrLinha = StrLinha & "|" & StrTemp
        
        If IsNull(dbNotasCorpo!aliquotaicms) = True Then
          If dbProdutos.RecordCount <> 0 Then
            dbProdutos.MoveFirst
            dbProdutos.Find "codigo=" & dbNotasCorpo!CodigoProduto
            If dbProdutos.EOF = False Then
              If dbProdutos!Aliquota = "FF" Then
                StrTemp = "18,00"
              Else
                StrTemp = dbProdutos!Aliquota / 100
              End If
            Else
              StrTemp = "0"
            End If
          Else
            StrTemp = "0"
            StrTemp = dbNotasCorpo!aliquotaicms
          End If
        Else
          StrTemp = "0"
          StrTemp = dbNotasCorpo!aliquotaicms
        End If
        StrTemp = Format(StrTemp, "0.00")
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        If dbNotasCorpo!aliquotaipi <> 0 Then
          StrTemp = dbNotasCorpo!aliquotaipi
          StrTemp = Format(StrTemp, "0.##")
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = ""
        If dbNotasCorpo!valoripi <> 0 Then
          StrTemp = dbNotasCorpo!valoripi
          StrTemp = Format(StrTemp, "0.##")
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        Print #A, StrLinha
        
ProcimoProduto:
        dbNotasCorpo.MoveNext
      Loop
      
      Linhas40 = Linhas40 + 1
      StrLinha = "40"
      
      StrTemp = "0,00"
      If IsNull(dbNotas!BaseICMS) = False Then
        StrTemp = Format(dbNotas!BaseICMS, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!ValorICMS) = False Then
        StrTemp = Format(dbNotas!ValorICMS, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!baseicmssubst) = False Then
        StrTemp = Format(dbNotas!baseicmssubst, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!ValorICMSSubst) = False Then
        StrTemp = Format(dbNotas!ValorICMSSubst, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!totaldosprodutos) = False Then
        StrTemp = Format(dbNotas!totaldosprodutos, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!ValorFrete) = False Then
        StrTemp = Format(dbNotas!ValorFrete, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!ValorSeguro) = False Then
        StrTemp = Format(dbNotas!ValorSeguro, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00" ' total desconto
      StrTemp = Format(0, "0.00")
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!valoripi) = False Then
        StrTemp = Format(dbNotas!valoripi, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!OutrasDespesas) = False Then
        StrTemp = Format(dbNotas!OutrasDespesas, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "0,00"
      If IsNull(dbNotas!ValorTotalDaNota) = False Then
        StrTemp = Format(dbNotas!ValorTotalDaNota, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "" 'Serviços sob não-incidência ou não tributados pelo ICMS
      If IsNull(dbNotas!servicoiss) = False Then
        StrTemp = Format(dbNotas!servicoiss, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "" 'Alíquota do ISS
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "" 'Valor Total do ISS
      If IsNull(dbNotas!servicototal) = False Then
        StrTemp = Format(dbNotas!servicototal, "0.00")
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      Print #A, StrLinha
      
      Linhas50 = Linhas50 + 1
      StrLinha = "50"
      
      '0 - por conta do emitente;
      '1 - por conta do destinatário;
      StrTemp = "1"
      If IsNull(dbNotas!FretePorConta) = False Then
        If dbNotas!FretePorConta = "2" Then
          StrTemp = "1"
        Else
          StrTemp = "0"
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = RemoveString(dbNotas!CNPJ)
      StrLinha = StrLinha & "|" & strtmep
      
      StrTemp = ""
      If IsNull(dbNotas!CNPJ2) = False Then
        StrTemp = dbNotas!CNPJ2
      End If
      StrTemp = RemoveString(StrTemp)
      If Len(StrTemp) = 11 Or Len(StrTemp) = 14 Then
        StrLinha = StrLinha & "|" & StrTemp
      Else
        StrLinha = StrLinha & "|"
      End If
      StrTemp = ""
      StrTemp = Left(dbNotas!nome2, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!IE2, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!Endereco2, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!UF3, 2)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!Placa, 8)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!UF2, 2)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!quantidade2, 15)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!Especie, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!Marca, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!Numero, 60)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Format(dbNotas!PesoLiquido, "0.000")
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Format(dbNotas!PesoBruto, "0.000")
      StrLinha = StrLinha & "|" & StrTemp
      
      Print #A, StrLinha
      
      Linhas60 = Linhas60 + 1
      StrLinha = "60"
      
      StrTemp = ""
      StrTemp = Left(dbNotas!dadosfatura, 256)
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = "" 'destinado ao fisco
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = ""
      StrTemp = Left(dbNotas!dadosadicionais, 5000)
      StrLinha = StrLinha & "|" & StrTemp
      
      Print #A, StrLinha
      
    End If
    dbNotasCorpo.Close
    dbNotas.MoveNext
  Loop
  
  
  StrLinha = "90|" & Format(Linhas20, "00000") & "|" & Format(Linhas30, "00000") & "|" & Format(Linhas40, "00000") & "|" & Format(Linhas50, "00000") & "|" & Format(Linhas60, "00000")
  Print #A, StrLinha
End If

dbNotas.Close
dbProdutos.Close
db.Close
Close #A
NFPModelo1 = True

End Function


Public Function PrecoAtual(ByVal CodigoProduto As Double, ByVal Dia As Date, ByVal CodigoTurno As Double, Optional Bico As Integer = 0) As Currency
Dim db As New ADODB.Connection
Dim DbPrecos As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim CodigoAlteracao As Double

db.Open CaminhoADO
dbTurnos.Open "select *from turnos order by horaini", db, adOpenKeyset, adLockOptimistic
If dbTurnos.RecordCount <> 0 Then
  dbTurnos.MoveFirst
  dbTurnos.Find "codigoturno=" & CodigoTurno
  If dbTurnos.EOF = True Then
    PrecoAtual = 0
    Exit Function
  End If
End If

If Bico <> 0 Then
  DbPrecos.Open "select alteracoes.*, turnos.* from alteracoes, turnos where turnos.codigoturno=alteracoes.codigoturno order by dataalteracao, horaini", db, adOpenKeyset, adLockOptimistic
  If DbPrecos.RecordCount <> 0 Then
    DbPrecos.MoveLast
    Do While DbPrecos.BOF = False
      If DbPrecos!dataalteracao <= Dia Then
        If DbPrecos!dataalteracao < Dia Then
          CodigoAlteracao = DbPrecos!codalteracao
          Exit Do
        Else
          If DbPrecos!HoraIni <= dbTurnos!HoraIni Then
            CodigoAlteracao = DbPrecos!codalteracao
            Exit Do
          Else
            GoTo Procimo
          End If
        End If
      End If
Procimo:
      DbPrecos.MovePrevious
    Loop
  End If
  If CodigoAlteracao = 0 Then
    DbPrecos.Close
    DbPrecos.Open "select bicos.precovenda from bicos where bico=" & Bico, db, adOpenKeyset, adLockOptimistic
    If DbPrecos.RecordCount <> 0 Then
      PrecoAtual = DbPrecos!PrecoVenda
    End If
  Else
    DbPrecos.Close
    DbPrecos.Open "select preco from alterabico where codalteracao=" & CodigoAlteracao & " and bico=" & Bico, db, adOpenKeyset, adLockOptimistic
    If DbPrecos.RecordCount <> 0 Then
      PrecoAtual = DbPrecos!Preco
    End If
  End If
Else
  DbPrecos.Open "select *from produtosaltera order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
  If DbPrecos.RecordCount <> 0 Then
    DbPrecos.MoveLast
    Do While DbPrecos.BOF = False
      If DbPrecos!DataCaixa <= Dia Then
        If DbPrecos!DataCaixa < Dia Then
          CodigoAlteracao = DbPrecos!codigoprodutoaltera
          Exit Do
        Else
          If DbPrecos!HoraIni <= dbTurnos!HoraIni Then
            CodigoAlteracao = DbPrecos!codigoprodutoaltera
            Exit Do
          Else
            CodigoAlteracao = 0
            Exit Do
          End If
        End If
      End If
      DbPrecos.MovePrevious
    Loop
  End If
  If CodigoAlteracao = 0 Then
    DbPrecos.Close
    DbPrecos.Open "select precovenda from produtos where codigoproduto=" & CodigoProduto, db, adOpenKeyset, adLockOptimistic
    If DbPrecos.RecordCount <> 0 Then
      PrecoAtual = DbPrecos!PrecoVenda
    End If
  Else
    DbPrecos.Close
    DbPrecos.Open "select precovenda from produtosalteradetalhe where codigoprodutoaltera=" & CodigoAlteracao & " and codigoproduto=" & CodigoProduto, db, adOpenKeyset, adLockOptimistic
    If DbPrecos.RecordCount <> 0 Then
      PrecoAtual = DbPrecos!PrecoVenda
    End If
  End If
End If

DbPrecos.Close
dbTurnos.Close
db.Close
End Function

Public Sub ExportaClienteMicrosffer()
Dim dbRemoto As New ADODB.Connection
Dim dbLocal As New ADODB.Connection
Dim adoClientes As New ADODB.Recordset, adoProg As New ADODB.Recordset
Dim adoConfig As New ADODB.Recordset
Dim adoHistorico As New ADODB.Recordset
Dim Bloqueado As Boolean, Resposta As Integer
Dim Bloqueio As Integer

Resposta = MsgBox("Deseja exportar agora?", vbYesNo)
If Resposta = vbNo Then Exit Sub

dbLocal.Open CaminhoADO
adoConfig.CursorLocation = adUseClient
adoConfig.Open "select servidorpista from config", dbLocal, adOpenKeyset, adLockOptimistic
If adoConfig.RecordCount <> 0 Then
  If IsNull(adoConfig!servidorpista) = True Then
    MsgBox "Servidor do pista ainda não foi configurado!"
    Exit Sub
  Else
    StrTemp = adoConfig!servidorpista
  End If
End If

adoClientes.CursorLocation = adUseClient
adoClientes.Open "select codigocliente, codigonoposto, mensalista, protestado from clientes", dbLocal, adOpenKeyset, adLockOptimistic

adoHistorico.CursorLocation = adUseClient
adoHistorico.Open "select *from clienteshistorico", dbLocal, adOpenKeyset, adLockOptimistic

On Error GoTo TrataErro
StrTemp = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=A30Sigpo;Data Source=" & StrTemp
dbRemoto.Open StrTemp
adoProg.CursorLocation = adUseClient
adoProg.Open "select cliente, bloqueado, dt_alter from a30prog_clie", dbRemoto, adOpenKeyset, adLockOptimistic

If adoClientes.RecordCount <> 0 Then
  If adoProg.RecordCount <> 0 Then
    adoProg.MoveLast
    adoProg.MoveFirst
    Do While adoProg.EOF = False
      adoClientes.MoveFirst
      adoClientes.Find "codigonoposto=" & CInt(adoProg!Cliente)
      If adoClientes.EOF = False Then
        If adoClientes!protestado = 0 Then
          Bloqueado = adoClientes!mensalista
          If Bloqueado = True Then
            Bloqueio = 0
          Else
            Bloqueio = 1
          End If
        Else
          Bloqueio = 1
        End If
        If adoProg!Bloqueado <> Bloqueio Then
          If Bloqueado = True Then
            adoProg!Bloqueado = 0
          Else
            adoProg!Bloqueado = 1
          End If
          adoProg!Dt_Alter = Format(Date, "yyyymmdd")
          adoProg.Update
          
          adoHistorico.AddNew
          adoHistorico!CodigoCliente = adoClientes!CodigoCliente
          adoHistorico!ativando = Bloqueado
          adoHistorico!Dia = Now
          adoHistorico!Usuario = Usuarios.Nome
          If Bloqueado = True Then
            adoHistorico!Descri = "Ativação do cliente!"
          Else
            adoHistorico!Descri = "Desativação do cliente!"
          End If
          adoHistorico.Update
        End If
      End If
      adoProg.MoveNext
    Loop
  End If
End If

MsgBox "Exportado com sucesso!"

  Exit Sub
  
TrataErro:
  MsgBox Err.Number & " - " & Err.Description
End Sub

Public Sub ProdutosMicrosffer()
Dim dbRemoto As New ADODB.Connection, dbLocal As New ADODB.Connection
Dim dbProdutoRemoto As New ADODB.Recordset
Dim dbProdutoLocal As New ADODB.Recordset
Dim dbGrupoICFRemoto As New ADODB.Recordset
Dim dbGrupoICFLocal As New ADODB.Recordset
Dim adoConfig As New ADODB.Recordset

On Error GoTo TrataErro
dbLocal.Open CaminhoADO
adoConfig.CursorLocation = adUseClient
adoConfig.Open "select servidorpista from config", dbLocal, adOpenKeyset, adLockOptimistic
If adoConfig.RecordCount <> 0 Then
  If IsNull(adoConfig!servidorpista) = True Then
    MsgBox "Servidor do pista ainda não foi configurado!"
    Exit Sub
  Else
    StrTemp = adoConfig!servidorpista
  End If
End If

StrTemp = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=A30Sigpo;Data Source=" & StrTemp
dbRemoto.Open StrTemp
dbProdutoRemoto.CursorLocation = adUseClient
dbProdutoRemoto.Open "Select produto, grupoif from a30produto", dbRemoto, adOpenKeyset, adLockOptimistic

dbProdutoLocal.CursorLocation = adUseClient
dbProdutoLocal.Open "select produtos.codigo, produtosgrupoif.codigogrupo from produtos, produtosgrupoif where produtos.codigogrupoif=produtosgrupoif.codigo", dbLocal, adOpenKeyset, adLockOptimistic


If dbProdutoRemoto.RecordCount <> 0 And dbProdutoLocal.RecordCount <> 0 Then
  dbProdutoRemoto.MoveLast
  dbProdutoRemoto.MoveFirst
  Do While dbProdutoRemoto.EOF = False
    dbProdutoLocal.MoveFirst
    dbProdutoLocal.Find "codigo='" & CDbl(dbProdutoRemoto!Produto) & "'"
    If dbProdutoLocal.EOF = False Then
      If dbProdutoRemoto!grupoif <> Trim(Str(dbProdutoLocal!CodigoGrupo)) Then
        dbProdutoRemoto!grupoif = dbProdutoLocal!CodigoGrupo
        dbProdutoRemoto.Update
      End If
    End If
    dbProdutoRemoto.MoveNext
  Loop
End If


dbProdutoLocal.Close
dbProdutoRemoto.Close
adoConfig.Close
dbLocal.Close
dbRemoto.Close

MsgBox "Atualização realizada com sucesso!"


Exit Sub

TrataErro:
MsgBox Err.Number & " - " & Err.Description

End Sub

Public Sub GravaCupons(ByVal StrTemp As String, ByVal dbVendasLeituraX As Adodc)
Dim dbProdutoGrupoIf As New ADODB.Recordset
Dim db As New ADODB.Connection

Dim CodigoCliente As Double, DataCupom As Date
Dim HoraCupom As Date, NumeroCupom As String, Placa As String
Dim Km As Double, Carro As String, QtdProduto As Double
Dim ValorTotal As Currency, CodigoGrupo As Double, Tributo As String
Dim strCategoria As String

db.Open CaminhoADO
dbProdutoGrupoIf.CursorLocation = adUseClient
dbProdutoGrupoIf.Open "select *from produtosgrupoif", db, adOpenKeyset, adLockOptimistic


'007|     27/12/2008|        132,315|         324,04|            102
On Error GoTo TrataErro
DataCupom = CDate(Trim(Mid(StrTemp, 5, 15)))
If Trim(Mid(StrTemp, 21, 15)) <> "" Then
  QtdProduto = Trim(Mid(StrTemp, 21, 15))
Else
  QtdProduto = 0
End If
If Trim(Mid(StrTemp, 37, 15)) <> "" Then
  ValorTotal = Trim(Mid(StrTemp, 37, 15))
Else
  ValorTotal = 0
End If
If Trim(Mid(StrTemp, 53)) <> "" Then
  CodigoGrupo = Trim(Mid(StrTemp, 53))
Else
  CodigoGrupo = 0
End If
If dbProdutoGrupoIf.RecordCount <> 0 Then
  dbProdutoGrupoIf.MoveFirst
  dbProdutoGrupoIf.Find "codigogrupo=" & CodigoGrupo
  If dbProdutoGrupoIf.EOF = False Then
    CodigoGrupo = dbProdutoGrupoIf!Codigo
    strCategoria = dbProdutoGrupoIf!CodigoGrupo & " " & dbProdutoGrupoIf!Descri
    'strCategoria = dbProdutoGrupoIf!Descri
  End If
End If

With dbVendasLeituraX
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
  Else
    .Recordset.Filter = "data=#" & DataCupom & "# and categoria='" & strCategoria & "'"
    If .Recordset.RecordCount = 0 Then
      .Recordset.AddNew
    End If
  End If
  .Recordset!Data = DataCupom
  .Recordset!LeituraXQtd = QtdProduto
  .Recordset!LeituraXValor = ValorTotal
  .Recordset!Categoria = strCategoria
  .Recordset.Update
End With

TrataErro:

dbProdutoGrupoIf.Close
db.Close

End Sub

Public Sub GravaComissoes(ByVal StrTemp As String, ByVal CodigoFechamento As Double)
Dim db As New ADODB.Connection
Dim dbComissoes As New ADODB.Recordset
Dim dbFuncionarios As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset

Dim Produto As String
Dim Bico As String
Dim Funcionario As String
Dim Qtd As Double
Dim VlUnitario As Currency
Dim VlTotal As Currency
Dim VlVendaC As Currency
Dim VlTotalC As Currency
Dim VlComissao As Currency

Dim CodigoFuncionario As Double
Dim Nome As String

Dim CodigoProduto As Double

Dim strSql As String
On Error GoTo TrataErro

db.Open CaminhoADO

dbFuncionarios.CursorLocation = adUseClient
dbFuncionarios.Open "select *from vendedores", db, adOpenKeyset, adLockOptimistic

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "select *from produtos", db, adOpenKeyset, adLockOptimistic

'008|         000572|               |         000512|                  2,00000|                 21,90000|                 43,80000|                 21,90000|                 43,80000|                  3,06600

Produto = Trim(Mid(StrTemp, 5, 15))
Bico = Trim(Mid(StrTemp, 21, 15))
Funcionario = Trim(Mid(StrTemp, 37, 15))
Qtd = CDbl(Mid(StrTemp, 53, 25))
VlUnitario = CCur(Mid(StrTemp, 79, 25))
VlTotal = CCur(Mid(StrTemp, 105, 25))
VlVendaC = CCur(Mid(StrTemp, 131, 25))
VlTotalC = CCur(Mid(StrTemp, 157, 25))
VlComissao = CCur(Mid(StrTemp, 183, 25))

CodigoFuncionario = 0
Nome = ""
CodigoProduto = 0

If Trim(Produto) <> "" Then
    If IsNumeric(Produto) = True Then
        If dbProdutos.RecordCount <> 0 Then
            dbProdutos.MoveFirst
            dbProdutos.Find "codigo=" & Produto
            If dbProdutos.EOF = False Then
                CodigoProduto = dbProdutos!CodigoProduto
            End If
        End If
    End If
End If

If Trim(Funcionario) <> "" Then
    If IsNumeric(Funcionario) = True Then
        If dbFuncionarios.RecordCount <> 0 Then
            dbFuncionarios.MoveFirst
            dbFuncionarios.Find "codigo=" & Trim(Funcionario)
            If dbFuncionarios.EOF = False Then
                CodigoFuncionario = dbFuncionarios!CodigoFuncionario
                Nome = dbFuncionarios!Nome
            End If
        End If
    End If
End If

If IsNumeric(Bico) = False Then
    Bico = "0"
End If
If IsNumeric(Funcionario) = False Then
    Funcionario = "0"
    Nome = " "
End If
strSql = "insert into comissoes (codigofechamento,CodigoProduto,Codigo,bico,CodigoFuncionario,funcionario,Nome,qtd,VlUnitario,VlTotal,VlVendaC,VlTotalC,VlComissao) values (" & _
           CodigoFechamento & "," & CodigoProduto & ",'" & Produto & "'," & Bico & "," & CodigoFuncionario & "," & Funcionario & ",'" & Nome & "'," & NumeroIngles(Qtd) & "," & NumeroIngles(VlUnitario) & _
           "," & NumeroIngles(VlTotal) & "," & NumeroIngles(VlVendaC) & "," & NumeroIngles(VlTotalC) & "," & NumeroIngles(VlComissao) & ")"

db.Execute strSql

TrataErro:

End Sub


Public Sub GravaCupons2(ByVal StrTemp As String)
Dim dbProdutoGrupoIf As New ADODB.Recordset
Dim db As New ADODB.Connection
Dim dbVendasLeituraX As New ADODB.Recordset

Dim CodigoCliente As Double, DataCupom As Date
Dim HoraCupom As Date, NumeroCupom As String, Placa As String
Dim Km As Double, Carro As String, QtdProduto As Double
Dim ValorTotal As Currency, CodigoGrupo As Double, Tributo As String
Dim strCategoria As String

db.Open CaminhoADO
dbProdutoGrupoIf.CursorLocation = adUseClient
dbProdutoGrupoIf.Open "select *from produtosgrupoif", db, adOpenKeyset, adLockOptimistic


'007|     27/12/2008|        132,315|         324,04|            102
On Error GoTo TrataErro
DataCupom = CDate(Trim(Mid(StrTemp, 5, 15)))

dbVendasLeituraX.CursorLocation = adUseClient
dbVendasLeituraX.Open "select *from cuponsfiscais where datacupom=#" & DataInglesa(DateAdd("d", -1, DataCupom)) & "#", db, adOpenKeyset, adLockOptimistic

If Trim(Mid(StrTemp, 21, 15)) <> "" Then
  QtdProduto = Trim(Mid(StrTemp, 21, 15))
Else
  QtdProduto = 0
End If
If Trim(Mid(StrTemp, 37, 15)) <> "" Then
  ValorTotal = Trim(Mid(StrTemp, 37, 15))
Else
  ValorTotal = 0
End If
If Trim(Mid(StrTemp, 53)) <> "" Then
  CodigoGrupo = Trim(Mid(StrTemp, 53))
Else
  CodigoGrupo = 0
End If
If dbProdutoGrupoIf.RecordCount <> 0 Then
  dbProdutoGrupoIf.MoveFirst
  dbProdutoGrupoIf.Find "codigogrupo=" & CodigoGrupo
  If dbProdutoGrupoIf.EOF = False Then
    CodigoGrupo = dbProdutoGrupoIf!Codigo
    strCategoria = dbProdutoGrupoIf!CodigoGrupo & " " & dbProdutoGrupoIf!Descri
  End If
End If

  If dbVendasLeituraX.RecordCount = 0 Then
    dbVendasLeituraX.AddNew
  Else
    dbVendasLeituraX.Filter = "data=#" & DataCupom & "# and categoria='" & strCategoria & "'"
    If dbVendasLeituraX.RecordCount = 0 Then
      dbVendasLeituraX.AddNew
    End If
  End If
  dbVendasLeituraX!Data = DataCupom
  dbVendasLeituraX!LeituraXQtd = QtdProduto
  dbVendasLeituraX!LeituraXValor = ValorTotal
  dbVendasLeituraX!Categoria = strCategoria
  dbVendasLeituraX.Update

TrataErro:

dbProdutoGrupoIf.Close
db.Close

End Sub


Public Function RegistraEstoque(ByVal DataCaixa As Date, ByVal CodigoTurno As Double, ByVal Turno As String, ByVal HoraIni As Date, ByVal CodigoProduto As Double, Optional Tanque As Integer = 0, Optional Entrada As Double = 0, Optional Saida As Double = 0, Optional Acerto As Double = 0) As Boolean
Dim db As New ADODB.Connection
Dim dbEstoque As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim Abertura As Double, Disponivel As Double

On Error GoTo TrataErro
RegistraEstoque = False

Disponivel = 0

db.Open CaminhoADO

dbEstoque.CursorLocation = adUseClient
dbEstoque.Open "Select *from produtosestoque where codigoproduto=" & CodigoProduto & " order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos", db, adOpenKeyset, adLockOptimistic

If dbProdutos.RecordCount = 0 Then
  Exit Function
End If
dbProdutos.MoveFirst
dbProdutos.Find "codigoproduto=" & CodigoProduto

If dbEstoque.RecordCount = 0 Then
  dbEstoque.AddNew
  Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
Else
  dbEstoque.Filter = "datacaixa=#" & DataInglesa(DataCaixa) & "# and codigoturno=" & CodigoTurno
  If dbEstoque.RecordCount = 0 Then
    dbEstoque.AddNew
    Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
  Else
    dbEstoque.MovePrevious
    If dbEstoque.BOF = True Then
      Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
    Else
      Abertura = dbEstoque!Disponivel
    End If
    dbEstoque.MoveNext
  End If
End If


dbEstoque!CodigoProduto = CodigoProduto
dbEstoque!Codigo = dbProdutos!Codigo
dbEstoque!Tanque = Tanque
dbEstoque!DataCaixa = DataCaixa
dbEstoque!CodigoTurno = CodigoTurno
dbEstoque!Turno = Turno
dbEstoque!HoraIni = HoraIni
dbEstoque!Combustivel = dbProdutos!Combustivel
dbEstoque!Abertura = Abertura
If IsNull(dbEstoque!Entrada) = True Then dbEstoque!Entrada = 0
If IsNull(dbEstoque!Saida) = True Then dbEstoque!Saida = 0
If IsNull(dbEstoque!Acerto) = True Then dbEstoque!Acerto = 0
If IsNull(dbEstoque!Diferenca) = True Then dbEstoque!Diferenca = 0
If IsNull(dbEstoque!Disponivel) = True Then dbEstoque!Disponivel = 0
dbEstoque!Entrada = dbEstoque!Entrada + Entrada
dbEstoque!Saida = dbEstoque!Saida + Saida
dbEstoque!Acerto = dbEstoque!Acerto + Acerto
If Disponivel = 0 Then
  dbEstoque!Disponivel = Abertura + dbEstoque!Entrada - dbEstoque!Saida + dbEstoque!Acerto
  dbEstoque!Abertura = Abertura
Else
  dbEstoque!Abertura = Disponivel - dbEstoque!Entrada + dbEstoque!Saida - dbEstoque!Acerto
  dbEstoque!Disponivel = Disponivel
End If
dbEstoque!dataalterado = Now
dbEstoque!Usuario = Usuarios.Nome
dbEstoque.Update

Abertura = dbEstoque!Disponivel
dbEstoque.MoveNext
Do While dbEstoque.EOF = False
  dbEstoque!Abertura = Abertura
  dbEstoque!Disponivel = Abertura + dbEstoque!Entrada - dbEstoque!Saida + dbEstoque!Acerto
  dbEstoque.Update
  Abertura = dbEstoque!Disponivel
  dbEstoque.MoveNext
Loop

RegistraEstoque = True
Exit Function

TrataErro:
  MsgBox Err.Number & " - " & Err.Description
  RegistraEstoque = False
End Function


Public Function EstoqueNoDia(ByVal DataCaixa As Date, ByVal CodigoTurno As Double, ByVal CodigoProduto As Double) As Double
Dim StrTemp As String, Sequencia As Double
Dim db As New ADODB.Connection
Dim dbFechamento As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbVendas As New ADODB.Recordset
Dim dbEntradas As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim SequenciaFinalizado As Double
Dim Estoque As Double


db.Open CaminhoADO
dbFechamento.CursorLocation = adUseClient
dbFechamento.Open "Select fechado, datacaixa, horaini, codigoturno, sequencia from fechamentodecaixa where datacaixa<=#" & DataInglesa(DataCaixa) & "# order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic
dbEntradas.CursorLocation = adUseClient
dbEntradas.Open "select datanota, codigoproduto, quantidade from qprodutosnotas where codigoproduto=" & CodigoProduto & " and datanota>#" & DataInglesa(DataCaixa) & "# order by codigoproduto", db, adOpenKeyset, adLockOptimistic
dbTurnos.CursorLocation = adUseClient
dbTurnos.Open "select *from turnos where codigoturno=" & CodigoTurno, db, adOpenKeyset, adLockOptimistic

If dbFechamento.RecordCount <> 0 Then
  dbFechamento.Find "fechado=-1"
  If dbFechamento.EOF = False Then
    SequenciaFinalizado = dbFechamento!Sequencia
  Else
    SequenciaFinalizado = 1
  End If
  dbFechamento.MoveFirst
  TempData = DataCaixa
  If dbFechamento!DataCaixa >= DataCaixa Then
    If dbFechamento!DataCaixa > DataCaixa Then
      dbFechamento.Find "datacaixa=#" & DataInglesa(TempData) & "#"
    End If
    If dbFechamento!HoraIni <= dbTurnos!HoraIni Then
      Sequencia = dbFechamento!Sequencia
    Else
      Sequencia = dbFechamento!Sequencia
      dbFechamento.Find "horaini<=#" & dbTurnos!HoraIni & "#"
      If dbFechamento.EOF = False Then
        If dbFechamento!DataCaixa < DataCaixa Then
          TempData = dbFechamento!DataCaixa
          dbFechamento.MoveFirst
          dbFechamento.Find "datacaixa=#" & DataInglesa(TempData) & "#"
          Sequencia = dbFechamento!Sequencia
        Else
          Sequencia = dbFechamento!Sequencia
        End If
      End If
    End If
  Else
    Sequencia = dbFechamento!Sequencia
  End If
Else
  Sequencia = 0
End If


dbProdutos.Open "Select codigoproduto, estoque from produtos where codigoproduto=" & CodigoProduto & " order by codigoproduto", db, adOpenKeyset, adLockOptimistic
If dbProdutos.RecordCount <> 0 Then
  Estoque = dbProdutos!Estoque
End If

'combustiveis
StrTemp = "select produtos.codigoproduto, produtos.descri, sum(encerrante-abertura) as estoquedia from qbicoencerrantes where produtos.codigoproduto=" & CodigoProduto & " and fechado=0 and sequencia>" & SequenciaFinalizado & " and sequencia<=" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    Estoque = dbProdutos!Estoque - dbVendas!estoquedia
    dbVendas.MoveNext
  Loop
End If

dbVendas.Close
StrTemp = "select produtos.codigoproduto, produtos.descri, sum(encerrante-abertura) as estoquedia from qbicoencerrantes where produtos.codigoproduto=" & CodigoProduto & " and fechado=-1 and sequencia>" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    Estoque = Estoque + dbVendas!estoquedia
    dbVendas.MoveNext
  Loop
End If



'não combustiveis
dbVendas.Close
StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where produtos.codigoproduto=" & CodigoProduto & " and fechado=0 and sequencia between " & SequenciaFinalizado & " and " & Sequencia - 1 & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    Estoque = dbProdutos!Estoque - dbVendas!estoquedia
    dbVendas.MoveNext
  Loop
End If

dbVendas.Close
StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where produtos.codigoproduto=" & CodigoProduto & " and fechado=-1 and sequencia>" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
dbVendas.Open StrTemp, db, adOpenKeyset, adLockOptimistic
If dbVendas.RecordCount <> 0 Then
  Do While dbVendas.EOF = False
    Estoque = Estoque + dbVendas!estoquedia
    dbVendas.MoveNext
  Loop
End If


If dbEntradas.RecordCount <> 0 Then
  Do While dbEntradas.EOF = False
    Estoque = Estoque - dbEntradas!Quantidade
    dbEntradas.MoveNext
  Loop
End If

dbFechamento.Close
dbProdutos.Close
dbVendas.Close
dbEntradas.Close
dbTurnos.Close

db.Close

EstoqueNoDia = Estoque
Exit Function

TrataErro:
  MsgBox Err.Number & " - " & Err.Description
  EstoqueNoDia = 0
End Function


Public Function EstoqueDesdeAData(ByVal DataCaixa As Date, ByVal CodigoProduto As Double)
Dim db As New ADODB.Connection
Dim dbVendas As New ADODB.Recordset
Dim dbEntradas As New ADODB.Recordset
Dim dbAcertos As New ADODB.Recordset
Dim dbEstoque As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim Disponivel As Double, Abertura As Double



db.Open CaminhoADO

db.Execute "update produtosestoque set abertura=0 where codigoproduto=" & CodigoProduto
db.Execute "update produtosestoque set entrada=0 where codigoproduto=" & CodigoProduto
db.Execute "update produtosestoque set saida=0 where codigoproduto=" & CodigoProduto
db.Execute "update produtosestoque set acerto=0 where codigoproduto=" & CodigoProduto
db.Execute "update produtosestoque set diferenca=0 where codigoproduto=" & CodigoProduto
db.Execute "update produtosestoque set disponivel=0 where codigoproduto=" & CodigoProduto

dbEstoque.CursorLocation = adUseClient
dbEstoque.Open "select *from produtosestoque where codigoproduto=" & CodigoProduto & " order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
dbTurnos.CursorLocation = adUseClient
dbTurnos.Open "select *from turnos", db, adOpenForwardOnly, adLockReadOnly
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select codigoproduto, codigo, estoque, combustivel from produtos where codigoproduto=" & CodigoProduto, db, adOpenForwardOnly, adLockReadOnly



If dbProdutos.RecordCount = 0 Then GoTo Termina
If dbTurnos.RecordCount = 0 Then GoTo Termina

dbVendas.CursorLocation = adUseClient
dbVendas.Open "select fechamentodecaixa.datacaixa, fechamentodecaixa.codigofechamento, fechamentodecaixa.codigoturno, fechamentodecaixa.turno, fechamentodecaixa.horaini, venda2.codigoproduto, venda2.quantidade from venda2, fechamentodecaixa where fechamentodecaixa.codigofechamento=venda2.codigofechamento and venda2.codigoproduto=" & CodigoProduto & " and fechamentodecaixa.fechado=-1 order by fechamentodecaixa.datacaixa, fechamentodecaixa.horaini", db, adOpenForwardOnly, adLockReadOnly
If dbVendas.RecordCount <> 0 Then
  dbVendas.MoveFirst
  Do While dbVendas.EOF = False
    If dbEstoque.RecordCount = 0 Then
      GoSub Adiciona
    Else
      dbEstoque.MoveFirst
      dbEstoque.Filter = "datacaixa=#" & dbVendas!DataCaixa & "# and codigoturno=" & dbVendas!CodigoTurno
      If dbEstoque.EOF = True Then
        GoSub Adiciona
      End If
    End If
    dbEstoque.Requery
    dbEstoque.Filter = ""
    dbEstoque.Filter = "datacaixa=#" & dbVendas!DataCaixa & "# and codigoturno=" & dbVendas!CodigoTurno
    dbEstoque!Saida = dbEstoque!Saida + dbVendas!Quantidade
    dbEstoque.Update
    
    'DoEvents
    
    dbVendas.MoveNext
  Loop
End If
dbVendas.Close
dbVendas.CursorLocation = adUseClient
dbVendas.Open "select fechamentodecaixa.datacaixa, fechamentodecaixa.codigofechamento, fechamentodecaixa.codigoturno, fechamentodecaixa.turno, fechamentodecaixa.horaini, bicoencerrantes.abertura, bicoencerrantes.encerrante, bicoencerrantes.codigoproduto,bicoencerrantes.Retorno, bicoencerrantes.tanque from bicoencerrantes, fechamentodecaixa where bicoencerrantes.codigofechamento=fechamentodecaixa.codigofechamento and bicoencerrantes.codigoproduto=" & CodigoProduto & " and fechamentodecaixa.fechado=-1 order by fechamentodecaixa.datacaixa, fechamentodecaixa.horaini", db, adOpenForwardOnly, adLockReadOnly
If dbVendas.RecordCount <> 0 Then
  dbVendas.MoveFirst
  Do While dbVendas.EOF = False
    dbEstoque.Filter = "datacaixa=#" & dbVendas!DataCaixa & "# and codigoturno=" & dbVendas!CodigoTurno
    If dbEstoque.EOF = True Then
      GoSub Adiciona
    End If
    dbEstoque.Requery
    dbEstoque.Filter = ""
    dbEstoque.Filter = "datacaixa=#" & dbVendas!DataCaixa & "# and codigoturno=" & dbVendas!CodigoTurno
    dbEstoque!Saida = dbEstoque!Saida + (dbVendas!Encerrante - dbVendas!Abertura)
    dbEstoque.Update
    
    'DoEvents
    
    dbVendas.MoveNext
  Loop
End If
dbEntradas.CursorLocation = adUseClient
dbEntradas.Open "select *from qprodutosnotas where codigoproduto=" & CodigoProduto & " and confirmado=-1 order by datanota", db, adOpenForwardOnly, adLockReadOnly
If dbEntradas.RecordCount <> 0 Then
  dbEntradas.MoveFirst
  Do While dbEntradas.EOF = False
    dbTurnos.MoveFirst
    If IsNull(dbEntradas!CodigoTurno) = False Then
      dbTurnos.Find "codigoturno=" & dbEntradas!CodigoTurno
      If dbTurnos.EOF = True Then dbTurnos.MoveFirst
    Else
      dbTurnos.MoveFirst
    End If
    dbEstoque.Filter = "datacaixa=#" & dbEntradas!datanota & "# and codigoturno=" & dbTurnos!CodigoTurno
    If dbEstoque.EOF = True Then
      GoSub Adiciona2
    End If
    dbEstoque.Requery
    dbEstoque.Filter = ""
    dbEstoque.Filter = "datacaixa=#" & dbEntradas!datanota & "# and codigoturno=" & dbTurnos!CodigoTurno
    dbEstoque!Entrada = dbEstoque!Entrada + dbEntradas!Quantidade
    dbEstoque.Update
    
    'DoEvents
    
    dbEntradas.MoveNext
  Loop
End If

dbAcertos.CursorLocation = adUseClient
dbAcertos.Open "select *from produtosacerto where codproduto=" & CodigoProduto & " order by datalancada", db, adOpenForwardOnly, adLockReadOnly
If dbAcertos.RecordCount <> 0 Then
  dbAcertos.MoveFirst
  Do While dbAcertos.EOF = False
    dbTurnos.MoveFirst
    dbEstoque.Filter = "datacaixa=#" & dbAcertos!datalancada & "# and codigoturno=" & dbTurnos!CodigoTurno
    If dbEstoque.EOF = True Then
      GoSub Adiciona3
    End If
    dbEstoque.Requery
    dbEstoque.Filter = ""
    dbEstoque.Filter = "datacaixa=#" & dbAcertos!datalancada & "# and codigoturno=" & dbTurnos!CodigoTurno
    dbEstoque!Acerto = dbEstoque!Acerto + dbAcertos!Valorutilizado
    dbEstoque.Update
    
    'DoEvents
    
    dbAcertos.MoveNext
  Loop
End If

dbEstoque.Filter = ""
dbEstoque.Requery

If dbEstoque.RecordCount <> 0 Then
  dbEstoque.MoveLast
  If Disponivel = 0 Then
    Disponivel = dbProdutos!Estoque
  End If
  Do While dbEstoque.BOF = False
    Abertura = Disponivel + dbEstoque!Saida - dbEstoque!Entrada - dbEstoque!Acerto
    
    dbEstoque!Disponivel = Disponivel
    dbEstoque!Abertura = Abertura
    dbEstoque.Update
    Disponivel = Abertura
    
    'DoEvents
    
    dbEstoque.MovePrevious
  Loop
End If






Termina:
  dbEstoque.Close
  dbVendas.Close
  dbEntradas.Close
  dbProdutos.Close
  dbTurnos.Close
  db.Close
  Unload frmMensagem
  
Exit Function

Adiciona:
  dbEstoque.AddNew
  dbEstoque!CodigoProduto = CodigoProduto
  dbEstoque!Codigo = dbProdutos!Codigo
  dbEstoque!Tanque = 0
  dbEstoque!DataCaixa = dbVendas!DataCaixa
  dbEstoque!Turno = dbVendas!Turno
  dbEstoque!CodigoTurno = dbVendas!CodigoTurno
  dbEstoque!HoraIni = dbVendas!HoraIni
  dbEstoque!Combustivel = dbProdutos!Combustivel
  dbEstoque!Abertura = 0
  dbEstoque!Entrada = 0
  dbEstoque!Saida = 0
  dbEstoque!Acerto = 0
  dbEstoque!Diferenca = 0
  dbEstoque!Disponivel = 0
  dbEstoque!dataalterado = Now
  dbEstoque!Usuario = Usuarios.Nome
  dbEstoque.Update
Return

Adiciona2:
  dbEstoque.AddNew
  dbEstoque!CodigoProduto = CodigoProduto
  dbEstoque!Codigo = dbProdutos!Codigo
  dbEstoque!Tanque = 0
  dbEstoque!DataCaixa = dbEntradas!datanota
  dbEstoque!Turno = dbTurnos!Descri
  dbEstoque!CodigoTurno = dbTurnos!CodigoTurno
  dbEstoque!HoraIni = dbTurnos!HoraIni
  dbEstoque!Combustivel = dbProdutos!Combustivel
  dbEstoque!Abertura = 0
  dbEstoque!Entrada = 0
  dbEstoque!Saida = 0
  dbEstoque!Acerto = 0
  dbEstoque!Diferenca = 0
  dbEstoque!Disponivel = 0
  dbEstoque!dataalterado = Now
  dbEstoque!Usuario = Usuarios.Nome
  dbEstoque.Update
Return

Adiciona3:
  dbEstoque.AddNew
  dbEstoque!CodigoProduto = CodigoProduto
  dbEstoque!Codigo = dbProdutos!Codigo
  dbEstoque!Tanque = 0
  dbEstoque!DataCaixa = dbAcertos!datalancada
  dbEstoque!Turno = dbTurnos!Descri
  dbEstoque!CodigoTurno = dbTurnos!CodigoTurno
  dbEstoque!HoraIni = dbTurnos!HoraIni
  dbEstoque!Combustivel = dbProdutos!Combustivel
  dbEstoque!Abertura = 0
  dbEstoque!Entrada = 0
  dbEstoque!Saida = 0
  dbEstoque!Acerto = 0
  dbEstoque!Diferenca = 0
  dbEstoque!Disponivel = 0
  dbEstoque!dataalterado = Now
  dbEstoque!Usuario = Usuarios.Nome
  dbEstoque.Update
Return

End Function

Public Function UltimoCaixa() As String
Dim db As New ADODB.Connection
Dim dbFechamentos As New ADODB.Recordset
Dim StrTemp As String

db.Open CaminhoADO
dbFechamentos.CursorLocation = adUseClient
dbFechamentos.Open "Select datacaixa, turno from fechamentodecaixa order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic

If dbFechamentos.RecordCount <> o Then
  StrTemp = "Digitado: " & Format(dbFechamentos!DataCaixa, "short date") & " - " & dbFechamentos!Turno
End If
dbFechamentos.Close

dbFechamentos.Open "Select datacaixa, turno from fechamentodecaixa where fechado=-1 order by datacaixa desc, horaini desc", db, adOpenKeyset, adLockOptimistic
If dbFechamentos.RecordCount <> o Then
  StrTemp = StrTemp & " | Finalizado: " & Format(dbFechamentos!DataCaixa, "short date") & " - " & dbFechamentos!Turno
End If
UltimoCaixa = StrTemp

db.Close
End Function

Public Sub AtualizaCartoesNoDia(ByVal CodigoCartao As Double)
Dim DataIni As Date, DataFim As Date, Dias As Integer, Mes As Boolean
Dim db As New ADODB.Connection
Dim dbCartoes As New ADODB.Recordset
Dim dbFormaDePG As New ADODB.Recordset

db.Open CaminhoADO
dbCartoes.CursorLocation = adUseClient
dbCartoes.Open "Select *from cartoes where codigocartao=" & CodigoCartao, db, adOpenForwardOnly, adLockReadOnly
If dbCartoes.RecordCount = 0 Then Exit Sub

dbFormaDePG.CursorLocation = adUseClient
dbFormaDePG.Open "Select *from formadepagamento where codigopagamento=" & dbCartoes!CodigoFormaPg, db, adOpenForwardOnly, adLockReadOnly

If dbFormaDePG.RecordCount = 0 Then Exit Sub

If dbFormaDePG!corte = False Then
  db.Execute "update formadepagamentorecebido2 set confirma='" & dbCartoes!CodigoCartao & "' where codigoformadepg=" & dbCartoes!CodigoFormaPg & " and data=#" & DataInglesa(dbCartoes!DataLanc) & "#"
Else
  DataFim = dbCartoes!DataLanc
  DataIni = DateAdd("d", -dbFormaDePG!diascorte, DataFim)
  db.Execute "update formadepagamentorecebido2 set confirma='" & dbCartoes!CodigoCartao & "' where codigoformadepg=" & dbCartoes!CodigoFormaPg & " and data between #" & DataInglesa(DataIni) & "# and #" & DataInglesa(DataFim) & "#"
End If

dbCartoes.Close
dbFormaDePG.Close
db.Close

End Sub

Public Function GeraNFE() As Boolean
Dim db As New ADODB.Connection
Dim dbNotas As New ADODB.Recordset
Dim dbNotasCorpo As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbPostos As New ADODB.Recordset
Dim dbMunicipios As New ADODB.Recordset
Dim NrNota As Double, Arquivo As String
Dim A As Integer
Dim StrLinha As String, StrTemp As String

GeraNFE = False

Arquivo = NomePosto & " " & Format(Now, "yyyy-mm-dd hh:nn") & ".txt"
A = FreeFile()

db.Open CaminhoADO
dbNotas.CursorLocation = adUseClient
dbNotas.Open "Select *from notas where eletronica=-1  order by notanr", db, adOpenKeyset, adLockOptimistic

dbNotasCorpo.CursorLocation = adUseClient
dbNotasCorpo.Open "Select *from notascorpo order by codigonota, codigonotacorpo", db, adOpenKeyset, adLockOptimistic

dbPostos.CursorLocation = adUseClient
dbPostos.Open "Select *from Postos", db, adOpenForwardOnly, adLockReadOnly

dbMunicipios.CursorLocation = adUseClient
dbMunicipios.Open "Select *from municipios order by nome", db, adOpenForwardOnly, adLockReadOnly

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos", db, adOpenForwardOnly, adLockReadOnly

If dbNotas.RecordCount = 0 Then
  NrNota = 0
Else
  dbNotas.MoveLast
  If IsNumeric(dbNotas!NrNota) = True Then
    NrNota = dbNotas!NrNota
  Else
    MsgBox "Erro ao gerar o número da prócima nota!"
    Exit Function
  End If
End If
dbNotas.Filter = "gerada=0"

If dbNotas.RecordCount = 0 Then
  MsgBox "Não existe nota para ser gerado o arquivo!"
  Exit Function
End If

Open Arquivo For Output As #A
StrLinha = "NOTA FISCAL|" & dbNotas.RecordCount
Print #A, StrLinha
StrLinha = "A|1.1.1|NFe"
Print #A, StrLinha

dbNotas.MoveLast
dbNotas.MoveFirst
Do While dbNotas.EOF = False
  NrNota = NrNota + 1
  StrTemp = ReadINI("Fiscal", "CodigoUFEmitente", "", App.Path & "\Posto.ini")
  
  If StrTemp = "" Then
    MsgBox "'Fiscal'->'CodigoUFEmitente' não cadastrado. Será assumido como Estado São Paulo=35, Municipio São Paulo=3550308"
    B = FreeFile()
    Open App.Path & "\posto.ini" For Append As #B
    Print #B, "[Fiscal]"
    Print #B, "CodigoUFEmitente=35"
    Print #B, "CodigoMunicipio=3550308"
    Close #B
    StrTemp = "35"
  End If
  
  'Inicio de B
  StrLinha = "B|" & StrTemp & "|" & dbNotas!NaturezaOP & "|1|55|0|" & NrNota & "|" & Format(dbnota!dataemissao, "yyyy-mm-dd") & "|" & Format(dbnota!datasaida, "yyyy-mm-dd") & "|"
  If dbNotas!Entrada = 0 Then
    StrTemp = "1"
  Else
    StrTemp = "0"
  End If
  StrLinha = StrLinha & StrTemp
  StrTemp = ReadINI("Fiscal", "CodigoMunicipio", "", App.Path & "\Posto.ini")
  If StrTemp = "" Then
    StrTemp = "3550308"
  End If
  StrLinha = StrLinha & "|" & StrTemp & "|1|1||2|1|3|1.1.1"
  
  Print #A, StrLinha
  
  'inicio de C
  StrLinha = "C|" & dbPostos!Nome & "||" & dbPostos!ie & "||||"
  Print #A, StrLinha
  
  StrLinha = "C02|" & dbPostos!CNPJ & "|"
  Print #A, StrLinha
  
  StrTemp = ""
  IntA = InStr(1, dbPostos!Endereco, ",")
  
  If IntA <> 0 Then
    StrTemp = Mid(dbPostos!Endereco, 1, IntA - 1)
    StrLinha = "C05|" & StrTemp
    StrTemp = "" 'numero1
    StrTemp = Mid(dbPostos!Endereco, IntA + 1)
    StrLinha = StrLinha & "|" & StrTemp
  Else
    StrTemp = Left(dbPostos!Endereco, 60)
    StrLinha = "C05|" & StrTemp
    StrTemp = "0" 'numero
    StrLinha = StrLinha & "|" & StrTemp
  End If
  
  StrLinha = StrLinha & "|" & dbPostos!complemento & "|" & dbPostos!bairro & "|"
  
  StrTemp = ReadINI("Fiscal", "CodigoMunicipio", "", App.Path & "\Posto.ini")
  If StrTemp = "" Then
    StrTemp = "3550308"
  End If
  
  StrLinha = StrLinha & StrTemp & "|" & dbPostos!municipio & "|" & dbPostos!Estado & "|" & dbPostos!CEP & "|1058|Brasil|" & dbPostos!Telefone & "|"
  
  Print #A, StrLinha
  
  
  StrLinha = "E|" & dbNotas!Nome & "|"
  If dbNotas!ie = "" Then
    StrTemp = "ISENTO"
  Else
    StrTemp = dbNotas!ie
  End If
  StrLinha = StrLinha & StrTemp & "||"
  
  Print #A, StrLinha
  
  StrTemp = RemoveString(dbNotas!CNPJ)
  If Len(StrTemp) > 11 Then
    StrLinha = "E02|" & Format(StrTemp, "00000000000000")
  Else
    StrLinha = "E03|" & Format(StrTemp, "00000000000")
  End If
  
  Print #A, StrLinha
  
  StrLinha = "E05|"
  If IntA <> 0 Then
    StrTemp = Mid(dbNotas!Endereco, 1, IntA - 1)
    StrLinha = StrLinha & "|" & StrTemp
    StrTemp = "" 'numero1
    StrTemp = Mid(dbNotas!Endereco, IntA + 1)
    StrLinha = StrLinha & "|" & StrTemp
  Else
    StrTemp = Left(dbNotas!Endereco, 60)
    StrLinha = StrLinha & "|" & StrTemp
    StrTemp = "0" 'numero
    StrLinha = StrLinha & "|" & StrTemp
  End If
  
  StrLinha = StrLinha & "||" & dbNotas!bairro
  
  If IsNull(dbNotas!municipio) = True Then
    StrTemp = "3550308"
  Else
    StrTemp = dbNotas!municipio
  End If
  StrLinha = StrLinha & "|" & StrTemp
  
  If dbMunicipios.RecordCount = 0 Then
    StrTemp = "São Paulo"
  Else
    dbMunicipios.MoveFirst
    dbMunicipios.Find "codigo='" & StrTemp & "'"
    If dbMunicipios.EOF = False Then
      StrTemp = dbMunicipios!Nome
    Else
      StrTemp = "São Paulo"
    End If
  End If
  StrLinha = StrLinha & "|" & StrTemp & "|" & dbNotas!uf & "|" & dbNotas!CEP & "|||" & dbNotas!fone
  
  Print #A, StrLinha
  
  dbNotasCorpo.Filter = "codigonota=" & dbNotas!CodigoNota
  
  If dbNotasCorpo.RecordCount <> 0 Then
    dbNotasCorpo.MoveLast
    dbNotasCorpo.MoveFirst
    
    Do While dbNotasCorpo.EOF = False
      
      StrLinha = "H|" & dbNotasCorpo.AbsolutePosition & "||"
      
      Print #A, StrLinha
      
      StrLinha = "I|" & dbNotasCorpo!CodigoProduto & "||" & dbNotasCorpo!descriproduto & "||||" & dbNotasCorpo!cfop & "|" & dbNotasCorpo!unidade & "|" & dbNotasCorpo!Quantidade & "|" & dbNotasCorpo!valorUnitario & "|" & dbNotasCorpo!ValorTotal & "||" & dbNotasCorpo!unidade & "|" & dbNotasCorpo!Quantidade & "|" & dbNotasCorpo!valorUnitario & "||||"
      
      Print #A, StrLinha
      
      Select Case dbNotasCorpo!clasfiscal
        Case "00"
          StrLinha = "N02|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|0|" & dbNotasCorpo!BaseICMS & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMS & "|"
        Case "10"
          StrLinha = "N03|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|4|" & dbNotasCorpo!BaseICMS & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMS & "|4|||" & dbNotasCorpo!baseicmssubst & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMSSubst & "|"
        Case "20"
          StrLinha = "N04|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|0|" & dbNotasCorpo!BaseICMS & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!BaseICMS & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMS & "|"
        Case "30"
          StrLinha = "N05|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|4|||" & dbNotasCorpo!baseicmssubst & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMSSubst & "|"
        Case "40", "41", "50"
          StrLinha = "N06|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|"
        Case "51"
          StrLinha = "N07|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "||||||"
        Case "60"
          StrLinha = "N08|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|" & dbNotasCorpo!baseicmssubst & "|" & dbNotasCorpo!ValorICMSSubst & "|"
        Case "70"
          StrLinha = "N09|" & dbNotasCorpo!Origem & "|" & dbNotasCorpo!clasfiscal & "|0|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMS & "|4|||" & dbNotasCorpo!baseicmssubst & "|" & dbNotasCorpo!aliquotaicms & "|" & dbNotasCorpo!ValorICMSSubst & "|"
        Case "90"
                
      End Select
      
      Print #A, StrLinha
      
      dbNotasCorpo.MoveNext
    Loop
    
  End If
  
  
  
  dbNotas.MoveNext
Loop


Close #A
GeraNFE = True
End Function

Public Function NotaExiste(ByVal NrNota As Double, ByVal CodigoNota As Double) As Boolean
Dim db As New ADODB.Connection
Dim dbNotas As New ADODB.Recordset

NotaExiste = True

db.Open CaminhoADO
dbNotas.Open "select *from notas where notanr=" & NrNota, db, adOpenForwardOnly, adLockReadOnly

If dbNotas.EOF = False Then
    If dbNotas.RecordCount > 1 Then
        NotaExiste = True
    Else
        If dbNotas!CodigoNota <> CodigoNota Then
            NotaExiste = True
        Else
            NotaExiste = False
        End If
    End If
Else
    NotaExiste = False
End If
dbNotas.Close
db.Close

End Function
