VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm mdiPosto 
   BackColor       =   &H8000000C&
   Caption         =   "Posto de Combustível"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   1275
   ClientWidth     =   10665
   Icon            =   "mdiPosto.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR1 
      Left            =   6240
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   4680
      Top             =   3360
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":150A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":1826
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":1B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":1E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":2186
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":24A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPosto.frx":27BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clientes"
            Object.ToolTipText     =   "Cadastro de Clientes de Cheque"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Contas"
            Object.ToolTipText     =   "Cadastro de Contas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Produtos"
            Object.ToolTipText     =   "Cadastro de Produtos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fecha1"
            Object.ToolTipText     =   "Fechamento de Caixa"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fecha2"
            Object.ToolTipText     =   "Conferência de valores"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PgAntecipado"
            Object.ToolTipText     =   "Pagamentos Antecipados"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ContasPg"
            Object.ToolTipText     =   "Contas a pagar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Deposito"
            Object.ToolTipText     =   "Depósito de cheques"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Concilia"
            Object.ToolTipText     =   "Conciliação Bancária"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transfere"
            Object.ToolTipText     =   "Transferência Bancária"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Seleciona"
            Object.ToolTipText     =   "Seleciona Posto"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7020
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6897
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Versão: "
            TextSave        =   "Versão: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1130
            MinWidth        =   1130
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "12/03/2016"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "15:01"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Dif. Status:"
            TextSave        =   "Dif. Status:"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCad 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCadBicos 
         Caption         =   "&Bombas de Combustível"
      End
      Begin VB.Menu mnuCadClientes 
         Caption         =   "Clientes de Nota"
      End
      Begin VB.Menu mnuCadClienteCheque 
         Caption         =   "Clientes de Cheque"
      End
      Begin VB.Menu mnuCadContas 
         Caption         =   "Co&ntas"
      End
      Begin VB.Menu mnuCadDespesasTipo 
         Caption         =   "Despesas/Tipo"
      End
      Begin VB.Menu mnuCadDespesaBanco 
         Caption         =   "Despesas Bancárias/Tipo"
      End
      Begin VB.Menu mnuCadFormaDePg 
         Caption         =   "&Forma de Pagamento"
      End
      Begin VB.Menu mnuCadFornecedor 
         Caption         =   "F&ornecedores"
      End
      Begin VB.Menu mnuCadVendedores 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mnuCadJuros 
         Caption         =   "&Juros"
      End
      Begin VB.Menu mnuCadPlanoDeConta 
         Caption         =   "Planos de Contas"
      End
      Begin VB.Menu mnuCadPostos 
         Caption         =   "P&ostos"
      End
      Begin VB.Menu mnuCadProdutos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnuCadProdutosFornecedores 
         Caption         =   "Produtos / Fornecedores"
      End
      Begin VB.Menu mnuCadTanque 
         Caption         =   "Tanques"
      End
      Begin VB.Menu mnuCadTurnos 
         Caption         =   "&Turnos / Pdvs"
      End
      Begin VB.Menu mnuCadConfigura 
         Caption         =   "Seleciona Posto"
      End
      Begin VB.Menu mnuCadConfiguracao 
         Caption         =   "Configuração"
      End
      Begin VB.Menu mnuBackUp 
         Caption         =   "Back Up"
      End
      Begin VB.Menu mnuCadH1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadAlteraSenha 
         Caption         =   "Alterar Senha"
      End
      Begin VB.Menu mnuCadComissoes 
         Caption         =   "Verifica Comissões"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCadAtualiza 
         Caption         =   "Atualização"
      End
      Begin VB.Menu mnuCadAtualizaSPED 
         Caption         =   "Atualização para SPED"
      End
      Begin VB.Menu mnuCadDownload 
         Caption         =   "Download"
      End
      Begin VB.Menu mnuCadH2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuControle 
      Caption         =   "C&ontrole"
      Begin VB.Menu mnuControleFechaDia 
         Caption         =   "Fechamento &Diário"
      End
      Begin VB.Menu mnuConferencia 
         Caption         =   "Conferência de Caixa"
      End
      Begin VB.Menu mnuControleRecebimentos 
         Caption         =   "Cartões Pendentes"
      End
      Begin VB.Menu mnuControlePgAntecipado 
         Caption         =   "Pagamentos Antecipados"
      End
      Begin VB.Menu mnuControleNotas 
         Caption         =   "Lançamento de Notas"
      End
      Begin VB.Menu mnuControleDespesaLanc 
         Caption         =   "Lançamento de Contas a Pagar"
      End
      Begin VB.Menu mnuControlePgDespesa 
         Caption         =   "Contas a Pagar"
      End
      Begin VB.Menu mnuControleContasAReceber 
         Caption         =   "Contas a Receber"
      End
      Begin VB.Menu mnuControleFaturaCliente 
         Caption         =   "Faturamento de Clientes"
      End
      Begin VB.Menu mnuControleCobra 
         Caption         =   "Cobrança de Clientes"
      End
      Begin VB.Menu mnuControleNotaFiscalAvulsa 
         Caption         =   "Nota Fiscal Avulsa"
      End
      Begin VB.Menu mnuControlePrecos 
         Caption         =   "Alteração de Preços"
      End
      Begin VB.Menu mnuControleAgua 
         Caption         =   "Controle de Água"
      End
      Begin VB.Menu mnuControleLuz 
         Caption         =   "Controle de Luz"
      End
      Begin VB.Menu mnuControleLavagem 
         Caption         =   "Controle de Lavagem"
      End
      Begin VB.Menu mnuControleVales 
         Caption         =   "Vales de Funcionários"
      End
      Begin VB.Menu mnuControleVendasLeituraX 
         Caption         =   "Vendas e Leitura X"
      End
      Begin VB.Menu mnuControleNFPModelo1 
         Caption         =   "Nota Fiscal Paulista Modelo 1"
      End
   End
   Begin VB.Menu mnuCheques 
      Caption         =   "Cheques"
      Begin VB.Menu mnuBancoDepositaCheque 
         Caption         =   "Depósito de Cheque"
      End
      Begin VB.Menu mnuBancoDevolucao 
         Caption         =   "Devolução de Cheques"
      End
      Begin VB.Menu mnuBancoCobraCheque 
         Caption         =   "Cobrança de Cheque"
      End
      Begin VB.Menu mnuChequeProtesto 
         Caption         =   "Protesto de Cheque"
      End
      Begin VB.Menu mnuChequeEmpresaDeCobranca 
         Caption         =   "Enviar para Empresa de Cobrança"
      End
      Begin VB.Menu mnuChequesResgatados 
         Caption         =   "Cheques e Boletos Resgatados"
      End
      Begin VB.Menu mnuBancoChequesData 
         Caption         =   "Cheques Por Data"
      End
   End
   Begin VB.Menu mnuBanco 
      Caption         =   "Banco"
      Begin VB.Menu mnuControleConcilia 
         Caption         =   "Conciliação"
      End
      Begin VB.Menu mnuBancoTransfere 
         Caption         =   "Transferência"
      End
   End
   Begin VB.Menu mnuRelat 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelatEst 
         Caption         =   "Produtos"
         Begin VB.Menu mnuRelatAcertoEstoque 
            Caption         =   "Acerto de Estoque"
         End
         Begin VB.Menu mnuRelatComprasProd 
            Caption         =   "Compras/Vendas"
         End
         Begin VB.Menu mnuRelatEstoquePorCaixa 
            Caption         =   "Estoque Por Caixa"
         End
         Begin VB.Menu mnuRelatExtratoProdutos 
            Caption         =   "Extrato de produtos"
         End
         Begin VB.Menu mnuRelatMovimentaProdutos 
            Caption         =   "Movimentação de Produtos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRelatCompras 
            Caption         =   "Produtos Comprados"
         End
         Begin VB.Menu mnuRelatVendaComissoes 
            Caption         =   "Venda / Lucro de Produtos Comissionados"
         End
         Begin VB.Menu mnuRelatVendas 
            Caption         =   "Vendas de Produtos"
         End
         Begin VB.Menu mnuRelatVendaDetalhada 
            Caption         =   "Venda Detalhada"
         End
         Begin VB.Menu mnuRelatVendaDiariaCombustivel 
            Caption         =   "Venda diária de Combustivel"
         End
         Begin VB.Menu mnuRelatMediaVenda 
            Caption         =   "Venda Média de Produtos"
         End
         Begin VB.Menu mnuRelatVendasPorDia 
            Caption         =   "Vendas Por Dia"
         End
      End
      Begin VB.Menu mnuRelatAnexoNotasCobranca 
         Caption         =   "Anexo de Notas para Cobrança"
      End
      Begin VB.Menu mnuRelatChequeCliente 
         Caption         =   "Cheques Por Cliente"
      End
      Begin VB.Menu mnuRelatDiferencaCaixa 
         Caption         =   "Diferença de Caixa"
      End
      Begin VB.Menu mnuRelatDifRecebimentos 
         Caption         =   "Diferença de Recebimentos"
      End
      Begin VB.Menu mnuRelatDifComb 
         Caption         =   "Diferença de Combustível"
      End
      Begin VB.Menu mnuRelatFormaDePg 
         Caption         =   "Forma de Pagamento"
      End
      Begin VB.Menu mnuRelatFormaDePgBordero 
         Caption         =   "Forma de Pagamento por Borderô"
      End
      Begin VB.Menu mnuRelatGalonagem 
         Caption         =   "Galonagem"
      End
      Begin VB.Menu mnuRelatGalonagemTotal 
         Caption         =   "Galonagem Total"
      End
      Begin VB.Menu mnuRelatProtesto 
         Caption         =   "Protestos de Cheques e Boletos"
      End
      Begin VB.Menu mnuRelatCadIncompleto 
         Caption         =   "Cadastro Incompleto"
      End
      Begin VB.Menu mnuRelatRetorno 
         Caption         =   "Retorno de Combustivel"
      End
      Begin VB.Menu mnuRelatFaturaCheque 
         Caption         =   "Faturamento de Cheques"
      End
      Begin VB.Menu mnuRelatKilometragem 
         Caption         =   "Kilometragem de Clientes"
      End
      Begin VB.Menu mnuRelatEstacionamento 
         Caption         =   "Controle de Estacionamento"
      End
      Begin VB.Menu mnuRelatContasAReceber 
         Caption         =   "Contas a Receber"
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "Administração"
      Begin VB.Menu mnuAdmConfirmaDespesa 
         Caption         =   "Confirmar Despesas"
      End
      Begin VB.Menu mnuAdmDespesasParaContabilidade 
         Caption         =   "Despesas Para Contabilidade"
      End
      Begin VB.Menu mnuAdmEstatus 
         Caption         =   "Estatus"
      End
      Begin VB.Menu mnuAdmBloqueiaFinaliza 
         Caption         =   "Bloqueia Finalização"
      End
      Begin VB.Menu mnuAdmAlteraPreco 
         Caption         =   "Alterar Preço de Combustível"
      End
      Begin VB.Menu mnuAdmTotalVenda 
         Caption         =   "Total de Venda"
      End
      Begin VB.Menu mnuAdmLMC 
         Caption         =   "LMC"
      End
      Begin VB.Menu mnuAdmUsuarios 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuAdmGrupos 
         Caption         =   "Grupos de Usuários"
      End
      Begin VB.Menu mnuRelatFaturaCliente 
         Caption         =   "Faturamento de Clientes"
      End
   End
End
Attribute VB_Name = "mdiPosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub VerificaNovaVersao()
CaminhoAtualiza = ReadINI("Atualiza", "CaminhoVersao", "", App.Path & "\Posto.ini")
If CaminhoAtualiza = "" Then
  A = FreeFile()
  Open App.Path & "\Posto.ini" For Append As #A
  Print #A, "[Atualiza]"
  Print #A, ";0 - não"
  Print #A, ";1 - sim"
  Print #A, "Internet = 0"
  Print #A, "CaminhoVersao=\\servidor01\rede\AtualizaPosto.ini"
  Close #A
End If
StrTemp = ReadINI("Atualiza", "Internet", "", App.Path & "\Posto.ini")
If StrTemp = 0 Then
  AtualizaInternet = False
Else
  AtualizaInternet = True
End If

If StrTemp <> 3 Then AtualizaSistema AtualizaInternet, CaminhoAtualiza

End Sub

Public Function VerificaPendencias()
Dim Ws As Workspace, db As Database, dbTemp As Recordset, dbTemp2 As Recordset
Dim TempData As Date, DiasTravar As String, Dias As Integer

On Error GoTo TrataErro

If mnuControleCobra.Enabled = True Then
  Set Ws = DBEngine.Workspaces(0)
  Set db = Ws.OpenDatabase(Caminho, , , Conectar)
  
  DiasTravar = ReadINI("Notas no Caixa", "Dias", "5", App.Path & "\Posto.ini")
  If DiasTravar <> "" Then
    Dias = -CInt(DiasTravar)
  End If
  TempData = DateAdd("d", Dias, Date)
  Select Case Weekday(TempData)
    Case 1
      TempData = DateAdd("d", -2, TempData)
    Case 7
      TempData = DateAdd("d", -1, TempData)
  End Select
  'Desativa os clientes que não pagaram.
  Set dbTemp = db.OpenRecordset("select *from clientescobranca where pago=0 and datafechamento<=#" & DataInglesa(TempData) & "# order by datafechamento")
  If dbTemp.RecordCount <> 0 Then
    Set dbTemp2 = db.OpenRecordset("select *from clientes where mensalista=-1")
    dbTemp.MoveLast
    dbTemp.MoveFirst
    Do While dbTemp.EOF = False
      If dbTemp2.RecordCount <> 0 Then
        dbTemp2.FindFirst "codigocliente=" & dbTemp!CodigoCliente
        If dbTemp2.NoMatch = False Then
          'MsgBox "O cliente " & dbTemp2!Nome & " está com boleto vencido e será desativado!"
          dbTemp2.Edit
          dbTemp2!mensalista = False
          On Error Resume Next
          dbTemp2!desativado = Now
          If Err.Number <> 0 Then
            AtualizaDbAntigo
          End If
          dbTemp2.Update
          Set dbTemp2 = db.OpenRecordset("select *from clientes where mensalista=-1")
        End If
      Else
        Exit Do
      End If
      dbTemp.MoveNext
    Loop
  End If
  
  'Reativa os clientes que pagaram.
  Set dbTemp = db.OpenRecordset("select *from clientescobranca where pago=0 and datafechamento<=#" & DataInglesa(TempData) & "# and protestado=0 order by datafechamento")
  If dbTemp.RecordCount <> 0 Then
    Set dbTemp2 = db.OpenRecordset("select *from clientes where mensalista=0")
    Do While dbTemp2.EOF = False
      dbTemp.FindFirst "codigocliente=" & dbTemp2!CodigoCliente
      If dbTemp.NoMatch = True Then
        If dbTemp!protestador = False Then
          dbTemp2.Edit
          dbTemp2!mensalista = True
          dbTemp2.Update
        End If
      End If
      dbTemp2.MoveNext
    Loop
  End If
  
  'Desativa os clientes que não abastecem + de 30 dias
  Set dbTemp = db.OpenRecordset("select ultimoabastecimento, protestado from clientes where protestado=0 and ultimoabastecimento<#" & DataInglesa(DateAdd("m", -3, Date)) & "#")
  If dbTemp.RecordCount <> 0 Then
    Do While dbTemp.EOF = False
      dbTemp.Edit
      dbTemp!protestado = True
      dbTemp.Update
      dbTemp.MoveNext
    Loop
  End If
End If

Exit Function
TrataErro:
  
End Function

Private Sub MDIForm_Activate()
'If Usuarios.Grupo.AdmEstatus = 2 Then
'  On Error Resume Next
'  frmDiferencaCombustivel.Show
'  frmDiferencaCombustivel.SetFocus
'  With frmDiferencaCombustivel
'    .Move Screen.Width - .Width - 100, Screen.Height - .Height - 2200
'  End With
'End If
End Sub

Private Sub MDIForm_Load()
Dim AtualizaInternet As Boolean, CaminhoAtualiza As String
Dim db As New ADODB.Connection
Dim dbPostos As New ADODB.Recordset

With StatusBar1
  .Panels(2).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
End With
With frmSplash
  .lblWarning.Caption = "Verificando se existe cliente com fatura atrasada!"
  .lblWarning.Refresh
End With
VerificaPendencias

With frmSplash
  .lblWarning.Caption = "Verificando diferença de status!"
  .lblWarning.Refresh
End With

With frmSplash
  .lblWarning.Caption = "Verificando novas versões do programa!"
  .lblWarning.Refresh
End With

VerificaNovaVersao



db.Open CaminhoADO
dbPostos.CursorLocation = adUseClient
dbPostos.Open "Select *from postos order by nome", db, adOpenDynamic, adLockOptimistic
On Error Resume Next
If dbPostos.RecordCount <> 0 Then
  ComissaoAcumulativa = dbPostos!ComissaoAcumulativa
End If
dbPostos.Close
db.Close
Timer1_Timer

End Sub

Private Sub mnuAdmAlteraPreco_Click()
Screen.MousePointer = vbHourglass
frmAdmAlteraPreco.Show
frmAdmAlteraPreco.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmBloqueiaFinaliza_Click()
Screen.MousePointer = vbHourglass
frmBloqueiaFinalizacao.Show
frmBloqueiaFinalizacao.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmConfirmaDespesa_Click()
Screen.MousePointer = vbHourglass
frmDespesasConfirma.Show
frmDespesasConfirma.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmDespesasParaContabilidade_Click()
Screen.MousePointer = vbHourglass
frmDespesasParaContabilidade.Show
frmDespesasParaContabilidade.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmEstatus_Click()
Screen.MousePointer = vbHourglass
'frmEstatus.Show
'frmEstatus.SetFocus
frmEstatus2.Show
frmEstatus2.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmGrupos_Click()
'Permissao = False
'frmPermissao2.Show vbModal
'If Permissao = False Then Exit Sub
Screen.MousePointer = vbHourglass
frmUsuariosGrupo.Show
frmUsuariosGrupo.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmLMC_Click()
Screen.MousePointer = vbHourglass
frmRelatLMC.Show
frmRelatLMC.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmTotalVenda_Click()
Screen.MousePointer = vbHourglass
frmRelatTotalVendido.Show
frmRelatTotalVendido.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAdmUsuarios_Click()
'Permissao = False
'frmPermissao2.Show vbModal
'If Permissao = False Then Exit Sub
Screen.MousePointer = vbHourglass
frmUsuarios.Show
frmUsuarios.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBackUp_Click()
frmBackup.Show vbModal
End Sub

Private Sub mnuBancoChequesData_Click()
Screen.MousePointer = vbHourglass
frmConciliaChequeTotalData.Show
frmConciliaChequeTotalData.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBancoCobraCheque_Click()
Screen.MousePointer = vbHourglass
frmConciliaChequeCobranca.Show
frmConciliaChequeCobranca.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBancoDepositaCheque_Click()
Screen.MousePointer = vbHourglass
frmConciliaDeposito.Show
frmConciliaDeposito.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBancoDevolucao_Click()
Screen.MousePointer = vbHourglass
frmConciliaDevolucao.Show
frmConciliaDevolucao.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBancoTransfere_Click()
Screen.MousePointer = vbHourglass
frmConciliaTransfere.Show
frmConciliaTransfere.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadAlteraSenha_Click()
Dim Ws As Workspace, db As Database, dbUsuarios As Recordset
Dim StrTemp As String, StrTemp2 As String

If Usuarios.Nome = "Administrador" Then
  MsgBox "O usuário atual não pode alterar a senha!", vbCritical, "Erro!"
  Exit Sub
End If

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(CaminhoUsuarios, , , Conectar)
Set dbUsuarios = db.OpenRecordset("select *from usuarios where nome='" & Usuarios.Nome & "'")

If dbUsuarios.RecordCount = 0 Then
  MsgBox "Erro na tabela de Usuários!", vbCritical, "Erro!"
  Exit Sub
End If

StrTemp = dbUsuarios("senha")
StrTemp = Criptografa(StrTemp, 225)

frmUsuarioAlteraSenha.Show vbModal

If frmUsuarioAlteraSenha.Confirma = False Then
  Unload frmUsuarioAlteraSenha
  Exit Sub
End If
With frmUsuarioAlteraSenha
  If StrTemp <> .txtAtual.Text Then
    MsgBox "A senha atual não confere!", vbCritical, "Erro!"
    Unload frmUsuarioAlteraSenha
    Exit Sub
  End If
  If Len(.txtNova.Text) < 6 Then
    MsgBox "A senha deve ter pelo menos 6 dígitos!", vbCritical, "Erro!"
    Unload frmUsuarioAlteraSenha
    Exit Sub
  End If
  StrTemp = Criptografa(.txtNova.Text, 225)
  StrTemp2 = Criptografa(.txtNova.Text, 224)
  dbUsuarios.Edit
  dbUsuarios("senha") = StrTemp
  dbUsuarios("confirma") = StrTemp2
  dbUsuarios.Update
End With
End Sub

Private Sub mnuCadAtualiza_Click()
Dim StrTemp As String, CaminhoAtual As String
Dim Resposta As Integer, CaminhoAtual2 As String

Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset


Resposta = MsgBox("Deseja atualizar somente o posto atual ou todos os postos?" & Chr(vbKeyReturn) & "Sim - atualizar somente o atual." & Chr(vbKeyReturn) & "Não - atualiza todos." & Chr(vbKeyReturn) & "Cancelar - não atualiza nenhum.", vbYesNoCancel)

If Resposta = vbCancel Then Exit Sub

If Resposta = vbYes Then
  AtualizaUsuarios
  AtualizaADO001
  AtualizaAdo002
  AtualizaAdo2
  Resposta = MsgBox("Deseja fazer atualização antiga?", vbYesNo)
  Select Case Resposta
    Case vbYes
      AtualizaDb
      AtualizaDbAntigo
  End Select
Else
  
  CaminhoAtual = Caminho
  CaminhoAtual2 = CaminhoADO
  
  db.Open CaminhoUsuariosAdo
  dbTemp.Open "select *from CaminhoNovo order by nome", db, adOpenKeyset
  
  If dbTemp.RecordCount <> 0 Then
    dbTemp.MoveLast
    dbTemp.MoveFirst
    CaminhoAtual = Caminho
    Do While dbTemp.EOF = False
      frmMensagem.Show
      frmMensagem.SetFocus
      frmMensagem.Refresh
      With frmMensagem.lblMensagem
        .Caption = "Atualizando tabelas de " & dbTemp!Nome & ". Aguarde..."
        .Refresh
      End With
      DoEvents
      If Dir(dbTemp!Dados) <> "" Then
        Caminho = dbTemp!Dados
        If Provedor = "SQLOLEDB.1" Then
          CaminhoAtual2 = "Provider=" & Provedor & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Maria Vitoria;Data Source=temvale17"
        Else
          CaminhoAtual2 = "Provider=" & Provedor & ";Data Source=" & Caminho & ";Persist Security Info=False"
        End If
        AtualizaADO001
        AtualizaAdo002
        AtualizaAdo2 CaminhoAtual2
      End If
      dbTemp.MoveNext
    Loop
    Unload frmMensagem
  End If
  
  dbTemp.Close
  db.Close
  
  
  Caminho = CaminhoAtual
  CaminhoAtual2 = CaminhoADO
  
  MsgBox "Tabelas Atualizadas."
  
End If
End Sub

Private Sub mnuCadAtualizaSPED_Click()
Screen.MousePointer = vbHourglass
AtualizaADOSPED
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadBicos_Click()
Screen.MousePointer = vbHourglass
frmCadBombas.Show
frmCadBombas.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadClienteCheque_Click()
Screen.MousePointer = vbHourglass
frmCadChequeCliente.Show
frmCadChequeCliente.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadClientes_Click()
Screen.MousePointer = vbHourglass
frmCadClientes.Show
frmCadClientes.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadComissoes_Click()
frmComissoes2.Show
frmComissoes2.SetFocus
End Sub

Private Sub mnuCadConfigura_Click()
Unload Me
Main
'frmSelecionaPosto.Show vbModal
'Screen.MousePointer = vbDefault
'Call MDIForm_Load
'frmUsuarioLogin.Show vbModal
'VerificaAcesso
End Sub

Private Sub mnuCadConfiguracao_Click()
frmCadConfigura.Show vbModal
End Sub

Private Sub mnuCadContas_Click()
Screen.MousePointer = vbHourglass
With frmCadContas
  .Show
  .SetFocus
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadDespesaBanco_Click()
Screen.MousePointer = vbHourglass
frmCadDespBanco.Show
frmCadDespBanco.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadDespesasTipo_Click()
Screen.MousePointer = vbHourglass
frmCadDespesasTipo.Show
frmCadDespesasTipo.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadDownload_Click()
Shell "http://temvale.selfip.com:9099/atualizapostodecombustivel.exe"
End Sub

Private Sub mnuCadFormaDePg_Click()
Screen.MousePointer = vbHourglass
frmCadFormaDePg.Show
frmCadFormaDePg.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadFornecedor_Click()
Screen.MousePointer = vbHourglass
frmCadFornecedores.Show
frmCadFornecedores.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadJuros_Click()
Screen.MousePointer = vbHourglass
frmCadJuros.Show
frmCadJuros.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadPlanoDeConta_Click()
Screen.MousePointer = vbHourglass
frmCadPlanoDeConta.Show
frmCadPlanoDeConta.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadPostos_Click()
Screen.MousePointer = vbHourglass
frmCadPosto.Show
frmCadPosto.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadProdutos_Click()
Screen.MousePointer = vbHourglass
frmCadProdutos.Show
frmCadProdutos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadProdutosFornecedores_Click()
Screen.MousePointer = vbHourglass
frmCadProdutosFornecedor.Show
frmCadProdutosFornecedor.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadSair_Click()
End
End Sub

Private Sub mnuCadTanque_Click()
Screen.MousePointer = vbHourglass
frmCadTanque.Show
frmCadTanque.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadTurnos_Click()
Screen.MousePointer = vbHourglass
frmCadTurnos.Show
frmCadTurnos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCadVendedores_Click()
Screen.MousePointer = vbHourglass
frmCadVendedor.Show
frmCadVendedor.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChequeEmpresaDeCobranca_Click()
Permissao = False
frmPermissao2.Show vbModal
If Permissao = False Then Exit Sub
Screen.MousePointer = vbHourglass
frmConciliaEmpresaCobranca.Show
frmConciliaEmpresaCobranca.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChequeProtesto_Click()
Screen.MousePointer = vbHourglass
frmConciliaProtesto.Show
frmConciliaProtesto.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChequesResgatados_Click()
Screen.MousePointer = vbHourglass
frmRelatResgatados.Show
frmRelatResgatados.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuConferencia_Click()
Screen.MousePointer = vbHourglass
frmFechamentoDeCaixaConfere.Show
frmFechamentoDeCaixaConfere.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleAgua_Click()
Screen.MousePointer = vbHourglass
frmControleAgua.Show
frmControleAgua.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleCobra_Click()
Screen.MousePointer = vbHourglass
frmClientesNotas.Show
frmClientesNotas.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleConcilia_Click()
Screen.MousePointer = vbHourglass
Dim Conciliacao As New frmConciliacaoNova
Conciliacao.Show
Conciliacao.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleContasAReceber_Click()
Screen.MousePointer = vbHourglass
frmCadastroContasAReceber.Show
frmCadastroContasAReceber.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleDespesaLanc_Click()
Screen.MousePointer = vbHourglass
frmDespesasLanc.Show
frmDespesasLanc.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleFaturaCliente_Click()
Screen.MousePointer = vbHourglass
frmFaturaClientes.Show
frmFaturaClientes.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleFechaDia_Click()
Screen.MousePointer = vbHourglass
'If Usuarios.Nome = "Usuário Master" Then
'  With frmFechamentoDeCaixa
'    .Show
'    .SetFocus
'  End With
'End If
'frmFechamentoDeCaixa.Show
frmFechamentoDeCaixaNovo.Show
frmFechamentoDeCaixaNovo.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleLavagem_Click()
Screen.MousePointer = vbHourglass
frmLavagem.Show
frmLavagem.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleLuz_Click()
Screen.MousePointer = vbHourglass
frmControleLuz.Show
frmControleLuz.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleNFPModelo1_Click()
frmNFPModelo1.Show vbModal
End Sub

Private Sub mnuControleNotaFiscalAvulsa_Click()
Screen.MousePointer = vbHourglass
frmNotaFiscalAvulsa.Show
frmNotaFiscalAvulsa.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleNotas_Click()
Screen.MousePointer = vbHourglass
frmPedidoDeCompra.Show
frmPedidoDeCompra.SetFocus
'FrmEntradaDeProdutos.Show
'FrmEntradaDeProdutos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControlePgAntecipado_Click()
Screen.MousePointer = vbHourglass
frmPgAntecipado.Show
frmPgAntecipado.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControlePgDespesa_Click()
Screen.MousePointer = vbHourglass
frmDespesasPagar.Show
frmDespesasPagar.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControlePrecos_Click()
Screen.MousePointer = vbHourglass
frmCadProdutosPreco.Show
frmCadProdutosPreco.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleRecebimentos_Click()
Screen.MousePointer = vbHourglass
frmCartoes.Show
frmCartoes.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleVales_Click()
Screen.MousePointer = vbHourglass
frmControleVales.Show
frmControleVales.SetFocus

'frmControlePagamento.Show
'frmControlePagamento.SetFocus

Screen.MousePointer = vbDefault
End Sub

Private Sub mnuControleVendasLeituraX_Click()
Screen.MousePointer = vbHourglass
frmVendasLeituraX.Show
frmVendasLeituraX.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatAcertoEstoque_Click()
Screen.MousePointer = vbHourglass
frmRelatAcertoEstoque.Show
frmRelatAcertoEstoque.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatAnexoNotasCobranca_Click()
Screen.MousePointer = vbHourglass
frmRelatAnexoNotas.Show
frmRelatAnexoNotas.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatCadIncompleto_Click()
Screen.MousePointer = vbHourglass
frmRelatClientesChequesIncompletos.Show
frmRelatClientesChequesIncompletos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatChequeCliente_Click()
Screen.MousePointer = vbHourglass
frmRelatChequesClientes.Show
frmRelatChequesClientes.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatCompras_Click()
Screen.MousePointer = vbHourglass
frmRelatCompras.Show
frmRelatCompras.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatComprasProd_Click()
Screen.MousePointer = vbHourglass
frmRelatCompra.Show
frmRelatCompra.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatContasAReceber_Click()
Screen.MousePointer = vbHourglass
frmRelatContasAReceber.Show
frmRelatContasAReceber.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatDifComb_Click()
Screen.MousePointer = vbHourglass
frmRelatDifComb.Show
frmRelatDifComb.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatDiferencaCaixa_Click()
Screen.MousePointer = vbHourglass
frmRelatDifCaixa.Show
frmRelatDifCaixa.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatDifRecebimentos_Click()
Screen.MousePointer = vbHourglass
frmRelatDifRecebimentos.Show
frmRelatDifRecebimentos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatEstacionamento_Click()
Screen.MousePointer = vbHourglass
frmRelatEstacionamento.Show
frmRelatEstacionamento.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatEstoquePorCaixa_Click()
Screen.MousePointer = vbhourglas
frmRelatPosicaoDoEstoque.Show
frmRelatPosicaoDoEstoque.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatExtratoProdutos_Click()
Screen.MousePointer = vbhourglas
frmRelatExtratoProdutos.Show
frmRelatExtratoProdutos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatFaturaCheque_Click()
Screen.MousePointer = vbHourglass
frmRelatFaturaCheque.Show
frmRelatFaturaCheque.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatFaturaCliente_Click()
Screen.MousePointer = vbHourglass
frmRelatFaturamentoClientes.Show
frmRelatFaturamentoClientes.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatFormaDePg_Click()
Screen.MousePointer = vbHourglass
frmRelatFormaDePg.Show
frmRelatFormaDePg.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatFormaDePgBordero_Click()
Screen.MousePointer = vbHourglass
frmRelatCartoesPorBordero.Show
frmRelatCartoesPorBordero.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatGalonagem_Click()
Screen.MousePointer = vbHourglass
frmRelatGalonagem.Show
frmRelatGalonagem.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatGalonagemTotal_Click()
Screen.MousePointer = vbHourglass
frmRelatGalonagemTotal.Show
frmRelatGalonagemTotal.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatKilometragem_Click()
Screen.MousePointer = vbHourglass
frmRelatKilometragem.Show
frmRelatKilometragem.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatMediaVenda_Click()
Screen.MousePointer = vbHourglass
frmRelatVendasMedia.Show
frmRelatVendasMedia.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatMovimentaProdutos_Click()
Dim MovimentaProdutos As New frmRelatMovimentacaoProdutos
Screen.MousePointer = vbHourglass
MovimentaProdutos.Show
MovimentaProdutos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatProtesto_Click()
Screen.MousePointer = vbHourglass
frmRelatProtestos.Show
frmRelatProtestos.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatRetorno_Click()
Screen.MousePointer = vbHourglass
frmRelatRetorno.Show
frmRelatRetorno.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatVendaComissoes_Click()
Screen.MousePointer = vbHourglass
frmRelatVendasComissao.Show
frmRelatVendasComissao.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatVendaDetalhada_Click()
Screen.MousePointer = vbHourglass
frmRelatVendasDetalhado.Show
frmRelatVendasDetalhado.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatVendaDiariaCombustivel_Click()
Screen.MousePointer = vbHourglass
frmRelatVendaDiaria.Show
frmRelatVendaDiaria.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatVendas_Click()
Screen.MousePointer = vbHourglass
frmRelatVendas.Show
frmRelatVendas.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRelatVendasPorDia_Click()
Screen.MousePointer = vbHourglass
frmRelatResumoDia.Show
frmRelatResumoDia.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()


StatusBar1.Panels(1).Text = UltimoCaixa

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Clientes"
    If mnuCadClienteCheque.Enabled = True Then
      Call mnuCadClienteCheque_Click
    End If
  Case "Fecha1"
    If mnuControleFechaDia.Enabled = True Then
      Call mnuControleFechaDia_Click
    End If
  Case "Fecha2"
    If mnuConferencia.Enabled = True Then
      Call mnuConferencia_Click
    End If
  Case "PgAntecipado"
    If mnuControlePgAntecipado.Enabled = True Then
      Call mnuControlePgAntecipado_Click
    End If
  Case "ContasPg"
    If mnuControlePgDespesa.Enabled = True Then
      Call mnuControlePgDespesa_Click
    End If
  Case "Deposito"
    If mnuBancoDepositaCheque.Enabled = True Then
      Call mnuBancoDepositaCheque_Click
    End If
  Case "Concilia"
    If mnuControleConcilia.Enabled = True Then
      Call mnuControleConcilia_Click
    End If
  Case "Transfere"
    If mnuBancoTransfere.Enabled = True Then
      Call mnuBancoTransfere_Click
    End If
  Case "Seleciona"
    Call mnuCadConfigura_Click
  Case "Contas"
    If mnuCadContas.Enabled = True Then
      Call mnuCadContas_Click
    End If
  Case "Produtos"
    If mnuCadProdutos.Enabled = True Then
      Call mnuCadProdutos_Click
    End If
End Select
End Sub

