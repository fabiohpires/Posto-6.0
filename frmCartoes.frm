VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCartoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartões Pendentes"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "frmCartoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExibeTrasfere 
      Caption         =   "Exibe Transferências"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Data dbCartoesTransfere 
      Caption         =   "dbCartoesTransfere"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Meus documentos\Projeto For Windows\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CartoesTransfereHistorico"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibe Cartões"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "c:\rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contas"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtOperacoes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
   Begin VB.Data dbCartoes 
      Caption         =   "dbCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "c:\rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FormaDePagamento"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdTransfere 
      Caption         =   "Transfere"
      Height          =   300
      Left            =   6120
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   135659521
      CurrentDate     =   38191
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   6600
      Picture         =   "frmCartoes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Imprimir"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Data qPendentesTotal 
      Caption         =   "qPendentesTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valorliquido) as liquido from cartoes where confirmado=0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qPendentes 
      Caption         =   "qPendentes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from Cartoes where confirmado=0 order by dataprevista"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmCartoes.frx":0EC4
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "frmCartoes.frx":0EDD
      TabIndex        =   11
      Top             =   3600
      Width           =   7335
   End
   Begin VB.Data qCartoesTotal 
      Caption         =   "qCartoesTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2333
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(cartoes.valorliquido) as liquido from cartoes where confirmado=0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qCartoes 
      Caption         =   "qCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2333
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmCartoes.frx":1DFC
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCartoes.frx":1F1F
      Height          =   2655
      Left            =   480
      OleObjectBlob   =   "frmCartoes.frx":1F36
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label5 
      Caption         =   "Operações:"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Dt Prevista:"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Valor Bruto:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblTotalPendente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
End
Attribute VB_Name = "frmCartoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

StrTemp = "Relatório de Cartões Pendentes"
Printer.FontSize = 16
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short Time")
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = qCartoes.Recordset!Descri
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Lançado"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Dt. Prevista"
Printer.CurrentX = 30
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 60
Printer.Print StrTemp;

StrTemp = "V. Bruto"
Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "V. Líquido"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub cmdExibeTrasfere_Click()
frmCartoesTransfere.Show vbModal
End Sub

Private Sub cmdExibir_Click()
Dim Bruto As Currency, Liquido As Currency

If qPendentes.Recordset.EOF = True Then
  MsgBox "Escolha um cartão pendente!"
  Exit Sub
End If
With frmCartoesFinalizacao.QFormaDePgRecebido
  Load frmCartoesFinalizacao
  .Connect = Conectar
  .DatabaseName = Caminho
  '.RecordSource = "select *from qformadepgrecebidofechamento2 where codigoformadepg=" & qPendentes.Recordset!CodigoFormaPg & " and FormaDePagamentoRecebido2.confirma='" & qPendentes.Recordset!CodigoCartao & "' and fechamentodiario=-1 order by datacaixa, horaini, valorbruto"
  StrTemp = "FormaDePagamentoRecebido2.*, FormaDePagamento.*, Contas.*, FechamentoDeCaixa.* FROM (FormaDePagamentoRecebido2 LEFT JOIN (Contas RIGHT JOIN FormaDePagamento ON Contas.CodigoConta = FormaDePagamento.CodigoConta) ON FormaDePagamentoRecebido2.CodigoFormaDePg = FormaDePagamento.CodigoPagamento) left  JOIN FechamentoDeCaixa ON FormaDePagamentoRecebido2.CodigoFechamento = FechamentoDeCaixa.CodigoFechamento"
  .RecordSource = "select " & StrTemp & " where FormaDePagamentoRecebido2.codigoformadepg=" & qPendentes.Recordset!CodigoFormaPg & " and data=#" & DataInglesa(qPendentes.Recordset!DataLanc) & "# and FormaDePagamentoRecebido2.fechamentodiario=-1 order by data, valorbruto"
  .Refresh
End With
With frmCartoesFinalizacao.qTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  StrTemp = "FormaDePagamentoRecebido2 left JOIN (Contas right JOIN FormaDePagamento ON Contas.CodigoConta = FormaDePagamento.CodigoConta) ON FormaDePagamentoRecebido2.CodigoFormaDePg = FormaDePagamento.CodigoPagamento"
  .RecordSource = "select sum(valorBruto) as Bruto, sum(Valor) as Liquido from " & StrTemp & " where FormaDePagamentoRecebido2.codigoformadepg=" & qPendentes.Recordset!CodigoFormaPg & " and data=#" & DataInglesa(qPendentes.Recordset!DataLanc) & "# and FormaDePagamentoRecebido2.fechamentodiario=-1"
  .Refresh
  If IsNull(.Recordset!Bruto) = False Then
    Bruto = .Recordset!Bruto
  End If
  If IsNull(.Recordset!Liquido) = False Then
    Liquido = .Recordset!Liquido
  End If
End With
If frmCartoesFinalizacao.QFormaDePgRecebido.Recordset.RecordCount = 0 Then
  If qPendentes.Recordset.EOF = True Then
    MsgBox "Escolha um cartão pendente!"
    Exit Sub
  End If
  With frmCartoesFinalizacao.QFormaDePgRecebido
    Load frmCartoesFinalizacao
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from qformadepgrecebidofechamento2 where codigoformadepg=" & qPendentes.Recordset!CodigoFormaPg & " and data=#" & DataInglesa(qPendentes.Recordset!DataLanc) & "# and fechamentodiario=-1 order by datacaixa, horaini, valorbruto"
    .Refresh
  End With
  With frmCartoesFinalizacao.qTotal
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select sum(valorBruto) as Bruto, sum(Valor) as Liquido from qformadepgrecebidofechamento2 where codigoformadepg=" & qPendentes.Recordset!CodigoFormaPg & " and data=#" & DataInglesa(qPendentes.Recordset!DataLanc) & "# and fechamentodiario=-1"
    .Refresh
    If IsNull(.Recordset!Bruto) = False Then
      Bruto = .Recordset!Bruto
    End If
    If IsNull(.Recordset!Liquido) = False Then
      Liquido = .Recordset!Liquido
    End If
  End With
End If
With frmCartoesFinalizacao
  .lblTotalBruto.Caption = Format(Bruto, "Currency")
  .lblTotalLiquido.Caption = Format(Liquido, "Currency")
End With
frmCartoesFinalizacao.Show vbModal
End Sub

Private Sub cmdImprime_Click()
Dim Largura As Double, Dia As Date
Dim StrTemp As String, Liquido As Currency, Bruto As Currency

Dim Resposta As Integer

Resposta = MsgBox("Deseja imprimir somente os totais?", vbYesNo)

If Resposta = vbYes Then
    
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    
    ImprimeGrid DBGrid1, Printer, qCartoes, 2, True, , , , , "Cartões a receber", , Chr(vbKeyReturn) & "Data Impresso:" & Format(Now, "short date") & "-" & Format(Now, "short time")
    Printer.EndDoc
    
Else
  With qPendentes
    .Refresh
    If .Recordset.RecordCount = 0 Then Exit Sub
    
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Arial"
    Largura = 190
    Dia = Now
    
    Cabeca Largura, Dia
    
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 25 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        StrTemp = "Página: " & Printer.Page
        Printer.CurrentX = 0
        Printer.Print StrTemp;
        
        StrTemp = Format(Bruto, "currency")
        Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(Liquido, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.CurrentY = 0
        Printer.NewPage
        Cabeca Largura, Dia
      End If
      
      StrTemp = Format(.Recordset!DataLanc, "Short Date")
      Printer.CurrentX = 0
      Printer.Print StrTemp;
      
      StrTemp = Format(.Recordset!DataPrevista, "Short Date")
      Printer.CurrentX = 30
      Printer.Print StrTemp;
      
      StrTemp = .Recordset!Conta
      Printer.CurrentX = 60
      Printer.Print StrTemp;
      
      Bruto = Bruto + .Recordset!ValorBruto
      StrTemp = Format(.Recordset!ValorBruto, "Currency")
      Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      Liquido = Liquido + .Recordset!valorliquido
      StrTemp = Format(.Recordset!valorliquido, "currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    StrTemp = "Página: " & Printer.Page
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(Bruto, "currency")
    Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(Liquido, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.EndDoc
  End With

End If
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTransfere_Click()
Dim Valor As Currency, Liquido As Currency, Dia As Date
Dim Taxa As Double, Qtd As Double, DiaOrigem As Date
Dim DataPrevista As Date
Dim BrutoOrigem As Currency, LiquidoOrigem As Currency
Dim BrutoDestino As Currency, LiquidoDestino As Currency
Dim CodigoOrigem As Double, CodigoFormadePg As Double

With qPendentes
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione qual o valor a ser debitado!"
    Exit Sub
  End If
  Dia = txtData.Value
  DiaOrigem = .Recordset!DataPrevista
  CodigoFormadePg = .Recordset!CodigoFormaPg
  CodigoOrigem = .Recordset!CodigoCartao
  .Recordset.FindFirst "dataprevista = #" & DataInglesa(Trim(Str(Dia))) & "# and confirmado=0"
  If .Recordset.NoMatch = True Then
    dbCartoes.Recordset.FindFirst "codigopagamento=" & .Recordset!CodigoFormaPg
    If dbCartoes.Recordset.NoMatch = False Then
      dbContas.Recordset.FindFirst "codigoconta=" & dbCartoes.Recordset!CodigoConta
      If dbContas.Recordset.NoMatch = True Then
        MsgBox "O cartão atual não possue uma conta destino"
        Exit Sub
      End If
      .Recordset.AddNew
      .Recordset!CodigoFormaPg = dbCartoes.Recordset!CodigoPagamento
      .Recordset!Grupo = dbCartoes.Recordset!Grupo
      .Recordset!Descri = dbCartoes.Recordset!Descri
      .Recordset!DataPrevista = Dia
      If dbCartoes.Recordset!reembolso > 0 Then
        If dbCartoes.Recordset!Mes = True Then
          DataPrevista = DateAdd("m", -dbCartoes.Recordset!reembolso, Dia)
        Else
          DataPrevista = DateAdd("d", -dbCartoes.Recordset!reembolso, Dia)
        End If
      Else
        DataPrevista = Dia
      End If
      .Recordset!DataLanc = DataPrevista
      .Recordset!ValorBruto = 0
      .Recordset!valorliquido = 0
      .Recordset!ValorRecebido = 0
      .Recordset!Diferenca = 0
      .Recordset!fechadiferenca = False
      .Recordset!CodigoConta = dbCartoes.Recordset!CodigoConta
      .Recordset!Conta = dbContas.Recordset!Descri
      .Recordset!fechataxa = False
      .Recordset.Update
    Else
      .Recordset.FindFirst "dataprevista = #" & DataInglesa(Trim(Str(DiaOrigem))) & "#"
      MsgBox "A data destino não foi encontrada!"
      txtData.SetFocus
      Exit Sub
    End If
  End If
  If IsNumeric(txtValor.Text) = False Then
    MsgBox "Digite um valor correto!"
    txtValor.SetFocus
    Exit Sub
  End If
  dbCartoes.Recordset.FindFirst "codigopagamento=" & .Recordset!CodigoFormaPg
  If dbCartoes.Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de forma de pagamento!"
    Exit Sub
  End If
  Taxa = dbCartoes.Recordset!DescontoPorcento / 100
  If dbCartoes.Recordset!descontoporoperacao <> 0 Then
    If IsNumeric(txtOperacoes.Text) = False Then
      MsgBox "Digite um valor correto!"
      txtOperacoes.SetFocus
      Exit Sub
    End If
    Qtd = txtOperacoes.Text * dbCartoes.Recordset!descontoporoperacao
  End If
  Valor = CCur(txtValor.Text)
  Liquido = Valor - (Valor * Taxa) - Qtd
  .Recordset.FindFirst "codigocartao = " & CodigoOrigem & " and confirmado=0"
  .Recordset.Edit
  BrutoOrigem = .Recordset!ValorBruto
  LiquidoOrigem = .Recordset!valorliquido
  .Recordset!ValorBruto = .Recordset!ValorBruto - Valor
  .Recordset!valorliquido = .Recordset!valorliquido - Liquido
  .Recordset.Update
  .Refresh
  .Recordset.FindFirst "dataprevista = #" & DataInglesa(Trim(Str(Dia))) & "# and confirmado=0"
  .Recordset.Edit
  BrutoDestino = .Recordset!ValorBruto
  LiquidoDestino = .Recordset!valorliquido
  .Recordset!ValorBruto = .Recordset!ValorBruto + Valor
  .Recordset!valorliquido = .Recordset!valorliquido + Liquido
  .Recordset.Update
  .Refresh
  .Recordset.FindFirst "dataprevista = #" & DataInglesa(Trim(Str(DiaOrigem))) & "# and confirmado=0"
End With

With dbCartoesTransfere
  .Recordset.AddNew
  .Recordset!CodigoFormadePg = CodigoFormadePg
  .Recordset!CodigoCartao = CodigoOrigem
  .Recordset!datatransfere = Now
  .Recordset!dataorigem = DiaOrigem
  .Recordset!datadestino = Dia
  .Recordset!ValorBruto = Valor
  .Recordset!valorliquido = Liquido
  .Recordset!BrutoOrigem = BrutoOrigem
  .Recordset!LiquidoOrigem = LiquidoOrigem
  .Recordset!BrutoDestino = BrutoDestino
  .Recordset!LiquidoDestino = LiquidoDestino
  .Recordset!Usuario = Usuarios.Nome
  .Recordset.Update
End With
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
If qPendentes.RecordSource = "select *from cartoes where descri='" & qCartoes.Recordset!Descri & "' and confirmado=0 order by " & DBGrid2.Columns(ColIndex).DataField Then
  qPendentes.RecordSource = "select *from cartoes where descri='" & qCartoes.Recordset!Descri & "' and confirmado=0 order by " & DBGrid2.Columns(ColIndex).DataField & " desc"
Else
  qPendentes.RecordSource = "select *from cartoes where descri='" & qCartoes.Recordset!Descri & "' and confirmado=0 order by " & DBGrid2.Columns(ColIndex).DataField
End If
qPendentes.Refresh
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyEscape
    Call cmdSair_Click
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
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
txtData.Value = Date
With qCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCartoesTransfere
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qCartoesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Liquido) = False Then
    lblTotal.Caption = Format(.Recordset!Liquido, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

If Usuarios.Nome = "Usuário Master" Then
  txtValor.Enabled = True
  txtOperacoes.Enabled = True
  txtData.Enabled = True
  cmdTransfere.Enabled = True
End If
Select Case Usuarios.Grupo.ControleCartoes
  Case 1 'Somente leitura
    'cmdTransfere.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub qCartoes_Reposition()
On Error Resume Next
With qPendentes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cartoes where CodigoFormaPg=" & qCartoes.Recordset!CodigoFormaPg & " and confirmado=0 order by datalanc"
  .Refresh
End With
With qPendentesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valorliquido) as total from cartoes where CodigoFormaPg=" & qCartoes.Recordset!CodigoFormaPg & " and confirmado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalPendente.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalPendente.Caption = Format(0, "Currency")
  End If
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
