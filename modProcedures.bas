Attribute VB_Name = "modProcedures"

Public Function ExtornaCobranca(ByVal CodigoConcilia As Double) As Boolean
Dim db As New ADODB.Connection
Dim dbConcilia As New ADODB.Recordset
Dim dbPendencias As New ADODB.Recordset
Dim DbClientes As New ADODB.Recordset
Dim dbContas As New ADODB.Recordset

Dim TempValor As Currency, Taxa As Double, Juros As Currency
Dim ValorRecebido As Currency, ValorDesconto As Currency
Dim Resposta As Integer, Diferenca As Currency
Dim CodigoCliente As Double, Valor As Currency, Obs As String

ExtornaCobranca = False

db.Open CaminhoADO

dbConcilia.CursorLocation = adUseClient
dbConcilia.Open "Select *from concilianova where codigoConciliaconta=" & CodigoConcilia, db, adOpenKeyset, adLockOptimistic

If dbConcilia.RecordCount = 0 Then
    MsgBox "Registro não localizado!"
    Exit Function
End If

dbPendencias.CursorLocation = adUseClient
dbPendencias.Open "select *from clientescobranca where codigocobranca=" & dbConcilia!NrDocumento, db, adOpenKeyset, adLockOptimistic

If dbPendencias.RecordCount = 0 Then
    MsgBox "Cobrança não localizada!"
    Exit Function
End If

If dbPendencias!fechames = True Then
    MsgBox "Já foi feito o fechamento desta cobrança!"
    Exit Function
End If
'If dbPendencias!fechaaluguel = True Then
'    MsgBox "Já foi feito o fechamento desta cobrança!"
'    Exit Function
'End If

DbClientes.CursorLocation = adUseClient
DbClientes.Open "select *from clientes where codigocliente=" & dbPendencias!CodigoCliente, db, adOpenKeyset, adLockOptimistic

If DbClientes.RecordCount = 0 Then
    MsgBox "Cliente não localizado!"
    Exit Function
End If

dbContas.CursorLocation = adUseClient
dbContas.Open "select *from contas where codigoconta=" & dbPendencias!CodigoFormadePg, db, adOpenKeyset, adLockOptimistic

If dbContas.RecordCount = 0 Then
    MsgBox "Conta não localizada!"
    Exit Function
End If

ValorRecebido = dbPendencias!valorpago
Juros = dbPendencias!Juros


dbContas!Saldo = dbContas!Saldo - ValorRecebido
dbContas.Update

dbPendencias!Pago = False
dbPendencias!valorpago = 0
dbPendencias!Juros = 0
dbPendencias!DataPagamento = Null
dbPendencias!CodigoFormadePg = 0
dbPendencias!Descri = Null
dbPendencias!fechames = False
dbPendencias.Update


DbClientes!TotalBoleto = DbClientes!TotalBoleto + ValorRecebido
DbClientes!Saldo = DbClientes!Limite - DbClientes!TotalNotas - DbClientes!TotalBoleto
DbClientes.Update

db.Execute "delete from concilianova where codigoconciliaconta=" & CodigoConcilia

dbConcilia.Close
DbClientes.Close
dbPendencias.Close
dbContas.Close

db.Close

ExtornaCobranca = True

End Function

Public Function ConfirmaNota(ByVal CodigoNota As Double, ByVal DataRecebida As Date, ByVal CodigoTurno As Double, ByVal FormaDePagamento As String, ByVal NrNota As String) As Boolean
Dim db As New ADODB.Connection
Dim dbTurnos As New ADODB.Recordset
Dim dbNotasCorpo As New ADODB.Recordset
Dim dbTanque As New ADODB.Recordset
Dim dbPosto As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbProdutosHistorico As New ADODB.Recordset
Dim dbMovimento  As New ADODB.Recordset
Dim dbStatus As New ADODB.Recordset
Dim dbDespesaLanc As New ADODB.Recordset
Dim dbNotas  As New ADODB.Recordset
Dim dbBloqueiaFechamento As New ADODB.Recordset


Dim EstoqueAntigo As Double, PrecoCompraAntigo As Currency
Dim EstoqueNovo As Double, PrecoCompraNovo As Currency
Dim Variacao As Currency, Variacao2 As Currency, Venda As Currency
Dim TempComissao As Currency, Aguardando As Boolean, Parcelas As Integer, DiasParcelas As Integer
Dim Resposta As Integer, LucroVariacao As Currency, Total As Currency


db.Open CaminhoADO

dbTurnos.Open "select *from turnos order by horaIni", db, adOpenKeyset, adLockOptimistic

dbNotas.CursorLocation = adUseClient
dbNotas.Open "select *from produtosnotas where codigoentrada=" & CodigoNota, db, adOpenKeyset, adLockOptimistic
dbNotasCorpo.CursorLocation = adUseClient
dbNotasCorpo.Open "Select *from produtosnotascorpo where codigoprodutonota=" & CodigoNota, db, adOpenKeyset, adLockOptimistic
dbTanque.CursorLocation = adUseClient
dbTanque.Open "select *from tanques order by tanque", db, adOpenKeyset, adLockOptimistic
dbPosto.CursorLocation = adUseClient
dbPosto.Open "select *from postos order by Nome", db, adOpenKeyset, adLockOptimistic
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos order by codigo", db, adOpenKeyset, adLockOptimistic
dbStatus.CursorLocation = adUseClient
dbStatus.Open "select *from status", db, adOpenKeyset, adLockOptimistic
dbDespesaLanc.CursorLocation = adUseClient
dbDespesaLanc.Open "select *from despesaslanc2", db, adOpenKeyset, adLockOptimistic

dbBloqueiaFechamento.CursorLocation = adUseClient
dbBloqueiaFechamento.Open "Select *from bloqueiafechamento", db, adOpenKeyset, adLockOptimistic

ConfirmaNota = False

With dbBloqueiaFechamento
  If dbBloqueiaFechamento.EOF = False Then
    If dbBloqueiaFechamento!Data1 <= DataRecebida And dbBloqueiaFechamento!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & dbBloqueiaFechamento!Data1
      Exit Function
    End If
  End If
End With

If dbTurnos.RecordCount = 0 Then
  MsgBox "Tabela de turnos vazia."
  Exit Function
End If
dbTurnos.MoveFirst
dbTurnos.Find "codigoturno=" & CodigoTurno
If dbTurnos.EOF = True Or dbTurnos.BOF = True Then
  MsgBox "Informe um turno válido!"
  Exit Function
End If

If FormaDePagamento = "" Then
  MsgBox "Selecione uma forma de pagamento!"
  Exit Function
End If

'If DateDiff("d", Date, DataRecebida) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar recebimento futuro!"
'    Exit Function
'  End If
'End If
'If DateDiff("d", Date, DataRecebida) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar recebimento com data anterior a 15 dias!"
'    Exit Function
'  End If
'End If

If Trim(NrNota) = "" Then
  MsgBox "Informe um número de nota!"
  Exit Function
End If
If dbNotasCorpo.RecordCount = 0 Then
  MsgBox "Para confirmar uma nota deve existir pelo menos um produto lançado!"
  Exit Function
End If
dbNotasCorpo.MoveFirst
Aguardando = False

Do While dbNotasCorpo.EOF = False
  If dbNotasCorpo!Tanque <> 0 Then
    'Verifica se não vai ultrapassar o estoque máximo!
    If dbTanque.RecordCount <> 0 Then
      dbTanque.MoveFirst
      dbTanque.Filter = "tanque=" & dbNotasCorpo!Tanque
      If dbTanque.RecordCount = 0 Then
        MsgBox "Tanque " & dbNotasCorpo!Tanque & " não Cadastrado!"
        Exit Function
      End If
    End If
  End If
  dbTanque.Filter = ""
  
  dbProdutos.MoveFirst
  dbProdutos.Find "codigoproduto=" & dbNotasCorpo!CodigoProduto
  If dbProdutos.EOF = True Then
    MsgBox "Produto " & dbNotasCorpo!Descri & " com erro na tabela de produtos!"
    Exit Function
  End If
  dbNotasCorpo.MoveNext
Loop
dbNotasCorpo.MoveFirst

Do While dbNotasCorpo.EOF = False
  If dbNotasCorpo!Tanque <> 0 Then
    'acrecenta no estoque do tanque
    If Aguardando = False Then
      If dbTanque.RecordCount <> 0 Then
        dbTanque.MoveFirst
        dbTanque.Filter = "codigoposto=" & dbPosto!codigoPosto & " and tanque=" & dbNotasCorpo!Tanque
        If dbTanque.RecordCount <> 0 Then
          dbTanque!Estoque = dbTanque!Estoque + dbNotasCorpo!Quantidade
          dbTanque.Update
        End If
      End If
    End If
  End If
  dbTanque.Filter = ""
  
  If dbProdutos.RecordCount <> 0 Then
    'Registra na tabela de produtos
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbNotasCorpo!CodigoProduto
    If dbProdutos.EOF = False Then
      PrecoCompraAntigo = dbProdutos!precocompra
      If IsNull(dbProdutos!Estoque) = False Then
        EstoqueAntigo = dbProdutos!Estoque
      Else
        EstoqueAntigo = 0
      End If
      PrecoCompraNovo = dbNotasCorpo!Total / dbNotasCorpo!Quantidade
      EstoqueNovo = EstoqueAntigo + dbNotasCorpo!Quantidade
      Variacao = (EstoqueAntigo * PrecoCompraNovo) - (EstoqueAntigo * PrecoCompraAntigo)
      
      If IsNull(dbProdutos!ValorEstoque) = True Then
        dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
      End If
      If IsNull(dbProdutos!PrecoMedio) = True Then
        dbProdutos!PrecoMedio = dbProdutos!precocompra
      End If
      If IsNull(dbProdutos!DifEstoque) = True Then
        dbProdutos!DifEstoque = 0
      End If
      If IsNull(dbProdutos!valordifestoque) = True Then
        dbProdutos!valordifestoque = 0
      End If
      If IsNull(dbProdutos!LucroMedio) = True Then
        dbProdutos!LucroMedio = 0
      End If
      
      ValorProduto = dbNotasCorpo!Total
      
      Total = Total + dbNotasCorpo!Total
      
      dbProdutos!ValorEstoque = dbProdutos!ValorEstoque + ValorProduto
      dbProdutos!Estoque = EstoqueNovo
      dbProdutos!precocompra = PrecoCompraNovo
      If IsNull(dbProdutos!lucrominimo) = True Then
        dbProdutos!lucrominimo = 0
      End If
      Venda = 0
      dbProdutos!Variacao = dbProdutos!Variacao + Variacao
      If IsNull(dbProdutos!qtdcomprado) = True Then dbProdutos!qtdcomprado = 0
      dbProdutos!qtdcomprado = dbProdutos!qtdcomprado + dbNotasCorpo!Quantidade
      If IsNull(dbProdutos!valorcomprado) = True Then dbProdutos!valorcomprado = 0
      dbProdutos!valorcomprado = dbProdutos!valorcomprado + (dbNotasCorpo!Quantidade * PrecoCompraNovo)
      dbProdutos.Update
    End If
  End If
  
  StrTemp = "insert into produtoshistorico (lancadoem,dataalteracao,codigoproduto,codigo,descriproduto,descrioperacao,precocompra,precovenda," & _
            "estoqueAnterior,Quantidade,EstoqueFinal) values (#" & DataInglesa(Now) & "#,#" & DataInglesa(DataRecebida) & "#," & dbProdutos!CodigoProduto & "," & _
            dbProdutos!Codigo & ",'" & dbProdutos!Descri & "','" & "Entrada de Nota: " & NrNota & " - " & CodigoNota & "'," & NumeroIngles(dbProdutos!precocompra) & "," & _
            NumeroIngles(dbProdutos!PrecoVenda) & "," & NumeroIngles(EstoqueAntigo) & "," & NumeroIngles(dbNotasCorpo!Quantidade) & "," & NumeroIngles(dbProdutos!Estoque) & ")"
  
  db.Execute StrTemp
  
  StrTemp = "insert into produtosentrada2 (data,codigoproduto,codigo,descri,precoantigo,preconovo,variaestoque,quantidade,valornota,tanque,codigonota,formadepg) values (" & _
            "#" & DataInglesa(DataRecebida) & "#," & dbNotasCorpo!CodigoProduto & "," & dbNotasCorpo!Codigo & ",'" & dbNotasCorpo!Descri & "'," & _
            NumeroIngles(PrecoCompraAntigo) & "," & NumeroIngles(PrecoCompraNovo) & "," & NumeroIngles(Variacao) & "," & _
            NumeroIngles(dbNotasCorpo!Quantidade) & "," & NumeroIngles(dbNotasCorpo!Total) & "," & NumeroIngles(dbNotasCorpo!Tanque) & ",'" & _
            CodigoNota & "','" & FormaDePagamento & "')"
  
  db.Execute StrTemp
  
  dbStatus!variacaoestoque = dbStatus!variacaoestoque + Variacao
  dbStatus.Update
  
  
  RegistraEstoque DataRecebida, dbTurnos!CodigoTurno, dbTurnos!Descri, dbTurnos!HoraIni, dbNotasCorpo!CodigoProduto, dbNotasCorpo!Tanque, dbNotasCorpo!Quantidade
  
  dbNotasCorpo!Aguardando = Aguardando
  dbNotasCorpo.Update
  dbNotasCorpo.MoveNext
Loop

Parcelas = dbNotas!Parcelas
If Parcelas = 0 Then Parcelas = 1
If Parcelas > 1 Then
  DiasParcelas = dbNotas!Dias
End If
For i = 1 To Parcelas
  dbDespesaLanc.AddNew
  dbDespesaLanc!CodigoFechamento = 0
  dbDespesaLanc!Origem = "Despesa"
  dbDespesaLanc!Data = DataRecebida
  dbDespesaLanc!Hora = Now
  If Parcelas > 1 Then
    If i = 1 Then
      dbDespesaLanc!Vencimento = dbNotas!Vencimento
    Else
      dbDespesaLanc!Vencimento = DateAdd("d", (i - 1) * DiasParcelas, dbNotas!Vencimento)
    End If
  Else
    dbDespesaLanc!Vencimento = dbNotas!Vencimento
  End If
  dbDespesaLanc!CodigoDespesa = -1
  dbDespesaLanc!Descri = "Compra de Produto"
  dbDespesaLanc!Obs = Left(dbNotas!fornecedor, 25) & Left("-Nota Nr.: " & NrNota, 25)
  dbDespesaLanc!Valor = -Total / Parcelas
  dbDespesaLanc!Produto = True
  dbDespesaLanc!fechamentodiario = True
  dbDespesaLanc!pgantecipado = dbNotas!pgantecipado
  dbDespesaLanc!codigoenviar = "1"
  dbDespesaLanc!formadepg = FormaDePagamento
  dbDespesaLanc.Update
Next i
  
  
dbNotas!Confirmado = True
dbNotas!NrNota = NrNota
dbNotas!datanota = DataRecebida
dbNotas!formadepg = FormaDePagamento
dbNotas!CodigoTurno = dbTurnos!CodigoTurno
dbNotas.Update
  
ConfirmaNota = True

End Function

Public Function DesconfirmaNota(ByVal CodigoNota As Double, ByVal DataRecebida As Date, ByVal CodigoTurno As Double, ByVal FormaDePagamento As String, ByVal NrNota As String) As Boolean
Dim db As New ADODB.Connection
Dim dbTurnos As New ADODB.Recordset
Dim dbNotasCorpo As New ADODB.Recordset
Dim dbTanque As New ADODB.Recordset
Dim dbPosto As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbProdutosHistorico As New ADODB.Recordset
Dim dbMovimento  As New ADODB.Recordset
Dim dbStatus As New ADODB.Recordset
Dim dbDespesaLanc As New ADODB.Recordset
Dim dbNotas  As New ADODB.Recordset


Dim EstoqueAntigo As Double, PrecoCompraAntigo As Currency
Dim EstoqueNovo As Double, PrecoCompraNovo As Currency
Dim Variacao As Currency, Variacao2 As Currency, Venda As Currency
Dim TempComissao As Currency, Aguardando As Boolean, Parcelas As Integer, DiasParcelas As Integer
Dim Resposta As Integer, LucroVariacao As Currency, Total As Currency


db.Open CaminhoADO

dbTurnos.Open "select *from turnos order by horaIni", db, adOpenKeyset, adLockOptimistic

dbNotas.CursorLocation = adUseClient
dbNotas.Open "select *from produtosnotas where codigoentrada=" & CodigoNota, db, adOpenKeyset, adLockOptimistic
dbNotasCorpo.CursorLocation = adUseClient
dbNotasCorpo.Open "Select *from produtosnotascorpo where codigoprodutonota=" & CodigoNota, db, adOpenKeyset, adLockOptimistic
dbTanque.CursorLocation = adUseClient
dbTanque.Open "select *from tanques order by tanque", db, adOpenKeyset, adLockOptimistic
dbPosto.CursorLocation = adUseClient
dbPosto.Open "select *from postos order by Nome", db, adOpenKeyset, adLockOptimistic
dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "Select *from produtos order by codigo", db, adOpenKeyset, adLockOptimistic
dbStatus.CursorLocation = adUseClient
dbStatus.Open "select *from status", db, adOpenKeyset, adLockOptimistic
dbDespesaLanc.CursorLocation = adUseClient
dbDespesaLanc.Open "select *from despesaslanc2", db, adOpenKeyset, adLockOptimistic

DesconfirmaNota = False

If dbTurnos.RecordCount = 0 Then
  MsgBox "Tabela de turnos vazia."
  Exit Function
End If
dbTurnos.MoveFirst
dbTurnos.Find "codigoturno=" & CodigoTurno
If dbTurnos.EOF = True Or dbTurnos.BOF = True Then
  MsgBox "Informe um turno válido!"
  Exit Function
End If

If FormaDePagamento = "" Then
  MsgBox "Selecione uma forma de pagamento!"
  Exit Function
End If

'If DateDiff("d", Date, DataRecebida) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar recebimento futuro!"
'    Exit Function
'  End If
'End If
'If DateDiff("d", Date, DataRecebida) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar recebimento com data anterior a 15 dias!"
'    Exit Function
'  End If
'End If

If Trim(NrNota) = "" Then
  MsgBox "Informe um número de nota!"
  Exit Function
End If
If dbNotasCorpo.RecordCount = 0 Then
  MsgBox "Para confirmar uma nota deve existir pelo menos um produto lançado!"
  Exit Function
End If
dbNotasCorpo.MoveFirst
Aguardando = False

Do While dbNotasCorpo.EOF = False
  If dbNotasCorpo!Tanque <> 0 Then
    'Verifica se não vai ultrapassar o estoque máximo!
    If dbTanque.RecordCount <> 0 Then
      dbTanque.MoveFirst
      dbTanque.Filter = "tanque=" & dbNotasCorpo!Tanque
      If dbTanque.RecordCount = 0 Then
        MsgBox "Tanque " & dbNotasCorpo!Tanque & " não Cadastrado!"
        Exit Function
      End If
    End If
  End If
  dbTanque.Filter = ""
  
  dbProdutos.MoveFirst
  dbProdutos.Find "codigoproduto=" & dbNotasCorpo!CodigoProduto
  If dbProdutos.EOF = True Then
    MsgBox "Produto " & dbNotasCorpo!Descri & " com erro na tabela de produtos!"
    Exit Function
  End If
  dbNotasCorpo.MoveNext
Loop
dbNotasCorpo.MoveFirst

Do While dbNotasCorpo.EOF = False
  If dbNotasCorpo!Tanque <> 0 Then
    'acrecenta no estoque do tanque
    If Aguardando = False Then
      If dbTanque.RecordCount <> 0 Then
        dbTanque.MoveFirst
        dbTanque.Filter = "codigoposto=" & dbPosto!codigoPosto & " and tanque=" & dbNotasCorpo!Tanque
        If dbTanque.RecordCount <> 0 Then
          dbTanque!Estoque = dbTanque!Estoque - dbNotasCorpo!Quantidade
          dbTanque.Update
        End If
      End If
    End If
  End If
  dbTanque.Filter = ""
  
  If dbProdutos.RecordCount <> 0 Then
    'Registra na tabela de produtos
    dbProdutos.MoveFirst
    dbProdutos.Find "codigoproduto=" & dbNotasCorpo!CodigoProduto
    If dbProdutos.EOF = False Then
      PrecoCompraAntigo = dbProdutos!precocompra
      If IsNull(dbProdutos!Estoque) = False Then
        EstoqueAntigo = dbProdutos!Estoque
      Else
        EstoqueAntigo = 0
      End If
      PrecoCompraNovo = dbNotasCorpo!Total / dbNotasCorpo!Quantidade
      EstoqueNovo = EstoqueAntigo - dbNotasCorpo!Quantidade
      Variacao = (EstoqueAntigo * PrecoCompraNovo) - (EstoqueAntigo * PrecoCompraAntigo)
      
      If IsNull(dbProdutos!ValorEstoque) = True Then
        dbProdutos!ValorEstoque = dbProdutos!precocompra * dbProdutos!Estoque
      End If
      If IsNull(dbProdutos!PrecoMedio) = True Then
        dbProdutos!PrecoMedio = dbProdutos!precocompra
      End If
      If IsNull(dbProdutos!DifEstoque) = True Then
        dbProdutos!DifEstoque = 0
      End If
      If IsNull(dbProdutos!valordifestoque) = True Then
        dbProdutos!valordifestoque = 0
      End If
      If IsNull(dbProdutos!LucroMedio) = True Then
        dbProdutos!LucroMedio = 0
      End If
      
      ValorProduto = dbNotasCorpo!Total
      
      Total = Total + dbNotasCorpo!Total
      
      dbProdutos!ValorEstoque = dbProdutos!ValorEstoque - ValorProduto
      dbProdutos!Estoque = EstoqueNovo
      dbProdutos!precocompra = PrecoCompraNovo
      If IsNull(dbProdutos!lucrominimo) = True Then
        dbProdutos!lucrominimo = 0
      End If
      Venda = 0
      dbProdutos!Variacao = dbProdutos!Variacao - Variacao
      If IsNull(dbProdutos!qtdcomprado) = True Then dbProdutos!qtdcomprado = 0
      dbProdutos!qtdcomprado = dbProdutos!qtdcomprado - dbNotasCorpo!Quantidade
      If IsNull(dbProdutos!valorcomprado) = True Then dbProdutos!valorcomprado = 0
      dbProdutos!valorcomprado = dbProdutos!valorcomprado - (dbNotasCorpo!Quantidade * PrecoCompraNovo)
      dbProdutos.Update
    End If
  End If
  
  db.Execute "insert into produtoshistorico (lancadoem,dataalteracao,codigoproduto,codigo,descriproduto,descrioperacao,precocompra,precovenda," & _
  "estoqueAnterior,Quantidade,EstoqueFinal) values (#" & DataInglesa(Now) & "#,#" & DataInglesa(DataRecebida) & "#," & dbProdutos!CodigoProduto & "," & _
  dbProdutos!Codigo & ",'" & dbProdutos!Descri & "','" & "Estorno de Nota: " & NrNota & " - " & CodigoNota & "'," & NumeroIngles(dbProdutos!precocompra) & "," & _
  NumeroIngles(dbProdutos!PrecoVenda) & "," & NumeroIngles(EstoqueAntigo) & "," & NumeroIngles(dbNotasCorpo!Quantidade) & "," & NumeroIngles(dbProdutos!Estoque) & ")"
  
  
  dbStatus!variacaoestoque = dbStatus!variacaoestoque - Variacao
  dbStatus.Update
  
  
  RegistraEstoque DataRecebida, dbTurnos!CodigoTurno, dbTurnos!Descri, dbTurnos!HoraIni, dbNotasCorpo!CodigoProduto, dbNotasCorpo!Tanque, -dbNotasCorpo!Quantidade
  
  dbNotasCorpo!Aguardando = 0
  dbNotasCorpo.Update
  dbNotasCorpo.MoveNext
Loop

db.Execute "delete from produtosentrada2 where codigonota=" & CodigoNota

Parcelas = dbNotas!Parcelas

If Parcelas = 0 Then Parcelas = 1

If Parcelas > 1 Then
  DiasParcelas = dbNotas!Dias
End If
For i = 1 To Parcelas
  dbDespesaLanc.AddNew
  dbDespesaLanc!CodigoFechamento = 0
  dbDespesaLanc!Origem = "Despesa"
  dbDespesaLanc!Data = DataRecebida
  dbDespesaLanc!Hora = Now
  If Parcelas > 1 Then
    If i = 1 Then
      dbDespesaLanc!Vencimento = dbNotas!Vencimento
    Else
      dbDespesaLanc!Vencimento = DateAdd("d", (i - 1) * DiasParcelas, dbNotas!Vencimento)
    End If
  Else
    dbDespesaLanc!Vencimento = dbNotas!Vencimento
  End If
  dbDespesaLanc!CodigoDespesa = -1
  dbDespesaLanc!Descri = "Estorno de Produto"
  dbDespesaLanc!Obs = Left(dbNotas!fornecedor, 25) & Left("-Nota Nr.: " & NrNota, 25)
  dbDespesaLanc!Valor = Total / Parcelas
  dbDespesaLanc!Produto = True
  dbDespesaLanc!fechamentodiario = True
  dbDespesaLanc!pgantecipado = dbNotas!pgantecipado
  dbDespesaLanc!codigoenviar = "1"
  dbDespesaLanc!formadepg = FormaDePagamento
  dbDespesaLanc.Update
Next i
  
  
dbNotas!Confirmado = False
dbNotas!NrNota = NrNota
dbNotas!datanota = DataRecebida
dbNotas!formadepg = FormaDePagamento
dbNotas!CodigoTurno = dbTurnos!CodigoTurno
dbNotas.Update
  
DesconfirmaNota = True

End Function


Public Function ProtestoAdicionaHistorico(ByVal DataLanc As Date, ByVal DataDocumento As Date, ByVal TipoDocumento As String, ByVal Status As String, ByVal Descri As String, ByVal Valor As Currency, ByVal Nome As String, ByVal Documento As String, Optional CodigoClienteNota As Double = 0, Optional CodigoClienteCheque As Double = 0) As Boolean
Dim db As New ADODB.Connection
ProtestoAdicionaHistorico = False
On Error Resume Next

db.Open CaminhoADO
db.Execute "insert into protestos(DataLanc,DataDocumento,TipoDocumento,Status,Descri,Valor,CodigoClienteNota,CodigoClienteCheque,Nome,CPF_CNPJ) values('" & DataLanc & "','" & DataDocumento & "','" & TipoDocumento & "','" & Status & "','" & Descri & "','" & Valor & "','" & CodigoClienteNota & "','" & CodigoClienteCheque & "','" & Nome & "','" & Documento & "')"
db.Close

If Err.Number = 0 Then
  ProtestoAdicionaHistorico = True
Else
  MsgBox "Erro: " & Err.Number & " - " & Err.Description
End If
End Function

Public Sub CorrigeTanque(ByVal CodigoProduto As Double)
Dim db As New ADODB.Connection, dbTanques As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset, dbAguardando As New ADODB.Recordset
Dim TotalPendente As Double, TotalEstoque As Double, TotalTanque As Double
Dim ValorACorrigir As Double

db.Open CaminhoADO
dbTanques.Open "select *from tanques where codigoproduto=" & CodigoProduto, db, adOpenKeyset, adLockOptimistic
dbProdutos.Open "select *from produtos where combustivel=-1 and codigoproduto=" & CodigoProduto, db, adOpenKeyset, adLockOptimistic
dbAguardando.Open "select *from produtosnotascorpo where aguardando=-1 and codigoproduto=" & CodigoProduto, db, adOpenKeyset, adLockOptimistic

If dbTanques.RecordCount <> 0 Then
  dbTanques.MoveLast
  dbTanques.MoveFirst
  Do While dbTanques.EOF = False
    TotalTanque = TotalTanque + dbTanques!Estoque
    dbTanques.MoveNext
  Loop
  TotalEstoque = dbProdutos!Estoque
  If dbAguardando.EOF = False And dbAguardando.BOF = False Then
    dbAguardando.MoveLast
    dbAguardando.MoveFirst
    Do While dbAguardando.EOF = False
      TotalPendente = TotalPendente + dbAguardando!Quantidade
      dbAguardando.MoveNext
    Loop
  End If
  ValorACorrigir = TotalEstoque - TotalPendente - TotalTanque
  dbTanques.MoveFirst
  dbTanques!Estoque = dbTanques!Estoque + ValorACorrigir
  dbTanques.Update
End If

End Sub

