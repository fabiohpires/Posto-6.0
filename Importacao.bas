Attribute VB_Name = "Importacao"
Public CaminhoImporta As String, strMdb As String, strSql As String
Public SoPrimeira As Boolean
Public Configura As ConfigIni

Public Type ConfigIni
  NotaNoCaixa As Integer
  NotaBloqueia As Integer
  ChequesNoCaixa As Integer
End Type

Public Function AbreCaixa(ByVal CaminhoAdo As String, ByVal PDV As Double, ByVal DataCaixa As Date, ByVal CodigoTurno As Double) As Double
Dim CodigoFechamento As Double

Dim db As New ADODB.Connection
Dim dbFechamentos As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim dbBicos As New ADODB.Recordset
Dim dbEncerrantes As New ADODB.Recordset
Dim dbTanques As New ADODB.Recordset
Dim dbDifComb As New ADODB.Recordset
Dim dbPdvs As New ADODB.Recordset


AbreCaixa = 0

'On Error GoTo TrataErro

db.Open CaminhoAdo

dbPdvs.CursorLocation = adUseClient
dbPdvs.Open "select *from pdvs where codigo='" & PDV & "'", db, adOpenKeyset, adLockOptimistic

If dbPdvs.RecordCount <> 0 Then
dbFechamentos.CursorLocation = adUseClient
dbFechamentos.Open "select *from fechamentodecaixa where datacaixa=#" & DataInglesa(DataCaixa) & "# and codigoturno=" & CodigoTurno & " and codigopdv=" & dbPdvs!codigopdv, db, adOpenKeyset, adLockOptimistic
Else
  dbPdvs.Close
  db.Close
  Exit Function
End If


If dbFechamentos.RecordCount = 0 Then
  dbTurnos.CursorLocation = adUseClient
  dbTurnos.Open "select *from turnos where codigoturno=" & CodigoTurno, db, adOpenKeyset, adLockOptimistic
  If dbTurnos.RecordCount = 0 Then
    dbTurnos.Close
    db.Close
    Exit Function
  End If
  
  dbFechamentos.AddNew
  dbFechamentos!DataCaixa = DataCaixa
  dbFechamentos!CodigoTurno = dbTurnos!CodigoTurno
  dbFechamentos!HoraIni = dbTurnos!HoraIni
  dbFechamentos!Turno = dbTurnos!Descri
  dbFechamentos!horafim = dbTurnos!horafim
  dbFechamentos!codigopdv = dbPdvs!codigopdv
  dbFechamentos.Update
  
  dbFechamentos.Requery
  
  dbTurnos.Close
End If

dbPdvs.Close

If dbFechamentos.RecordCount = 0 Then
  db.Close
  Exit Function
End If
If dbFechamentos!fechado = True Then
  db.Close
  Exit Function
End If

CodigoFechamento = dbFechamentos!CodigoFechamento

dbEncerrantes.CursorLocation = adUseClient
dbEncerrantes.Open "select *from bicoencerrantes where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

If dbEncerrantes.RecordCount = 0 Then
  dbBicos.CursorLocation = adUseClient
  dbBicos.Open "Select *from bicos order by bico", db, adOpenKeyset, adLockOptimistic
  
  If dbBicos.RecordCount <> 0 Then
    dbBicos.MoveLast
    dbBicos.MoveFirst
    Do While dbBicos.EOF = False
      
      dbEncerrantes.AddNew
      dbEncerrantes!CodigoFechamento = CodigoFechamento
      dbEncerrantes!Bico = dbBicos!Bico
      dbEncerrantes!CodigoProduto = dbBicos!CodigoProduto
      dbEncerrantes!Tanque = dbBicos!Tanque
      dbEncerrantes.Update
      
      dbBicos.MoveNext
    Loop
    
    dbBicos.Close
    
  End If
End If

dbEncerrantes.Close

dbDifComb.CursorLocation = adUseClient
dbDifComb.Open "Select *from diferencacombustivel where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

If dbDifComb.RecordCount = 0 Then
  dbTanques.CursorLocation = adUseClient
  dbTanques.Open "select tanques.*, Produtos.descri from tanques inner join produtos on tanques.codigoproduto=produtos.codigoproduto", db, adOpenKeyset, adLockOptimistic
  If dbTanques.RecordCount <> 0 Then
    dbTanques.MoveLast
    dbTanques.MoveFirst
    Do While dbTanques.EOF = False
      
      dbDifComb.AddNew
      dbDifComb!CodigoFechamento = CodigoFechamento
      dbDifComb!CodigoProduto = dbTanques("codigoproduto")
      dbDifComb!Descri = dbTanques!Descri
      dbDifComb!tanquenr = dbTanques!Tanque
      dbDifComb.Update
      
      dbTanques.MoveNext
    Loop
    dbTanques.Close
    
  End If
  
End If

dbDifComb.Close

AbreCaixa = CodigoFechamento

dbFechamentos.Close
TrataErro:

db.Close

End Function

Public Sub Importar(ByVal CaminhoAdo As String, ByVal DataCaixa As Date, ByVal PDV As String, ByVal Turno As String, ByVal CodigoTurno As Double)
Dim CodigoFechamento As Double
Dim Dia As Date, strEncerrantes As String, intArquivo As Integer
Dim StrTemp As String, SoPrimeira As Boolean
Dim Codigo As String, Descri As String, Tipo As String, Valor As Currency
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency
Dim DescontoPorcento As Currency
Dim Tanque As Integer, Estoque As Double
Dim Bico As Integer, Encerrante As Double, Encontrou As Boolean, Abertura As Double
Dim Preco As Currency, Qtd As Double, Funcionario As Integer
Dim CodigoConta As String, DesteCaixaQtd As Double, DesteCaixaValor As Currency

Dim CodigoCliente As Double, Cupom As String, Placa As String
Dim Km As String, Veiculo As String, ValorTotal As Currency
Dim CodigoProduto As Double, valorUnitario As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim Autorizar As Boolean, Motivo As String, Autorizado As Boolean
Dim DataHoraCaixa As Date, AlteraAnterior As Double, AlteraBico As Double


Dim db As New ADODB.Connection, dbSql As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset
Dim dbVendasLeituraX As New ADODB.Recordset
Dim dbImportacao As New ADODB.Recordset
Dim dbDespesasTipo As New ADODB.Recordset
Dim dbFormaDePg As New ADODB.Recordset
Dim DbClientes As New ADODB.Recordset
Dim dbClientesCarros As New ADODB.Recordset
Dim dbProdutos As New ADODB.Recordset
Dim dbTotalNotas As New ADODB.Recordset
Dim dbTotalCobranca As New ADODB.Recordset
Dim dbClientesProdutos As New ADODB.Recordset

Dim dbFechamentos As New ADODB.Recordset
Dim dbEncerrantes As New ADODB.Recordset
Dim qProdutosAltera As New ADODB.Recordset
Dim dbVendedores As New ADODB.Recordset
Dim dbVendas As New ADODB.Recordset
Dim dbDifComb As New ADODB.Recordset
Dim dbPdvs As New ADODB.Recordset
Dim qPrecoCombustivel As New ADODB.Recordset

CodigoFechamento = AbreCaixa(CaminhoAdo, PDV, DataCaixa, CodigoTurno)

If CodigoFechamento = 0 Then
  Exit Sub
End If

db.Open CaminhoAdo

db.Execute "delete *from importacaoerros where codigofechamento=" & CodigoFechamento

dbFechamentos.CursorLocation = adUseClient
dbFechamentos.Open "select *from fechamentodecaixa where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

dbEncerrantes.CursorLocation = adUseClient
dbEncerrantes.Open "select *from bicoencerrantes where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

dbDespesasTipo.CursorLocation = adUseClient
dbDespesasTipo.Open "select *from despesatipo", db, adOpenForwardOnly, adLockReadOnly

dbFormaDePg.CursorLocation = adUseClient
dbFormaDePg.Open "select *from formadepagamento", db, adOpenForwardOnly, adLockReadOnly

DbClientes.CursorLocation = adUseClient
DbClientes.Open "select *from clientes", db, adOpenKeyset, adLockOptimistic

dbClientesCarros.CursorLocation = adUseClient
dbClientesCarros.Open "select *from clientescarros", db, adOpenForwardOnly, adLockReadOnly

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "select *from produtos", db, adOpenForwardOnly, adLockReadOnly

dbTotalNotas.CursorLocation = adUseClient
dbTotalNotas.Open "select codigocliente, sum(valorprevisto) as total from clientesnota2 where confirmado=0 group by codigocliente", db, adOpenForwardOnly, adLockReadOnly

dbTotalCobranca.CursorLocation = adUseClient
dbTotalCobranca.Open "select codigocliente, sum(valor) as total from clientescobranca where pago=0 group by codigocliente", db, adOpenForwardOnly, adLockReadOnly

dbClientesProdutos.CursorLocation = adUseClient
dbClientesProdutos.Open "select *from clientesprodutos", db, adOpenForwardOnly, adLockReadOnly

dbConfig.CursorLocation = adUseClient
dbConfig.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly

dbVendedores.CursorLocation = adUseClient
dbVendedores.Open "select *from vendedores", db, adOpenForwardOnly, adLockReadOnly

dbVendas.CursorLocation = adUseClient
dbVendas.Open "select *from venda2 where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

dbDifComb.CursorLocation = adUseClient
dbDifComb.Open "select *from diferencacombustivel where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

dbPdvs.CursorLocation = adUseClient
dbPdvs.Open "select *from pdvs", db, adOpenKeyset, adLockOptimistic

DataHoraCaixa = dbFechamentos!DataCaixa & " " & dbFechamentos!HoraIni

With qProdutosAltera
  .CursorLocation = adUseClient
  .Open "select CodigoProdutoAltera, (datacaixa+horaini) as Data from produtosaltera group by CodigoProdutoAltera, (datacaixa+horaini) order by (datacaixa+horaini) desc", db, adOpenKeyset, adLockOptimistic
  If .RecordCount <> 0 Then
    .MoveFirst
    .Find "data<=#" & DataHoraCaixa & "#"
    If .EOF = True Then
      AlteraAnterior = 0
    Else
      AlteraAnterior = qProdutosAltera!codigoprodutoaltera
    End If
  Else
    AlteraAnterior = 0
  End If
  .Close
  .Open "select produtosalteradetalhe.*, produtos.* from produtosalteradetalhe right join produtos on produtosalteradetalhe.codigoproduto=produtos.codigoproduto where codigoprodutoaltera=" & AlteraAnterior & " order by produtos.codigo", db, adOpenKeyset, adLockOptimistic
End With

With qPrecoCombustivel
  .CursorLocation = adUseClient
  .Open "SELECT Alteracoes.CodAlteracao, Alteracoes.DataAlteracao, Turnos.Descri, Turnos.HoraIni FROM Alteracoes LEFT JOIN Turnos ON Alteracoes.codigoTurno = Turnos.CodigoTurno GROUP BY Alteracoes.CodAlteracao, Alteracoes.DataAlteracao, Turnos.Descri, Turnos.HoraIni order by dataalteracao desc, horaini desc", db, adOpenKeyset, adLockOptimistic
  If .RecordCount <> 0 Then
    .MoveFirst
    .Find "dataalteracao<=#" & DataCaixa & "#"
    If .EOF = True Then
      AlteraAnterior = 0
    Else
      Do While DataHoraCaixa <= qPrecoCombustivel!dataalteracao + qPrecoCombustivel!HoraIni
        If qProdutosAltera.EOF = True Then Exit Do
        qPrecoCombustivel.MoveNext
      Loop
      If qPrecoCombustivel.EOF = False Then
        AlteraBico = qPrecoCombustivel!codalteracao
      Else
        AlteraBico = 0
      End If
    End If
  Else
    AlteraBico = 0
  End If
  .Close
  .Open "select * from alterabico where codalteracao=" & AlteraBico & " order by bico", db, adOpenKeyset, adLockOptimistic
End With



dbSql.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig!ftp
dbImportacao.CursorLocation = adUseClient
On Error Resume Next
  
  dbImportacao.Open "select *from caixas where datacaixa='" & DataCaixa & "' and turno='" & Turno & "' and codigoposto='" & dbConfig!porta & "' and planodeconta='" & PDV & "' order by linhaexportada", dbSql, adOpenForwardOnly, adLockReadOnly
  
  If Err.Number <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
  End If
  
  On Error GoTo 0
  
  If dbImportacao.RecordCount = 0 Then
    GoTo Sair
  End If
  dbImportacao.MoveLast
  dbImportacao.MoveFirst
  
  
  
  SoPrimeira = False
  If ApagaRegistros(CaminhoAdo, CodigoFechamento) = False Then
    'MsgBox "Este caixa não pode ser importado a segunda parte porque existe registro já gravado!"
    SoPrimeira = True
  End If
  
  Do While dbImportacao.EOF = False
    StrTemp = dbImportacao!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
      Case "001"
        'Grava os encerrantes
        Bico = CInt(Mid(StrTemp, 5, 6))
        Encerrante = CDbl(Mid(StrTemp, 29, 16))
        Abertura = CDbl(Mid(StrTemp, 12, 16))
        If dbPdvs.RecordCount > 1 Then
          DesteCaixaQtd = CDbl(Mid(StrTemp, 46, 16))
          DesteCaixaValor = CDbl(Mid(StrTemp, 63, 16))
        End If
        If Encerrante > 1000000 Then
          If Abertura > 1000000 Then
            Do While Encerrante > 1000000
              Encerrante = Encerrante - 1000000
            Loop
          End If
        End If
        qPrecoCombustivel.MoveFirst
        qPrecoCombustivel.Find "bico=" & Bico
        With dbEncerrantes
          If .RecordCount <> 0 Then
            .MoveFirst
            .Find "bico=" & Bico
            If .EOF = True Then
              'MsgBox "Bico " & Bico & " cadastrado no posto mas não localizado no sistema."
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,bico) values (" & CodigoFechamento & ",'Bico','Bico não cadastrado'," & Bico & ")"
              Encontrou = False
            Else
              Encontrou = True
              dbEncerrantes!Abertura = Abertura
              dbEncerrantes!Encerrante = Encerrante
              If Len(StrTemp) > 47 Then
                dbEncerrantes!DesteCaixaQtd = DesteCaixaQtd
                dbEncerrantes!DesteCaixaValor = DesteCaixaValor
              Else
                dbEncerrantes!DesteCaixaQtd = Encerrante - Abertura
                dbEncerrantes!DesteCaixaValor = dbEncerrantes!DesteCaixaQtd * qPrecoCombustivel!Preco
              End If
              .Update
              'CalculaBicos ColIndex
            End If
          End If
        End With
      Case "002"
        'Grava Venda
        If Trim(Mid(StrTemp, 18, 6)) <> "" Then
          Bico = CInt(Mid(StrTemp, 18, 6))
        Else
          Bico = 0
        End If
        Preco = CCur(Mid(StrTemp, 38, 12))
        StrTemp2 = Mid(StrTemp, 5, 12)
        If IsNumeric(StrTemp2) = False Then
          StrTemp2 = RemoveString(StrTemp2)
        End If
        Codigo = CDbl(StrTemp2)
        Qtd = CDbl(Mid(StrTemp, 25, 12))
        If Trim(Mid(StrTemp, 64)) <> "" Then
          Funcionario = CInt(Mid(StrTemp, 64))
        Else
          Funcionario = 0
        End If
        
        If Bico = 0 Then
          If qProdutosAltera.RecordCount <> 0 Then
            qProdutosAltera.MoveFirst
            qProdutosAltera.Find "produtos.codigo=" & Codigo
          End If
          If qProdutosAltera.EOF = True Then
            db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto não cadastrado'," & Codigo & ")"
            GoTo naoIncuirProduto
          Else
            If qProdutosAltera("produtosalteradetalhe.PrecoVenda") <> Preco Then
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera("produtosalteradetalhe.PrecoVenda")) & ")"
            End If
          End If
        Else
          If qPrecoCombustivel.RecordCount <> 0 Then
            With qPrecoCombustivel
              .MoveFirst
              .Find "bico=" & Bico
              If .EOF = True Then
                db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera("produtosalteradetalhe.PrecoVenda")) & ")"
              Else
                If qPrecoCombustivel!Preco <> Preco Then
                  db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera("produtosalteradetalhe.PrecoVenda")) & ")"
                End If
              End If
            End With
          End If
        End If
        If Funcionario <> 0 Then
          If dbVendedores.RecordCount <> 0 Then
            dbVendedores.MoveFirst
            dbVendedores.Find "codigo=" & Funcionario
            If dbVendedores.EOF = True Then
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,funcionario,qtd) values (" & dbFechamentos!CodigoFechamento & ",'Funcionario','Funcionário não cadastrado'," & Codigo & "," & Funcionario & "," & NumeroIngles(Qtd) & ")"
              GoTo naoIncuirProduto
            End If
          End If
        Else
          If Bico = 0 Then
            If qProdutosAltera!ComissaoValor <> 0 Or qProdutosAltera!Comissao <> 0 Then
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigofuncionario,qtd) values (" & dbFechamentos!CodigoFechamento & ",'Funcionario','Funcionário não informado'," & Codigo & "," & Funcionario & "," & NumeroIngles(Qtd) & ")"
              GoTo naoIncuirProduto
            End If
          End If
        End If
        If Bico = 0 Then
          If qProdutosAltera.EOF = True Then
            db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto não cadastrado'," & Codigo & ")"
            GoTo naoIncuirProduto
          Else
            If qProdutosAltera("produtosalteradetalhe.PrecoVenda") <> Preco Then
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(Preco) & "," & NumeroIngles(qProdutosAltera("produtosalteradetalhe.PrecoVenda")) & ")"
            End If
            
            With qProdutosAltera
              If qProdutosAltera!Comissao <> 0 Then
                Comissao = (qProdutosAltera("produtosalteradetalhe.PrecoVenda") * Qtd) * (qProdutosAltera!Comissao)
              End If
              If qProdutosAltera!ComissaoValor <> 0 Then
                Comissao = Comissao + (qProdutosAltera!ComissaoValor * Qtd)
              End If
            End With
            
            dbVendas.AddNew
            dbVendas!CodigoFechamento = CodigoFechamento
            dbVendas!Hora = Now
            dbVendas!Data = DataCaixa
            dbVendas!CodigoProduto = qProdutosAltera("produtos.CodigoProduto")
            dbVendas!CodProduto = qProdutosAltera("produtos.Codigo")
            dbVendas!Descri = qProdutosAltera("produtos.Descri")
            dbVendas!Quantidade = Qtd
            dbVendas!valorUnitario = qProdutosAltera("produtosalteradetalhe.PrecoVenda")
            dbVendas!ValorTotal = qProdutosAltera("produtosalteradetalhe.PrecoVenda") * Qtd
            If Funcionario = 0 Then
              dbVendas!Codigovendedor = 0
              dbVendas!CodigoPagamento = 0
            Else
              dbVendas!Codigovendedor = Funcionario
              dbVendas!CodigoPagamento = dbVendedores!Codigovendedor
            End If
            dbVendas!ValorComissao = Comissao
            dbVendas.Update
            
          End If
        End If
      
naoIncuirProduto:
      Case "003"
        'notas de clientes
        If SoPrimeira = False Then
          On Error GoTo 0
          PrecoDif = False
          If Len(StrTemp) > 120 Then
            CodigoCliente = CDbl(Mid(StrTemp, 5, 12))
            Cupom = RemoveString(Trim(Mid(StrTemp, 18, 12)))
            Placa = Mid(StrTemp, 31, 9)
            Km = Mid(StrTemp, 41, 15)
            Veiculo = Mid(StrTemp, 57, 25)
            Qtd = Mid(StrTemp, 83, 15)
            ValorTotalDif = Mid(StrTemp, 99, 15)
            If Len(StrTemp) > 179 Then
              ValorUnitarioDif = CCur(Mid(StrTemp, 179, 15))
              If ValorUnitarioDif = 0 Then
                ValorUnitarioDif = CCur(Format(ValorTotalDif / Qtd, "0.000"))
              End If
            Else
              ValorUnitarioDif = CCur(Format(ValorTotalDif / Qtd, "0.000"))
            End If
            StrTemp2 = Mid(StrTemp, 115, 15)
            If IsNumeric(StrTemp2) = False Then
              StrTemp2 = RemoveString(StrTemp2)
            End If
            CodigoProduto = StrTemp2
            If Len(StrTemp) > 130 Then
              If IsNumeric(Mid(StrTemp, 131, 15)) = True Then
                'LucroDif = Mid(StrTemp, 131, 15)
                If IsNumeric(Mid(StrTemp, 147, 15)) = True Then
                  valorUnitario = Mid(StrTemp, 147, 15)
                End If
                If IsNumeric(Mid(StrTemp, 163, 15)) = False Then
                  ValorTotal = valorUnitario * Qtd
                Else
                  ValorTotal = Mid(StrTemp, 163, 15)
                End If
              Else
                ValorTotal = ValorTotalDif
                valorUnitario = ValorUnitarioDif
                LucroDif = 0
              End If
            Else
              ValorTotal = ValorTotalDif
              valorUnitario = ValorUnitarioDif
              LucroDif = 0
            End If
            Autorizar = False
            Autorizado = False
            Motivo = ""
            LucroDif = ValorTotal - ValorTotalDif
            If IsNumeric(Cupom) = False Then
              Cupom = 0
            End If
            DbClientes.MoveFirst
            DbClientes.Find "codigonoposto=" & CodigoCliente
            If DbClientes.EOF = True Then
              'MsgBox "Código de cliente de nota " & CodigoCliente & " não encontrado!"
              'GravaBloqueado CodigoCliente, "Não encontrado", Cupom, ValorTotal, "Cliente não localizado"
              db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente não cadastrado'," & CodigoCliente & ")"
              GoTo SairDoCliente
            Else
              If dbProdutos.RecordCount <> 0 Then
                If DbClientes!protestado = True Then
                  'MsgBox "Cliente bloqueado!"
                  Autorizar = True
                  Autorizado = True
                  Motivo = "Bloqueado/Protestado"
                End If
                
                dbProdutos.MoveFirst
                dbProdutos.Find "codigo=" & CodigoProduto
                If dbProdutos.EOF = True Then
                  'MsgBox "Código do produto " & CodigoProduto & " não cadastrado!"
                  'GravaBloqueado CodigoCliente, "Código de produto não encontrado", Cupom, ValorTotal, "Cliente não localizado"
                  db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigonoposto) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cupom " & Cupom & " com produto não cadastrado'," & CodigoCliente & "," & CodigoProduto & ")"
                  GoTo Sair
                Else
                  If dbProdutos!Combustivel = True Then
                    dbEncerrantes.MoveFirst
                    dbEncerrantes.Find "codigoproduto=" & dbProdutos!CodigoProduto
                    Preco = PrecoAtual(dbProdutos!CodigoProduto, dbFechamentos!DataCaixa, dbFechamentos!CodigoTurno, CaminhoAdo, dbEncerrantes!Bico)
                  Else
                    Preco = PrecoAtual(dbProdutos!CodigoProduto, dbFechamentos!DataCaixa, dbFechamentos!CodigoTurno, CaminhoAdo)
                  End If
                End If
                If DbClientes!mensalista = False Then
                  If DbClientes!desativado < dbFechamentos!DataCaixa Then
                    'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está desativado! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                    'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                    db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & DbClientes!CodigoCliente & ")"
                    If Configura.NotaBloqueia = 0 Then
                      Autorizar = True
                      Autorizado = False
                      Motivo = "Desativado"
                    End If
                  End If
                End If
                If DbClientes!limitar = True Then
                  If IsNull(DbClientes!Limite) = False Then
                    Limite = CCur(ValorTotal)
                    dbTotalNotas.Requery
                    If dbTotalNotas.RecordCount <> 0 Then
                      dbTotalNotas.MoveFirst
                      dbTotalNotas.Find "codigocliente=" & CodigoCliente
                      If dbTotalNotas.EOF = False Then
                        If IsNull(dbTotalNotas!Total) = False Then
                          Limite = Limite + dbTotalNotas!Total
                        End If
                      End If
                    End If
                    
                    dbTotalCobranca.Requery
                    If dbTotalCobranca.RecordCount <> 0 Then
                      dbTotalCobranca.MoveFirst
                      dbTotalCobranca.Find "codigocliente=" & CodigoCliente
                      If dbTotalCobranca.EOF = False Then
                        If IsNull(dbTotalCobranca!Total) = False Then
                          Limite = Limite + dbTotalCobranca!Total
                        End If
                      End If
                    End If
                    If Limite > DbClientes!Limite Then
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " ultrapassará o limite dele! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
                      'If Resposta = vbNo Then GoTo SairDoCliente
                      Autorizar = True
                      Autorizado = False
                      Motivo = "Ultrapassou Limite"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes!CodigoCliente & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(ValorTotal) & ")"
                    End If
                  Else
                    'MsgBox "O cliente " & DbClientes!Nome & " esta marcado para ser limitado mas não possue valor definido!"
                    'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Marcado para limitar mas não possue valor a ser limitado"
                    Autorizar = True
                    Motivo = "Sem Limite"
                    db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente marcado para limitar mas sem limite cadastrado'," & CodigoCliente & "," & DbClientes!CodigoCliente & ")"
                  End If
                End If
                If DbClientes!diapagamento <> 0 Then
                  If DbClientes!diapagamento >= 28 Then
                    DataPrevista = CDate(Format(UltimoDiaDoMes(Month(dbFechamentos!DataCaixa), Year(dbFechamentos!DataCaixa)), "00") & "/" & Month(dbFechamentos!DataCaixa) & "/" & Year(dbFechamentos!DataCaixa))
                  Else
                    DataPrevista = CDate(Format(DbClientes!diapagamento, "00") & "/" & Month(dbFechamentos!DataCaixa) & "/" & Year(dbFechamentos!DataCaixa))
                  End If
                Else
                  DataPrevista = DateAdd("m", 1, dbFechamentos!DataCaixa)
                End If
                If DataPrevista < dbFechamentos!DataCaixa Then
                  DataPrevista = DateAdd("m", 1, DataPrevista)
                End If
                dbClientesProdutos.Filter = ""
                If dbClientesProdutos.RecordCount <> 0 Then
                  dbClientesProdutos.MoveFirst
                  dbClientesProdutos.Filter = "codigocliente=" & DbClientes!CodigoCliente & " and codproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(txtData.Value) & "#"
                  If dbClientesProdutos.EOF = False Then
                    If dbClientesProdutos!validade = txtData.Value Then
                      If dbClientesProdutos!HoraIni >= dbFechamentos!HoraIni Then
                        PrecoDif = True
                      End If
                    Else
                      PrecoDif = True
                    End If
                  End If
                  If PrecoDif = True Then
                    If dbClientesProdutos!Preco <> 0 Then
                      TempValorPagar = Qtd * dbClientesProdutos!Preco
                    Else
                      TempValorPagar = Qtd * Preco
                      If dbClientesProdutos!Porcento <> 0 Then
                        TempValorPagar = TempValorPagar * dbClientesProdutos!Porcento
                      End If
                    End If
                    If dbClientesProdutos!valorasomar <> 0 Then
                      TempValorPagar = TempValorPagar + (Qtd * dbClientesProdutos!valorasomar)
                    End If
                    TempDif = TempValorPagar - ValorTotal
                    If TempDif > 0.2 Or TempDif < -0.2 Then
                      'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                      'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
                      'If Resposta = vbNo Then GoTo SairDoCliente
                      Autorizar = True
                      Autorizado = False
                      Motivo = "Preço Diferenciado"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal) & "," & NumeroIngles(TempValorPagar) & ")"
                    End If
                  Else
                    'ValorUnitarioDif = Qtd * valorUnitario
                    TempDif = (ValorUnitarioDif * Qtd) - ValorTotal
                    If TempDif > 0.01 Or TempDif < -0.01 Then
                      'MsgBox "Preço unitário incorreto!"
                      db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes!CodigoCliente & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(ValorUnitarioDif * Qtd) & ")"
                      GoTo SairDoCliente
                    End If
                  End If
                Else
                  TempDif = Preco - (ValorTotal / Qtd)
                  If TempDif > 0.2 Or TempDif < -0.02 Then
                    'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                    'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
                    'If Resposta = vbNo Then GoTo SairDoCliente
                    Autorizar = True
                    Autorizado = False
                    Motivo = "Preço incorreto!"
                    db.Execute "insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos!CodigoFechamento & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes!CodigoCliente & "," & NumeroIngles(ValorTotal / Qtd) & "," & NumeroIngles(Preco) & ")"
                  End If
                End If
                A = Fix(valorUnitario)
                If Qtd = 0 Then
                  Qtd = ValorTotal / ValorUnitarioDif
                End If
              End If
            End If
          End If
          
          dbClientesCarros.Filter = "placa='" & Trim(Placa) & "'"
          
          StrTemp = "insert into clientesnota2 (codigofechamento,codigocliente,nome,datalanc,dataprevista,valorprevisto,Data,"
          If Trim(Cupom) <> "" Then
            StrTemp = StrTemp & "Cupom,"
          End If
          StrTemp = StrTemp & "Km,Placa,"
          On Error Resume Next
          If dbClientesCarros.EOF = False And dbClientesCarros.BOF = False Then
            StrTemp = StrTemp & "codigocarro,"
          End If
          On Error GoTo 0
          StrTemp = StrTemp & "Litros,Consumo,CodigoProduto,valorUnitario,Qtd,ValorUnitarioDif,ValorTotalDif,LucroDif,Autorizar,Autorizado,Motivo) values ("
          
          StrTemp = StrTemp & dbFechamentos!CodigoFechamento & "," & DbClientes!CodigoCliente & ",'" & DbClientes!Nome & "',#" & DataInglesa(Date) & " " & Time & "#,#" & DataInglesa(DataPrevista) & "#," & NumeroIngles(ValorTotal) & ",#" & DataInglesa(dbFechamentos!DataCaixa) & "#,"
          If Trim(Cupom) <> "" Then
            StrTemp = StrTemp & Trim(Cupom) & ","
          End If
          If Trim(Km) = "" Then Km = 0
          StrTemp = StrTemp & NumeroIngles(Trim(Km)) & ",'" & Trim(Placa) & "',"
          On Error Resume Next
          If dbClientesCarros.EOF = False And dbClientesCarros.BOF = False Then
            StrTemp = StrTemp & dbClientesCarros!codigocarro & ","
          End If
          On Error GoTo 0
          If Consumo = "" Then
            Consumo = 0
          End If
          StrTemp = StrTemp & NumeroIngles(Qtd) & "," & NumeroIngles(Consumo) & "," & CodigoProduto & "," & NumeroIngles(valorUnitario) & "," & NumeroIngles(Qtd) & "," & NumeroIngles(ValorUnitarioDif) & "," & NumeroIngles(ValorTotalDif) & "," & NumeroIngles(LucroDif) & "," & Autorizar & "," & Autorizado & ",'" & Motivo & "')"
          
          db.Execute StrTemp
        
          If IsNull(DbClientes!UltimoAbastecimento) = True Then
            DbClientes!UltimoAbastecimento = dbFechamentos!DataCaixa
          End If
          If DbClientes!UltimoAbastecimento < dbFechamentos!DataCaixa Then
            DbClientes!UltimoAbastecimento = dbFechamentos!DataCaixa
          End If
          db.Execute "update clientes set TotalNotas=TotalNotas+" & NumeroIngles(ValorTotal) & " where codigocliente=" & CodigoCliente
          db.Execute "update clientes set saldo=limite-totalnotas-totalboleto where codigocliente=" & CodigoCliente
        End If
SairDoCliente:
        
      Case "004"
        'grava estoque dos tanques
        Tanque = Mid(StrTemp, 5, 5)
        StrTemp2 = Mid(StrTemp, 11)
        For I = 1 To Len(StrTemp2)
          If Mid(StrTemp2, I, 1) <> 0 Then
            StrTemp2 = Mid(StrTemp2, I)
            Exit For
          End If
        Next I
        If IsNumeric(StrTemp2) = True Then
          Estoque = CDbl(StrTemp2)
        Else
          Estoque = 0
        End If
        
        With dbDifComb
          If .RecordCount <> 0 Then
            .MoveFirst
            .Find "tanquenr=" & Tanque
            If .EOF = False Then
              dbDifComb!Tanque = Estoque
              .Update
            End If
          End If
        End With
      Case "005"
        'forma de pagamento recebido
        If SoPrimeira = False Then
          If dbFormaDePg.RecordCount <> 0 Then
            Codigo = CDbl(Trim(Mid(StrTemp, 5, 15)))
            Valor = CCur(Mid(StrTemp, 37))
            dbFormaDePg.MoveFirst
            dbFormaDePg.Find "codigonoposto='" & Trim(Codigo) & "'"
            If dbFormaDePg.EOF = False Then
              Tarifa = dbFormaDePg!descontovalor
              Operacao = dbFormaDePg!descontoporoperacao
              Porcento = dbFormaDePg!DescontoPorcento / 100
              
              ValorBruto = Valor
              
              If Porcento <> 0 Then
                DescontoPorcento = ValorBruto * Porcento
              End If
              
              Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao
              
              If dbFormaDePg!CodigoConta = 0 Then
                MsgBox "A forma de pagamento " & dbFormaDePg!Descri & " está sem conta destino!"
              Else
                db.Execute "insert into formadepagamentorecebido2 (codigofechamento,codigoformadepg,descri,valorbruto,valordescoper,valordesctarifa,valordesconto,valor,operacoes,data,hora) values (" & dbFechamentos!CodigoFechamento & "," & dbFormaDePg!CodigoPagamento & ",'" & dbFormaDePg!Descri & "'," & NumeroIngles(ValorBruto) & "," & NumeroIngles(Operacao) & "," & NumeroIngles(Tarifa) & "," & NumeroIngles(DescontoPorcento) & "," & NumeroIngles(Liquido) & "," & TotalOper & ",#" & DataInglesa(dbFechamentos!DataCaixa) & "#,#" & Now & "#)"
              End If
            End If
          End If
        End If
      Case "006"
        'despesas
        If SoPrimeira = False Then
          If dbDespesasTipo.RecordCount <> 0 Then
            Codigo = Trim(Mid(StrTemp, 5, 15))
            Descri = Trim(Mid(StrTemp, 21, 50))
            Tipo = Trim(Mid(StrTemp, 72, 5))
            Valor = CCur(Mid(StrTemp, 78))
            
            If Tipo = "PAG" Then
              Valor = Valor * -1
            End If
            dbDespesasTipo.MoveFirst
            dbDespesasTipo.Find "codigonoposto='" & Codigo & "'"
            If dbDespesasTipo.EOF = False Then
              db.Execute "insert into despesaslanc2 (codigofechamento,origem,data,vencimento,hora,codigoconta,conta,codigodespesa,descri,obs,compensado,valor,valorpago) values (" & dbFechamentos!CodigoFechamento & ",'Fechamento',#" & DataInglesa(dbFechamentos!DataCaixa) & "#,#" & DataInglesa(dbFechamentos!DataCaixa) & "#,#" & Now & "#,-1,'Fechamento de Caixa'," & dbDespesasTipo("codigodespesa") & ",'" & dbDespesasTipo("descri") & "','" & Descri & "',-1," & NumeroIngles(Valor) & "," & NumeroIngles(Valor) & ")"
            End If
          End If
        End If
      Case "007"
        GravaCupons2 StrTemp, CaminhoAdo
      Case "008"
        GravaComissoes StrTemp, CodigoFechamento, CaminhoAdo
      Case "998"
        'GravaResultado StrTemp
        
        '998|     2100000000|1,54
        
        CodigoConta = Trim(Mid(StrTemp, 5, 15))
        Valor = CCur(Trim(Mid(StrTemp, 21)))
        
        db.Execute "insert into fechamentodecaixapista (codigofechamento,codigoconta,valor) values (" & dbFechamentos!CodigoFechamento & "," & CodigoConta & "," & NumeroIngles(Valor) & ")"
        
    End Select
    dbImportacao.MoveNext
  Loop

Sair:

db.Execute "update importacaoerros set dataimportado=#" & DataInglesa(DataCaixa) & " " & Format(Time, "short time") & "# where dataimportado is null"

dbConfig.Close
'dbVendasLeituraX.Close
dbDespesasTipo.Close
dbFormaDePg.Close
'DbClientes.Close
dbClientesCarros.Close
dbProdutos.Close
dbTotalNotas.Close
dbTotalCobranca.Close
dbClientesProdutos.Close
dbImportacao.Close
dbSql.Close
db.Close

End Sub

Public Sub Verifica(ByVal Horas As Integer, ByVal dbImportar As Object)
Dim dbSql As New ADODB.Connection
Dim db As New ADODB.Connection
Dim dbFechamentos As New ADODB.Recordset
Dim dbCaixas As New ADODB.Recordset
Dim dbPdvs As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim CodigoPosto As String

Dim DataCaixa As Date, Turno As String
Dim ProcimaImportacao As Date
Dim CodigoTurno As Double

'On Error GoTo TrataErro

ProcimaImportacao = DateAdd("h", Horas, Now)


With dbImportar
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      If IsNull(.Recordset!ultimaimportacao) = True Then
        .Recordset!ultimaimportacao = CDate("01/01/2009 00:00:00")
        .Recordset.Update
      End If
      If .Recordset!ultimaimportacao <= ProcimaImportacao Then
        If Trim(.Recordset!localdb) <> "" Then
          CaminhoAdo = strMdb & .Recordset!localdb
          db.Open CaminhoAdo

          dbPdvs.CursorLocation = adUseClient
          dbPdvs.Open "select *from pdvs", db, adOpenKeyset, adLockOptimistic
          
          dbTurnos.CursorLocation = adUseClient
          dbTurnos.Open "select *from turnos order by horaini", db, adOpenKeyset, adLockOptimistic

          dbFechamentos.CursorLocation = adUseClient
          dbFechamentos.Open "select *from fechamentodecaixa order by datacaixa, horaini", db, adOpenKeyset, adLockOptimistic
          If dbFechamentos.RecordCount = 0 Then
            If dbTurnos.RecordCount <> 0 Then
              dbTurnos.MoveFirst
              DataCaixa = "01/01/2009"
              Turno = dbTurnos!Descri
              CodigoTurno = dbTurnos!CodigoTurno
            Else
              DataCaixa = "01/01/2009"
              Turno = "01"
              CodigoTurno = 1
            End If
          Else
            dbFechamentos.MoveLast
            DataCaixa = dbFechamentos!DataCaixa
            Turno = dbFechamentos!Turno
            CodigoTurno = dbFechamentos!CodigoTurno
          End If
          dbFechamentos.Close
          dbFechamentos.Open "select *from config", db, adOpenKeyset, adLockOptimistic
          If dbFechamentos.RecordCount <> 0 Then
            strSql = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbFechamentos!ftp
            CodigoPosto = dbFechamentos!porta
          Else
            strSql = ""
          End If
          dbFechamentos.Close
          
          If strSql <> "" Then
            dbSql.Open strSql
            dbCaixas.CursorLocation = adUseClient
            dbCaixas.Open "select datacaixa, turno, planodeconta from caixas where codigoposto='" & CodigoPosto & "' and datacaixa>='" & DataCaixa & "' group by datacaixa, turno, planodeconta order by datacaixa, turno, planodeconta", dbSql, adOpenKeyset, adLockOptimistic
            
            If dbCaixas.RecordCount <> 0 Then
              Do While dbCaixas!Turno <= Turno And dbCaixas!DataCaixa = DataCaixa
                dbCaixas.MoveNext
                If dbCaixas.EOF = True Then Exit Do
              Loop
              Do While dbCaixas.EOF = False
                If dbCaixas!DataCaixa = DataCaixa Then
                  If dbCaixas!Turno > Turno Then
                    dbPdvs.MoveFirst
                    dbPdvs.Find "codigo='" & dbCaixas!planodeconta & "'"
                    dbTurnos.MoveFirst
                    dbTurnos.Find "descri='" & dbCaixas!Turno & "'"
                    If dbPdvs.EOF = False And dbTurnos.EOF = False Then
                      Importar CaminhoAdo, dbCaixas!DataCaixa, dbPdvs!Codigo, dbTurnos!Descri, dbTurnos!CodigoTurno
                    End If
                  End If
                Else
                  dbPdvs.MoveFirst
                  dbPdvs.Find "codigo='" & dbCaixas!planodeconta & "'"
                  dbTurnos.MoveFirst
                  dbTurnos.Find "descri='" & dbCaixas!Turno & "'"
                  If dbPdvs.EOF = False And dbTurnos.EOF = False Then
                    Importar CaminhoAdo, dbCaixas!DataCaixa, dbPdvs!Codigo, dbTurnos!Descri, dbTurnos!CodigoTurno
                  End If
                End If
                dbCaixas.MoveNext
              Loop
              dbPdvs.Close
              db.Close
            End If
            dbCaixas.Close
            dbSql.Close
          End If
          
        End If
      End If
      .Recordset!ultimaimportacao = Now
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With

Exit Sub

TrataErro:

MsgBox "Erro: " & Err.Number & " - " & Err.Description


End Sub


Public Sub GravaCupons2(ByVal StrTemp As String, ByVal CaminhoAdo As String)
Dim dbProdutoGrupoIf As New ADODB.Recordset
Dim db As New ADODB.Connection
Dim dbVendasLeituraX As New ADODB.Recordset

Dim CodigoCliente As Double, DataCupom As Date
Dim HoraCupom As Date, NumeroCupom As String, Placa As String
Dim Km As Double, Carro As String, QtdProduto As Double
Dim ValorTotal As Currency, CodigoGrupo As Double, Tributo As String
Dim strCategoria As String

db.Open CaminhoAdo
dbProdutoGrupoIf.CursorLocation = adUseClient
dbProdutoGrupoIf.Open "select *from produtosgrupoif", db, adOpenKeyset, adLockOptimistic


'007|     27/12/2008|        132,315|         324,04|            102
'On Error GoTo TrataErro
DataCupom = CDate(Trim(Mid(StrTemp, 5, 15)))

dbVendasLeituraX.CursorLocation = adUseClient
dbVendasLeituraX.Open "select *from vendasleiturax where data=#" & DataInglesa(DateAdd("d", -1, DataCupom)) & "#", db, adOpenKeyset, adLockOptimistic

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
  dbVendasLeituraX!leituraxqtd = QtdProduto
  dbVendasLeituraX!leituraxvalor = ValorTotal
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

'On Error GoTo TrataErro
RegistraEstoque = False

Disponivel = 0

db.Open CaminhoAdo

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


Public Sub GravaComissoes(ByVal StrTemp As String, ByVal CodigoFechamento As Double, ByVal CaminhoAdo As String)
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
'On Error GoTo TrataErro

db.Open CaminhoAdo

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
                CodigoFuncionario = dbFuncionarios!Codigovendedor
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

Public Function PrecoAtual(ByVal CodigoProduto As Double, ByVal Dia As Date, ByVal CodigoTurno As Double, ByVal CaminhoAdo As String, Optional Bico As Integer = 0) As Currency
Dim db As New ADODB.Connection
Dim DbPrecos As New ADODB.Recordset
Dim dbTurnos As New ADODB.Recordset
Dim CodigoAlteracao As Double

db.Open CaminhoAdo
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


Public Function ApagaRegistros(ByVal CaminhoAdo As String, ByVal CodigoFechamento As Double, Optional RemovendoCaxa As Boolean = False) As Boolean

Dim SoPrimeira As Boolean
Dim db As New ADODB.Connection
Dim dbClientesNotas As New ADODB.Recordset
Dim dbFormaDePgRecebido As New ADODB.Recordset
Dim dbDespesasLanc As New ADODB.Recordset
Dim DbClientes As New ADODB.Recordset

Dim dbFechamentos As New ADODB.Recordset

db.Open CaminhoAdo

dbFechamentos.CursorLocation = adUseClient
dbFechamentos.Open "select *from fechamentodecaixa where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic

SoPrimeira = False
ApagaRegistros = False



If dbFechamentos!notaconferida = True Then
  SoPrimeira = True
End If

dbClientesNotas.CursorLocation = adUseClient
dbClientesNotas.Open "select *from clientesnota2 where codigofechamento=" & CodigoFechamento, db, adOpenKeyset, adLockOptimistic
dbClientesNotas.Filter = "confirmado=-1"
If dbClientesNotas.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbClientesNotas.Filter = ""

dbFormaDePgRecebido.CursorLocation = adUseClient
dbFormaDePgRecebido.Open "select fechamentodiario from formadepagamentorecebido2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If dbFormaDePgRecebido.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbFormaDePgRecebido.Close

dbDespesasLanc.CursorLocation = adUseClient
dbDespesasLanc.Open "select fechamentodiario from despesaslanc2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, adOpenForwardOnly, adLockReadOnly
If dbDespesasLanc.RecordCount <> 0 Then
  SoPrimeira = True
End If
dbDespesasLanc.Close

If SoPrimeira = False Then
  ApagaRegistros = True
End If


If RemovendoCaxa = True Then
  If SoPrimeira = True Then
    ApagaRegistros = False
    Exit Function
  End If
End If

db.Execute "delete from venda2 where codigofechamento=" & CodigoFechamento
db.Execute "delete from comissoes where codigofechamento=" & CodigoFechamento

If SoPrimeira = False Then
  With dbClientesNotas
    DbClientes.CursorLocation = adUseClient
    DbClientes.Open "select *from clientes", db, adOpenKeyset, adLockOptimistic
    
    If dbClientesNotas.RecordCount <> 0 Then
      Do While dbClientesNotas.EOF = False
        DbClientes.MoveFirst
        DbClientes.Find "codigocliente=" & dbClientesNotas!CodigoCliente
        DbClientes!TotalNotas = DbClientes!TotalNotas - dbClientesNotas!ValorPrevisto
        DbClientes!Saldo = DbClientes!Limite - DbClientes!TotalNotas - DbClientes!TotalBoleto
        DbClientes.Update
        
        dbClientesNotas.MoveNext
      Loop
      db.Execute "delete *from clientesnota2 where codigofechamento=" & CodigoFechamento
    End If
  End With
  dbClientesNotas.Close
  DbClientes.Close
  
  db.Execute "delete *from formadepagamentorecebido2 where codigofechamento=" & CodigoFechamento
  
  db.Execute "delete *from despesaslanc2 where codigofechamento=" & CodigoFechamento
  
End If

db.Execute "delete *from fechamentodecaixapista where codigofechamento=" & CodigoFechamento

db.Close

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


db.Open CaminhoAdo
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


