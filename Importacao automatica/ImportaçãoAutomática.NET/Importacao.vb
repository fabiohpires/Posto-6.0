Option Strict Off
Option Explicit On
Module Importacao
	Public strMdb, CaminhoImporta, strSql As String
	Public SoPrimeira As Boolean
	Public Configura As ConfigIni
    Public CaminhoAdo As String

	Public Structure ConfigIni
		Dim NotaNoCaixa As Short
		Dim NotaBloqueia As Short
		Dim ChequesNoCaixa As Short
	End Structure
	
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
		
		db.Open(CaminhoAdo)
		
		dbPdvs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbPdvs.Open("select *from pdvs where codigo='" & PDV & "'", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If dbPdvs.RecordCount <> 0 Then
			dbFechamentos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			dbFechamentos.Open("select *from fechamentodecaixa where datacaixa=#" & DataInglesa(CStr(DataCaixa)) & "# and codigoturno=" & CodigoTurno & " and codigopdv=" & dbPdvs.Fields("codigopdv").Value, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		Else
			dbPdvs.Close()
			db.Close()
			Exit Function
		End If
		
		
		If dbFechamentos.RecordCount = 0 Then
			dbTurnos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			dbTurnos.Open("select *from turnos where codigoturno=" & CodigoTurno, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If dbTurnos.RecordCount = 0 Then
				dbTurnos.Close()
				db.Close()
				Exit Function
			End If
			
			dbFechamentos.AddNew()
			dbFechamentos.Fields("DataCaixa").Value = DataCaixa
			dbFechamentos.Fields("CodigoTurno").Value = dbTurnos.Fields("CodigoTurno").Value
			dbFechamentos.Fields("HoraIni").Value = dbTurnos.Fields("HoraIni").Value
			dbFechamentos.Fields("Turno").Value = dbTurnos.Fields("Descri").Value
			dbFechamentos.Fields("horafim").Value = dbTurnos.Fields("horafim").Value
			dbFechamentos.Fields("codigopdv").Value = dbPdvs.Fields("codigopdv").Value
			dbFechamentos.Update()
			
			dbFechamentos.Requery()
			
			dbTurnos.Close()
		End If
		
		dbPdvs.Close()
		
		If dbFechamentos.RecordCount = 0 Then
			db.Close()
			Exit Function
		End If
		If dbFechamentos.Fields("fechado").Value = True Then
			db.Close()
			Exit Function
		End If
		
		CodigoFechamento = dbFechamentos.Fields("CodigoFechamento").Value
		
		dbEncerrantes.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbEncerrantes.Open("select *from bicoencerrantes where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If dbEncerrantes.RecordCount = 0 Then
			dbBicos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			dbBicos.Open("Select *from bicos order by bico", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			
			If dbBicos.RecordCount <> 0 Then
				dbBicos.MoveLast()
				dbBicos.MoveFirst()
				Do While dbBicos.EOF = False
					
					dbEncerrantes.AddNew()
					dbEncerrantes.Fields("CodigoFechamento").Value = CodigoFechamento
					dbEncerrantes.Fields("Bico").Value = dbBicos.Fields("Bico").Value
					dbEncerrantes.Fields("CodigoProduto").Value = dbBicos.Fields("CodigoProduto").Value
					dbEncerrantes.Fields("Tanque").Value = dbBicos.Fields("Tanque").Value
					dbEncerrantes.Update()
					
					dbBicos.MoveNext()
				Loop 
				
				dbBicos.Close()
				
			End If
		End If
		
		dbEncerrantes.Close()
		
		dbDifComb.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbDifComb.Open("Select *from diferencacombustivel where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If dbDifComb.RecordCount = 0 Then
			dbTanques.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			dbTanques.Open("select tanques.*, Produtos.descri from tanques inner join produtos on tanques.codigoproduto=produtos.codigoproduto", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If dbTanques.RecordCount <> 0 Then
				dbTanques.MoveLast()
				dbTanques.MoveFirst()
				Do While dbTanques.EOF = False
					
					dbDifComb.AddNew()
					dbDifComb.Fields("CodigoFechamento").Value = CodigoFechamento
					dbDifComb.Fields("CodigoProduto").Value = dbTanques.Fields("codigoproduto").Value
					dbDifComb.Fields("Descri").Value = dbTanques.Fields("Descri").Value
					dbDifComb.Fields("tanquenr").Value = dbTanques.Fields("Tanque").Value
					dbDifComb.Update()
					
					dbTanques.MoveNext()
				Loop 
				dbTanques.Close()
				
			End If
			
		End If
		
		dbDifComb.Close()
		
		AbreCaixa = CodigoFechamento
		
		dbFechamentos.Close()
TrataErro: 
		
		db.Close()
		
	End Function
	
	Public Sub Importar(ByVal CaminhoAdo As String, ByVal DataCaixa As Date, ByVal PDV As String, ByVal Turno As String, ByVal CodigoTurno As Double)
        Dim I As Double
        Dim Consumo As Double
        Dim A As Double
        Dim TempDif As Double
        Dim txtData As Date
        Dim DataPrevista As Date
        Dim Limite As Double
        Dim Comissao As Double
        Dim StrTemp2 As String
		Dim CodigoFechamento As Double
		Dim Dia As Date
		Dim strEncerrantes As String
		Dim intArquivo As Short
		Dim StrTemp As String
		Dim SoPrimeira As Boolean
		Dim Descri, Codigo, Tipo As String
		Dim Valor As Decimal
		Dim Tarifa, ValorBruto, Operacao As Decimal
		Dim TotalOper, Porcento As Double
		Dim Liquido As Decimal
		Dim DescontoPorcento As Decimal
		Dim Tanque As Short
		Dim Estoque As Double
		Dim Bico As Short
		Dim Encerrante, Abertura As Double
		Dim Encontrou As Boolean
		Dim Preco As Decimal
		Dim Qtd As Double
		Dim Funcionario As Short
		Dim CodigoConta As String
		Dim DesteCaixaQtd As Double
		Dim DesteCaixaValor As Decimal
		
		Dim CodigoCliente As Double
		Dim Cupom, Placa As String
		Dim Km, Veiculo As String
		Dim ValorTotal As Decimal
		Dim CodigoProduto As Double
		Dim valorUnitario As Decimal
		Dim ValorTotalDif, ValorUnitarioDif, LucroDif As Decimal
		Dim PrecoDif As Boolean
		Dim TempValorPagar As Decimal
		Dim Autorizar, Autorizado As Boolean
		Dim Motivo As String
		Dim DataHoraCaixa As Date
		Dim AlteraAnterior, AlteraBico As Double
		
		
		Dim db As New ADODB.Connection
		Dim dbSql As New ADODB.Connection
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
		
		CodigoFechamento = AbreCaixa(CaminhoAdo, CDbl(PDV), DataCaixa, CodigoTurno)
		
		If CodigoFechamento = 0 Then
			Exit Sub
		End If
		
		db.Open(CaminhoAdo)
		
		db.Execute("delete *from importacaoerros where codigofechamento=" & CodigoFechamento)
		
		dbFechamentos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFechamentos.Open("select *from fechamentodecaixa where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbEncerrantes.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbEncerrantes.Open("select *from bicoencerrantes where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbDespesasTipo.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbDespesasTipo.Open("select *from despesatipo", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbFormaDePg.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFormaDePg.Open("select *from formadepagamento", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		DbClientes.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		DbClientes.Open("select *from clientes", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbClientesCarros.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbClientesCarros.Open("select *from clientescarros", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbProdutos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbProdutos.Open("select *from produtos", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbTotalNotas.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbTotalNotas.Open("select codigocliente, sum(valorprevisto) as total from clientesnota2 where confirmado=0 group by codigocliente", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbTotalCobranca.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbTotalCobranca.Open("select codigocliente, sum(valor) as total from clientescobranca where pago=0 group by codigocliente", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbClientesProdutos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbClientesProdutos.Open("select *from clientesprodutos", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbConfig.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbConfig.Open("select *from config", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbVendedores.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbVendedores.Open("select *from vendedores", db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		dbVendas.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbVendas.Open("select *from venda2 where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbDifComb.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbDifComb.Open("select *from diferencacombustivel where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbPdvs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbPdvs.Open("select *from pdvs", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		DataHoraCaixa = CDate(dbFechamentos.Fields("DataCaixa").Value & " " & dbFechamentos.Fields("HoraIni").Value)
		
        Placa = ""
        Cupom = ""
        Km = ""
        Motivo = ""


        With qProdutosAltera
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .Open("select CodigoProdutoAltera, (datacaixa+horaini) as Data from produtosaltera group by CodigoProdutoAltera, (datacaixa+horaini) order by (datacaixa+horaini) desc", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            If .RecordCount <> 0 Then
                .MoveFirst()
                .Find("data<=#" & DataHoraCaixa & "#")
                If .EOF = True Then
                    AlteraAnterior = 0
                Else
                    AlteraAnterior = qProdutosAltera.Fields("codigoprodutoaltera").Value
                End If
            Else
                AlteraAnterior = 0
            End If
            .Close()
            .Open("select produtosalteradetalhe.*, produtos.* from produtosalteradetalhe right join produtos on produtosalteradetalhe.codigoproduto=produtos.codigoproduto where codigoprodutoaltera=" & AlteraAnterior & " order by produtos.codigo", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        End With
		
		With qPrecoCombustivel
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.Open("SELECT Alteracoes.CodAlteracao, Alteracoes.DataAlteracao, Turnos.Descri, Turnos.HoraIni FROM Alteracoes LEFT JOIN Turnos ON Alteracoes.codigoTurno = Turnos.CodigoTurno GROUP BY Alteracoes.CodAlteracao, Alteracoes.DataAlteracao, Turnos.Descri, Turnos.HoraIni order by dataalteracao desc, horaini desc", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If .RecordCount <> 0 Then
				.MoveFirst()
				.Find("dataalteracao<=#" & DataCaixa & "#")
				If .EOF = True Then
					AlteraAnterior = 0
				Else
                    Do While DataHoraCaixa <= CDate(qPrecoCombustivel.Fields("dataalteracao").Value & " " & qPrecoCombustivel.Fields("HoraIni").Value)
                        If qProdutosAltera.EOF = True Then Exit Do
                        qPrecoCombustivel.MoveNext()
                    Loop
					If qPrecoCombustivel.EOF = False Then
						AlteraBico = qPrecoCombustivel.Fields("codalteracao").Value
					Else
						AlteraBico = 0
					End If
				End If
			Else
				AlteraBico = 0
			End If
			.Close()
			.Open("select * from alterabico where codalteracao=" & AlteraBico & " order by bico", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		End With
		
		
		
		dbSql.Open("Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig.Fields("ftp").Value)
		dbImportacao.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		On Error Resume Next
		
		dbImportacao.Open("select *from caixas where datacaixa='" & DataCaixa & "' and turno='" & Turno & "' and codigoposto='" & dbConfig.Fields("porta").Value & "' and planodeconta='" & PDV & "' order by linhaexportada", dbSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			MsgBox(Err.Number & " - " & Err.Description)
		End If
		
		On Error GoTo 0
		
		If dbImportacao.RecordCount = 0 Then
			GoTo Sair
		End If
		dbImportacao.MoveLast()
		dbImportacao.MoveFirst()
		
		
		
		SoPrimeira = False
		If ApagaRegistros(CaminhoAdo, CodigoFechamento) = False Then
			'MsgBox "Este caixa não pode ser importado a segunda parte porque existe registro já gravado!"
			SoPrimeira = True
		End If
		
		Do While dbImportacao.EOF = False
			StrTemp = dbImportacao.Fields("linhaexportada").Value
			System.Windows.Forms.Application.DoEvents()
			Select Case Mid(StrTemp, 1, 3)
				Case "001"
					'Grava os encerrantes
					Bico = CShort(Mid(StrTemp, 5, 6))
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
                            Do While Abertura > 1000000
                                Abertura -= 1000000
                            Loop
                        End If
                    End If
                    
                    qPrecoCombustivel.MoveFirst()
					qPrecoCombustivel.Find("bico=" & Bico)
					With dbEncerrantes
						If .RecordCount <> 0 Then
							.MoveFirst()
							.Find("bico=" & Bico)
							If .EOF = True Then
								'MsgBox "Bico " & Bico & " cadastrado no posto mas não localizado no sistema."
								db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,bico) values (" & CodigoFechamento & ",'Bico','Bico não cadastrado'," & Bico & ")")
								Encontrou = False
							Else
								Encontrou = True
								dbEncerrantes.Fields("Abertura").Value = Abertura
								dbEncerrantes.Fields("Encerrante").Value = Encerrante
								If Len(StrTemp) > 47 Then
									dbEncerrantes.Fields("DesteCaixaQtd").Value = DesteCaixaQtd
									dbEncerrantes.Fields("DesteCaixaValor").Value = DesteCaixaValor
								Else
									dbEncerrantes.Fields("DesteCaixaQtd").Value = Encerrante - Abertura
									dbEncerrantes.Fields("DesteCaixaValor").Value = dbEncerrantes.Fields("DesteCaixaQtd").Value * qPrecoCombustivel.Fields("Preco").Value
								End If
								.Update()
								'CalculaBicos ColIndex
							End If
						End If
					End With
				Case "002"
					'Grava Venda
					If Trim(Mid(StrTemp, 18, 6)) <> "" Then
						Bico = CShort(Mid(StrTemp, 18, 6))
					Else
						Bico = 0
					End If
					Preco = CDec(Mid(StrTemp, 38, 12))
					'UPGRADE_WARNING: Couldn't resolve default property of object StrTemp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					StrTemp2 = Mid(StrTemp, 5, 12)
					If IsNumeric(StrTemp2) = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object RemoveString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object StrTemp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        StrTemp2 = RemoveString(StrTemp2)
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object StrTemp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Codigo = CStr(CDbl(StrTemp2))
					Qtd = CDbl(Mid(StrTemp, 25, 12))
					If Trim(Mid(StrTemp, 64)) <> "" Then
						Funcionario = CShort(Mid(StrTemp, 64))
					Else
						Funcionario = 0
					End If
					
					If Bico = 0 Then
						If qProdutosAltera.RecordCount <> 0 Then
							qProdutosAltera.MoveFirst()
							qProdutosAltera.Find("produtos.codigo=" & Codigo)
						End If
						If qProdutosAltera.EOF = True Then
							db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto não cadastrado'," & Codigo & ")")
							GoTo naoIncuirProduto
						Else
							If qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value <> Preco Then
								db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(CStr(Preco)) & "," & NumeroIngles(qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value) & ")")
							End If
						End If
					Else
						If qPrecoCombustivel.RecordCount <> 0 Then
							With qPrecoCombustivel
								.MoveFirst()
								.Find("bico=" & Bico)
								If .EOF = True Then
									db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(CStr(Preco)) & "," & NumeroIngles(qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value) & ")")
								Else
									If qPrecoCombustivel.Fields("Preco").Value <> Preco Then
										db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(CStr(Preco)) & "," & NumeroIngles(qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value) & ")")
									End If
								End If
							End With
						End If
					End If
					If Funcionario <> 0 Then
						If dbVendedores.RecordCount <> 0 Then
							dbVendedores.MoveFirst()
							dbVendedores.Find("codigo=" & Funcionario)
							If dbVendedores.EOF = True Then
								db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,funcionario,qtd) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Funcionario','Funcionário não cadastrado'," & Codigo & "," & Funcionario & "," & NumeroIngles(CStr(Qtd)) & ")")
								GoTo naoIncuirProduto
							End If
						End If
					Else
						If Bico = 0 Then
							If qProdutosAltera.Fields("ComissaoValor").Value <> 0 Or qProdutosAltera.Fields("Comissao").Value <> 0 Then
								db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigofuncionario,qtd) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Funcionario','Funcionário não informado'," & Codigo & "," & Funcionario & "," & NumeroIngles(CStr(Qtd)) & ")")
								GoTo naoIncuirProduto
							End If
						End If
					End If
					If Bico = 0 Then
						If qProdutosAltera.EOF = True Then
							db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto não cadastrado'," & Codigo & ")")
							GoTo naoIncuirProduto
						Else
							If qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value <> Preco Then
								db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigonoposto,codigoproduto,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Produto','Produto com preço errado'," & Codigo & "," & Codigo & "," & NumeroIngles(CStr(Preco)) & "," & NumeroIngles(qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value) & ")")
							End If
							
                            With qProdutosAltera
                                If IsDBNull(qProdutosAltera.Fields("Comissao").Value) = False Then
                                    If qProdutosAltera.Fields("Comissao").Value <> 0 Then
                                        Comissao = (qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value * Qtd) * (qProdutosAltera.Fields("Comissao")).Value
                                    End If
                                End If

                                If IsDBNull(qProdutosAltera.Fields("ComissaoValor").Value) = False Then
                                    If qProdutosAltera.Fields("ComissaoValor").Value <> 0 Then
                                        Comissao = Comissao + (qProdutosAltera.Fields("ComissaoValor").Value * Qtd)
                                    End If
                                End If
                            End With
							
							dbVendas.AddNew()
							dbVendas.Fields("CodigoFechamento").Value = CodigoFechamento
							dbVendas.Fields("Hora").Value = Now
							dbVendas.Fields("Data").Value = DataCaixa
							dbVendas.Fields("CodigoProduto").Value = qProdutosAltera.Fields("produtos.CodigoProduto").Value
							dbVendas.Fields("CodProduto").Value = qProdutosAltera.Fields("produtos.Codigo").Value
							dbVendas.Fields("Descri").Value = qProdutosAltera.Fields("produtos.Descri").Value
							dbVendas.Fields("Quantidade").Value = Qtd
							dbVendas.Fields("valorUnitario").Value = qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value
							dbVendas.Fields("ValorTotal").Value = qProdutosAltera.Fields("produtosalteradetalhe.PrecoVenda").Value * Qtd
							If Funcionario = 0 Then
								dbVendas.Fields("Codigovendedor").Value = 0
								dbVendas.Fields("CodigoPagamento").Value = 0
							Else
								dbVendas.Fields("Codigovendedor").Value = Funcionario
								dbVendas.Fields("CodigoPagamento").Value = dbVendedores.Fields("Codigovendedor").Value
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object Comissao. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							dbVendas.Fields("ValorComissao").Value = Comissao
							dbVendas.Update()
							
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
							'UPGRADE_WARNING: Couldn't resolve default property of object RemoveString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Cupom = RemoveString(Trim(Mid(StrTemp, 18, 12)))
							Placa = Mid(StrTemp, 31, 9)
							Km = Mid(StrTemp, 41, 15)
							Veiculo = Mid(StrTemp, 57, 25)
							Qtd = CDbl(Mid(StrTemp, 83, 15))
							ValorTotalDif = CDec(Mid(StrTemp, 99, 15))
							If Len(StrTemp) > 179 Then
								ValorUnitarioDif = CDec(Mid(StrTemp, 179, 15))
								If ValorUnitarioDif = 0 Then
									ValorUnitarioDif = CDec(VB6.Format(ValorTotalDif / Qtd, "0.000"))
								End If
							Else
								ValorUnitarioDif = CDec(VB6.Format(ValorTotalDif / Qtd, "0.000"))
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
										valorUnitario = CDec(Mid(StrTemp, 147, 15))
									End If
									If IsNumeric(Mid(StrTemp, 163, 15)) = False Then
										ValorTotal = valorUnitario * Qtd
									Else
										ValorTotal = CDec(Mid(StrTemp, 163, 15))
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
								Cupom = CStr(0)
                            End If
                            If DbClientes.RecordCount <> 0 Then
                                DbClientes.MoveFirst()
                                DbClientes.Find("codigonoposto=" & CodigoCliente)
                            End If
							If DbClientes.EOF = True Then
                                'MsgBox "Código de cliente de nota " & CodigoCliente & " não encontrado!"
                                'GravaBloqueado CodigoCliente, "Não encontrado", Cupom, ValorTotal, "Cliente não localizado"
                                db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente não cadastrado'," & CodigoCliente & ")")
                                GoTo SairDoCliente
                            Else
                                If dbProdutos.RecordCount <> 0 Then
                                    If DbClientes.Fields("protestado").Value = True Then
                                        'MsgBox "Cliente bloqueado!"
                                        Autorizar = True
                                        Autorizado = True
                                        Motivo = "Bloqueado/Protestado"
                                    End If

                                    dbProdutos.MoveFirst()
                                    dbProdutos.Find("codigo=" & CodigoProduto)
                                    If dbProdutos.EOF = True Then
                                        'MsgBox "Código do produto " & CodigoProduto & " não cadastrado!"
                                        'GravaBloqueado CodigoCliente, "Código de produto não encontrado", Cupom, ValorTotal, "Cliente não localizado"
                                        db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigonoposto) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cupom " & Cupom & " com produto não cadastrado'," & CodigoCliente & "," & CodigoProduto & ")")
                                        GoTo Sair
                                    Else
                                        If dbProdutos.Fields("Combustivel").Value = True Then
                                            dbEncerrantes.MoveFirst()
                                            dbEncerrantes.Find("codigoproduto=" & dbProdutos.Fields("CodigoProduto").Value)
                                            Preco = PrecoAtual(dbProdutos.Fields("CodigoProduto").Value, dbFechamentos.Fields("DataCaixa").Value, dbFechamentos.Fields("CodigoTurno").Value, CaminhoAdo, dbEncerrantes.Fields("Bico").Value)
                                        Else
                                            Preco = PrecoAtual(dbProdutos.Fields("CodigoProduto").Value, dbFechamentos.Fields("DataCaixa").Value, dbFechamentos.Fields("CodigoTurno").Value, CaminhoAdo)
                                        End If
                                    End If
                                    If DbClientes.Fields("mensalista").Value = False Then
                                        If DbClientes.Fields("desativado").Value < dbFechamentos.Fields("DataCaixa").Value Then
                                            'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está desativado! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                                            'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Cliente Desativado"
                                            db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente Bloqueado'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & ")")
                                            If Configura.NotaBloqueia = 0 Then
                                                Autorizar = True
                                                Autorizado = False
                                                Motivo = "Desativado"
                                            End If
                                        End If
                                    End If
                                    If DbClientes.Fields("limitar").Value = True Then
                                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                        If IsDBNull(DbClientes.Fields("Limite").Value) = False Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object Limite. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            Limite = CDec(ValorTotal)
                                            dbTotalNotas.Requery()
                                            If dbTotalNotas.RecordCount <> 0 Then
                                                dbTotalNotas.MoveFirst()
                                                dbTotalNotas.Find("codigocliente=" & CodigoCliente)
                                                If dbTotalNotas.EOF = False Then
                                                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                                    If IsDBNull(dbTotalNotas.Fields("Total").Value) = False Then
                                                        'UPGRADE_WARNING: Couldn't resolve default property of object Limite. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                        Limite = Limite + dbTotalNotas.Fields("Total").Value
                                                    End If
                                                End If
                                            End If

                                            dbTotalCobranca.Requery()
                                            If dbTotalCobranca.RecordCount <> 0 Then
                                                dbTotalCobranca.MoveFirst()
                                                dbTotalCobranca.Find("codigocliente=" & CodigoCliente)
                                                If dbTotalCobranca.EOF = False Then
                                                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                                    If IsDBNull(dbTotalCobranca.Fields("Total").Value) = False Then
                                                        'UPGRADE_WARNING: Couldn't resolve default property of object Limite. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                        Limite = Limite + dbTotalCobranca.Fields("Total").Value
                                                    End If
                                                End If
                                            End If
                                            If Limite > DbClientes.Fields("Limite").Value Then
                                                'Resposta = MsgBox("O cliente " & DbClientes!Nome & " ultrapassará o limite dele! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                                                'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Ultrapassou o limite estipulado"
                                                'If Resposta = vbNo Then GoTo SairDoCliente
                                                Autorizar = True
                                                Autorizado = False
                                                Motivo = "Ultrapassou Limite"
                                                'UPGRADE_WARNING: Couldn't resolve default property of object Limite. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,limitenadata,valorbloqueado) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & "," & NumeroIngles(Limite - ValorTotal) & "," & NumeroIngles(CStr(ValorTotal)) & ")")
                                            End If
                                        Else
                                            'MsgBox "O cliente " & DbClientes!Nome & " esta marcado para ser limitado mas não possue valor definido!"
                                            'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Marcado para limitar mas não possue valor a ser limitado"
                                            Autorizar = True
                                            Motivo = "Sem Limite"
                                            db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente marcado para limitar mas sem limite cadastrado'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & ")")
                                        End If
                                    End If
                                    If DbClientes.Fields("diapagamento").Value <> 0 Then
                                        If DbClientes.Fields("diapagamento").Value >= 28 Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object UltimoDiaDoMes(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            'UPGRADE_WARNING: Couldn't resolve default property of object DataPrevista. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            DataPrevista = UltimoDiaDoMes(dbFechamentos.Fields("DataCaixa").Value)
                                        Else
                                            'UPGRADE_WARNING: Couldn't resolve default property of object DataPrevista. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            DataPrevista = CDate(VB6.Format(DbClientes.Fields("diapagamento").Value, "00") & "/" & Month(dbFechamentos.Fields("DataCaixa").Value) & "/" & Year(dbFechamentos.Fields("DataCaixa").Value))
                                        End If
                                    Else
                                        DataPrevista = DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, dbFechamentos.Fields("DataCaixa").Value)
                                    End If
                                    If DataPrevista < dbFechamentos.Fields("DataCaixa").Value Then
                                        DataPrevista = DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, DataPrevista)
                                    End If
                                    dbClientesProdutos.Filter = ""
                                    If dbClientesProdutos.RecordCount <> 0 Then
                                        dbClientesProdutos.MoveFirst()
                                        dbClientesProdutos.Filter = "codigocliente=" & DbClientes.Fields("CodigoCliente").Value & " and codproduto=" & CodigoProduto & " and validade>=#" & DataInglesa(dbFechamentos.Fields("datacaixa").Value) & "#"
                                        If dbClientesProdutos.EOF = False Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object txtData.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            If dbClientesProdutos.Fields("validade").Value = dbFechamentos.Fields("datacaixa").Value Then
                                                If dbClientesProdutos.Fields("HoraIni").Value >= dbFechamentos.Fields("HoraIni").Value Then
                                                    PrecoDif = True
                                                End If
                                            Else
                                                PrecoDif = True
                                            End If
                                        End If
                                        If PrecoDif = True Then
                                            If dbClientesProdutos.Fields("Preco").Value <> 0 Then
                                                TempValorPagar = Qtd * dbClientesProdutos.Fields("Preco").Value
                                            Else
                                                TempValorPagar = Qtd * Preco
                                                If dbClientesProdutos.Fields("Porcento").Value <> 0 Then
                                                    TempValorPagar = TempValorPagar * dbClientesProdutos.Fields("Porcento").Value
                                                End If
                                            End If
                                            If dbClientesProdutos.Fields("valorasomar").Value <> 0 Then
                                                TempValorPagar = TempValorPagar + (Qtd * dbClientesProdutos.Fields("valorasomar").Value)
                                            End If
                                            'UPGRADE_WARNING: Couldn't resolve default property of object TempDif. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            TempDif = TempValorPagar - ValorTotal
                                            If TempDif > 0.2 Or TempDif < -0.2 Then
                                                'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                                                'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço diferenciado incorreto!"
                                                'If Resposta = vbNo Then GoTo SairDoCliente
                                                Autorizar = True
                                                Autorizado = False
                                                Motivo = "Preço Diferenciado"
                                                db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & "," & NumeroIngles(CStr(ValorTotal)) & "," & NumeroIngles(CStr(TempValorPagar)) & ")")
                                            End If
                                        Else
                                            'ValorUnitarioDif = Qtd * valorUnitario
                                            'UPGRADE_WARNING: Couldn't resolve default property of object TempDif. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            TempDif = (ValorUnitarioDif * Qtd) - ValorTotal
                                            If TempDif > 0.01 Or TempDif < -0.01 Then
                                                'MsgBox "Preço unitário incorreto!"
                                                db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & "," & NumeroIngles(CStr(ValorTotalDif)) & "," & NumeroIngles(CStr(ValorUnitarioDif * Qtd)) & ")")
                                                GoTo SairDoCliente
                                            End If
                                        End If
                                    Else
                                        'UPGRADE_WARNING: Couldn't resolve default property of object TempDif. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        TempDif = Preco - (ValorTotal / Qtd)
                                        If TempDif > 0.2 Or TempDif < -0.02 Then
                                            'Resposta = MsgBox("O cliente " & DbClientes!Nome & " está com o produto diferenciado com valor incorreto! Deseja incluir esta nota?", vbYesNo + vbDefaultButton2)
                                            'GravaBloqueado DbClientes!CodigoCliente, DbClientes!Nome, Cupom, ValorTotal, "Produto " & CodigoProduto & " com preço incorreto!"
                                            'If Resposta = vbNo Then GoTo SairDoCliente
                                            Autorizar = True
                                            Autorizado = False
                                            Motivo = "Preço incorreto!"
                                            db.Execute("insert into importacaoerros (codigofechamento,tipo,Descri,codigoclientenoposto,codigoclientesistema,valorposto,valorsistema) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Cliente','Cliente ultrapassou o limite'," & CodigoCliente & "," & DbClientes.Fields("CodigoCliente").Value & "," & NumeroIngles(CStr(ValorTotal / Qtd)) & "," & NumeroIngles(CStr(Preco)) & ")")
                                        End If
                                    End If
                                    'UPGRADE_WARNING: Couldn't resolve default property of object A. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
						
						'UPGRADE_WARNING: Couldn't resolve default property of object DataPrevista. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						StrTemp = StrTemp & dbFechamentos.Fields("CodigoFechamento").Value & "," & DbClientes.Fields("CodigoCliente").Value & ",'" & DbClientes.Fields("Nome").Value & "',#" & DataInglesa(CStr(Today)) & " " & TimeOfDay & "#,#" & DataInglesa(DataPrevista) & "#," & NumeroIngles(CStr(ValorTotal)) & ",#" & DataInglesa(dbFechamentos.Fields("DataCaixa").Value) & "#,"
						If Trim(Cupom) <> "" Then
							StrTemp = StrTemp & Trim(Cupom) & ","
						End If
						If Trim(Km) = "" Then Km = CStr(0)
						StrTemp = StrTemp & NumeroIngles(Trim(Km)) & ",'" & Trim(Placa) & "',"
						On Error Resume Next
						If dbClientesCarros.EOF = False And dbClientesCarros.BOF = False Then
							StrTemp = StrTemp & dbClientesCarros.Fields("codigocarro").Value & ","
						End If
						On Error GoTo 0
						'UPGRADE_WARNING: Couldn't resolve default property of object Consumo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

                        'UPGRADE_WARNING: Couldn't resolve default property of object Consumo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						StrTemp = StrTemp & NumeroIngles(CStr(Qtd)) & "," & NumeroIngles(Consumo) & "," & CodigoProduto & "," & NumeroIngles(CStr(valorUnitario)) & "," & NumeroIngles(CStr(Qtd)) & "," & NumeroIngles(CStr(ValorUnitarioDif)) & "," & NumeroIngles(CStr(ValorTotalDif)) & "," & NumeroIngles(CStr(LucroDif)) & "," & Autorizar & "," & Autorizado & ",'" & Motivo & "')"
						
						db.Execute(StrTemp)
						
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDbNull(DbClientes.Fields("UltimoAbastecimento").Value) = True Then
							DbClientes.Fields("UltimoAbastecimento").Value = dbFechamentos.Fields("DataCaixa").Value
						End If
						If DbClientes.Fields("UltimoAbastecimento").Value < dbFechamentos.Fields("DataCaixa").Value Then
							DbClientes.Fields("UltimoAbastecimento").Value = dbFechamentos.Fields("DataCaixa").Value
						End If
						db.Execute("update clientes set TotalNotas=TotalNotas+" & NumeroIngles(CStr(ValorTotal)) & " where codigocliente=" & CodigoCliente)
						db.Execute("update clientes set saldo=limite-totalnotas-totalboleto where codigocliente=" & CodigoCliente)
					End If
SairDoCliente: 
					
				Case "004"
					'grava estoque dos tanques
					Tanque = CShort(Mid(StrTemp, 5, 5))
					'UPGRADE_WARNING: Couldn't resolve default property of object StrTemp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    StrTemp2 = Mid(StrTemp, 11)
                    If Trim(StrTemp2) <> "" Then
                        For I = 1 To Len(StrTemp2)
                            If CDbl(Mid(Trim(StrTemp2), I, 1)) <> 0 Then
                                StrTemp2 = Mid(Trim(StrTemp2), I)
                                Exit For
                            End If
                        Next I
                    End If
					
					If IsNumeric(StrTemp2) = True Then
						'UPGRADE_WARNING: Couldn't resolve default property of object StrTemp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Estoque = CDbl(StrTemp2)
					Else
						Estoque = 0
					End If
					
					With dbDifComb
						If .RecordCount <> 0 Then
							.MoveFirst()
							.Find("tanquenr=" & Tanque)
							If .EOF = False Then
								dbDifComb.Fields("Tanque").Value = Estoque
								.Update()
							End If
						End If
					End With
				Case "005"
					'forma de pagamento recebido
					If SoPrimeira = False Then
						If dbFormaDePg.RecordCount <> 0 Then
							Codigo = CStr(CDbl(Trim(Mid(StrTemp, 5, 15))))
							Valor = CDec(Mid(StrTemp, 37))
							dbFormaDePg.MoveFirst()
							dbFormaDePg.Find("codigonoposto='" & Trim(Codigo) & "'")
							If dbFormaDePg.EOF = False Then
								Tarifa = dbFormaDePg.Fields("descontovalor").Value
								Operacao = dbFormaDePg.Fields("descontoporoperacao").Value
								Porcento = dbFormaDePg.Fields("DescontoPorcento").Value / 100
								
								ValorBruto = Valor
								
								If Porcento <> 0 Then
									DescontoPorcento = ValorBruto * Porcento
								End If
								
								Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao
								
								If dbFormaDePg.Fields("CodigoConta").Value = 0 Then
									MsgBox("A forma de pagamento " & dbFormaDePg.Fields("Descri").Value & " está sem conta destino!")
								Else
									db.Execute("insert into formadepagamentorecebido2 (codigofechamento,codigoformadepg,descri,valorbruto,valordescoper,valordesctarifa,valordesconto,valor,operacoes,data,hora) values (" & dbFechamentos.Fields("CodigoFechamento").Value & "," & dbFormaDePg.Fields("CodigoPagamento").Value & ",'" & dbFormaDePg.Fields("Descri").Value & "'," & NumeroIngles(CStr(ValorBruto)) & "," & NumeroIngles(CStr(Operacao)) & "," & NumeroIngles(CStr(Tarifa)) & "," & NumeroIngles(CStr(DescontoPorcento)) & "," & NumeroIngles(CStr(Liquido)) & "," & TotalOper & ",#" & DataInglesa(dbFechamentos.Fields("DataCaixa").Value) & "#,#" & Now & "#)")
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
							Valor = CDec(Mid(StrTemp, 78))
							
							If Tipo = "PAG" Then
								Valor = Valor * -1
							End If
							dbDespesasTipo.MoveFirst()
							dbDespesasTipo.Find("codigonoposto='" & Codigo & "'")
							If dbDespesasTipo.EOF = False Then
								db.Execute("insert into despesaslanc2 (codigofechamento,origem,data,vencimento,hora,codigoconta,conta,codigodespesa,descri,obs,compensado,valor,valorpago) values (" & dbFechamentos.Fields("CodigoFechamento").Value & ",'Fechamento',#" & DataInglesa(dbFechamentos.Fields("DataCaixa").Value) & "#,#" & DataInglesa(dbFechamentos.Fields("DataCaixa").Value) & "#,#" & Now & "#,-1,'Fechamento de Caixa'," & dbDespesasTipo.Fields("codigodespesa").Value & ",'" & dbDespesasTipo.Fields("descri").Value & "','" & Descri & "',-1," & NumeroIngles(CStr(Valor)) & "," & NumeroIngles(CStr(Valor)) & ")")
							End If
						End If
					End If
				Case "007"
					GravaCupons2(StrTemp, CaminhoAdo)
				Case "008"
					GravaComissoes(StrTemp, CodigoFechamento, CaminhoAdo)
				Case "998"
					'GravaResultado StrTemp
					
					'998|     2100000000|1,54
					
					CodigoConta = Trim(Mid(StrTemp, 5, 15))
					Valor = CDec(Trim(Mid(StrTemp, 21)))
					
					db.Execute("insert into fechamentodecaixapista (codigofechamento,codigoconta,valor) values (" & dbFechamentos.Fields("CodigoFechamento").Value & "," & CodigoConta & "," & NumeroIngles(CStr(Valor)) & ")")
					
			End Select
			dbImportacao.MoveNext()
		Loop 
		
Sair: 
		
		db.Execute("update importacaoerros set dataimportado=#" & DataInglesa(CStr(DataCaixa)) & " " & VB6.Format(TimeOfDay, "short time") & "# where dataimportado is null")
		
		dbConfig.Close()
		'dbVendasLeituraX.Close
		dbDespesasTipo.Close()
		dbFormaDePg.Close()
		'DbClientes.Close
		dbClientesCarros.Close()
		dbProdutos.Close()
		dbTotalNotas.Close()
		dbTotalCobranca.Close()
		dbClientesProdutos.Close()
		dbImportacao.Close()
		dbSql.Close()
		db.Close()
		
	End Sub
	
    
	Public Sub GravaCupons2(ByVal StrTemp As String, ByVal CaminhoAdo As String)
		Dim dbProdutoGrupoIf As New ADODB.Recordset
		Dim db As New ADODB.Connection
		Dim dbVendasLeituraX As New ADODB.Recordset
		
		Dim DataCupom As Date
        Dim QtdProduto As Double
		Dim ValorTotal As Decimal
		Dim CodigoGrupo As Double
		Dim strCategoria As String
		
		db.Open(CaminhoAdo)
		dbProdutoGrupoIf.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbProdutoGrupoIf.Open("select *from produtosgrupoif", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		
		'007|     27/12/2008|        132,315|         324,04|            102
		'On Error GoTo TrataErro
		DataCupom = CDate(Trim(Mid(StrTemp, 5, 15)))
		
		dbVendasLeituraX.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbVendasLeituraX.Open("select *from vendasleiturax where data=#" & DataInglesa(CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DataCupom))) & "#", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If Trim(Mid(StrTemp, 21, 15)) <> "" Then
			QtdProduto = CDbl(Trim(Mid(StrTemp, 21, 15)))
		Else
			QtdProduto = 0
		End If
		If Trim(Mid(StrTemp, 37, 15)) <> "" Then
			ValorTotal = CDec(Trim(Mid(StrTemp, 37, 15)))
		Else
			ValorTotal = 0
		End If
		If Trim(Mid(StrTemp, 53)) <> "" Then
			CodigoGrupo = CDbl(Trim(Mid(StrTemp, 53)))
		Else
			CodigoGrupo = 0
        End If
        strCategoria = ""
		If dbProdutoGrupoIf.RecordCount <> 0 Then
			dbProdutoGrupoIf.MoveFirst()
			dbProdutoGrupoIf.Find("codigogrupo=" & CodigoGrupo)
			If dbProdutoGrupoIf.EOF = False Then
				CodigoGrupo = dbProdutoGrupoIf.Fields("Codigo").Value
				strCategoria = dbProdutoGrupoIf.Fields("CodigoGrupo").Value & " " & dbProdutoGrupoIf.Fields("Descri").Value
			End If
		End If
		
		If dbVendasLeituraX.RecordCount = 0 Then
			dbVendasLeituraX.AddNew()
		Else
			dbVendasLeituraX.Filter = "data=#" & DataCupom & "# and categoria='" & strCategoria & "'"
			If dbVendasLeituraX.RecordCount = 0 Then
				dbVendasLeituraX.AddNew()
			End If
		End If
		dbVendasLeituraX.Fields("Data").Value = DataCupom
		dbVendasLeituraX.Fields("leituraxqtd").Value = QtdProduto
		dbVendasLeituraX.Fields("leituraxvalor").Value = ValorTotal
		dbVendasLeituraX.Fields("Categoria").Value = strCategoria
		dbVendasLeituraX.Update()
		
TrataErro: 
		
		dbProdutoGrupoIf.Close()
		db.Close()
		
	End Sub
	
	
	Public Function RegistraEstoque(ByVal DataCaixa As Date, ByVal CodigoTurno As Double, ByVal Turno As String, ByVal HoraIni As Date, ByVal CodigoProduto As Double, Optional ByRef Tanque As Short = 0, Optional ByRef Entrada As Double = 0, Optional ByRef Saida As Double = 0, Optional ByRef Acerto As Double = 0) As Boolean
		Dim db As New ADODB.Connection
		Dim dbEstoque As New ADODB.Recordset
		Dim dbProdutos As New ADODB.Recordset
		Dim Abertura, Disponivel As Double
		
		'On Error GoTo TrataErro
		RegistraEstoque = False
		
		Disponivel = 0
		
		db.Open(CaminhoAdo)
		
		dbEstoque.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbEstoque.Open("Select *from produtosestoque where codigoproduto=" & CodigoProduto & " order by datacaixa, horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		dbProdutos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbProdutos.Open("Select *from produtos", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If dbProdutos.RecordCount = 0 Then
			Exit Function
		End If
		dbProdutos.MoveFirst()
		dbProdutos.Find("codigoproduto=" & CodigoProduto)
		
		If dbEstoque.RecordCount = 0 Then
			dbEstoque.AddNew()
			Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
		Else
			dbEstoque.Filter = "datacaixa=#" & DataInglesa(CStr(DataCaixa)) & "# and codigoturno=" & CodigoTurno
			If dbEstoque.RecordCount = 0 Then
				dbEstoque.AddNew()
				Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
			Else
				dbEstoque.MovePrevious()
				If dbEstoque.BOF = True Then
					Disponivel = EstoqueNoDia(DataCaixa, CodigoTurno, CodigoProduto)
				Else
					Abertura = dbEstoque.Fields("Disponivel").Value
				End If
				dbEstoque.MoveNext()
			End If
		End If
		
		
		dbEstoque.Fields("CodigoProduto").Value = CodigoProduto
		dbEstoque.Fields("Codigo").Value = dbProdutos.Fields("Codigo").Value
		dbEstoque.Fields("Tanque").Value = Tanque
		dbEstoque.Fields("DataCaixa").Value = DataCaixa
		dbEstoque.Fields("CodigoTurno").Value = CodigoTurno
		dbEstoque.Fields("Turno").Value = Turno
		dbEstoque.Fields("HoraIni").Value = HoraIni
		dbEstoque.Fields("Combustivel").Value = dbProdutos.Fields("Combustivel").Value
		dbEstoque.Fields("Abertura").Value = Abertura
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(dbEstoque.Fields("Entrada").Value) = True Then dbEstoque.Fields("Entrada").Value = 0
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(dbEstoque.Fields("Saida").Value) = True Then dbEstoque.Fields("Saida").Value = 0
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(dbEstoque.Fields("Acerto").Value) = True Then dbEstoque.Fields("Acerto").Value = 0
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(dbEstoque.Fields("Diferenca").Value) = True Then dbEstoque.Fields("Diferenca").Value = 0
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(dbEstoque.Fields("Disponivel").Value) = True Then dbEstoque.Fields("Disponivel").Value = 0
		dbEstoque.Fields("Entrada").Value = dbEstoque.Fields("Entrada").Value + Entrada
		dbEstoque.Fields("Saida").Value = dbEstoque.Fields("Saida").Value + Saida
		dbEstoque.Fields("Acerto").Value = dbEstoque.Fields("Acerto").Value + Acerto
		If Disponivel = 0 Then
			dbEstoque.Fields("Disponivel").Value = Abertura + dbEstoque.Fields("Entrada").Value - dbEstoque.Fields("Saida").Value + dbEstoque.Fields("Acerto").Value
			dbEstoque.Fields("Abertura").Value = Abertura
		Else
			dbEstoque.Fields("Abertura").Value = Disponivel - dbEstoque.Fields("Entrada").Value + dbEstoque.Fields("Saida").Value - dbEstoque.Fields("Acerto").Value
			dbEstoque.Fields("Disponivel").Value = Disponivel
		End If
		dbEstoque.Fields("dataalterado").Value = Now
		'UPGRADE_WARNING: Couldn't resolve default property of object Usuarios.Nome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dbEstoque.Fields("Usuario").Value = "Automático"
		dbEstoque.Update()
		
		Abertura = dbEstoque.Fields("Disponivel").Value
		dbEstoque.MoveNext()
		Do While dbEstoque.EOF = False
			dbEstoque.Fields("Abertura").Value = Abertura
			dbEstoque.Fields("Disponivel").Value = Abertura + dbEstoque.Fields("Entrada").Value - dbEstoque.Fields("Saida").Value + dbEstoque.Fields("Acerto").Value
			dbEstoque.Update()
			Abertura = dbEstoque.Fields("Disponivel").Value
			dbEstoque.MoveNext()
		Loop 
		
		RegistraEstoque = True
		Exit Function
		
TrataErro: 
		MsgBox(Err.Number & " - " & Err.Description)
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
		Dim VlUnitario As Decimal
		Dim VlTotal As Decimal
		Dim VlVendaC As Decimal
		Dim VlTotalC As Decimal
		Dim VlComissao As Decimal
		
		Dim CodigoFuncionario As Double
		Dim Nome As String
		
		Dim CodigoProduto As Double
		
		Dim strSql As String
		'On Error GoTo TrataErro
		
		db.Open(CaminhoAdo)
		
		dbFuncionarios.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFuncionarios.Open("select *from vendedores", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		dbProdutos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbProdutos.Open("select *from produtos", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		'008|         000572|               |         000512|                  2,00000|                 21,90000|                 43,80000|                 21,90000|                 43,80000|                  3,06600
		
		Produto = Trim(Mid(StrTemp, 5, 15))
		Bico = Trim(Mid(StrTemp, 21, 15))
		Funcionario = Trim(Mid(StrTemp, 37, 15))
		Qtd = CDbl(Mid(StrTemp, 53, 25))
		VlUnitario = CDec(Mid(StrTemp, 79, 25))
		VlTotal = CDec(Mid(StrTemp, 105, 25))
		VlVendaC = CDec(Mid(StrTemp, 131, 25))
		VlTotalC = CDec(Mid(StrTemp, 157, 25))
		VlComissao = CDec(Mid(StrTemp, 183, 25))
		
		CodigoFuncionario = 0
		Nome = ""
		CodigoProduto = 0
		
		If Trim(Produto) <> "" Then
			If IsNumeric(Produto) = True Then
				If dbProdutos.RecordCount <> 0 Then
					dbProdutos.MoveFirst()
					dbProdutos.Find("codigo=" & Produto)
					If dbProdutos.EOF = False Then
						CodigoProduto = dbProdutos.Fields("CodigoProduto").Value
					End If
				End If
			End If
		End If
		
		If Trim(Funcionario) <> "" Then
			If IsNumeric(Funcionario) = True Then
				If dbFuncionarios.RecordCount <> 0 Then
					dbFuncionarios.MoveFirst()
					dbFuncionarios.Find("codigo=" & Trim(Funcionario))
					If dbFuncionarios.EOF = False Then
						CodigoFuncionario = dbFuncionarios.Fields("Codigovendedor").Value
						Nome = dbFuncionarios.Fields("Nome").Value
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
		strSql = "insert into comissoes (codigofechamento,CodigoProduto,Codigo,bico,CodigoFuncionario,funcionario,Nome,qtd,VlUnitario,VlTotal,VlVendaC,VlTotalC,VlComissao) values (" & CodigoFechamento & "," & CodigoProduto & ",'" & Produto & "'," & Bico & "," & CodigoFuncionario & "," & Funcionario & ",'" & Nome & "'," & NumeroIngles(CStr(Qtd)) & "," & NumeroIngles(CStr(VlUnitario)) & "," & NumeroIngles(CStr(VlTotal)) & "," & NumeroIngles(CStr(VlVendaC)) & "," & NumeroIngles(CStr(VlTotalC)) & "," & NumeroIngles(CStr(VlComissao)) & ")"
		
		db.Execute(strSql)
		
TrataErro: 
		
	End Sub
	
	Public Function PrecoAtual(ByVal CodigoProduto As Double, ByVal Dia As Date, ByVal CodigoTurno As Double, ByVal CaminhoAdo As String, Optional ByRef Bico As Short = 0) As Decimal
		Dim db As New ADODB.Connection
		Dim DbPrecos As New ADODB.Recordset
		Dim dbTurnos As New ADODB.Recordset
		Dim CodigoAlteracao As Double
		
		db.Open(CaminhoAdo)
		dbTurnos.Open("select *from turnos order by horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbTurnos.RecordCount <> 0 Then
			dbTurnos.MoveFirst()
			dbTurnos.Find("codigoturno=" & CodigoTurno)
			If dbTurnos.EOF = True Then
				PrecoAtual = 0
				Exit Function
			End If
		End If
		
		If Bico <> 0 Then
			DbPrecos.Open("select alteracoes.*, turnos.* from alteracoes, turnos where turnos.codigoturno=alteracoes.codigoturno order by dataalteracao, horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If DbPrecos.RecordCount <> 0 Then
				DbPrecos.MoveLast()
				Do While DbPrecos.BOF = False
					If DbPrecos.Fields("dataalteracao").Value <= Dia Then
						If DbPrecos.Fields("dataalteracao").Value < Dia Then
							CodigoAlteracao = DbPrecos.Fields("codalteracao").Value
							Exit Do
						Else
							If DbPrecos.Fields("HoraIni").Value <= dbTurnos.Fields("HoraIni").Value Then
								CodigoAlteracao = DbPrecos.Fields("codalteracao").Value
								Exit Do
							Else
								GoTo Procimo
							End If
						End If
					End If
Procimo: 
					DbPrecos.MovePrevious()
				Loop 
			End If
			If CodigoAlteracao = 0 Then
				DbPrecos.Close()
				DbPrecos.Open("select bicos.precovenda from bicos where bico=" & Bico, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
				If DbPrecos.RecordCount <> 0 Then
					PrecoAtual = DbPrecos.Fields("PrecoVenda").Value
				End If
			Else
				DbPrecos.Close()
				DbPrecos.Open("select preco from alterabico where codalteracao=" & CodigoAlteracao & " and bico=" & Bico, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
				If DbPrecos.RecordCount <> 0 Then
					PrecoAtual = DbPrecos.Fields("Preco").Value
				End If
			End If
		Else
			DbPrecos.Open("select *from produtosaltera order by datacaixa, horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If DbPrecos.RecordCount <> 0 Then
				DbPrecos.MoveLast()
				Do While DbPrecos.BOF = False
					If DbPrecos.Fields("DataCaixa").Value <= Dia Then
						If DbPrecos.Fields("DataCaixa").Value < Dia Then
							CodigoAlteracao = DbPrecos.Fields("codigoprodutoaltera").Value
							Exit Do
						Else
							If DbPrecos.Fields("HoraIni").Value <= dbTurnos.Fields("HoraIni").Value Then
								CodigoAlteracao = DbPrecos.Fields("codigoprodutoaltera").Value
								Exit Do
							Else
								CodigoAlteracao = 0
								Exit Do
							End If
						End If
					End If
					DbPrecos.MovePrevious()
				Loop 
			End If
			If CodigoAlteracao = 0 Then
				DbPrecos.Close()
				DbPrecos.Open("select precovenda from produtos where codigoproduto=" & CodigoProduto, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
				If DbPrecos.RecordCount <> 0 Then
					PrecoAtual = DbPrecos.Fields("PrecoVenda").Value
				End If
			Else
				DbPrecos.Close()
				DbPrecos.Open("select precovenda from produtosalteradetalhe where codigoprodutoaltera=" & CodigoAlteracao & " and codigoproduto=" & CodigoProduto, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
				If DbPrecos.RecordCount <> 0 Then
					PrecoAtual = DbPrecos.Fields("PrecoVenda").Value
				End If
			End If
		End If
		
		DbPrecos.Close()
		dbTurnos.Close()
		db.Close()
	End Function
	
	
	Public Function ApagaRegistros(ByVal CaminhoAdo As String, ByVal CodigoFechamento As Double, Optional ByRef RemovendoCaxa As Boolean = False) As Boolean
		
		Dim SoPrimeira As Boolean
		Dim db As New ADODB.Connection
		Dim dbClientesNotas As New ADODB.Recordset
		Dim dbFormaDePgRecebido As New ADODB.Recordset
		Dim dbDespesasLanc As New ADODB.Recordset
		Dim DbClientes As New ADODB.Recordset
		
		Dim dbFechamentos As New ADODB.Recordset
		
		db.Open(CaminhoAdo)
		
		dbFechamentos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFechamentos.Open("select *from fechamentodecaixa where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		SoPrimeira = False
		ApagaRegistros = False
		
		
		
		If dbFechamentos.Fields("notaconferida").Value = True Then
			SoPrimeira = True
		End If
		
		dbClientesNotas.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbClientesNotas.Open("select *from clientesnota2 where codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		dbClientesNotas.Filter = "confirmado=-1"
		If dbClientesNotas.RecordCount <> 0 Then
			SoPrimeira = True
		End If
		dbClientesNotas.Filter = ""
		
		dbFormaDePgRecebido.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFormaDePgRecebido.Open("select fechamentodiario from formadepagamentorecebido2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If dbFormaDePgRecebido.RecordCount <> 0 Then
			SoPrimeira = True
		End If
		dbFormaDePgRecebido.Close()
		
		dbDespesasLanc.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbDespesasLanc.Open("select fechamentodiario from despesaslanc2 where fechamentodiario=-1 and codigofechamento=" & CodigoFechamento, db, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If dbDespesasLanc.RecordCount <> 0 Then
			SoPrimeira = True
		End If
		dbDespesasLanc.Close()
		
		If SoPrimeira = False Then
			ApagaRegistros = True
		End If
		
		
		If RemovendoCaxa = True Then
			If SoPrimeira = True Then
				ApagaRegistros = False
				Exit Function
			End If
		End If
		
		db.Execute("delete from venda2 where codigofechamento=" & CodigoFechamento)
		db.Execute("delete from comissoes where codigofechamento=" & CodigoFechamento)
		
		If SoPrimeira = False Then
			With dbClientesNotas
				DbClientes.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				DbClientes.Open("select *from clientes", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
				
				If dbClientesNotas.RecordCount <> 0 Then
					Do While dbClientesNotas.EOF = False
						DbClientes.MoveFirst()
                        DbClientes.Find("codigocliente=" & dbClientesNotas.Fields("CodigoCliente").Value)
                        If IsDBNull(dbClientesNotas.Fields("ValorPrevisto").Value) = False And IsDBNull(DbClientes.Fields("TotalNotas").Value) = False Then
                            DbClientes.Fields("TotalNotas").Value = DbClientes.Fields("TotalNotas").Value - dbClientesNotas.Fields("ValorPrevisto").Value
                            DbClientes.Fields("Saldo").Value = DbClientes.Fields("Limite").Value - DbClientes.Fields("TotalNotas").Value - DbClientes.Fields("TotalBoleto").Value
                            DbClientes.Update()
                        End If
                        
                        dbClientesNotas.MoveNext()
                    Loop
					db.Execute("delete *from clientesnota2 where codigofechamento=" & CodigoFechamento)
				End If
			End With
			dbClientesNotas.Close()
			DbClientes.Close()
			
			db.Execute("delete *from formadepagamentorecebido2 where codigofechamento=" & CodigoFechamento)
			
			db.Execute("delete *from despesaslanc2 where codigofechamento=" & CodigoFechamento)
			
		End If
		
		db.Execute("delete *from fechamentodecaixapista where codigofechamento=" & CodigoFechamento)
		
		db.Close()
		
	End Function
	
	
	
	Public Function EstoqueNoDia(ByVal DataCaixa As Date, ByVal CodigoTurno As Double, ByVal CodigoProduto As Double) As Double
        Dim TempData As Date
		Dim StrTemp As String
		Dim Sequencia As Double
		Dim db As New ADODB.Connection
		Dim dbFechamento As New ADODB.Recordset
		Dim dbProdutos As New ADODB.Recordset
		Dim dbVendas As New ADODB.Recordset
		Dim dbEntradas As New ADODB.Recordset
		Dim dbTurnos As New ADODB.Recordset
		Dim SequenciaFinalizado As Double
		Dim Estoque As Double
		
		
		db.Open(CaminhoAdo)
		dbFechamento.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbFechamento.Open("Select fechado, datacaixa, horaini, codigoturno, sequencia from fechamentodecaixa where datacaixa<=#" & DataInglesa(CStr(DataCaixa)) & "# order by datacaixa desc, horaini desc", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		dbEntradas.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbEntradas.Open("select datanota, codigoproduto, quantidade from qprodutosnotas where codigoproduto=" & CodigoProduto & " and datanota>#" & DataInglesa(CStr(DataCaixa)) & "# order by codigoproduto", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		dbTurnos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		dbTurnos.Open("select *from turnos where codigoturno=" & CodigoTurno, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If dbFechamento.RecordCount <> 0 Then
			dbFechamento.Find("fechado=-1")
			If dbFechamento.EOF = False Then
				SequenciaFinalizado = dbFechamento.Fields("Sequencia").Value
			Else
				SequenciaFinalizado = 1
			End If
			dbFechamento.MoveFirst()
			TempData = DataCaixa
			If dbFechamento.Fields("DataCaixa").Value >= DataCaixa Then
				If dbFechamento.Fields("DataCaixa").Value > DataCaixa Then
					dbFechamento.Find("datacaixa=#" & DataInglesa(TempData) & "#")
				End If
				If dbFechamento.Fields("HoraIni").Value <= dbTurnos.Fields("HoraIni").Value Then
					Sequencia = dbFechamento.Fields("Sequencia").Value
				Else
					Sequencia = dbFechamento.Fields("Sequencia").Value
					dbFechamento.Find("horaini<=#" & dbTurnos.Fields("HoraIni").Value & "#")
					If dbFechamento.EOF = False Then
						If dbFechamento.Fields("DataCaixa").Value < DataCaixa Then
							TempData = dbFechamento.Fields("DataCaixa").Value
							dbFechamento.MoveFirst()
							dbFechamento.Find("datacaixa=#" & DataInglesa(TempData) & "#")
							Sequencia = dbFechamento.Fields("Sequencia").Value
						Else
							Sequencia = dbFechamento.Fields("Sequencia").Value
						End If
					End If
				End If
			Else
				Sequencia = dbFechamento.Fields("Sequencia").Value
			End If
		Else
			Sequencia = 0
		End If
		
		
		dbProdutos.Open("Select codigoproduto, estoque from produtos where codigoproduto=" & CodigoProduto & " order by codigoproduto", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbProdutos.RecordCount <> 0 Then
			Estoque = dbProdutos.Fields("Estoque").Value
		End If
		
		'combustiveis
		StrTemp = "select produtos.codigoproduto, produtos.descri, sum(encerrante-abertura) as estoquedia from qbicoencerrantes where produtos.codigoproduto=" & CodigoProduto & " and fechado=0 and sequencia>" & SequenciaFinalizado & " and sequencia<=" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
		dbVendas.Open(StrTemp, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbVendas.RecordCount <> 0 Then
			Do While dbVendas.EOF = False
				Estoque = dbProdutos.Fields("Estoque").Value - dbVendas.Fields("estoquedia").Value
				dbVendas.MoveNext()
			Loop 
		End If
		
		dbVendas.Close()
		StrTemp = "select produtos.codigoproduto, produtos.descri, sum(encerrante-abertura) as estoquedia from qbicoencerrantes where produtos.codigoproduto=" & CodigoProduto & " and fechado=-1 and sequencia>" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
		dbVendas.Open(StrTemp, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbVendas.RecordCount <> 0 Then
			Do While dbVendas.EOF = False
				Estoque = Estoque + dbVendas.Fields("estoquedia").Value
				dbVendas.MoveNext()
			Loop 
		End If
		
		
		
		'não combustiveis
		dbVendas.Close()
		StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where produtos.codigoproduto=" & CodigoProduto & " and fechado=0 and sequencia between " & SequenciaFinalizado & " and " & Sequencia - 1 & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
		dbVendas.Open(StrTemp, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbVendas.RecordCount <> 0 Then
			Do While dbVendas.EOF = False
				Estoque = dbProdutos.Fields("Estoque").Value - dbVendas.Fields("estoquedia").Value
				dbVendas.MoveNext()
			Loop 
		End If
		
		dbVendas.Close()
		StrTemp = "select produtos.codigoproduto, produtos.descri, sum(quantidade) as estoquedia from qprodutosVendaCaixa where produtos.codigoproduto=" & CodigoProduto & " and fechado=-1 and sequencia>" & Sequencia & " group by produtos.codigoproduto, produtos.descri order by produtos.codigoproduto"
		dbVendas.Open(StrTemp, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		If dbVendas.RecordCount <> 0 Then
			Do While dbVendas.EOF = False
				Estoque = Estoque + dbVendas.Fields("estoquedia").Value
				dbVendas.MoveNext()
			Loop 
		End If
		
		
		If dbEntradas.RecordCount <> 0 Then
			Do While dbEntradas.EOF = False
				Estoque = Estoque - dbEntradas.Fields("Quantidade").Value
				dbEntradas.MoveNext()
			Loop 
		End If
		
		dbFechamento.Close()
		dbProdutos.Close()
		dbVendas.Close()
		dbEntradas.Close()
		dbTurnos.Close()
		
		db.Close()
		
		EstoqueNoDia = Estoque
		Exit Function
		
TrataErro: 
		MsgBox(Err.Number & " - " & Err.Description)
		EstoqueNoDia = 0
	End Function
End Module