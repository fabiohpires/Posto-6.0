Attribute VB_Name = "ModPrincipal"

Public Sub CriaMDB()
Dim Db As New ADODB.Connection
Dim Catalogo As New ADOX.Catalog ', Tabela As New ADOX.Table
Catalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Caixas.mdb"
'With Tabela
'    .Name = "Caixa"
'    .Columns.Append "CodigoPosto", adVarWChar, 10
'    .Columns.Append "NomePosto", adVarWChar, 50
'    .Columns.Append "DataCaixa", adDate
'    .Columns.Append "Turno", adVarWChar, 2
'End With
'Catalogo.Tables.Append Tabela
Set Catalogo = Nothing
'Set Tabela = Nothing

Db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Caixas.mdb"
Db.Execute "create table Caixas(CodigoCaixa counter, CodigoPosto text(20), NomePosto Text(50), Datacaixa datetime, Turno Text(15))"
Db.Execute "create table Cupons(CodigoCupom counter, CodigoCaixa double, DataCupom datetime, NumeroCupom text(9), ValorTotal currency, ValorDesconto Currency, ValorTroco currency, Cancelado bit)"
Db.Execute "Create Table CuponsDetalhe(CodigoCupomDetalhe counter, CodigoCaixa double, CodigoCupom double, CodigoProduto Text(20), Produto Text(50), Quantidade double, ValorUnitario Currency, ValorBruto Currency, ValorDescontoItem currency, ValorDescontoRateadoItem Currency, ValorTotalItem Currency, Cancelado bit)"
Db.Execute "Create Table CuponsFormas(CodigoCupomForma counter, CodigoCaixa double, CodigoCupom double, CodigoFormaPagamento text(20), Descri Text(50), CodigoAutorizadora Text(20), DescriAutorizadora Text(50), Parcelas integer, NSU text(12), CodigoModalidade Text(20), DescriModalidade Text(50))"

Db.Close
End Sub
