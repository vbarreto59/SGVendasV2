<%
on error resume next

If Len(StrConn) = 0 Then
    <!--#include file="conexao.asp"-->
End If

If Len(StrConnSales) = 0 Then
    <!--#include file="conSunSales.asp"-->
End If




'======== Executar querys dentro do banco - 14 08 2025 =========='
Set rsAjustaData = Server.CreateObject ("ADODB.Command")
rsAjustaData.ActiveConnection = StrConnSales
rsAjustaData.CommandType = 4 'adCmdStoredProc

'Limpa a tabela de resumo ComissaoSaldo
rsAjustaData.CommandText = "qryDelComissao"
rsAjustaData.Execute()

'------------ corretores '
rsAjustaData.CommandText = "qryAddComisVendaCorretor"
rsAjustaData.Execute()

'------------ diretores '
'incluir as vendas com as respectivas comissões'
rsAjustaData.CommandText = "qryAddComisVendaDiretor"
rsAjustaData.Execute()

'------------ Gerentes '
'incluir as vendas com as respectivas comissões'
rsAjustaData.CommandText = "qryAddComisVendaGerente"
rsAjustaData.Execute()



' adiciona todas as comissoes pagas '
rsAjustaData.CommandText = "qryAddComisPaga"
rsAjustaData.Execute()

'Atualizar Nomes de Comissao a pagar'
' adiciona todas as comissoes pagas '
rsAjustaData.CommandText = "qryAtuNomes"
rsAjustaData.Execute()

'=== Mudanças em 08 11 2025 Resumo de vendas e comissões ===='

'limpar e popular a tabela VENDA_TEMP'
rsAjustaData.CommandText = "qryDelVTemp"
'rsAjustaData.Execute()


rsAjustaData.CommandText = "qryAddDiretoriaVTemp"
'rsAjustaData.Execute()

rsAjustaData.CommandText = "qryAddGerenciaVTemp"
'rsAjustaData.Execute()

rsAjustaData.CommandText = "qryAddCorretorVTemp"
'rsAjustaData.Execute()

rsAjustaData.CommandText = "qryAtuComissAPagarVendasID"
rsAjustaData.Execute()


' Atualiza o Empreend_ID na tabela COMISSOES_A_PAGAR com base em Vendas 11 11 25
rsAjustaData.CommandType = 1 'adCmdText (Para executar SQL direto)
rsAjustaData.CommandText = "UPDATE Vendas INNER JOIN COMISSOES_A_PAGAR ON Vendas.Id = COMISSOES_A_PAGAR.ID_Venda SET COMISSOES_A_PAGAR.Empreend_ID = [Vendas].[Empreend_Id];"
rsAjustaData.Execute()
rsAjustaData.CommandType = 4 'Retorna ao padrão adCmdStoredProc (Embora seja o último comando de execução)

'=============== 22 11 2025 Atualizar Goordenadas Geo-Mapa
rsAjustaData.CommandType = 1 'adCmdText (Para executar SQL direto)
rsAjustaData.CommandText = "UPDATE GeoMapa INNER JOIN Vendas ON GeoMapa.Localidade = Vendas.Localidade SET Vendas.Localizacao = [GeoMapa].[Localizacao];"
rsAjustaData.Execute()
rsAjustaData.CommandType = 4 'Retorna ao padrão adCmdStoredProc (Embora seja o último comando de execução)

'================================================================'
' Fecha ambas as conexões
conn.Close
Set conn = Nothing

connSales.Close
Set connSales = Nothing

'Response.Write " Atulizado!"
%>