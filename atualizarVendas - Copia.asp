<%
on error resume next

' Verifica se as connection strings já foram definidas.
' Se não, inclui os arquivos para defini-las.
If Len(StrConn) = 0 Then
    ' ATENÇÃO: A sintaxe de inclusão correta é esta, sem o ' na frente
    ' A linha abaixo inclui o arquivo "conexao.asp"
    <!--#include file="conexao.asp"-->
End If

If Len(StrConnSales) = 0 Then
    ' A linha abaixo inclui o arquivo "conSunSales.asp"
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




'================================================================'
' Fecha ambas as conexões
conn.Close
Set conn = Nothing

connSales.Close
Set connSales = Nothing

'Response.Write " Atulizado!"
%>

