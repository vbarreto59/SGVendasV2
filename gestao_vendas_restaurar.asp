<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<%
' =======================================================
' === PÁGINA PARA RESTAURAR UMA VENDA EXCLUÍDA ===
' O campo 'Excluido' será alterado de -1 para 0 (ou False)
' =======================================================

' Variáveis
Dim conn, sql, vendaID, mensagem

' 1. Obter o ID da venda da QueryString
vendaID = Request.QueryString("id")

' 2. Validação do ID
If IsEmpty(vendaID) Or Not IsNumeric(vendaID) Then
    mensagem = "ERRO: ID de venda inválido."
    Response.Redirect "gestao_vendas_list2x.asp?mensagem=" & Server.URLEncode(mensagem)
    Response.End
End If

' 3. Conexão com o Banco de Dados
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open StrConnSales
If Err.Number <> 0 Then
    Response.Write "<h2>ERRO DE CONEXÃO COM O BANCO DE DADOS!</h2>"
    Response.Write "<p>Detalhes: " & Err.Description & "</p>"
    Response.End
End If
On Error Goto 0

' 4. Comando SQL para Restaurar (Atualizar o campo Excluido para 0)
' **ATENÇÃO:** O valor 0 (zero) é comumente usado para False em MS Access e SQL Server com Bit,
' mas se o seu banco usar outro valor para "Não Excluído" (Ex: True/False, 1/0), ajuste-o aqui.
sql = "UPDATE Vendas SET Excluido = 0, DataExclusao = NULL, UsuarioExclusao = NULL WHERE ID = " & vendaID

' 5. Executar o comando SQL
On Error Resume Next
conn.Execute sql

If Err.Number <> 0 Then
    ' Em caso de erro na execução
    mensagem = "ERRO: Falha ao restaurar a venda ID " & vendaID & ". Detalhes: " & Err.Description
Else
    ' Sucesso na restauração
    mensagem = "✅ Venda ID " & vendaID & " restaurada com sucesso!"
End If
On Error Goto 0

' 6. Fechar a conexão
If conn.State = 1 Then conn.Close
Set conn = Nothing

' 7. Redirecionar de volta para a lista de vendas excluídas com a mensagem
Response.Redirect "gestao_vendas_list2x.asp?mensagem=" & Server.URLEncode(mensagem)
%>