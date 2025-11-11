<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% 
If Len(StrConnSales) = 0 Then 
%>
    <!--#include file="conSunSales.asp"-->
<%
End If
%>

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Obter dados do POST
Dim idPagamento, idVenda
idPagamento = Request.Form("id_pagamento")
idVenda = Request.Form("id_venda")

If idPagamento = "" Or Not IsNumeric(idPagamento) Or idVenda = "" Or Not IsNumeric(idVenda) Then
    Response.Write "ERRO: Parâmetros inválidos."
    Response.End
End If

' Conexão com o banco
Dim connSales, sqlExcluir
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

On Error Resume Next

' Excluir o pagamento (Exclusão física usando DELETE)
sqlExcluir = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Pagamento = " & idPagamento
connSales.Execute sqlExcluir

If Err.Number = 0 Then
    ' SUCESSO - Retorna texto para o AJAX
    Response.Write "SUCESSO"
Else
    ' ERRO - Retorna mensagem de erro
    Response.Write "ERRO: " & Err.Description
End If

On Error GoTo 0

' Fechar conexão
If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If

Response.End
%>