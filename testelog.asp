<!--#include file="conSunSales.asp"-->
<!--#include file="registra_log.asp"-->

<%

' No login
Call InserirLog ("Sistema", "LOGIN", "Usuário autenticado com sucesso")

' Ao inserir dados
Call InserirLog ("tbl_clientes", "INSERT", "Cliente cadastrado")

' Ao atualizar dados
Call InserirLog ("tbl_pedidos", "UPDATE", "Pedido # atualizado")

' Ao excluir dados
Call InserirLog ("tbl_itens", "DELETE", "Item ID excluído")

' Em caso de erro
Call InserirLog ("Sistema", "ERRO", "Erro ao processar pedido: ")

Response.write "Registros inseridos!"
%>