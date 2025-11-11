<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
' Configuração
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

On Error Resume Next

Set conn = Server.CreateObject("ADODB.Connection")
' Recordset (rs) não é mais necessário para este tipo de operação
Set cmd = Server.CreateObject("ADODB.Command")

' Abrir conexão
conn.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "<p style='color: red;'>Erro ao conectar ao banco de dados: " & Err.Description & "</p>"
    Response.End
End If

' Variáveis para armazenar o número de registros afetados
Dim sql
Dim affectedRows ' Total de registros excluídos de Vendas
Dim affectedComissoes ' Total de registros excluídos de COMISSOES_A_PAGAR
Dim affectedPagamentos ' Total de registros excluídos de PAGAMENTOS_COMISSOES

' Iniciar transação para garantir integridade dos dados
conn.BeginTrans

' --- 1. Excluir TODOS os registros de Vendas ---
sql = "DELETE FROM Vendas"

Set cmd.ActiveConnection = conn
cmd.CommandText = sql
' O método Execute retorna o número de linhas afetadas no parâmetro
cmd.Execute affectedRows 

If Err.Number <> 0 Then
    conn.RollbackTrans
    Response.Write "<p style='color: red;'>Erro ao excluir registros de Vendas: " & Err.Description & "</p>"
    Response.End
End If

' --- 2. Excluir TODOS os registros de COMISSOES_A_PAGAR ---
sql = "DELETE FROM COMISSOES_A_PAGAR"

cmd.CommandText = sql
cmd.Execute affectedComissoes

If Err.Number <> 0 Then
    conn.RollbackTrans
    Response.Write "<p style='color: red;'>Erro ao excluir registros de COMISSOES_A_PAGAR: " & Err.Description & "</p>"
    Response.End
End If


' --- 3. Excluir TODOS os registros de PAGAMENTOS_COMISSOES ---
sql = "DELETE FROM PAGAMENTOS_COMISSOES"

cmd.CommandText = sql
cmd.Execute affectedPagamentos

If Err.Number <> 0 Then
    conn.RollbackTrans
    Response.Write "<p style='color: red;'>Erro ao excluir registros de PAGAMENTOS_COMISSOES: " & Err.Description & "</p>"
    Response.End
End If



' --- 4. Excluir VENDAS_TEMP ---
sql = "DELETE FROM VENDA_TEMP"

cmd.CommandText = sql
cmd.Execute affectedPagamentos

If Err.Number <> 0 Then
    conn.RollbackTrans
    Response.Write "<p style='color: red;'>Erro ao excluir registros de VENDA_TEMP: " & Err.Description & "</p>"
    Response.End
End If

' Commit da transação se todas as exclusões foram bem-sucedidas
conn.CommitTrans

' Resumo da execução
Response.Write "<h2>Resumo da Execução (Exclusão Total)</h2>"
Response.Write "<div style='padding: 15px; background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; border-radius: 5px; margin-bottom: 20px;'>"
Response.Write "<h3>⚠️ Exclusão Concluída!</h3>"
Response.Write "<p style='font-weight: bold;'>TODOS os registros das três tabelas foram excluídos com sucesso.</p>"
Response.Write "</div>"
Response.Write "<p><strong>Registros excluídos de Vendas:</strong> " & affectedRows & "</p>"
Response.Write "<p><strong>Registros excluídos de COMISSOES_A_PAGAR:</strong> " & affectedComissoes & "</p>"
Response.Write "<p><strong>Registros excluídos de PAGAMENTOS_COMISSOES:</strong> " & affectedPagamentos & "</p>"


' Fechar conexões
If conn.State = 1 Then conn.Close
Set cmd = Nothing
Set conn = Nothing

If Err.Number <> 0 Then
    Response.Write "<p style='color: red;'>Ocorreu um erro geral: " & Err.Description & "</p>"
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Limpeza Total de Dados</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f8f9fa; }
        h1 { color: #dc3545; }
        h2 { color: #dc3545; margin-top: 20px; border-bottom: 2px solid #dc3545; padding-bottom: 5px; }
        p { font-size: 1em; }
        strong { color: #343a40; }
        form { margin: 15px 0; }
        input[type="submit"] { 
            background: #dc3545; 
            color: white; 
            padding: 10px 15px; 
            border: none; 
            cursor: pointer; 
            font-weight: bold;
            border-radius: 5px;
            transition: background 0.3s;
        }
        input[type="submit"]:hover { background: #c82333; }
    </style>
</head>
<body>
    <h1>Rotina de Limpeza - EXCLUSÃO TOTAL DE DADOS</h1>
    <p><strong>ATENÇÃO:</strong> Este script executa a exclusão **incondicional** de **TODOS** os dados nas tabelas Vendas, COMISSOES_A_PAGAR e PAGAMENTOS_COMISSOES.</p>
    <p><strong>Data da execução:</strong> <%=Now()%></p>
    
    <%
    ' Botão para executar novamente
    Response.Write "<form method='post'>"
    Response.Write "<input type='submit' value='Executar Limpeza Total Novamente'>"
    Response.Write "</form>"
    %>
</body>
</html>