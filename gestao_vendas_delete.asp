<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
    <!--#include file="conSunSales.asp"-->


<%
' Verifica se a sessão do usuário existe
If Session("Usuario") = "" Then
    Response.Write "<script>"
    Response.Write "alert('Sua sessão expirou e é necessário fazer novo login.');"
    Response.Write "window.location.href = 'http://localhost/ImobVendas/gestao_login.asp';"
    Response.Write "</script>"
    Response.End
End If

' Verifica se foi passado o ID da venda
Dim vendaId
vendaId = Request.QueryString("id")
If vendaId = "" Then
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=ID da venda não fornecido."
End If

' Cria a conexão
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

' Processa a exclusão
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim sql
    
    ' Obtém o ID do usuário logado da sessão
    Dim usuarioExclusao
    usuarioExclusao = Session("Usuario")
    
    ' Constrói a consulta SQL para a exclusão lógica
    sql = "UPDATE Vendas SET Excluido = -1, DataExclusao = NOW(), UsuarioExclusao = '" & Replace(usuarioExclusao, "'", "''") & "' WHERE ID = " & vendaId
    
    conn.Execute(sql)

    sql = "UPDATE COMISSOES_A_PAGAR SET Excluido = -1 WHERE ID_Venda = " & vendaId
    conn.Execute(sql)
    
    sql = "UPDATE PAGAMENTOS_COMISSOES SET Excluido = -1 WHERE ID_Venda = " & vendaId
    conn.Execute(sql)


'============================= LOG ============================================'
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    set objMail = server.createobject("CDONTS.NewMail")
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    objMail.Subject = "SV-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    objMail.MailFormat = 0
    objMail.Body = "Exclusão lógica de Venda. ID: " & vendaId
    objMail.Send
    set objMail = Nothing
end if 
'============================= ATUALIZANDO O BANCO DE DADOS ==================='






    ' Redireciona para a página de listagem com mensagem de sucesso
    conn.Close
    Set conn = Nothing
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=Venda excluída com sucesso!"
End If

' Busca os dados da venda para exibir a confirmação
Dim rsVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
rsVenda.Open "SELECT * FROM Vendas WHERE ID = " & vendaId, conn

If rsVenda.EOF Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=Venda não encontrada."
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excluir Venda</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #807777;
            color: #fff;
            padding: 20px;
        }
        .card {
            background-color: #fff;
            color: #000;
            margin-bottom: 20px;
        }
        .card-header {
            background-color: #f8f9fa;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container" style="padding-top: 70px;">
        <div class="card text-center">
            <div class="card-header bg-danger text-white">
                <h3 class="my-2"><i class="fas fa-trash-alt"></i> Confirmar Exclusão</h3>
            </div>
            <div class="card-body">
                <p>Você tem certeza que deseja excluir a venda da unidade <strong><%= rsVenda("Unidade") %></strong> do empreendimento <strong><%= rsVenda("Empreend_ID") %></strong>, realizada em <strong><%= FormatDateTime(rsVenda("DataVenda"), 2) %></strong>?</p>
                <div class="alert alert-warning" role="alert">
                    <strong>Atenção:</strong> Esta ação não pode ser desfeita. A venda será marcada como excluída, mas permanecerá no banco de dados.
                </div>
                <form method="post">
                    <input type="hidden" name="vendaId" value="<%= vendaId %>">
                    <a href="gestao_vendas_list2r.asp" class="btn btn-secondary me-2"><i class="fas fa-times"></i> Cancelar</a>
                    <button type="submit" class="btn btn-danger"><i class="fas fa-trash-alt"></i> Excluir Venda</button>
                </form>
            </div>
        </div>
    </div>
</body>
</html>

<%
' Fecha conexões
If IsObject(rsVenda) Then
    rsVenda.Close
    Set rsVenda = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If
%>