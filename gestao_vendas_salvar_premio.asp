<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
' ====================================================================
' Script para Salvar um Novo Pagamento de Prêmio - Otimizado
' ====================================================================
Response.Buffer = True
Response.Expires = -1
On Error GoTo 0 ' Habilita tratamento de erro explícito

' O caminho do banco de dados principal é necessário para os UPDATES
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)

' O erro pode ocorrer se a string não contiver ';', então garantimos que seja apenas o caminho
If InStr(dbSunnyPath, ";") > 0 Then
    dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
End If

' Funções de formatação de número (da nossa conversa anterior)
Function FormatNumberForSQL(sValue)
    sValue = Replace(sValue, ".", "")
    sValue = Replace(sValue, ",", ".")
    FormatNumberForSQL = sValue
End Function

' Obter dados do formulário
Dim idComissao, idVenda, userId, valorPago, dataPagamento, statusPagamento, obs, recipientType
idComissao = Request.Form("ID_Comissao")
idVenda = Request.Form("ID_Venda")
userId = Request.Form("UserId")
valorPago = Request.Form("ValorPago")
dataPagamento = Request.Form("DataPagamento")
statusPagamento = Request.Form("Status")
obs = Request.Form("Obs")
recipientType = Request.Form("RecipientType")

' Limpar e formatar o valor monetário para o banco de dados
valorPago = FormatNumberForSQL(valorPago)

' Validação básica
If Not IsNumeric(idComissao) Or idComissao = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: ID da comissão inválido."
End If
If Not IsNumeric(idVenda) Or idVenda = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: ID da venda inválido."
End If
If Not IsNumeric(userId) Or userId = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: ID do usuário inválido."
End If
If Not IsNumeric(valorPago) Or CDbl(valorPago) <= 0 Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: Valor do prêmio inválido."
End If
If dataPagamento = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: Data do pagamento não informada."
End If
If statusPagamento = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: Status do pagamento não informado."
End If
If recipientType = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro: Tipo de recebedor não informado."
End If

' ----------------------------------------------------------------------
' INSERIR PAGAMENTO NA TABELA PAGAMENTOS_COMISSOES (USANDO connSales)
' ----------------------------------------------------------------------
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Localizar nome do usuário
sql = "SELECT * FROM Usuarios WHERE UserId=" & userId
Set rs = conn.Execute(sql)
If Not rs.EOF Then
    vNomeUsuario = rs("Nome")
Else
    vNomeUsuario = "Usuário não encontrado"
End If
rs.Close

' Inserir pagamento de prêmio
sqlInsert = "INSERT INTO PAGAMENTOS_COMISSOES (ID_Venda, UsuariosUserId, UsuariosNome, DataPagamento, ValorPago, Status, Obs, TipoRecebedor, TipoPagamento) VALUES (" & _
            CInt(idVenda) & ", " & _
            CInt(userId) & ", '" & Replace(vNomeUsuario, "'", "''") & "', '" & dataPagamento & "', " & _
            (valorPago) & ", '" & Replace(statusPagamento, "'", "''") & "', '" & Replace(obs, "'", "''") & "', '" & Replace(recipientType, "'", "''") & "', 'Premiação')"            

connSales.Execute sqlInsert

If Err.Number <> 0 Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao salvar pagamento do prêmio: " & Err.Description
End If

' ----------------------------------------------------------------------
' ATUALIZAÇÕES COMPLEMENTARES COM CROSS-DATABASE (USANDO connSales)
' ----------------------------------------------------------------------
Dim sqlUpdate
Dim adodb_path
adodb_path = "[;DATABASE=" & dbSunnyPath & "]"

' UPDATE para Diretorias
sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Diretorias ON PAGAMENTOS_COMISSOES.UserId = Diretorias.DiretoriaId SET PAGAMENTOS_COMISSOES.UsuariosUserId = [Diretorias].[UserId], PAGAMENTOS_COMISSOES.UsuariosNome = [Diretorias].[Nome] WHERE (((PAGAMENTOS_COMISSOES.TipoRecebedor)='diretoria') AND ((PAGAMENTOS_COMISSOES.TipoPagamento)='Premiação'));"
connSales.Execute sqlUpdate

If Err.Number <> 0 Then
    ' Não redireciona por erro, apenas registra no log se necessário
    ' Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao atualizar dados de Diretoria: " & Err.Description
End If

' UPDATE para Gerencias
sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Gerencias ON PAGAMENTOS_COMISSOES.UserId = Gerencias.GerenciaId SET PAGAMENTOS_COMISSOES.UsuariosUserId = [Gerencias].[UserId], PAGAMENTOS_COMISSOES.UsuariosNome = [Gerencias].[Nome] WHERE (((PAGAMENTOS_COMISSOES.TipoRecebedor)='gerencia') AND ((PAGAMENTOS_COMISSOES.TipoPagamento)='Premiação'));"
connSales.Execute sqlUpdate

If Err.Number <> 0 Then
    ' Não redireciona por erro, apenas registra no log se necessário
    ' Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao atualizar dados de Gerencia: " & Err.Description
End If

' UPDATE para Usuarios (Corretor)
sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Usuarios ON PAGAMENTOS_COMISSOES.UserId = Usuarios.UserId SET PAGAMENTOS_COMISSOES.UsuariosUserId = [Usuarios].[UserId], PAGAMENTOS_COMISSOES.UsuariosNome = [Usuarios].[Nome] WHERE (((PAGAMENTOS_COMISSOES.TipoRecebedor)='corretor') AND ((PAGAMENTOS_COMISSOES.TipoPagamento)='Premiação'));"
connSales.Execute sqlUpdate

If Err.Number <> 0 Then
    ' Não redireciona por erro, apenas registra no log se necessário
    ' Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao atualizar dados de Corretor: " & Err.Description
End If

' ----------------------------------------------------------------------
' REDIRECIONAMENTO FINAL E LIMPEZA
' ----------------------------------------------------------------------
Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Pagamento do prêmio salvo com sucesso!"

' Fecha e destrói os objetos de conexão
If Not connSales Is Nothing Then If connSales.State = adStateOpen Then connSales.Close
If Not conn Is Nothing Then If conn.State = adStateOpen Then conn.Close
Set connSales = Nothing
Set conn = Nothing
%>