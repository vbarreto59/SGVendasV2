<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<% 'funcional mas salvando 500 como 50k'
' ====================================================================
' Script para Salvar TODOS os Pagamentos de uma Venda - Versão 1
' ====================================================================
Response.Buffer = True
Response.Expires = -1
On Error GoTo 0

' Configurações de banco de dados
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
If InStr(dbSunnyPath, ";") > 0 Then
    dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
End If

' Função para formatar números para SQL - CORRIGIDA
Function FormatNumberForSQL(sValue)
    On Error Resume Next
    If IsNumeric(sValue) Then
        ' Se já é número, converte diretamente
        FormatNumberForSQL = CDbl(sValue)
    Else
        ' Remove R$, pontos e converte vírgula para ponto
        sValue = Replace(sValue, "R$", "")
        sValue = Replace(sValue, ".", "")
        sValue = Replace(sValue, ",", ".")
        sValue = Trim(sValue)
        If IsNumeric(sValue) Then
            sValue = sValue/1
            FormatNumberForSQL = CDbl(sValue)
        Else
            FormatNumberForSQL = 0
        End If
    End If
    If Err.Number <> 0 Then
        FormatNumberForSQL = 0
    End If
    On Error GoTo 0
End Function

' Função para formatar número para string SQL - NOVA FUNÇÃO
Function FormatNumberForSQLString(valor)
    Dim valorFormatado
    On Error Resume Next
    
    ' Garante que é um número
    If Not IsNumeric(valor) Then
        valor = 0
    End If
    
    ' Formata como string com ponto decimal
    valorFormatado = Replace(FormatNumber(valor, 2), ",", ".")
    
    ' Remove formatação de milhar se houver
    valorFormatado = Replace(valorFormatado, ".", "", 1, 1)
    
    FormatNumberForSQLString = valorFormatado
    On Error GoTo 0
End Function

' Obter dados do formulário
Dim idComissao, idVenda, dataPagamento, statusPagamento, obs
idComissao = Request.Form("ID_Comissao")
idVenda = Request.Form("ID_Venda")
dataPagamento = Request.Form("DataPagamento")
statusPagamento = Request.Form("Status")
obs = Request.Form("Obs")

' Validação básica
If Not IsNumeric(idVenda) Or idVenda = "" Then
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Erro: ID da venda inválido."
End If
If dataPagamento = "" Then
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Erro: Data do pagamento não informada."
End If
If statusPagamento = "" Then
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Erro: Status do pagamento não informado."
End If

' ----------------------------------------------------------------------
' CONEXÕES E BUSCA DOS DADOS DA VENDA
' ----------------------------------------------------------------------
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Buscar dados completos da venda
Dim sqlVenda, rsVenda
sqlVenda = "SELECT " & _
           "v.ID, v.NomeEmpreendimento, v.Unidade, " & _
           "v.ValorLiqDiretoria, v.PremioDiretoria, " & _
           "v.ValorLiqGerencia, v.PremioGerencia, " & _
           "v.ValorLiqCorretor, v.PremioCorretor, " & _
           "c.UserIdDiretoria, c.UserIdGerencia, c.UserIdCorretor " & _
           "FROM Vendas AS v " & _
           "LEFT JOIN COMISSOES_A_PAGAR AS c ON v.ID = c.ID_Venda " & _
           "WHERE v.ID = " & idVenda

Set rsVenda = connSales.Execute(sqlVenda)

If rsVenda.EOF Then
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Erro: Venda não encontrada."
End If

' ----------------------------------------------------------------------
' BUSCAR PAGAMENTOS JÁ REALIZADOS PARA CALCULAR SALDOS
' ----------------------------------------------------------------------
Function GetTotalPago(idVenda, userId, tipoRecebedor, tipoPagamento)
    Dim sql, rs, total
    total = 0
    
    sql = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
          "WHERE ID_Venda = " & idVenda & " AND UsuariosUserId = " & userId & _
          " AND TipoRecebedor = '" & tipoRecebedor & "' AND TipoPagamento = '" & tipoPagamento & "'"
    
    Set rs = connSales.Execute(sql)
    If Not rs.EOF And Not IsNull(rs("TotalPago")) Then
        total = CDbl(rs("TotalPago"))
    End If
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    GetTotalPago = total
End Function

' Buscar totais já pagos
Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
Dim totalPremioPagoDiretoria, totalPremioPagoGerencia, totalPremioPagoCorretor

' Inicializar valores
totalPagoDiretoria = 0
totalPagoGerencia = 0
totalPagoCorretor = 0
totalPremioPagoDiretoria = 0
totalPremioPagoGerencia = 0
totalPremioPagoCorretor = 0

' Verificar se os UserIds existem antes de buscar pagamentos
If Not IsNull(rsVenda("UserIdDiretoria")) And rsVenda("UserIdDiretoria") <> "" Then
    totalPagoDiretoria = GetTotalPago(idVenda, rsVenda("UserIdDiretoria"), "diretoria", "Comissão")
    totalPremioPagoDiretoria = GetTotalPago(idVenda, rsVenda("UserIdDiretoria"), "diretoria", "Premiação")
End If

If Not IsNull(rsVenda("UserIdGerencia")) And rsVenda("UserIdGerencia") <> "" Then
    totalPagoGerencia = GetTotalPago(idVenda, rsVenda("UserIdGerencia"), "gerencia", "Comissão")
    totalPremioPagoGerencia = GetTotalPago(idVenda, rsVenda("UserIdGerencia"), "gerencia", "Premiação")
End If

If Not IsNull(rsVenda("UserIdCorretor")) And rsVenda("UserIdCorretor") <> "" Then
    totalPagoCorretor = GetTotalPago(idVenda, rsVenda("UserIdCorretor"), "corretor", "Comissão")
    totalPremioPagoCorretor = GetTotalPago(idVenda, rsVenda("UserIdCorretor"), "corretor", "Premiação")
End If

' ----------------------------------------------------------------------
' CALCULAR SALDOS A PAGAR
' ----------------------------------------------------------------------
Dim saldoDiretoriaComissao, saldoGerenciaComissao, saldoCorretorComissao
Dim saldoDiretoriaPremio, saldoGerenciaPremio, saldoCorretorPremio

' Inicializar valores
saldoDiretoriaComissao = 0
saldoGerenciaComissao = 0
saldoCorretorComissao = 0
saldoDiretoriaPremio = 0
saldoGerenciaPremio = 0
saldoCorretorPremio = 0

' Calcular valores totais (tratando nulos)
If Not IsNull(rsVenda("ValorLiqDiretoria")) Then saldoDiretoriaComissao = CDbl(rsVenda("ValorLiqDiretoria"))
If Not IsNull(rsVenda("ValorLiqGerencia")) Then saldoGerenciaComissao = CDbl(rsVenda("ValorLiqGerencia"))
If Not IsNull(rsVenda("ValorLiqCorretor")) Then saldoCorretorComissao = CDbl(rsVenda("ValorLiqCorretor"))

If Not IsNull(rsVenda("PremioDiretoria")) Then saldoDiretoriaPremio = CDbl(rsVenda("PremioDiretoria"))
If Not IsNull(rsVenda("PremioGerencia")) Then saldoGerenciaPremio = CDbl(rsVenda("PremioGerencia"))
If Not IsNull(rsVenda("PremioCorretor")) Then saldoCorretorPremio = CDbl(rsVenda("PremioCorretor"))

' Subtrair pagamentos já realizados
saldoDiretoriaComissao = saldoDiretoriaComissao - totalPagoDiretoria
saldoGerenciaComissao = saldoGerenciaComissao - totalPagoGerencia
saldoCorretorComissao = saldoCorretorComissao - totalPagoCorretor
saldoDiretoriaPremio = saldoDiretoriaPremio - totalPremioPagoDiretoria
saldoGerenciaPremio = saldoGerenciaPremio - totalPremioPagoGerencia
saldoCorretorPremio = saldoCorretorPremio - totalPremioPagoCorretor

' Garantir que valores negativos sejam zero
If saldoDiretoriaComissao < 0 Then saldoDiretoriaComissao = 0
If saldoGerenciaComissao < 0 Then saldoGerenciaComissao = 0
If saldoCorretorComissao < 0 Then saldoCorretorComissao = 0
If saldoDiretoriaPremio < 0 Then saldoDiretoriaPremio = 0
If saldoGerenciaPremio < 0 Then saldoGerenciaPremio = 0
If saldoCorretorPremio < 0 Then saldoCorretorPremio = 0

' ----------------------------------------------------------------------
' REALIZAR TODOS OS PAGAMENTOS PENDENTES
' ----------------------------------------------------------------------
Dim transacaoIniciada, pagamentosRealizados
transacaoIniciada = False
pagamentosRealizados = 0

On Error Resume Next

' Iniciar transação
connSales.BeginTrans
transacaoIniciada = True

' Função para buscar nome do usuário
Function GetNomeUsuario(userId, tipo)
    Dim sql, rs, nome
    nome = "Não Encontrado"
    
    If IsNull(userId) Or userId = "" Then
        GetNomeUsuario = nome
        Exit Function
    End If
    
    Select Case tipo
        Case "diretoria"
            sql = "SELECT Nome FROM Diretorias WHERE UserId = " & userId
        Case "gerencia"
            sql = "SELECT Nome FROM Gerencias WHERE UserId = " & userId
        Case "corretor"
            sql = "SELECT Nome FROM Usuarios WHERE UserId = " & userId
    End Select
    
    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        nome = rs("Nome")
    End If
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    GetNomeUsuario = nome
End Function

' Função para inserir pagamento - CORRIGIDA
Function InserirPagamento(idVenda, userId, nomeUsuario, valor, tipoRecebedor, tipoPagamento)
    Dim sqlInsert, valorFormatado
    
    ' Verificar se há valor para pagar
    If valor <= 0 Then
        InserirPagamento = True
        Exit Function
    End If
    
    ' Verificar se userId é válido
    If IsNull(userId) Or userId = "" Or userId = 0 Then
        InserirPagamento = False
        Exit Function
    End If
    
    ' Formatar valor corretamente para SQL
    valor = valor/100
    valorFormatado = FormatNumberForSQLString(valor)
    
    ' Preparar valores para SQL
    Dim dataPagamentoSQL, statusSQL, obsSQL, nomeUsuarioSQL, tipoRecebedorSQL, tipoPagamentoSQL
    dataPagamentoSQL = "'" & dataPagamento & "'"
    statusSQL = "'" & Replace(statusPagamento, "'", "''") & "'"
    obsSQL = "'" & Replace(obs, "'", "''") & "'"
    nomeUsuarioSQL = "'" & Replace(nomeUsuario, "'", "''") & "'"
    tipoRecebedorSQL = "'" & Replace(tipoRecebedor, "'", "''") & "'"
    tipoPagamentoSQL = "'" & Replace(tipoPagamento, "'", "''") & "'"
    
    sqlInsert = "INSERT INTO PAGAMENTOS_COMISSOES " & _
               "(ID_Venda, UsuariosUserId, UsuariosNome, DataPagamento, ValorPago, Status, Obs, TipoRecebedor, TipoPagamento) " & _
               "VALUES (" & idVenda & ", " & userId & ", " & nomeUsuarioSQL & ", " & _
               dataPagamentoSQL & ", " & valorFormatado & ", " & statusSQL & ", " & _
               obsSQL & ", " & tipoRecebedorSQL & ", " & tipoPagamentoSQL & ")"
    
    connSales.Execute sqlInsert
    
    If Err.Number = 0 Then
        InserirPagamento = True
    Else
        Response.Write "<!-- Erro no INSERT: " & Err.Description & " -->"
        Response.Write "<!-- SQL: " & sqlInsert & " -->"
        InserirPagamento = False
    End If
End Function

' DEBUG: Mostrar valores calculados
Response.Write "<!-- DEBUG - Saldos Calculados -->"
Response.Write "<!-- Diretoria Comissão: " & saldoDiretoriaComissao & " -->"
Response.Write "<!-- Gerência Comissão: " & saldoGerenciaComissao & " -->"
Response.Write "<!-- Corretor Comissão: " & saldoCorretorComissao & " -->"
Response.Write "<!-- Diretoria Prêmio: " & saldoDiretoriaPremio & " -->"
Response.Write "<!-- Gerência Prêmio: " & saldoGerenciaPremio & " -->"
Response.Write "<!-- Corretor Prêmio: " & saldoCorretorPremio & " -->"

' PAGAMENTOS DE COMISSÃO
If saldoDiretoriaComissao > 0 Then
    Dim nomeDiretor
    If Not IsNull(rsVenda("UserIdDiretoria")) Then
        nomeDiretor = GetNomeUsuario(rsVenda("UserIdDiretoria"), "diretoria")
        If InserirPagamento(idVenda, rsVenda("UserIdDiretoria"), nomeDiretor, saldoDiretoriaComissao, "diretoria", "Comissão") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Diretoria Comissão: OK -->"
        Else
            Response.Write "<!-- Pagamento Diretoria Comissão: FALHOU -->"
        End If
    End If
End If

If saldoGerenciaComissao > 0 Then
    Dim nomeGerente
    If Not IsNull(rsVenda("UserIdGerencia")) Then
        nomeGerente = GetNomeUsuario(rsVenda("UserIdGerencia"), "gerencia")
        If InserirPagamento(idVenda, rsVenda("UserIdGerencia"), nomeGerente, saldoGerenciaComissao, "gerencia", "Comissão") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Gerência Comissão: OK -->"
        Else
            Response.Write "<!-- Pagamento Gerência Comissão: FALHOU -->"
        End If
    End If
End If

If saldoCorretorComissao > 0 Then
    Dim nomeCorretor
    If Not IsNull(rsVenda("UserIdCorretor")) Then
        nomeCorretor = GetNomeUsuario(rsVenda("UserIdCorretor"), "corretor")
        If InserirPagamento(idVenda, rsVenda("UserIdCorretor"), nomeCorretor, saldoCorretorComissao, "corretor", "Comissão") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Corretor Comissão: OK -->"
        Else
            Response.Write "<!-- Pagamento Corretor Comissão: FALHOU -->"
        End If
    End If
End If

' PAGAMENTOS DE PRÊMIO
If saldoDiretoriaPremio > 0 Then
    If Not IsNull(rsVenda("UserIdDiretoria")) Then
        saldoDiretoriaPremio = saldoDiretoriaPremio/100
        If InserirPagamento(idVenda, rsVenda("UserIdDiretoria"), nomeDiretor, saldoDiretoriaPremio, "diretoria", "Premiação") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Diretoria Prêmio: OK -->"
        Else
            Response.Write "<!-- Pagamento Diretoria Prêmio: FALHOU -->"
        End If
    End If
End If

If saldoGerenciaPremio > 0 Then
    If Not IsNull(rsVenda("UserIdGerencia")) Then
        saldoGerenciaPremio = saldoGerenciaPremio/100
        If InserirPagamento(idVenda, rsVenda("UserIdGerencia"), nomeGerente, saldoGerenciaPremio, "gerencia", "Premiação") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Gerência Prêmio: OK -->"
        Else
            Response.Write "<!-- Pagamento Gerência Prêmio: FALHOU -->"
        End If
    End If
End If

If saldoCorretorPremio > 0 Then
    If Not IsNull(rsVenda("UserIdCorretor")) Then
        saldoCorretorPremio = saldoCorretorPremio/100 
        If InserirPagamento(idVenda, rsVenda("UserIdCorretor"), nomeCorretor, saldoCorretorPremio, "corretor", "Premiação") Then
            pagamentosRealizados = pagamentosRealizados + 1
            Response.Write "<!-- Pagamento Corretor Prêmio: OK -->"
        Else
            Response.Write "<!-- Pagamento Corretor Prêmio: FALHOU -->"
        End If
    End If
End If

' ----------------------------------------------------------------------
' FINALIZAR TRANSAÇÃO E ATUALIZAÇÕES
' ----------------------------------------------------------------------
Response.Write "<!-- Total de Pagamentos Realizados: " & pagamentosRealizados & " -->"
Response.Write "<!-- Erro Number: " & Err.Number & " -->"
Response.Write "<!-- Erro Description: " & Err.Description & " -->"

If Err.Number = 0 Then
    connSales.CommitTrans
    transacaoIniciada = False
    
    ' Atualizar status na tabela COMISSOES_A_PAGAR
    If Not IsNull(idComissao) And idComissao <> "" Then
        Dim sqlUpdateStatus
        sqlUpdateStatus = "UPDATE COMISSOES_A_PAGAR SET StatusPagamento = 'PAGA' WHERE ID_Comissoes = " & idComissao
        connSales.Execute sqlUpdateStatus
    End If
    
    ' Atualizações cross-database
    Dim sqlUpdate, adodb_path
    adodb_path = "[;DATABASE=" & dbSunnyPath & "]"

    sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Diretorias ON PAGAMENTOS_COMISSOES.UsuariosUserId = Diretorias.UserId SET PAGAMENTOS_COMISSOES.UsuariosNome = [Diretorias].[Nome] WHERE PAGAMENTOS_COMISSOES.TipoRecebedor='diretoria' AND PAGAMENTOS_COMISSOES.ID_Venda=" & idVenda
    'connSales.Execute sqlUpdate

    sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Gerencias ON PAGAMENTOS_COMISSOES.UsuariosUserId = Gerencias.UserId SET PAGAMENTOS_COMISSOES.UsuariosNome = [Gerencias].[Nome] WHERE PAGAMENTOS_COMISSOES.TipoRecebedor='gerencia' AND PAGAMENTOS_COMISSOES.ID_Venda=" & idVenda
    'connSales.Execute sqlUpdate

    sqlUpdate = "UPDATE PAGAMENTOS_COMISSOES INNER JOIN " & adodb_path & ".Usuarios ON PAGAMENTOS_COMISSOES.UsuariosUserId = Usuarios.UserId SET PAGAMENTOS_COMISSOES.UsuariosNome = [Usuarios].[Nome] WHERE PAGAMENTOS_COMISSOES.TipoRecebedor='corretor' AND PAGAMENTOS_COMISSOES.ID_Venda=" & idVenda
    'connSales.Execute sqlUpdate
    
    ' Redirecionamento de sucesso
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Sucesso: " & pagamentosRealizados & " pagamentos realizados para a venda " & idVenda & "!"
Else
    If transacaoIniciada Then
        connSales.RollbackTrans
    End If
    Response.Redirect "gestao_vendas_comissoes_pag_todos.asp?mensagem=Erro ao processar pagamentos: " & Server.URLEncode(Err.Description)
End If

' ----------------------------------------------------------------------
' LIMPEZA
' ----------------------------------------------------------------------
If Not rsVenda Is Nothing Then rsVenda.Close
Set rsVenda = Nothing

If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close
If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
Set connSales = Nothing
Set conn = Nothing
%>