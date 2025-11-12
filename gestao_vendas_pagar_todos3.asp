<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
' ====================================================================
' Script para Pagamento AUTOMÁTICO e Exibição de Saldo de Venda - Versão 2
' Recebe ID_Venda e processa todos os pagamentos pendentes (Comissão e Prêmio)
' para Diretoria, Gerência e Corretor. Se não houver pendências, exibe o histórico.
' ====================================================================
Response.Buffer = True
Response.Expires = -1
On Error GoTo 0

' ----------------------------------------------------------------------
' VARIÁVEIS DE AMBIENTE
' ----------------------------------------------------------------------
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
If InStr(dbSunnyPath, ";") > 0 Then
    dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
End If

Dim idVenda, dataPagamento, statusPagamento, obsPagamento, usuarioSessao ' Variável adicionada
Dim mensagemErro, mensagemSucesso
mensagemErro = ""
mensagemSucesso = ""

' Obter ID da Venda (preferência QueryString ou Form)
idVenda = Request.QueryString("id")
If idVenda = "" Then idVenda = Request.Form("ID")

' Configurações fixas para o pagamento automático
dataPagamento = Date()
statusPagamento = "PAGO"
obsPagamento = "Pagamento automático de saldo restante via script."
usuarioSessao = Session("Usuario") ' CAPTURA O NOME DO USUÁRIO DA SESSÃO

' ----------------------------------------------------------------------
' FUNÇÕES AUXILIARES
' ----------------------------------------------------------------------

' Função para formatar número para string SQL (com ponto decimal)
Function FormatNumberForSQLString(valor)
    Dim valorFormatado
    On Error Resume Next
    
    If Not IsNumeric(valor) Then valor = 0
    
    ' Formata como string com 2 casas decimais (ex: 1234.56)
    valorFormatado = Replace(FormatNumber(valor, 2), ",", "@@@") ' Troca vírgula por placeholder
    valorFormatado = Replace(valorFormatado, ".", "")           ' Remove separador de milhar (ponto)
    valorFormatado = Replace(valorFormatado, "@@@", ".")        ' Troca placeholder por ponto decimal
    
    FormatNumberForSQLString = valorFormatado
    On Error GoTo 0
End Function

' Função para buscar o total já pago
Function GetTotalPago(conn, idVenda, userId, tipoRecebedor, tipoPagamento)
    Dim sql, rs, total
    total = 0.00
    
    ' Usar CDbl(userId) para evitar problemas de tipo na query
    If Not IsNull(userId) And userId <> "" And IsNumeric(userId) Then
        sql = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
              "WHERE ID_Venda = " & idVenda & " AND UsuariosUserId = " & CDbl(userId) & _
              " AND TipoRecebedor = '" & tipoRecebedor & "' AND TipoPagamento = '" & tipoPagamento & "'"
        
        Set rs = conn.Execute(sql)
        If Not rs.EOF And Not IsNull(rs("TotalPago")) Then
            total = CDbl(rs("TotalPago"))
        End If
        If Not rs Is Nothing Then rs.Close
        Set rs = Nothing
    End If
    
    GetTotalPago = total
End Function

' Função para buscar nome do usuário (simulação simples)
Function GetNomeUsuario(conn, userId, tipo)
    Dim sql, rs, nome
    nome = "Usuário ID " & userId & " (" & tipo & ")"
    
    If IsNull(userId) Or userId = "" Or Not IsNumeric(userId) Then
        GetNomeUsuario = "ID Inválido"
        Exit Function
    End If
    
    ' Nota: As tabelas 'Diretorias', 'Gerencias', 'Usuarios' devem estar acessíveis pela 'conn'
    Select Case tipo
        Case "diretoria"
            sql = "SELECT Nome FROM Diretorias WHERE UserId = " & CDbl(userId)
        Case "gerencia"
            sql = "SELECT Nome FROM Gerencias WHERE UserId = " & CDbl(userId)
        Case "corretor"
            sql = "SELECT Nome FROM Usuarios WHERE UserId = " & CDbl(userId)
    End Select
    
    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        nome = rs("Nome")
    End If
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    GetNomeUsuario = nome
End Function

' Função para inserir pagamento - AGORA INCLUI O USUÁRIO DA SESSÃO
Function InserirPagamento(connSales, userId, idVenda, nomeUsuario, valor, dataPagamento, statusPagamento, obsPagamento, tipoRecebedor, tipoPagamento, usuarioSessao)
    Dim sqlInsert, valorFormatado
    InserirPagamento = False
    
    If valor <= 0.0001 Then Exit Function ' Quase zero
    If IsNull(userId) Or userId = "" Or Not IsNumeric(userId) Then Exit Function
    
    valorFormatado = FormatNumberForSQLString(valor)
    
    Dim dataPagamentoSQL, statusSQL, obsSQL, nomeUsuarioSQL, tipoRecebedorSQL, tipoPagamentoSQL, usuarioSessaoSQL
    
    ' Preparar valores para SQL
    dataPagamentoSQL = "'" & dataPagamento & "'"
    statusSQL = "'" & Replace(statusPagamento, "'", "''") & "'"
    obsSQL = "'" & Replace(obsPagamento, "'", "''") & "'"
    nomeUsuarioSQL = "'" & Replace(nomeUsuario, "'", "''") & "'"
    tipoRecebedorSQL = "'" & Replace(tipoRecebedor, "'", "''") & "'"
    tipoPagamentoSQL = "'" & Replace(tipoPagamento, "'", "''") & "'"
    usuarioSessaoSQL = "'" & Replace(usuarioSessao, "'", "''") & "'" ' NOVO CAMPO
    
    sqlInsert = "INSERT INTO PAGAMENTOS_COMISSOES " & _
                "(ID_Venda, UsuariosUserId, UsuariosNome, DataPagamento, ValorPago, Status, Obs, TipoRecebedor, TipoPagamento, Usuario) " & _
                "VALUES (" & CDbl(idVenda) & ", " & CDbl(userId) & ", " & nomeUsuarioSQL & ", " & _
                dataPagamentoSQL & ", " & valorFormatado & ", " & statusSQL & ", " & _
                obsSQL & ", " & tipoRecebedorSQL & ", " & tipoPagamentoSQL & ", " & usuarioSessaoSQL & ")" ' NOVO VALOR
    
    On Error Resume Next
    connSales.Execute sqlInsert
    
    If Err.Number = 0 Then
        InserirPagamento = True
    Else
        Response.Write "<!-- Erro ao inserir pagamento: " & Err.Description & ". SQL: " & sqlInsert & " -->"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------
' CONEXÕES E PROCESSAMENTO
' ----------------------------------------------------------------------
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Dim pagamentosRealizados
pagamentosRealizados = 0

If IsNumeric(idVenda) And idVenda <> "" Then
    idVenda = CDbl(idVenda)
    
    ' 1. Buscar dados da venda e IDs dos usuários
    Dim sqlVenda, rsVenda
    ' CORREÇÃO AQUI: Usando v.ValorUnidade conforme solicitado
    sqlVenda = "SELECT " & _
               "v.ID, v.NomeEmpreendimento, v.Unidade, v.ValorUnidade, v.DataVenda, " & _
               "v.ValorLiqDiretoria, v.PremioDiretoria, " & _
               "v.ValorLiqGerencia, v.PremioGerencia, " & _
               "v.ValorLiqCorretor, v.PremioCorretor, " & _
               "c.UserIdDiretoria, c.UserIdGerencia, c.UserIdCorretor " & _
               "FROM Vendas AS v " & _
               "LEFT JOIN COMISSOES_A_PAGAR AS c ON v.ID = c.ID_Venda " & _
               "WHERE v.ID = " & idVenda
    
    Set rsVenda = connSales.Execute(sqlVenda)
    
    If rsVenda.EOF Then
        mensagemErro = "Erro: Venda com ID " & idVenda & " não encontrada."
    Else
        ' Iniciar Transação
        Dim transacaoIniciada
        transacaoIniciada = False
        On Error Resume Next
        connSales.BeginTrans
        transacaoIniciada = True
        
        ' 2. Variáveis de Venda
        Dim nomeEmpreendimento, unidade, valorVenda
        nomeEmpreendimento = rsVenda("NomeEmpreendimento")
        unidade = rsVenda("Unidade")
        ' Usando rsVenda("ValorUnidade") para a exibição (o valor da comissão é ValorLiq...)
        valorVenda = FormatCurrency(rsVenda("ValorUnidade")) 
        
        ' 3. Calcular Saldos
        Dim totalPagoDiretoriaC, totalPagoDiretoriaP, totalPagoGerenciaC, totalPagoGerenciaP, totalPagoCorretorC, totalPagoCorretorP
        Dim saldoDiretoriaC, saldoDiretoriaP, saldoGerenciaC, saldoGerenciaP, saldoCorretorC, saldoCorretorP
        Dim totalComissaoDiretoria, totalComissaoGerencia, totalComissaoCorretor
        Dim totalPremioDiretoria, totalPremioGerencia, totalPremioCorretor
        
        If Not IsNull(rsVenda("ValorLiqDiretoria")) Then totalComissaoDiretoria = CDbl(rsVenda("ValorLiqDiretoria"))
        If Not IsNull(rsVenda("ValorLiqGerencia")) Then totalComissaoGerencia = CDbl(rsVenda("ValorLiqGerencia"))
        If Not IsNull(rsVenda("ValorLiqCorretor")) Then totalComissaoCorretor = CDbl(rsVenda("ValorLiqCorretor"))
        
        If Not IsNull(rsVenda("PremioDiretoria")) Then totalPremioDiretoria = CDbl(rsVenda("PremioDiretoria"))
        If Not IsNull(rsVenda("PremioGerencia")) Then totalPremioGerencia = CDbl(rsVenda("PremioGerencia"))
        If Not IsNull(rsVenda("PremioCorretor")) Then totalPremioCorretor = CDbl(rsVenda("PremioCorretor"))
        
        ' Buscar totais pagos
        totalPagoDiretoriaC = GetTotalPago(connSales, idVenda, rsVenda("UserIdDiretoria"), "diretoria", "Comissão")
        totalPagoDiretoriaP = GetTotalPago(connSales, idVenda, rsVenda("UserIdDiretoria"), "diretoria", "Premiação")
        totalPagoGerenciaC = GetTotalPago(connSales, idVenda, rsVenda("UserIdGerencia"), "gerencia", "Comissão")
        totalPagoGerenciaP = GetTotalPago(connSales, idVenda, rsVenda("UserIdGerencia"), "gerencia", "Premiação")
        totalPagoCorretorC = GetTotalPago(connSales, idVenda, rsVenda("UserIdCorretor"), "corretor", "Comissão")
        totalPagoCorretorP = GetTotalPago(connSales, idVenda, rsVenda("UserIdCorretor"), "corretor", "Premiação")
        
        ' Calcular saldos
        saldoDiretoriaC = totalComissaoDiretoria - totalPagoDiretoriaC
        saldoDiretoriaP = totalPremioDiretoria - totalPagoDiretoriaP
        saldoGerenciaC = totalComissaoGerencia - totalPagoGerenciaC
        saldoGerenciaP = totalPremioGerencia - totalPagoGerenciaP
        saldoCorretorC = totalComissaoCorretor - totalPagoCorretorC
        saldoCorretorP = totalPremioCorretor - totalPagoCorretorP
        
        ' Garantir que não se pague valores negativos (máximo a pagar é o saldo)
        If saldoDiretoriaC < 0 Then saldoDiretoriaC = 0
        If saldoDiretoriaP < 0 Then saldoDiretoriaP = 0
        If saldoGerenciaC < 0 Then saldoGerenciaC = 0
        If saldoGerenciaP < 0 Then saldoGerenciaP = 0
        If saldoCorretorC < 0 Then saldoCorretorC = 0
        If saldoCorretorP < 0 Then saldoCorretorP = 0
        
        ' 4. Realizar Pagamentos Pendentes
        Dim userIdDiretoria, userIdGerencia, userIdCorretor
        userIdDiretoria = rsVenda("UserIdDiretoria")
        userIdGerencia = rsVenda("UserIdGerencia")
        userIdCorretor = rsVenda("UserIdCorretor")
        
        ' Diretoria (Comissão)
        If saldoDiretoriaC > 0 And IsNumeric(userIdDiretoria) Then
            Dim nomeDiretor
            nomeDiretor = GetNomeUsuario(conn, userIdDiretoria, "diretoria")
            If InserirPagamento(connSales, userIdDiretoria, idVenda, nomeDiretor, saldoDiretoriaC, dataPagamento, statusPagamento, obsPagamento, "diretoria", "Comissão", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Diretoria (Premiação)
        If saldoDiretoriaP > 0 And IsNumeric(userIdDiretoria) Then
            If InserirPagamento(connSales, userIdDiretoria, idVenda, nomeDiretor, saldoDiretoriaP, dataPagamento, statusPagamento, obsPagamento, "diretoria", "Premiação", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Gerência (Comissão)
        If saldoGerenciaC > 0 And IsNumeric(userIdGerencia) Then
            Dim nomeGerente
            nomeGerente = GetNomeUsuario(conn, userIdGerencia, "gerencia")
            If InserirPagamento(connSales, userIdGerencia, idVenda, nomeGerente, saldoGerenciaC, dataPagamento, statusPagamento, obsPagamento, "gerencia", "Comissão", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Gerência (Premiação)
        If saldoGerenciaP > 0 And IsNumeric(userIdGerencia) Then
            If InserirPagamento(connSales, userIdGerencia, idVenda, nomeGerente, saldoGerenciaP, dataPagamento, statusPagamento, obsPagamento, "gerencia", "Premiação", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Corretor (Comissão)
        If saldoCorretorC > 0 And IsNumeric(userIdCorretor) Then
            Dim nomeCorretor
            nomeCorretor = GetNomeUsuario(conn, userIdCorretor, "corretor")
            If InserirPagamento(connSales, userIdCorretor, idVenda, nomeCorretor, saldoCorretorC, dataPagamento, statusPagamento, obsPagamento, "corretor", "Comissão", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Corretor (Premiação)
        If saldoCorretorP > 0 And IsNumeric(userIdCorretor) Then
            If InserirPagamento(connSales, userIdCorretor, idVenda, nomeCorretor, saldoCorretorP, dataPagamento, statusPagamento, obsPagamento, "corretor", "Premiação", usuarioSessao) Then
                pagamentosRealizados = pagamentosRealizados + 1
            End If
        End If
        
        ' Finalizar Transação
        If Err.Number = 0 Then
            connSales.CommitTrans
            transacaoIniciada = False
            
            If pagamentosRealizados > 0 Then
                mensagemSucesso = "Sucesso! " & pagamentosRealizados & " pagamentos pendentes processados para a Venda " & idVenda & ". (Registrado por: " & usuarioSessao & ")"
            Else
                mensagemSucesso = "Venda #" & idVenda & " consultada. Todos os pagamentos já haviam sido realizados."
            End If
        Else
            If transacaoIniciada Then connSales.RollbackTrans
            mensagemErro = "Erro fatal durante o processamento do pagamento: " & Err.Description
        End If
        
    End If ' Fim: If rsVenda.EOF
    
    If Not rsVenda Is Nothing Then rsVenda.Close
    Set rsVenda = Nothing
    
ElseIf idVenda <> "" Then
    mensagemErro = "Erro: O ID da Venda fornecido ('" & idVenda & "') é inválido."
End If

' ----------------------------------------------------------------------
' INÍCIO DA SAÍDA HTML
' ----------------------------------------------------------------------
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Processamento de Pagamento de Venda</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f7f9; color: #333; margin: 20px; }
        .container { max-width: 900px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1 { color: #007bff; text-align: center; margin-bottom: 20px; }
        .form-group { margin-bottom: 20px; display: flex; align-items: center; }
        .form-group label { flex: 0 0 120px; font-weight: bold; }
        .form-group input[type="text"] { flex-grow: 1; padding: 10px; border: 1px solid #ccc; border-radius: 6px; }
        .btn-submit { background-color: #28a745; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; transition: background-color 0.3s; margin-left: 20px; }
        .btn-submit:hover { background-color: #218838; }
        .alert { padding: 15px; border-radius: 6px; margin-bottom: 20px; font-weight: bold; }
        .alert-success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .alert-danger { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; border: 1px solid #ddd; text-align: left; }
        th { background-color: #007bff; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .section-title { color: #007bff; border-bottom: 2px solid #ccc; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; font-size: 1.2em; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Pagamento Automático de Comissões e Prêmios</h1>

        <% If mensagemSucesso <> "" Then %>
            <div class="alert alert-success"><%= mensagemSucesso %></div>
        <% End If %>

        <% If mensagemErro <> "" Then %>
            <div class="alert alert-danger"><%= mensagemErro %></div>
        <% End If %>

        <form method="post" action="gestao_vendas_pagar_todos3.asp">
            <div class="form-group">
                <label for="ID_Venda">ID da Venda:</label>
                <input type="text" id="ID" name="ID" value="<%= idVenda %>" required placeholder="Ex: 1234">
                <input type="submit" class="btn-submit" value="Processar Pagamento">
            </div>
        </form>

        <% ' Condição para exibir dados da venda e pagamentos, mesmo que não haja novos pagamentos %>
        <% If IsNumeric(idVenda) And idVenda <> "" And mensagemErro = "" Then %>
            
            <% 
            ' 5. Exibir Dados da Venda
            ' Re-executa a query da venda para garantir dados atualizados
            ' Usando v.ValorUnidade conforme solicitado
            sqlVenda = "SELECT " & _
               "v.ID, v.NomeEmpreendimento, v.Unidade, v.ValorUnidade, v.DataVenda, " & _
               "v.ValorLiqDiretoria, v.PremioDiretoria, " & _
               "v.ValorLiqGerencia, v.PremioGerencia, " & _
               "v.ValorLiqCorretor, v.PremioCorretor, " & _
               "c.UserIdDiretoria, c.UserIdGerencia, c.UserIdCorretor " & _
               "FROM Vendas AS v " & _
               "LEFT JOIN COMISSOES_A_PAGAR AS c ON v.ID = c.ID_Venda " & _
               "WHERE v.ID = " & idVenda
               
            Set rsVendaDisplay = connSales.Execute(sqlVenda)
            If Not rsVendaDisplay.EOF Then 
            %>
                <h2 class="section-title">Dados da Venda #<%= idVenda %></h2>
                <table>
                    <tr>
                        <th>Empreendimento</th>
                        <td><%= rsVendaDisplay("NomeEmpreendimento") %></td>
                        <th>Unidade</th>
                        <td><%= rsVendaDisplay("Unidade") %></td>
                    </tr>
                    <tr>
                        <th>Valor Total Venda/Unidade</th>
                        <td><%= FormatCurrency(rsVendaDisplay("ValorUnidade")) %></td> ' Exibindo ValorUnidade
                        <th>Data Venda</th>
                        <td><%= FormatDateTime(rsVendaDisplay("DataVenda"), 2) %></td>
                    </tr>
                </table>
            <% 
            End If
            If Not rsVendaDisplay Is Nothing Then rsVendaDisplay.Close
            Set rsVendaDisplay = Nothing
            %>

            <%
            ' 6. Exibir Dados de Pagamento
            Dim sqlPagamentos, rsPagamentos
            sqlPagamentos = "SELECT UsuariosNome, TipoRecebedor, TipoPagamento, ValorPago, DataPagamento, Status, Obs, Usuario " & _
                            "FROM PAGAMENTOS_COMISSOES " & _
                            "WHERE ID_Venda = " & idVenda & " ORDER BY DataPagamento DESC, TipoRecebedor"
            
            Set rsPagamentos = connSales.Execute(sqlPagamentos)
            %>

                <h2 class="section-title">Registros de Pagamento (ID Venda: <%= idVenda %>)</h2>
                <% If rsPagamentos.EOF Then %>
                    <p>Nenhum registro de pagamento encontrado para esta venda.</p>
                <% Else %>
                    <table>
                        <tr>
                            <th>Recebedor (Tipo)</th>
                            <th>Nome</th>
                            <th>Tipo Pagamento</th>
                            <th>Valor Pago</th>
                            <th>Data</th>
                            <th>Status</th>
                            <th>Registrado Por</th>
                            <th>Obs</th>
                        </tr>
                        <% Do While Not rsPagamentos.EOF %>
                            <tr>
                                <td><%= UCase(rsPagamentos("TipoRecebedor")) %></td>
                                <td><%= rsPagamentos("UsuariosNome") %></td>
                                <td><%= rsPagamentos("TipoPagamento") %></td>
                                <td><%= FormatCurrency(rsPagamentos("ValorPago")) %></td>
                                <td><%= FormatDateTime(rsPagamentos("DataPagamento"), 2) %></td>
                                <td><%= rsPagamentos("Status") %></td>
                                <td><%= rsPagamentos("Usuario") %></td>
                                <td><%= rsPagamentos("Obs") %></td>
                            </tr>
                        <% 
                        rsPagamentos.MoveNext
                        Loop 
                        %>
                    </table>
                <% End If %>

            <% 
            If Not rsPagamentos Is Nothing Then rsPagamentos.Close
            Set rsPagamentos = Nothing
        End If 
        %>

    </div>
</body>
</html>
<%
' ----------------------------------------------------------------------
' LIMPEZA FINAL DA CONEXÃO
' ----------------------------------------------------------------------
If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close
If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
Set connSales = Nothing
Set conn = Nothing
%>