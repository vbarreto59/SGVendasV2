<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% 'funcional 04 11 2025'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<!--#include file="AtualizarVendas.asp"-->
<!--#include file="gestao_atu_localizacao.asp"-->
<!--#include file="gestao_header.inc"-->

<!--#include file="usr_acoes_v4GVendas.inc"-->
<!--#include file="atualizarVendas.asp"-->
<!--#include file="atualizarVendas2.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"    
%>

<%
'============================= LOG ============================================'
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    set objMail = server.createobject("CDONTS.NewMail")
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    objMail.Subject = "SV-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    objMail.MailFormat = 0
    objMail.Body = "Página de Vendas (Gestão Vendas)"
    objMail.Send
    set objMail = Nothing
end if 
'============================= ATUALIZANDO O BANCO DE DADOS ==================='
%>

<% 
'Modificação para separar banco de dados em 08 08 2025'
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
%>

<%
'============================= ATUALIZANDO O BANCO DE DADOS ============================================================================'
Response.Buffer = True
Response.Expires = -1
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute(sqlUpdate1)

sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId];"
connSales.Execute(sqlUpdate2)

sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute(sqlUpdateCorretor)

sql = "UPDATE Vendas SET Semestre = SWITCH(Trimestre IN (1, 2), 1, Trimestre IN (3, 4), 2) WHERE Trimestre IS NOT NULL;"
On Error Resume Next
connSales.Execute sql
If Err.Number <> 0 Then
    Response.Write "Ocorreu um erro ao atualizar o campo Semestre: " & Err.Description
End If
On Error GoTo 0
%>

<%
' Função para remover números e asteriscos de uma string
Function RemoverNumeros(texto)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    RemoverNumerosEAsteriscos = regex.Replace(texto, "")
    RemoverNumeros = Trim(Replace(RemoverNumerosEAsteriscos, "  ", " "))
End Function

Dim mensagem
mensagem = Request.QueryString("mensagem")

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales
Set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT Vendas.*, Usuarios.Nome AS UsuarioNome FROM Vendas LEFT JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId WHERE (Vendas.Excluido <> -1 OR Vendas.Excluido IS NULL) ORDER BY Vendas.ID DESC;"
rs.CursorLocation = 3
rs.CursorType = 1
rs.LockType = 1
rs.Open sql, connSales

' Variáveis para cálculos
Dim totalValorHtml, totalComissaoHtml, totalVendas
Dim totalComissoesPagas, totalComissoesAPagar
Dim totalVendasPagas, totalVendasPendentes
totalValorHtml = 0
totalComissaoHtml = 0
totalVendas = 0
totalComissoesPagas = 0
totalComissoesAPagar = 0
totalVendasPagas = 0
totalVendasPendentes = 0

If Not rs.EOF Then
    Do While Not rs.EOF
        totalValorHtml = totalValorHtml + CDbl(rs("ValorUnidade"))
        totalComissaoHtml = totalComissaoHtml + CDbl(rs("ValorComissaoGeral"))
        totalVendas = totalVendas + 1
        
        ' Calcular status de pagamento para KPIs
        Dim sqlPagamentos, rsPagamentos
        Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
        Dim pagoDiretoria, pagoGerencia, pagoCorretor
        
        totalPagoDiretoria = 0
        totalPagoGerencia = 0
        totalPagoCorretor = 0
        pagoDiretoria = False
        pagoGerencia = False
        pagoCorretor = False

        sqlPagamentos = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " ORDER BY DataPagamento ASC;"
        Set rsPagamentos = connSales.Execute(sqlPagamentos)

        If Not rsPagamentos.EOF Then
            Do While Not rsPagamentos.EOF
                Select Case LCase(rsPagamentos("TipoRecebedor"))
                    Case "diretoria"
                        totalPagoDiretoria = totalPagoDiretoria + CDbl(rsPagamentos("ValorPago"))
                    Case "gerencia"
                        totalPagoGerencia = totalPagoGerencia + CDbl(rsPagamentos("ValorPago"))
                    Case "corretor"
                        totalPagoCorretor = totalPagoCorretor + CDbl(rsPagamentos("ValorPago"))
                End Select
                rsPagamentos.MoveNext
            Loop
        End If
        
        If Not rsPagamentos Is Nothing Then
            rsPagamentos.Close
            Set rsPagamentos = Nothing
        End If

        If rs("ValorDiretoria") > 0 And totalPagoDiretoria >= CDbl(rs("ValorDiretoria")) Then pagoDiretoria = True
        If rs("ValorDiretoria") = 0 Then pagoDiretoria = True
        If rs("ValorGerencia") > 0 And totalPagoGerencia >= CDbl(rs("ValorGerencia")) Then pagoGerencia = True
        If rs("ValorGerencia") = 0 Then pagoGerencia = True
        If rs("ValorCorretor") > 0 And totalPagoCorretor >= CDbl(rs("ValorCorretor")) Then pagoCorretor = True
        If rs("ValorCorretor") = 0 Then pagoCorretor = True

        ' Acumular totais para KPIs
        If pagoDiretoria And pagoGerencia And pagoCorretor Then
            totalComissoesPagas = totalComissoesPagas + CDbl(rs("ValorComissaoGeral"))
            totalVendasPagas = totalVendasPagas + 1
        Else
            totalComissoesAPagar = totalComissoesAPagar + CDbl(rs("ValorComissaoGeral"))
            totalVendasPendentes = totalVendasPendentes + 1
        End If
        
        rs.MoveNext
    Loop
    rs.MoveFirst
End If

' Calcular percentuais
Dim percentualPagas, percentualAPagar
If totalComissaoHtml > 0 Then
    percentualPagas = (totalComissoesPagas / totalComissaoHtml) * 100
    percentualAPagar = (totalComissoesAPagar / totalComissaoHtml) * 100
Else
    percentualPagas = 0
    percentualAPagar = 0
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Vendas | Sistema</title>
    <meta http-equiv="refresh" content="300">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --accent: #e74c3c;
            --success: #27ae60;
            --warning: #f39c12;
            --light-bg: #f8f9fa;
            --dark-text: #2c3e50;
            --light-text: #ecf0f1;
            --card-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            --hover-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
        }
        
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            color: var(--dark-text);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding-top: 100px;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1rem 0;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            margin-bottom: 2rem;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
        }
        
        .app-title {
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.5rem;
        }
        
        .main-content {
            margin-top: 20px;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: var(--card-shadow);
            transition: transform 0.3s, box-shadow 0.3s;
            margin-bottom: 1.5rem;
            overflow: hidden;
        }
        
        .card:hover {
            transform: translateY(-5px);
            box-shadow: var(--hover-shadow);
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1rem 1.5rem;
            font-weight: 600;
        }
        
        .stats-card {
            text-align: center;
            padding: 1.5rem;
            position: relative;
        }
        
        .stats-icon {
            font-size: 2rem;
            margin-bottom: 0.5rem;
            opacity: 0.8;
        }
        
        .stats-value {
            font-size: 1rem;
            font-weight: 700;
            margin: 0.5rem 0;
        }
        
        .stats-label {
            font-size: 0.9rem;
            color: #6c757d;
            font-weight: 500;
        }
        
        .stats-percent {
            font-size: 0.8rem;
            font-weight: 600;
            padding: 0.2rem 0.5rem;
            border-radius: 15px;
            margin-top: 0.5rem;
            display: inline-block;
        }
        
        .percent-success {
            background-color: #d4edda;
            color: #155724;
        }
        
        .percent-warning {
            background-color: #fff3cd;
            color: #856404;
        }
        
        .kpi-progress {
            height: 6px;
            background-color: #e9ecef;
            border-radius: 3px;
            margin-top: 0.5rem;
            overflow: hidden;
        }
        
        .progress-bar-pagas {
            background: linear-gradient(to right, var(--success), #20c997);
        }
        
        .progress-bar-apagar {
            background: linear-gradient(to right, var(--warning), #fd7e14);
        }
        
        .table-responsive {
            border-radius: 0 0 12px 12px;
        }
        
        .table th {
            background-color: var(--primary);
            color: white;
            font-weight: 600;
            border: none;
            padding: 1rem 0.75rem;
        }
        
        .table td {
            padding: 0.75rem;
            vertical-align: middle;
            border-color: #e9ecef;
        }
        
        .linha-paga {
            background-color: #e3f2fd !important;
        }
        
        .linha-pendente {
            background-color: #ffebee !important;
        }
        
        .btn-primary {
            background-color: var(--secondary);
            border-color: var(--secondary);
        }
        
        .btn-success {
            background-color: var(--success);
            border-color: var(--success);
        }
        
        .btn-warning {
            background-color: var(--warning);
            border-color: var(--warning);
            color: white;
        }
        
        .btn-danger {
            background-color: var(--accent);
            border-color: var(--accent);
        }
        
        .badge {
            font-size: 0.75em;
            padding: 0.4em 0.6em;
        }
        
        .badge-comissao {
            background-color: var(--secondary);
        }
        
        .badge-pago {
            background-color: var(--success);
        }
        
        .action-buttons {
            display: flex;
            gap: 5px;
            flex-wrap: wrap;
        }
        
        .filter-section {
            background-color: white;
            border-radius: 12px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            box-shadow: var(--card-shadow);
        }
        
        .mobile-card {
            background: white;
            border-radius: 10px;
            padding: 1.25rem;
            margin-bottom: 1rem;
            box-shadow: var(--card-shadow);
        }
        
        .mobile-card-pago {
            border-left: 4px solid var(--secondary);
            background-color: #e3f2fd;
        }
        
        .mobile-card-pendente {
            border-left: 4px solid var(--accent);
            background-color: #ffebee;
        }
        
        .mobile-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.75rem;
            border-bottom: 1px solid #e9ecef;
            padding-bottom: 0.5rem;
        }
        
        .mobile-card-title {
            font-weight: 600;
            color: var(--primary);
            margin: 0;
        }
        
        .mobile-card-body {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0.5rem;
        }
        
        .mobile-card-field {
            margin-bottom: 0.5rem;
        }
        
        .mobile-card-label {
            font-weight: 600;
            font-size: 0.8rem;
            color: #6c757d;
        }
        
        .mobile-card-value {
            font-size: 0.9rem;
        }
        
        @media (max-width: 767.98px) {
            body {
                padding-top: 90px;
            }
            
            .app-header {
                padding: 0.8rem 0;
            }
            
            .app-title {
                font-size: 1.3rem;
            }
            
            .desktop-table {
                display: none;
            }
            
            .mobile-cards {
                display: block;
            }
            
            .stats-card {
                padding: 1rem;
            }
            
            .stats-value {
                font-size: 1.5rem;
            }
            
            .stats-icon {
                font-size: 1.5rem;
            }
        }
        
        @media (min-width: 768px) {
            .mobile-cards {
                display: none;
            }
        }
        
        .alert-custom {
            border-radius: 10px;
            border-left: 4px solid var(--success);
        }
        
        .status-pago {
            background-color: var(--success);
            color: white;
        }
        .status-pendente {
            background-color: red;
            color: white;
        }        
        
        .comissao-info {
            font-size: 0.8rem;
            color: #6c757d;
        }
        
        .premio-pago {
            color: var(--success);
            font-weight: bold;
        }
    </style>
</head>

<body>
    <header class="app-header">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 class="app-title"><i class="fas fa-chart-line"></i> Gestão de Vendas</h1>
                </div>
                <div class="col-md-6 text-end">
                    <div class="btn-group btn-group-responsive">
                        <button type="button" onclick="window.close();" class="btn btn-light btn-sm">
                            <i class="fas fa-times me-1"></i>Fechar
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <div class="container-fluid main-content">
        <% If mensagem <> "" Then %>
            <div class="alert alert-success alert-custom alert-dismissible fade show">
                <i class="fas fa-check-circle me-2"></i><%= mensagem %>
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            </div>
        <% End If %>
        
        <div class="row mb-4">
            <div class="col-md-2 col-sm-4 mb-2">
                <div class="p-2 border rounded text-center bg-light">
                    <div class="text-primary" style="font-size: 0.8em;"><i class="fas fa-shopping-cart mr-1"></i> <%= totalVendas %></div>
                    <div class="text-muted" style="font-size: 0.8em;">Vendas</div>
                </div>
            </div>
            <div class="col-md-2 col-sm-4 mb-2">
                <div class="p-2 border rounded text-center bg-light">
                    <div class="text-success font-weight-bold" style="font-size: 0.8em;"><i class="fas fa-money-bill-wave mr-1"></i> R$ <%= FormatNumber(totalValorHtml, 2) %></div>
                    <div class="text-muted" style="font-size: 0.8em;">Valor Total</div>
                </div>
            </div>
            <div class="col-md-2 col-sm-4 mb-2">
                <div class="p-2 border rounded text-center bg-light">
                    <div class="text-info font-weight-bold" style="font-size: 0.8em;"><i class="fas fa-percentage mr-1"></i> R$ <%= FormatNumber(totalComissaoHtml, 2) %></div>
                    <div class="text-muted" style="font-size: 0.8em;">Total Comissões</div>
                </div>
            </div>

            <div class="col-md-3 col-sm-6 mb-2">
                <div class="p-2 border rounded text-center bg-light">
                    <div class="text-success font-weight-bold" style="font-size: 0.8em;">
                        <i class="fas fa-check-circle mr-1"></i> R$ <%= FormatNumber(totalComissoesPagas, 2) %>
                        <span class="badge badge-success ml-1"><%= FormatNumber(percentualPagas, 1) %>%</span>
                    </div>
                    <div class="progress mt-1" style="height: 5px;">
                        <div class="progress-bar bg-success" role="progressbar" style="width: <%= percentualPagas %>%"></div>
                    </div>
                    <div class="text-muted" style="font-size: 0.8em;">Comissões Pagas (<%= totalVendasPagas %> Vendas)</div>
                </div>
            </div>
            <div class="col-md-3 col-sm-6 mb-2">
                <div class="p-2 border rounded text-center bg-light">
                    <div class="text-warning font-weight-bold" style="font-size: 0.8em;">
                        <i class="fas fa-clock mr-1"></i> R$ <%= FormatNumber(totalComissoesAPagar, 2) %>
                        <span class="badge badge-warning ml-1"><%= FormatNumber(percentualAPagar, 1) %>%</span>
                    </div>
                    <div class="progress mt-1" style="height: 5px;">
                        <div class="progress-bar bg-warning" role="progressbar" style="width: <%= percentualAPagar %>%"></div>
                    </div>
                    <div class="text-muted" style="font-size: 0.8em;">Comissões a Pagar (<%= totalVendasPendentes %> Pendentes)</div>
                </div>
            </div>
        </div>

        <div class="filter-section">
            <div class="row">
                <div class="col-md-8">
                    <h5 class="mb-3"><i class="fas fa-filter me-2"></i>Filtros e Ações</h5>
                </div>
                <div class="col-md-4 text-end">
                    <div class="btn-group">
                        <a href="gestao_vendas_create2.asp" class="btn btn-info btn-sm" target="_blank">
                            <i class="fas fa-plus me-1"></i>Nova Venda
                        </a>
                        <a href="gestao_vendas_gerenc_comissoes.asp" class="btn btn-primary btn-sm" target="_blank">
                            <i class="fas fa-money-bill-wave me-1"></i>Comissões 1
                        </a>
                        <a href="gestao_vendas_comissoes_pag_todos.asp" class="btn btn-primary btn-sm" target="_blank">
                            <i class="fas fa-money-bill-wave me-1"></i>Comissões 2
                        </a>       

                        <a href="gestao_vendas_inserir_comissao_todos1.asp" class="btn btn-primary btn-sm" target="_blank">
                            <i class="fas fa-money-bill-wave me-1"></i>Inserir Todas
                        </a>  

                        <a href="gestao_vendas_list_excluidos.asp" class="btn btn-warning btn-sm" target="_blank">
                            <i class="fas fa-trash-restore me-1"></i>Excluídos
                        </a>
                        <%if Session("Usuario")="BARRETO" Then%>
                            <div class="btn-group">
                                <button type="button" class="btn btn-info btn-sm dropdown-toggle" data-bs-toggle="dropdown">
                                    <i class="fas fa-tools me-1"></i>Utilitários
                                </button>
                                <ul class="dropdown-menu">
                                    <li><a class="dropdown-item" href="inserirVendasTeste2.asp" target="_blank"><i class="fas fa-plus me-1"></i>Inserir Testes</a></li>
                                    <li><a class="dropdown-item" href="excluir_testes.asp" target="_blank"><i class="fas fa-trash me-1"></i>Excluir Testes</a></li>
                                    <li><a class="dropdown-item" href="tool_excluir_tudo.asp" target="_blank"><i class="fas fa-trash me-1"></i>Excluir Vendas</a></li>   

                                    <li><a class="dropdown-item" href="tool_visualizar_log.asp" target="_blank"><i class="fas fa-trash me-1"></i>Log Sistema</a></li>  
                                </ul>
                            </div>
                        <%end if%>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="card">
            <div class="card-header">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0"><i class="fas fa-list me-2"></i>Lista de Vendas</h5>
                    <div>
                        <span class="badge bg-success me-2"><i class="fas fa-check me-1"></i><%= totalVendasPagas %> Pagas</span>
                        <span class="badge bg-warning me-2"><i class="fas fa-clock me-1"></i><%= totalVendasPendentes %> Pendentes</span>
                        <span class="badge bg-light text-dark"><%= totalVendas %> Total</span>
                    </div>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="desktop-table">
                    <div class="table-responsive">
                        <table id="tabelaVendas" class="table table-hover" style="width:100%">
                            <thead>
                                <tr>
                                    <th>Data/ID</th>
                                    <th>Status</th>
                                    <th>Trimestre</th>
                                    
                                    <th>Empreendimento</th>
                                    <th>Unidade</th>
                                    <th>Diretoria</th>
                                    <th>Gerência</th>
                                    <th>Corretor</th>
                                    <th>Valor (R$)</th>
                                    <th>Comissão</th>
                                    <th>Registro</th>
                                    <th width="180">Ações</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If Not rs.EOF Then
                                    rs.MoveFirst

                                    Do While Not rs.EOF
                                    
                                        totalPagoDiretoria = 0
                                        totalPagoGerencia = 0
                                        totalPagoCorretor = 0
                                        dataPagamentoDiretoria = ""
                                        dataPagamentoGerencia = ""
                                        dataPagamentoCorretor = ""
                                        tooltipDiretoria = ""
                                        tooltipGerencia = ""
                                        tooltipCorretor = ""
                                        pagoDiretoria = False
                                        pagoGerencia = False
                                        pagoCorretor = False

                                        sqlPagamentos = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " ORDER BY DataPagamento ASC;"
                                        Set rsPagamentos = connSales.Execute(sqlPagamentos)

                                        If Not rsPagamentos.EOF Then
                                            Do While Not rsPagamentos.EOF
                                                Dim detalhePagamento
                                                detalhePagamento = "Data: " & FormatDateTime(rsPagamentos("DataPagamento"), 2) & " | Valor: R$ " & FormatNumber(rsPagamentos("ValorPago"), 2) & " | Status: " & rsPagamentos("Status")
                                                Select Case LCase(rsPagamentos("TipoRecebedor"))
                                                    Case "diretoria"
                                                        If tooltipDiretoria <> "" Then tooltipDiretoria = tooltipDiretoria & Chr(13)
                                                        tooltipDiretoria = tooltipDiretoria & detalhePagamento
                                                        
                                                        If Not IsNull(rsPagamentos("ValorPago")) And IsNumeric(rsPagamentos("ValorPago")) Then
                                                            totalPagoDiretoria = totalPagoDiretoria + CDbl(rsPagamentos("ValorPago"))
                                                        End If
                                                        
                                                        dataPagamentoDiretoria = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                                    Case "gerencia"
                                                        If tooltipGerencia <> "" Then tooltipGerencia = tooltipGerencia & Chr(13)
                                                        tooltipGerencia = tooltipGerencia & detalhePagamento
                                                        
                                                        If Not IsNull(rsPagamentos("ValorPago")) And IsNumeric(rsPagamentos("ValorPago")) Then
                                                            totalPagoGerencia = totalPagoGerencia + CDbl(rsPagamentos("ValorPago"))
                                                        End If
                                                        
                                                        dataPagamentoGerencia = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                                    Case "corretor"
                                                        If tooltipCorretor <> "" Then tooltipCorretor = tooltipCorretor & Chr(13)
                                                        tooltipCorretor = tooltipCorretor & detalhePagamento
                                                        
                                                        If Not IsNull(rsPagamentos("ValorPago")) And IsNumeric(rsPagamentos("ValorPago")) Then
                                                            totalPagoCorretor = totalPagoCorretor + CDbl(rsPagamentos("ValorPago"))
                                                        End If
                                                        
                                                        dataPagamentoCorretor = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                                End Select
                                                rsPagamentos.MoveNext
                                            Loop
                                        End If
                                        
                                        If Not rsPagamentos Is Nothing Then
                                            rsPagamentos.Close
                                            Set rsPagamentos = Nothing
                                        End If

                                        Dim dblValorDiretoria, dblValorGerencia, dblValorCorretor
                                        
                                        If Not IsNull(rs("ValorDiretoria")) And IsNumeric(rs("ValorDiretoria")) Then dblValorDiretoria = CDbl(rs("ValorDiretoria")) Else dblValorDiretoria = 0
                                        If Not IsNull(rs("ValorGerencia")) And IsNumeric(rs("ValorGerencia")) Then dblValorGerencia = CDbl(rs("ValorGerencia")) Else dblValorGerencia = 0
                                        If Not IsNull(rs("ValorCorretor")) And IsNumeric(rs("ValorCorretor")) Then dblValorCorretor = CDbl(rs("ValorCorretor")) Else dblValorCorretor = 0
                                        
                                        If dblValorDiretoria > 0 And totalPagoDiretoria >= dblValorDiretoria Then pagoDiretoria = True
                                        If dblValorDiretoria = 0 Then pagoDiretoria = True
                                        If dblValorGerencia > 0 And totalPagoGerencia >= dblValorGerencia Then pagoGerencia = True
                                        If dblValorGerencia = 0 Then pagoGerencia = True
                                        If dblValorCorretor > 0 And totalPagoCorretor >= dblValorCorretor Then pagoCorretor = True
                                        If dblValorCorretor = 0 Then pagoCorretor = True

                                        Dim comissaoText
                                        comissaoText = FormatNumber(rs("ComissaoPercentual"), 2) & "%"
                                        If Not IsNull(rs("ValorComissaoGeral")) And CDbl(rs("ValorComissaoGeral")) > 0 Then
                                            comissaoText = comissaoText & " (R$ " & FormatNumber(rs("ValorComissaoGeral"), 2) & ")"
                                        End If

                                        vAno = Right(rs("AnoVenda"), 2)

                                        Dim rsComissaoCheck, comissaoExiste
                                        Set rsComissaoCheck = Server.CreateObject("ADODB.Recordset")
                                        rsComissaoCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(rs("ID")), connSales
                                        comissaoExiste = Not rsComissaoCheck.EOF
                                        rsComissaoCheck.Close
                                        Set rsComissaoCheck = Nothing
                                        
                                        Dim linhaClasse
                                        If pagoDiretoria And pagoGerencia And pagoCorretor Then
                                            linhaClasse = "linha-paga"
                                        Else
                                            linhaClasse = "linha-pendente"
                                        End If
                                %>
                                <tr class="<%= linhaClasse %>">
                                    <td>
                                        <%= rs("AnoVenda") & "-" & Right("0"&rs("MesVenda"),2) & "-" & Right("0"&rs("DiaVenda"),2) %>
                                        <br><span class="fw-bold"><%= rs("ID") %></span>
                                    </td>
                                    <td>
                                        <% If pagoDiretoria And pagoGerencia And pagoCorretor Then %>
                                            <span class="badge status-pago" title="Comissões pagas">PAGO</span>
                                        <% Else %>  
                                            <span class="badge status-pendente" title="Comissões pendentes">PENDENTE</span>
                                        <% End If %>
                                    </td>  
                                    <td>
                                        <small class="text-muted"><%= vAno & "T" & rs("Trimestre") %></small>
                                    </td>
                                    
                                    <td>
                                        <strong><%= rs("Empreend_ID") %>-<%= RemoverNumeros(rs("NomeEmpreendimento")) %></strong>
                                        <br><small class="text-muted"><%= RemoverNumeros(rs("Localidade")) %></small>
                                    </td>
                                    <td><%= rs("Unidade") %></td>
                                    <!-- ############################ -->
                                    <% ' --------------------------------------- %>
<% ' COLUNA DIRETORIA: Comissao + Prêmio     %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Diretoria") %></div>
    <small class="comissao-info">
        <% If pagoDiretoria Then %>
            <span class="badge bg-success">
                <i class="fas fa-check"></i>
            </span>
        <% End If %> 
        R$ <%= FormatNumber(rs("ValorDiretoria"), 2) %>
        
        <% ' Verificação e Exibição do PRÊMIO DIRETORIA %>
        <% If Not IsNull(rs("premioDiretoria")) Then %>
            <% If IsNumeric(rs("premioDiretoria")) And CDbl(rs("premioDiretoria")) > 0 Then %>         
                <br>
                <span class="text-primary fw-bold">
                    <% ' VERIFICAR SE O PRÊMIO DA DIRETORIA FOI PAGO %>
                    <%
                    Dim premioPagoDiretoria
                    premioPagoDiretoria = False
                    
                    ' CORREÇÃO: Usando o campo TipoPagamento = 'Premiação'
                    sqlPagamentosPremio = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " AND TipoRecebedor = 'diretoria' AND TipoPagamento = 'Premiação'"
                    Set rsPagamentosPremio = connSales.Execute(sqlPagamentosPremio)
                    
                    If Not rsPagamentosPremio.EOF Then
                        Dim totalPagoPremioDiretoria
                        totalPagoPremioDiretoria = 0
                        Do While Not rsPagamentosPremio.EOF
                            ' Soma todos os pagamentos do tipo Premiação para diretoria
                            If Not IsNull(rsPagamentosPremio("ValorPago")) And IsNumeric(rsPagamentosPremio("ValorPago")) Then
                                totalPagoPremioDiretoria = totalPagoPremioDiretoria + CDbl(rsPagamentosPremio("ValorPago"))
                            End If
                            rsPagamentosPremio.MoveNext
                        Loop
                        
                        ' Verifica se o total pago em premiações é maior ou igual ao prêmio devido
                        If totalPagoPremioDiretoria >= CDbl(rs("premioDiretoria")) Then
                            premioPagoDiretoria = True
                        End If
                    End If
                    
                    If Not rsPagamentosPremio Is Nothing Then
                        rsPagamentosPremio.Close
                        Set rsPagamentosPremio = Nothing
                    End If
                    %>
                    
                    <% If premioPagoDiretoria Then %>
                        <span class="badge bg-success me-1">
                            <i class="fas fa-check"></i>
                        </span>
                    <% End If %>
                    <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioDiretoria"), 2) %>
                </span>
            <% End If %>
        <% End If %>
    </small>
</td>

<% ' --------------------------------------- %>
<% ' COLUNA GERÊNCIA: Comissao + Prêmio      %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Gerencia") %></div>
    <small class="comissao-info">
        <% If pagoGerencia Then %>
            <span class="badge bg-success">
                <i class="fas fa-check"></i>
            </span>
        <% End If %>              
        R$ <%= FormatNumber(rs("ValorGerencia"), 2) %>
        
        <% ' Verificação e Exibição do PRÊMIO GERÊNCIA %>
        <% If Not IsNull(rs("premioGerencia")) Then %>
            <% If IsNumeric(rs("premioGerencia")) And CDbl(rs("premioGerencia")) > 0 Then %>
                <br>
                <span class="text-primary fw-bold">
                    <% ' VERIFICAR SE O PRÊMIO DA GERÊNCIA FOI PAGO %>
                    <%
                    Dim premioPagoGerencia
                    premioPagoGerencia = False
                    
                    ' CORREÇÃO: Usando o campo TipoPagamento = 'Premiação'
                    sqlPagamentosPremio = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " AND TipoRecebedor = 'gerencia' AND TipoPagamento = 'Premiação'"
                    Set rsPagamentosPremio = connSales.Execute(sqlPagamentosPremio)
                    
                    If Not rsPagamentosPremio.EOF Then
                        Dim totalPagoPremioGerencia
                        totalPagoPremioGerencia = 0
                        Do While Not rsPagamentosPremio.EOF
                            ' Soma todos os pagamentos do tipo Premiação para gerencia
                            If Not IsNull(rsPagamentosPremio("ValorPago")) And IsNumeric(rsPagamentosPremio("ValorPago")) Then
                                totalPagoPremioGerencia = totalPagoPremioGerencia + CDbl(rsPagamentosPremio("ValorPago"))
                            End If
                            rsPagamentosPremio.MoveNext
                        Loop
                        
                        ' Verifica se o total pago em premiações é maior ou igual ao prêmio devido
                        If totalPagoPremioGerencia >= CDbl(rs("premioGerencia")) Then
                            premioPagoGerencia = True
                        End If
                    End If
                    
                    If Not rsPagamentosPremio Is Nothing Then
                        rsPagamentosPremio.Close
                        Set rsPagamentosPremio = Nothing
                    End If
                    %>
                    
                    <% If premioPagoGerencia Then %>
                        <span class="badge bg-success me-1">
                            <i class="fas fa-check"></i>
                        </span>
                    <% End If %>
                    <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioGerencia"), 2) %>
                </span>
            <% End If %>
        <% End If %>
    </small>
</td>

<% ' --------------------------------------- %>
<% ' COLUNA CORRETOR: Comissao + Prêmio      %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Corretor") %></div>
    <small class="comissao-info">
        <% If pagoCorretor Then %>
            <span class="badge bg-success">
                <i class="fas fa-check"></i>
            </span>
        <% End If %>             
        R$ <%= FormatNumber(rs("ValorCorretor"), 2) %>
        
        <% ' Verificação e Exibição do PRÊMIO CORRETOR %>
        <% If Not IsNull(rs("premioCorretor")) Then %>
            <% If IsNumeric(rs("premioCorretor")) And CDbl(rs("premioCorretor")) > 0 Then %>
                <br>                        
                <span class="text-primary fw-bold">
                    <% ' VERIFICAR SE O PRÊMIO DO CORRETOR FOI PAGO %>
                    <%
                    Dim premioPagoCorretor
                    premioPagoCorretor = False
                    
                    ' CORREÇÃO: Usando o campo TipoPagamento = 'Premiação'
                    sqlPagamentosPremio = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " AND TipoRecebedor = 'corretor' AND TipoPagamento = 'Premiação'"
                    Set rsPagamentosPremio = connSales.Execute(sqlPagamentosPremio)
                    
                    If Not rsPagamentosPremio.EOF Then
                        Dim totalPagoPremioCorretor
                        totalPagoPremioCorretor = 0
                        Do While Not rsPagamentosPremio.EOF
                            ' Soma todos os pagamentos do tipo Premiação para corretor
                            If Not IsNull(rsPagamentosPremio("ValorPago")) And IsNumeric(rsPagamentosPremio("ValorPago")) Then
                                totalPagoPremioCorretor = totalPagoPremioCorretor + CDbl(rsPagamentosPremio("ValorPago"))
                            End If
                            rsPagamentosPremio.MoveNext
                        Loop
                        
                        ' Verifica se o total pago em premiações é maior ou igual ao prêmio devido
                        If totalPagoPremioCorretor >= CDbl(rs("premioCorretor")) Then
                            premioPagoCorretor = True
                        End If
                    End If
                    
                    If Not rsPagamentosPremio Is Nothing Then
                        rsPagamentosPremio.Close
                        Set rsPagamentosPremio = Nothing
                    End If
                    %>
                    
                    <% If premioPagoCorretor Then %>
                        <span class="badge bg-success me-1">
                            <i class="fas fa-check"></i>
                        </span>
                    <% End If %>
                    <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioCorretor"), 2) %>
                </span>
            <% End If %>
        <% End If %>
    </small>
</td>
                                    <!-- ############################ -->
                                    
                                    <td class="text-end fw-bold" data-order="<%= rs("ValorUnidade") %>">
                                        <%= FormatNumber(rs("ValorUnidade"), 2) %>
                                    </td>
                                    <td class="text-end">
                                        <span class="badge badge-comissao"><%= comissaoText %></span>
                                    </td>
                                    <td>
                                        <small>
                                            <%= FormatDateTime(rs("DataRegistro"),2) %>
                                            <br>por <%= rs("Usuario") %>
                                        </small>
                                    </td>
                                    <td class="text-center">
                                        <div class="action-buttons">
                                            <a href="gestao_vendas_update2.asp?id=<%= rs("id") %>" class="btn btn-warning btn-sm" title="Editar">
                                                <i class="fas fa-edit"></i>
                                            </a>
                                            <% If Not comissaoExiste Then %>
                                                <a href="gestao_vendas_inserir_comissao1.asp?id=<%= rs("id") %>" class="btn btn-primary btn-sm" title="Inserir Comissão">
                                                    <i class="fas fa-hand-holding-usd"></i>
                                                </a>
                                            <% End If %>
                                            <% 'If UCase(Session("Usuario")) = "BARRETO" Then %>
                                                <a href="gestao_vendas_delete.asp?id=<%= rs("id") %>" class="btn btn-danger btn-sm" title="Excluir" onclick="return confirm('Confirma exclusão desta venda?');">
                                                    <i class="fas fa-trash"></i>
                                                </a>
                                            <% 'End If %>
                                        </div>
                                    </td>
                                </tr>
                                <%
                                        rs.MoveNext
                                    Loop
                                End If
                                %>
                            </tbody>
                            <tfoot>
                                <tr class="table-light">
                                    <th colspan="9" class="text-end">Totais:</th>
                                    <th id="totalValor" class="text-end">R$ <%= FormatNumber(totalValorHtml, 2) %></th>
                                    <th id="totalComissao" class="text-end">R$ <%= FormatNumber(totalComissaoHtml, 2) %></th>
                                    <th colspan="2"></th>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
                
                <!-- Conteúdo Mobile -->
                <div class="mobile-cards p-3" id="mobileCardsContainer">
                    <!-- O conteúdo mobile permanece o mesmo -->
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
    $(document).ready(function () {
        function initDataTable() {
            if (!$.fn.DataTable.isDataTable('#tabelaVendas')) {
                $('#tabelaVendas').DataTable({
                    language: {
                        url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
                    },
                    pageLength: 10,
                    order: [[0, "desc"]],
                    responsive: true,
                    dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6"f>>rt<"row"<"col-sm-12 col-md-6"i><"col-sm-12 col-md-6"p>>'
                });
            }
        }

        function checkScreenSize() {
            if (window.matchMedia('(max-width: 767.98px)').matches) {
                $('.desktop-table').hide();
                $('.mobile-cards').show();
            } else {
                $('.mobile-cards').hide();
                $('.desktop-table').show();
                initDataTable();
            }
        }

        checkScreenSize();
        $(window).on('resize', checkScreenSize);
    });
    </script>
</body>
</html>
<%
' Fechar conexões
If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
End If

If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>