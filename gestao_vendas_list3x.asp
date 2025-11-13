<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
 If Session("Usuario") = "" Then
    Response.redirect "gestao_login.asp"
 end if   
%>
<% 'funcional 04 11 2025'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

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
'connSales.Execute(sqlUpdate1)

sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId];"
'connSales.Execute(sqlUpdate2)

sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) SET Vendas.Corretor = Usuarios.Nome;"
'connSales.Execute(sqlUpdateCorretor)

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
' ##############################################################################################
If Not rs.EOF Then
    Do While Not rs.EOF
        totalValorHtml = totalValorHtml + CDbl(rs("ValorUnidade"))
        
        ' 🆕 CORREÇÃO: Calcular comissão total corretamente
        Dim comissaoTotalVenda
        comissaoTotalVenda = 0
        If Not IsNull(rs("ValorDiretoria")) And IsNumeric(rs("ValorDiretoria")) Then comissaoTotalVenda = comissaoTotalVenda + CDbl(rs("ValorDiretoria"))
        If Not IsNull(rs("ValorGerencia")) And IsNumeric(rs("ValorGerencia")) Then comissaoTotalVenda = comissaoTotalVenda + CDbl(rs("ValorGerencia"))
        If Not IsNull(rs("ValorCorretor")) And IsNumeric(rs("ValorCorretor")) Then comissaoTotalVenda = comissaoTotalVenda + CDbl(rs("ValorCorretor"))
        
        totalComissaoHtml = totalComissaoHtml + comissaoTotalVenda
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
            vPagoTotal = CDBL(totalPagoDiretoria)+CDBL(totalPagoGerencia)+CDBL(totalPagoCorretor)
        End If
        
        If Not rsPagamentos Is Nothing Then
            rsPagamentos.Close
            Set rsPagamentos = Nothing
        End If


        If rs("ValorLiqDiretoria") > 0 And totalPagoDiretoria >= CDbl(rs("ValorLiqDiretoria")) Then pagoDiretoria = True
        If rs("ValorLiqDiretoria") = 0 Then pagoDiretoria = True
        If rs("ValorLiqGerencia") > 0 And totalPagoGerencia >= CDbl(rs("ValorLiqGerencia")) Then pagoGerencia = True
        If rs("ValorLiqGerencia") = 0 Then pagoGerencia = True
        If rs("ValorLiqCorretor") > 0 And totalPagoCorretor >= CDbl(rs("ValorLiqCorretor")) Then pagoCorretor = True
        If rs("ValorLiqCorretor") = 0 Then pagoCorretor = True

        ' Acumular totais para KPIs
        If pagoDiretoria And pagoGerencia And pagoCorretor Then
            totalComissoesPagas = totalComissoesPagas + comissaoTotalVenda
            totalVendasPagas = totalVendasPagas + 1
        Else
            totalComissoesAPagar = totalComissoesAPagar + comissaoTotalVenda
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
            background-color: #46F334;
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
        
        .desconto-info {
            font-size: 0.75rem;
            color: #6c757d;
        }
        
        .valor-liquido {
            font-weight: bold;
            color: var(--success);
        }
        
        .valor-desconto {
            color: var(--accent);
            font-size: 0.8rem;
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
        
        /* 🆕 NOVO ESTILO PARA CHECK AO LADO DOS VALORES PAGOS */
        .valor-pago {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        
        .check-pago {
            color: var(--success);
            font-size: 0.9em;
        }
        
        .comissao-total {
            font-weight: bold;
            color: var(--primary);
        }
    </style>


<style>
    body {
        /* Define a escala de 0.8 (80%) */
        transform: scale(0.8); 
        
        /* Define o ponto de origem para o canto superior esquerdo */
        transform-origin: 0 0; 
        
        /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
        /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
        width: calc(100% / 0.8); 
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
                    <div class="text-info font-weight-bold" style="font-size: 0.8em;">
                        <i class="fas fa-percentage mr-1"></i> R$ <%= FormatNumber(totalComissaoHtml, 2) %>
                    </div>
                    <div class="text-muted" style="font-size: 0.8em;">
                        Total Comissões
                        <br><small>(Diretoria+Gerência+Corretor)</small>
                    </div>
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
<div class="col-md-4 text-end d-flex justify-content-end align-items-center">
    
    <a href="gestao_vendas_create3.asp" class="btn btn-info btn-sm me-2">
        <i class="fas fa-plus me-1"></i>Nova
    </a>
    
    <a href="gestao_vendas_gerenc_comissoes3.asp" class="btn btn-primary btn-sm me-2" target="_blank">
        <i class="fas fa-money-bill-wave me-1"></i>Pagar
    </a>

        <a href="gestao_pessoas_vendas.asp" class="btn btn-success btn-sm me-2" target="_blank">
        <i class="fas fa-money-bill-wave me-1"></i>Destinatários
    </a>


    <a href="gestao_vendas_list_excluidos.asp" class="btn btn-warning btn-sm me-2" target="_blank">
        <i class="fas fa-trash-restore me-1"></i>Excluídos
    </a>
    
    <%if Session("Usuario")="BARRETO" Then%>
        <div class="btn-group me-2">
            <button type="button" class="btn btn-info btn-sm dropdown-toggle" data-bs-toggle="dropdown">
                <i class="fas fa-tools me-1"></i>Utilitários
            </button>
            <ul class="dropdown-menu">
                <li><a class="dropdown-item" href="inserirVendasTeste2.asp" target="_blank"><i class="fas fa-plus me-1"></i>Inserir Testes</a></li>
                <li><a class="dropdown-item" href="excluir_testes.asp" target="_blank"><i class="fas fa-trash me-1"></i>Excluir Testes</a></li>
                <li><a class="dropdown-item" href="tool_excluir_tudo.asp" target="_blank"><i class="fas fa-trash me-1"></i>Excluir Vendas</a></li>    
                <li><a class="dropdown-item" href="tool_visualizar_log.asp" target="_blank"><i class="fas fa-trash me-1"></i>Log Sistema</a></li>  
                <li><a class="dropdown-item" href="tool_venda_criar_json.asp" target="_blank"><i class="fas fa-trash me-1"></i>Gerar Json</a></li>  
                <li><a class="dropdown-item" href="gestao_vendas_inserir_comissao_todos1.asp" target="_blank"><i class="fas fa-trash me-1"></i>Inserir Todos</a></li>  
            </ul>
        </div>
    <%end if%>
    
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
                                    
                                    <th>Diretoria</th>
                                    <th>Gerência</th>
                                    <th>Corretor</th>
                                    <th>Valor (R$)</th>
                                    
                                    
                                    
                                    <th width="180">Ações</th>
                                </tr>
                            </thead>
                            <!-- ######################################### -->
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

                                        'logica de pagamentos prêmios'

                                        ' ======================== LÓGICA CORRIGIDA PARA PRÊMIOS 08 11 2025========================

                                        Dim premioPagoDiretoria, premioPagoGerencia, premioPagoCorretor
                                        premioPagoDiretoria = False
                                        premioPagoGerencia = False
                                        premioPagoCorretor = False

                                        ' Verificar se os PRÊMIOS foram pagos (TipoPagamento = 'Premiação')
                                        sqlPremios = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " AND TipoPagamento = 'Premiação'"
                                        Set rsPremios = connSales.Execute(sqlPremios)

                                        If Not rsPremios.EOF Then
                                            Do While Not rsPremios.EOF
                                                Select Case LCase(rsPremios("TipoRecebedor"))
                                                    Case "diretoria"
                                                        premioPagoDiretoria = True
                                                    Case "gerencia"
                                                        premioPagoGerencia = True
                                                    Case "corretor"
                                                        premioPagoCorretor = True
                                                End Select
                                                rsPremios.MoveNext
                                            Loop
                                        End If

                                        If Not rsPremios Is Nothing Then
                                            rsPremios.Close
                                            Set rsPremios = Nothing
                                        End If

                                        ' ======================== FIM DA LÓGICA DOS PRÊMIOS ========================

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
                                        Dim dblDescontoDiretoria, dblDescontoGerencia, dblDescontoCorretor
                                        Dim dblValorLiqDiretoria, dblValorLiqGerencia, dblValorLiqCorretor
                                        Dim dblDescontoPerc, dblDescontoBruto
                                        
                                        ' Valores brutos das comissões
                                        If Not IsNull(rs("ValorDiretoria")) And IsNumeric(rs("ValorDiretoria")) Then dblValorDiretoria = CDbl(rs("ValorDiretoria")) Else dblValorDiretoria = 0
                                        If Not IsNull(rs("ValorGerencia")) And IsNumeric(rs("ValorGerencia")) Then dblValorGerencia = CDbl(rs("ValorGerencia")) Else dblValorGerencia = 0
                                        If Not IsNull(rs("ValorCorretor")) And IsNumeric(rs("ValorCorretor")) Then dblValorCorretor = CDbl(rs("ValorCorretor")) Else dblValorCorretor = 0
                                        
                                        ' 🆕 CORREÇÃO: Calcular comissão total da venda
                                        Dim comissaoTotalVendaLinha
                                        comissaoTotalVendaLinha = dblValorDiretoria + dblValorGerencia + dblValorCorretor
                                        
                                        ' Valores de desconto
                                        If Not IsNull(rs("DescontoDiretoria")) And IsNumeric(rs("DescontoDiretoria")) Then dblDescontoDiretoria = CDbl(rs("DescontoDiretoria")) Else dblDescontoDiretoria = 0
                                        If Not IsNull(rs("DescontoGerencia")) And IsNumeric(rs("DescontoGerencia")) Then dblDescontoGerencia = CDbl(rs("DescontoGerencia")) Else dblDescontoGerencia = 0
                                        If Not IsNull(rs("DescontoCorretor")) And IsNumeric(rs("DescontoCorretor")) Then dblDescontoCorretor = CDbl(rs("DescontoCorretor")) Else dblDescontoCorretor = 0
                                        
                                        ' Valores líquidos
                                        If Not IsNull(rs("ValorLiqDiretoria")) And IsNumeric(rs("ValorLiqDiretoria")) Then dblValorLiqDiretoria = CDbl(rs("ValorLiqDiretoria")) Else dblValorLiqDiretoria = dblValorDiretoria - dblDescontoDiretoria
                                        If Not IsNull(rs("ValorLiqGerencia")) And IsNumeric(rs("ValorLiqGerencia")) Then dblValorLiqGerencia = CDbl(rs("ValorLiqGerencia")) Else dblValorLiqGerencia = dblValorGerencia - dblDescontoGerencia
                                        If Not IsNull(rs("ValorLiqCorretor")) And IsNumeric(rs("ValorLiqCorretor")) Then dblValorLiqCorretor = CDbl(rs("ValorLiqCorretor")) Else dblValorLiqCorretor = dblValorCorretor - dblDescontoCorretor
                                        
                                        ' Percentual e valor total do desconto
                                        If Not IsNull(rs("DescontoPerc")) And IsNumeric(rs("DescontoPerc")) Then dblDescontoPerc = CDbl(rs("DescontoPerc")) Else dblDescontoPerc = 0
                                        If Not IsNull(rs("DescontoBruto")) And IsNumeric(rs("DescontoBruto")) Then dblDescontoBruto = CDbl(rs("DescontoBruto")) Else dblDescontoBruto = dblDescontoDiretoria + dblDescontoGerencia + dblDescontoCorretor
                                        
                                            If dblValorLiqDiretoria > 0 And totalPagoDiretoria >= dblValorLiqDiretoria Then pagoDiretoria = True
                                            If dblValorLiqDiretoria = 0 Then pagoDiretoria = True

                                            ' Aplicar a mesma correção para Gerência e Corretor:
                                            If dblValorLiqGerencia > 0 And totalPagoGerencia >= dblValorLiqGerencia Then pagoGerencia = True
                                            If dblValorLiqGerencia = 0 Then pagoGerencia = True

                                            If dblValorLiqCorretor > 0 And totalPagoCorretor >= dblValorLiqCorretor Then pagoCorretor = True
                                            If dblValorLiqCorretor = 0 Then pagoCorretor = True

                                        Dim comissaoText
                                        comissaoText = FormatNumber(rs("ComissaoPercentual"), 2) & "%"
                                        ' 🆕 CORREÇÃO: Usar o cálculo correto da comissão total
                                        If comissaoTotalVendaLinha > 0 Then
                                            comissaoText = comissaoText & " (R$ " & FormatNumber(comissaoTotalVendaLinha, 2) & ")"
                                        End If

                                        vAno = Right(rs("AnoVenda"), 2)

                                        Dim rsComissaoCheck, comissaoExiste
                                        Set rsComissaoCheck = Server.CreateObject("ADODB.Recordset")
                                        rsComissaoCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(rs("ID")), connSales
                                        comissaoExiste = Not rsComissaoCheck.EOF
                                        rsComissaoCheck.Close
                                        Set rsComissaoCheck = Nothing
                                        
                                        Dim linhaClasse
                                        If totalPagoDiretoria>0 And totalPagoGerencia>0 And totalPagoCorretor>0 Then
                                            linhaClasse = "linha-paga"
                                        Else
                                            linhaClasse = "linha-pendente"
                                        End If
                                %>
                                <tr class="<%= linhaClasse %>">
                                <td style="font-size: 12px;">
                                    <small>
                                        DTV: <%= rs("AnoVenda") & "-" & Right("0"&rs("MesVenda"),2) & "-" & Right("0"&rs("DiaVenda"),2) %>
                                        <br><span class="fw-bold">ID: <%= rs("ID") %><br></span>
                                        
                                        <%= FormatDateTime(rs("DataRegistro"),2) %>
                                           <% If Trim(rs("Usuario"))="" Then
                                               vUsuario = "SILVIOBF" 
                                            Else
                                               vUsuario = rs("Usuario")
                                            End if   
                                        %>
                                        <br>por <%= vUsuario %>
                                    </small>
                                </td>
                                    <td>
                                        <% If totalPagoDiretoria>0  And totalPagoGerencia > 0 And totalPagoCorretor >0Then %>
                                            <span class="badge status-pago" title="Comissões pagas">PAGA</span>
                                        <% Else %>  
                                            <span class="badge status-pendente" title="Comissões pendentes">PENDENTE</span>
                                        <% End If %>
                                    </td>  
                                    <td>
                                        <small class="text-muted"><%= vAno & "T" & rs("Trimestre") %></small>
                                    </td>
                                    
                                    <td>
                                        <strong><%= rs("Empreend_ID") %>-<%= RemoverNumeros(rs("NomeEmpreendimento")) %></strong>
                                        <br><small class="text-muted"><%= RemoverNumeros(rs("Localidade")) %></small><br>Cliente:
                                        <small class="text-muted"><strong><%= UCase(rs("NomeCliente")) %></strong></small>
                                        <small class="text-muted"><strong><br>Unid: <%= UCase(rs("Unidade")) %></strong></small>
                                        
                                    </td>
                                    
                                    <!-- ############################ -->
                                    <% ' --------------------------------------- %>
<% ' COLUNA DIRETORIA: Comissao + Prêmio + Desconto + Líquido %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Diretoria") %></div>
    <div class="fw-bold"><%= rs("NomeDiretor") %></div>
    <small class="comissao-info">
        <!-- Valor Bruto com badge PAGA se pago -->
        <div class="valor-pago d-flex align-items-center gap-2 mb-1">
            R$ <%= FormatNumber(dblValorDiretoria, 2) %>
        </div>
        
        <!-- Desconto -->
        <% If dblDescontoDiretoria > 0 Then %>
            <div class="valor-desconto">
                <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoDiretoria, 2) %>
            </div>
        <% End If %>
        
        <!-- Valor Líquido -->
        <div class="valor-liquido d-flex align-items-center gap-2">
            <% If pagoDiretoria AND totalPagoDiretoria >0 Then %>
                <span class="badge bg-success badge-sm">PAGA</span>
            <% End If %>
            <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqDiretoria, 2) %>
        </div>
        
        <% ' Verificação e Exibição do PRÊMIO DIRETORIA %>
        <% If Not IsNull(rs("premioDiretoria")) Then %>
            <% If IsNumeric(rs("premioDiretoria")) And CDbl(rs("premioDiretoria")) > 0 Then %>
                <div class="text-primary fw-bold mt-1">
                    <div class="valor-pago d-flex align-items-center gap-2">
                        <% If premioPagoDiretoria Then %>
                            <span class="badge bg-success badge-sm">PAGA</span>
                        <% End If %>
                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioDiretoria"), 2) %>
                    </div>
                </div>
            <% End If %>
        <% End If %>
    </small>
</td>

<% ' --------------------------------------- %>
<% ' COLUNA GERÊNCIA: Comissao + Prêmio + Desconto + Líquido %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Gerencia") %></div>
    <div class="fw-bold"><%= rs("NomeGerente") %></div>
    <small class="comissao-info">
        <!-- Valor Bruto com badge PAGA se pago -->
        <div class="valor-pago d-flex align-items-center gap-2 mb-1">

            R$ <%= FormatNumber(dblValorGerencia, 2) %>
        </div>
        
        <!-- Desconto -->
        <% If dblDescontoGerencia > 0 Then %>
            <div class="valor-desconto">
                <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoGerencia, 2) %>
            </div>
        <% End If %>
        
        <!-- Valor Líquido Gerência-->
        <div class="valor-liquido d-flex align-items-center gap-2">
            <% If pagoGerencia AND totalPagoGerencia >0 Then %>
           <span class="badge bg-success badge-sm">PAGA</span>
            <% End If %>
           <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqGerencia, 2) %>
        </div>
        
        <% ' Verificação e Exibição do PRÊMIO GERÊNCIA %>
        <% If Not IsNull(rs("premioGerencia")) Then %>
            <% If IsNumeric(rs("premioGerencia")) And CDbl(rs("premioGerencia")) > 0 Then %>
                <div class="text-primary fw-bold mt-1">
                    <div class="valor-pago d-flex align-items-center gap-2">
                        <% If premioPagoGerencia > 0 Then  %>
                            <span class="badge bg-success badge-sm">PAGA</span>
                        <% End If %>
                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioGerencia"), 2) %>
                    </div>
                </div>
            <% End If %>
        <% End If %>
    </small>
</td>

<% ' --------------------------------------- %>
<% ' COLUNA CORRETOR: Comissao + Prêmio + Desconto + Líquido %>
<% ' --------------------------------------- %>
<td>
    <div class="fw-bold"><%= rs("Corretor") %></div>
    <small class="comissao-info">
        <!-- Valor Bruto com badge PAGA se pago -->
        <div class="valor-pago d-flex align-items-center gap-2 mb-1">

            R$ <%= FormatNumber(dblValorCorretor, 2) %>
        </div>
        
        <!-- Desconto -->
        <% If dblDescontoCorretor > 0 Then %>
            <div class="valor-desconto">
                <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoCorretor, 2) %>
            </div>
        <% End If %>
        
        <!-- Valor Líquido -->
        <div class="valor-liquido d-flex align-items-center gap-2">
            <% If pagoCorretor AND totalPagoCorretor > 0 Then %>
                <span class="badge bg-success badge-sm">PAGA</span>
            <% End If %>
            <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqCorretor, 2) %>
        </div>
        
        <% ' Verificação e Exibição do PRÊMIO CORRETOR %>
        <% If Not IsNull(rs("premioCorretor")) Then %>
            <% If IsNumeric(rs("premioCorretor")) And CDbl(rs("premioCorretor")) > 0 Then %>
                <div class="text-primary fw-bold mt-1">
                    <div class="valor-pago d-flex align-items-center gap-2">

                        <% If premioPagoCorretor Then %>
                            <span class="badge bg-success badge-sm">PAGA</span>
                        <% End If %>
                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(rs("premioCorretor"), 2) %>
                    </div>
                </div>
            <% End If %>
        <% End If %>
    </small>
</td>
                                    <!-- ############################ -->
                                    
                                    <td class="text-end fw-bold" data-order="<%= rs("ValorUnidade") %>">
                                        <%= FormatNumber(rs("ValorUnidade"), 2) %>
                                         <span class="badge badge-comissao comissao-total">
                                            <%= comissaoText %>
                                        </span>
                                        <% If dblDescontoPerc > 0 Then 
                                               %>

                                            <div class="desconto-info">
                                                <strong><%= FormatNumber(dblDescontoPerc, 2) %>%</strong>
                                                <br>
                                                <small>Total: R$ <%= FormatNumber(dblDescontoBruto, 2) %></small>
                                                <% If Not IsNull(rs("DescontoDescricao")) And rs("DescontoDescricao") <> "" Then %>
                                                    <br>
                                                    <small title="<%= rs("DescontoDescricao") %>">
                                                        <i class="fas fa-info-circle"></i>
                                                    </small>
                                                <% End If %>
                                            </div>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </td>


                                    <td class="text-center">
                                        <div class="action-buttons">
                                            <a href="gestao_vendas_update3.asp?id=<%= rs("id") %>" class="btn btn-warning btn-sm" title="Editar">
                                                <i class="fas fa-edit"></i>
                                            </a>


                                           <!-- ------------------------ -->
                                            <a href="gestao_vendas_pagar_todos3.asp?id=<%= rs("id") %>" class="btn btn-success btn-sm" title="Pagar" target="_blank">
                                                    <i class="fas fa-dollar"></i>
                                            </a>

                                            <a href="gestao_vendas_ver_pagamentos.asp?id=<%= rs("id") %>" class="btn btn-primary btn-sm" title="Ver Pagamentos" target="_blank">
                                                    <i class="fas fa-eye"></i>
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
                            <!-- ##################################### -->
                            <tfoot>
                                <tr class="table-light">
                                    <th colspan="6" class="text-end">Totais:</th>
                                    <th id="totalValor" class="text-end">R$ <%= FormatNumber(totalValorHtml, 2) %></th>
                                    <th id="totalComissao" class="text-end">R$ <%= FormatNumber(totalComissaoHtml, 2) %></th>
                                    <th colspan="3"></th>
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
            <small>
    <%
        Response.Write Session("Usuario") & " " & Date() & " " & Time()
    %>

   </small>
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
                    pageLength: 50,
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