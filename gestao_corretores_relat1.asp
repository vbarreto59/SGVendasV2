<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar ao banco de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

' ===============================================
' OBTER PARÂMETROS
' ===============================================

Dim corretorId, corretorNome, filtroAno
corretorId = Request.QueryString("corretorid")
corretorNome = Request.QueryString("corretor")
filtroAno = Request.QueryString("ano")

If corretorId = "" Then
    Response.Write "<div class='alert alert-danger'>Nenhum corretor especificado.</div>"
    Response.End
End If

' Se ano não foi especificado, usar ano atual
If filtroAno = "" Then
    filtroAno = Year(Date())
End If

' ===============================================
' OBTER INFORMAÇÕES BÁSICAS DO CORRETOR
' ===============================================

Dim sqlInfoCorretor, rsInfoCorretor, nomeCorretor, diretoriaCorretor, gerenciaCorretor

' Buscar informações do corretor usando CorretorId numérico
sqlInfoCorretor = "SELECT TOP 1 Corretor, Diretoria, Gerencia FROM Vendas WHERE CorretorId = " & corretorId & " AND Excluido = 0"

Set rsInfoCorretor = Server.CreateObject("ADODB.Recordset")
rsInfoCorretor.Open sqlInfoCorretor, connSales

If Not rsInfoCorretor.EOF Then
    nomeCorretor = rsInfoCorretor("Corretor")
    diretoriaCorretor = rsInfoCorretor("Diretoria")
    gerenciaCorretor = rsInfoCorretor("Gerencia")
Else
    ' Se não encontrou, usar o nome passado como parâmetro
    nomeCorretor = corretorNome
    diretoriaCorretor = "N/A"
    gerenciaCorretor = "N/A"
End If

If rsInfoCorretor.State = 1 Then rsInfoCorretor.Close
Set rsInfoCorretor = Nothing

' ===============================================
' OBTER DADOS DETALHADOS DO CORRETOR
' ===============================================

' Consulta para obter vendas detalhadas do corretor usando CorretorId numérico
Dim sqlVendas, rsVendas
sqlVendas = "SELECT * FROM Vendas WHERE CorretorId = " & corretorId & " AND Excluido = 0 ORDER BY DataVenda DESC"

Set rsVendas = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rsVendas.Open sqlVendas, connSales

If Err.Number <> 0 Then
    Response.Write "Erro na consulta de vendas: " & Err.Description & "<br>"
    Response.Write "SQL: " & Server.HTMLEncode(sqlVendas)
    Response.End
End If
On Error GoTo 0

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Jan"
arrMesesNome(2) = "Fev"
arrMesesNome(3) = "Mar"
arrMesesNome(4) = "Abr"
arrMesesNome(5) = "Mai"
arrMesesNome(6) = "Jun"
arrMesesNome(7) = "Jul"
arrMesesNome(8) = "Ago"
arrMesesNome(9) = "Set"
arrMesesNome(10) = "Out"
arrMesesNome(11) = "Nov"
arrMesesNome(12) = "Dez"

' Função para remover números e asteriscos
Function RemoverNumeros(texto)
    If IsNull(texto) Then
        RemoverNumeros = ""
        Exit Function
    End If
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    Dim textoLimpo
    textoLimpo = regex.Replace(texto, "")
    RemoverNumeros = Trim(Replace(textoLimpo, "  ", " "))
End Function

' ===============================================
' CÁLCULOS DE TOTAIS E ESTATÍSTICAS
' ===============================================

' Variáveis para cálculos totais
Dim totalValor, totalComissaoGeral, totalVendas
Dim totalComissoesPagas, totalComissoesAPagar
Dim totalVendasPagas, totalVendasPendentes
Dim totalValorCorretor, totalComissaoCorretor
Dim primeiroVenda, ultimaVenda, periodoAtuacao
Dim empreendimentosTrabalhados

totalValor = 0
totalComissaoGeral = 0
totalVendas = 0
totalComissoesPagas = 0
totalComissoesAPagar = 0
totalVendasPagas = 0
totalVendasPendentes = 0
totalValorCorretor = 0
totalComissaoCorretor = 0

' Dicionário para contar empreendimentos distintos
Set empreendimentosTrabalhados = Server.CreateObject("Scripting.Dictionary")

' Calcular totais e estatísticas
If Not rsVendas.EOF Then
    rsVendas.MoveFirst
    
    ' Inicializar datas
    primeiroVenda = rsVendas("DataVenda")
    ultimaVenda = rsVendas("DataVenda")
    
    Do While Not rsVendas.EOF
        totalValor = totalValor + CDbl(rsVendas("ValorUnidade"))
        totalComissaoGeral = totalComissaoGeral + CDbl(rsVendas("ValorComissaoGeral"))
        totalValorCorretor = totalValorCorretor + CDbl(rsVendas("ValorCorretor"))
        totalComissaoCorretor = totalComissaoCorretor + CDbl(rsVendas("ValorCorretor"))
        totalVendas = totalVendas + 1
        
        ' Atualizar datas
        If rsVendas("DataVenda") < primeiroVenda Then primeiroVenda = rsVendas("DataVenda")
        If rsVendas("DataVenda") > ultimaVenda Then ultimaVenda = rsVendas("DataVenda")
        
        ' Contar empreendimentos distintos
        Dim empreendimentoKey
        empreendimentoKey = rsVendas("Empreend_ID") & "|" & rsVendas("NomeEmpreendimento")
        If Not empreendimentosTrabalhados.Exists(empreendimentoKey) Then
            empreendimentosTrabalhados.Add empreendimentoKey, 1
        End If
        
        ' Verificar status de pagamento para totais
        Dim sqlPagamentosTotais, rsPagamentosTotais
        Dim totalPagoDiretoriaTotais, totalPagoGerenciaTotais, totalPagoCorretorTotais
        Dim pagoDiretoriaTotais, pagoGerenciaTotais, pagoCorretorTotais
        
        totalPagoDiretoriaTotais = 0
        totalPagoGerenciaTotais = 0
        totalPagoCorretorTotais = 0
        pagoDiretoriaTotais = False
        pagoGerenciaTotais = False
        pagoCorretorTotais = False

        sqlPagamentosTotais = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rsVendas("ID") & " ORDER BY DataPagamento ASC;"
        Set rsPagamentosTotais = connSales.Execute(sqlPagamentosTotais)

        If Not rsPagamentosTotais.EOF Then
            Do While Not rsPagamentosTotais.EOF
                Select Case LCase(rsPagamentosTotais("TipoRecebedor"))
                    Case "diretoria"
                        totalPagoDiretoriaTotais = totalPagoDiretoriaTotais + CDbl(rsPagamentosTotais("ValorPago"))
                    Case "gerencia"
                        totalPagoGerenciaTotais = totalPagoGerenciaTotais + CDbl(rsPagamentosTotais("ValorPago"))
                    Case "corretor"
                        totalPagoCorretorTotais = totalPagoCorretorTotais + CDbl(rsPagamentosTotais("ValorPago"))
                End Select
                rsPagamentosTotais.MoveNext
            Loop
        End If
        
        If Not rsPagamentosTotais Is Nothing Then
            rsPagamentosTotais.Close
            Set rsPagamentosTotais = Nothing
        End If

        If rsVendas("ValorDiretoria") > 0 And totalPagoDiretoriaTotais >= CDbl(rsVendas("ValorDiretoria")) Then pagoDiretoriaTotais = True
        If rsVendas("ValorDiretoria") = 0 Then pagoDiretoriaTotais = True
        If rsVendas("ValorGerencia") > 0 And totalPagoGerenciaTotais >= CDbl(rsVendas("ValorGerencia")) Then pagoGerenciaTotais = True
        If rsVendas("ValorGerencia") = 0 Then pagoGerenciaTotais = True
        If rsVendas("ValorCorretor") > 0 And totalPagoCorretorTotais >= CDbl(rsVendas("ValorCorretor")) Then pagoCorretorTotais = True
        If rsVendas("ValorCorretor") = 0 Then pagoCorretorTotais = True

        ' Acumular totais para KPIs
        If pagoDiretoriaTotais And pagoGerenciaTotais And pagoCorretorTotais Then
            totalComissoesPagas = totalComissoesPagas + CDbl(rsVendas("ValorComissaoGeral"))
            totalVendasPagas = totalVendasPagas + 1
        Else
            totalComissoesAPagar = totalComissoesAPagar + CDbl(rsVendas("ValorComissaoGeral"))
            totalVendasPendentes = totalVendasPendentes + 1
        End If
        
        rsVendas.MoveNext
    Loop
    rsVendas.MoveFirst
End If

' Calcular período de atuação
If totalVendas > 0 Then
    periodoAtuacao = DateDiff("d", primeiroVenda, ultimaVenda) & " dias"
Else
    periodoAtuacao = "N/A"
End If

' Calcular percentuais
Dim percentualPagas, percentualAPagar
If totalComissaoGeral > 0 Then
    percentualPagas = (totalComissoesPagas / totalComissaoGeral) * 100
    percentualAPagar = (totalComissoesAPagar / totalComissaoGeral) * 100
Else
    percentualPagas = 0
    percentualAPagar = 0
End If

' Calcular ticket médio
Dim ticketMedioGeral, ticketMedioCorretor
If totalVendas > 0 Then
    ticketMedioGeral = totalValor / totalVendas
    ticketMedioCorretor = totalValorCorretor / totalVendas
Else
    ticketMedioGeral = 0
    ticketMedioCorretor = 0
End If

' Calcular média de vendas por mês
Dim mediaVendasMes
If totalVendas > 0 And periodoAtuacao <> "N/A" Then
    Dim totalMeses
    totalMeses = DateDiff("m", primeiroVenda, ultimaVenda) + 1
    If totalMeses > 0 Then
        mediaVendasMes = FormatNumber(totalVendas / totalMeses, 1)
    Else
        mediaVendasMes = totalVendas
    End If
Else
    mediaVendasMes = 0
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detalhes do Corretor - <%= nomeCorretor %></title>
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
            padding-top: 20px;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1rem 0;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            margin-bottom: 2rem;
            border-radius: 10px;
        }
        
        .app-title {
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.5rem;
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
            transform: translateY(-2px);
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
            font-size: 1.8rem;
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
        
        .corretor-info {
            background: white;
            border-radius: 10px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            box-shadow: var(--card-shadow);
        }
        
        .back-button {
            background-color: #6c757d;
            border-color: #6c757d;
            color: white;
        }
        
        .back-button:hover {
            background-color: #5a6268;
            border-color: #545b62;
            color: white;
        }
        
        .corretor-id {
            font-size: 0.8rem;
            color: #6c757d;
            background-color: #f8f9fa;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            margin-left: 0.5rem;
        }
        
        .data-pagamento {
            font-size: 0.7rem;
            color: #28a745;
            font-style: italic;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <!-- Cabeçalho -->
        <div class="app-header">
            <div class="container-fluid">
                <div class="row align-items-center">
                    <div class="col-md-8">
                        <h1 class="app-title">
                            <i class="fas fa-user-tie"></i> Detalhes do Corretor: <%= nomeCorretor %>
                            <span class="corretor-id">ID: <%= corretorId %></span>
                        </h1>
                    </div>
                    <div class="col-md-4 text-end">
                        <button type="button" onclick="window.history.back();" class="btn btn-light btn-sm">
                            <i class="fas fa-arrow-left me-1"></i>Voltar
                        </button>
                        <button type="button" onclick="window.print();" class="btn btn-light btn-sm">
                            <i class="fas fa-print me-1"></i>Imprimir
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- Informações do Corretor -->
        <div class="corretor-info">
            <div class="row">
                <div class="col-md-12">
                    <h3 class="text-primary mb-3"><i class="fas fa-chart-line me-2"></i>Resumo Geral - Corretor ID: <%= corretorId %></h3>
                    <div class="row">
                        <div class="col-md-3">
                            <p><strong>Nome:</strong> <%= nomeCorretor %></p>
                            <p><strong>Diretoria:</strong> <%= diretoriaCorretor %></p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>Gerência:</strong> <%= gerenciaCorretor %></p>
                            <p><strong>Período de Atuação:</strong> <%= periodoAtuacao %></p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>Primeira Venda:</strong> <%= FormatDateTime(primeiroVenda, 2) %></p>
                            <p><strong>Última Venda:</strong> <%= FormatDateTime(ultimaVenda, 2) %></p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>Empreendimentos:</strong> <%= empreendimentosTrabalhados.Count %></p>
                            <p><strong>Média/Mês:</strong> <%= mediaVendasMes %> vendas</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- KPIs Principais -->
        <div class="row mb-4">
            <div class="col-md-2">
                <div class="card stats-card">
                    <div class="stats-icon text-primary">
                        <i class="fas fa-shopping-cart"></i>
                    </div>
                    <div class="stats-value text-primary"><%= totalVendas %></div>
                    <div class="stats-label">Total de Vendas</div>
                </div>
            </div>
            <div class="col-md-2">
                <div class="card stats-card">
                    <div class="stats-icon text-success">
                        <i class="fas fa-money-bill-wave"></i>
                    </div>
                    <div class="stats-value text-success">R$ <%= FormatNumber(totalValor, 2) %></div>
                    <div class="stats-label">VGV Total</div>
                </div>
            </div>
            <div class="col-md-2">
                <div class="card stats-card">
                    <div class="stats-icon text-info">
                        <i class="fas fa-percentage"></i>
                    </div>
                    <div class="stats-value text-info">R$ <%= FormatNumber(totalComissaoCorretor, 2) %></div>
                    <div class="stats-label">Comissão do Corretor</div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card stats-card">
                    <div class="stats-icon text-success">
                        <i class="fas fa-check-circle"></i>
                    </div>
                    <div class="stats-value text-success">R$ <%= FormatNumber(totalComissoesPagas, 2) %></div>
                    <div class="stats-label">Comissões Pagas</div>
                    <div class="stats-percent percent-success">
                        <%= FormatNumber(percentualPagas, 1) %>%
                    </div>
                    <div class="kpi-progress">
                        <div class="progress-bar-pagas h-100" style="width: <%= percentualPagas %>%"></div>
                    </div>
                    <div class="stats-label mt-1">
                        <small><%= totalVendasPagas %> vendas quitadas</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card stats-card">
                    <div class="stats-icon text-warning">
                        <i class="fas fa-clock"></i>
                    </div>
                    <div class="stats-value text-warning">R$ <%= FormatNumber(totalComissoesAPagar, 2) %></div>
                    <div class="stats-label">Comissões a Pagar</div>
                    <div class="stats-percent percent-warning">
                        <%= FormatNumber(percentualAPagar, 1) %>%
                    </div>
                    <div class="kpi-progress">
                        <div class="progress-bar-apagar h-100" style="width: <%= percentualAPagar %>%"></div>
                    </div>
                    <div class="stats-label mt-1">
                        <small><%= totalVendasPendentes %> vendas pendentes</small>
                    </div>
                </div>
            </div>
        </div>

        <!-- KPIs Secundários -->
        <div class="row mb-4">
            <div class="col-md-4">
                <div class="card stats-card">
                    <div class="stats-icon text-secondary">
                        <i class="fas fa-chart-bar"></i>
                    </div>
                    <div class="stats-value text-secondary">R$ <%= FormatNumber(ticketMedioGeral, 2) %></div>
                    <div class="stats-label">Ticket Médio (VGV)</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card stats-card">
                    <div class="stats-icon text-secondary">
                        <i class="fas fa-user-tie"></i>
                    </div>
                    <div class="stats-value text-secondary">R$ <%= FormatNumber(ticketMedioCorretor, 2) %></div>
                    <div class="stats-label">Ticket Médio (Corretor)</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card stats-card">
                    <div class="stats-icon text-secondary">
                        <i class="fas fa-building"></i>
                    </div>
                    <div class="stats-value text-secondary"><%= empreendimentosTrabalhados.Count %></div>
                    <div class="stats-label">Empreendimentos</div>
                </div>
            </div>
        </div>

        <!-- Tabela de Vendas Detalhadas -->
        <div class="card">
            <div class="card-header">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0"><i class="fas fa-list me-2"></i>Vendas Detalhadas - Corretor ID: <%= corretorId %></h5>
                    <div>
                        <span class="badge bg-success me-2"><i class="fas fa-check me-1"></i><%= totalVendasPagas %> Pagas</span>
                        <span class="badge bg-warning me-2"><i class="fas fa-clock me-1"></i><%= totalVendasPendentes %> Pendentes</span>
                        <span class="badge bg-light text-dark"><%= totalVendas %> Total</span>
                    </div>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table id="tabelaVendas" class="table table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Status</th>
                                <th>Período</th>
                                <th>Data Venda</th>
                                <th>Empreendimento</th>
                                <th>Unidade</th>
                                <th>Diretoria</th>
                                <th>Gerência</th>
                                <th>Corretor</th>
                                <th>Valor (R$)</th>
                                <th>Comissão</th>
                                <th>Registro</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsVendas.EOF Then
                                Do While Not rsVendas.EOF
                                    
                                    ' Variáveis locais para cada venda (sem conflito com as variáveis globais)
                                    Dim sqlPagamentosVenda, rsPagamentosVenda
                                    Dim totalPagoDiretoriaVenda, totalPagoGerenciaVenda, totalPagoCorretorVenda
                                    Dim dataPagamentoDiretoriaVenda, dataPagamentoGerenciaVenda, dataPagamentoCorretorVenda
                                    Dim tooltipDiretoriaVenda, tooltipGerenciaVenda, tooltipCorretorVenda
                                    Dim pagoDiretoriaVenda, pagoGerenciaVenda, pagoCorretorVenda
                                    
                                    totalPagoDiretoriaVenda = 0
                                    totalPagoGerenciaVenda = 0
                                    totalPagoCorretorVenda = 0
                                    dataPagamentoDiretoriaVenda = ""
                                    dataPagamentoGerenciaVenda = ""
                                    dataPagamentoCorretorVenda = ""
                                    tooltipDiretoriaVenda = ""
                                    tooltipGerenciaVenda = ""
                                    tooltipCorretorVenda = ""
                                    pagoDiretoriaVenda = False
                                    pagoGerenciaVenda = False
                                    pagoCorretorVenda = False

                                    sqlPagamentosVenda = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rsVendas("ID") & " ORDER BY DataPagamento ASC;"
                                    Set rsPagamentosVenda = connSales.Execute(sqlPagamentosVenda)

                                    If Not rsPagamentosVenda.EOF Then
                                        Do While Not rsPagamentosVenda.EOF
                                            Dim detalhePagamentoVenda
                                            detalhePagamentoVenda = "Data: " & FormatDateTime(rsPagamentosVenda("DataPagamento"), 2) & " | Valor: R$ " & FormatNumber(rsPagamentosVenda("ValorPago"), 2) & " | Status: " & rsPagamentosVenda("Status")
                                            Select Case LCase(rsPagamentosVenda("TipoRecebedor"))
                                                Case "diretoria"
                                                    If tooltipDiretoriaVenda <> "" Then tooltipDiretoriaVenda = tooltipDiretoriaVenda & Chr(13)
                                                    tooltipDiretoriaVenda = tooltipDiretoriaVenda & detalhePagamentoVenda
                                                    totalPagoDiretoriaVenda = totalPagoDiretoriaVenda + CDbl(rsPagamentosVenda("ValorPago"))
                                                    dataPagamentoDiretoriaVenda = FormatDateTime(rsPagamentosVenda("DataPagamento"), 2)
                                                Case "gerencia"
                                                    If tooltipGerenciaVenda <> "" Then tooltipGerenciaVenda = tooltipGerenciaVenda & Chr(13)
                                                    tooltipGerenciaVenda = tooltipGerenciaVenda & detalhePagamentoVenda
                                                    totalPagoGerenciaVenda = totalPagoGerenciaVenda + CDbl(rsPagamentosVenda("ValorPago"))
                                                    dataPagamentoGerenciaVenda = FormatDateTime(rsPagamentosVenda("DataPagamento"), 2)
                                                Case "corretor"
                                                    If tooltipCorretorVenda <> "" Then tooltipCorretorVenda = tooltipCorretorVenda & Chr(13)
                                                    tooltipCorretorVenda = tooltipCorretorVenda & detalhePagamentoVenda
                                                    totalPagoCorretorVenda = totalPagoCorretorVenda + CDbl(rsPagamentosVenda("ValorPago"))
                                                    dataPagamentoCorretorVenda = FormatDateTime(rsPagamentosVenda("DataPagamento"), 2)
                                            End Select
                                            rsPagamentosVenda.MoveNext
                                        Loop
                                    End If
                                    
                                    If Not rsPagamentosVenda Is Nothing Then
                                        rsPagamentosVenda.Close
                                        Set rsPagamentosVenda = Nothing
                                    End If

                                    If rsVendas("ValorDiretoria") > 0 And totalPagoDiretoriaVenda >= CDbl(rsVendas("ValorDiretoria")) Then pagoDiretoriaVenda = True
                                    If rsVendas("ValorDiretoria") = 0 Then pagoDiretoriaVenda = True
                                    If rsVendas("ValorGerencia") > 0 And totalPagoGerenciaVenda >= CDbl(rsVendas("ValorGerencia")) Then pagoGerenciaVenda = True
                                    If rsVendas("ValorGerencia") = 0 Then pagoGerenciaVenda = True
                                    If rsVendas("ValorCorretor") > 0 And totalPagoCorretorVenda >= CDbl(rsVendas("ValorCorretor")) Then pagoCorretorVenda = True
                                    If rsVendas("ValorCorretor") = 0 Then pagoCorretorVenda = True

                                    Dim comissaoText
                                    comissaoText = FormatNumber(rsVendas("ComissaoPercentual"), 2) & "%"
                                    If rsVendas("ValorComissaoGeral") > 0 Then
                                        comissaoText = comissaoText & " (R$ " & FormatNumber(rsVendas("ValorComissaoGeral"), 2) & ")"
                                    End If

                                    Dim vAno
                                    vAno = Right(rsVendas("AnoVenda"), 2)

                                    ' Verifica se a venda já possui comissão cadastrada
                                    Dim rsComissaoCheck, comissaoExiste
                                    Set rsComissaoCheck = Server.CreateObject("ADODB.Recordset")
                                    rsComissaoCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(rsVendas("ID")), connSales
                                    comissaoExiste = Not rsComissaoCheck.EOF
                                    rsComissaoCheck.Close
                                    Set rsComissaoCheck = Nothing
                                    
                                    ' Definir classe CSS baseada no status de pagamento
                                    Dim linhaClasse
                                    If pagoDiretoriaVenda And pagoGerenciaVenda And pagoCorretorVenda Then
                                        linhaClasse = "linha-paga"
                                    Else
                                        linhaClasse = "linha-pendente"
                                    End If
                            %>
                            <tr class="<%= linhaClasse %>">
                                <td>
                                    <span class="fw-bold"><%= rsVendas("ID") %></span>
                                </td>
                                <td>
                                    <% If pagoDiretoriaVenda And pagoGerenciaVenda And pagoCorretorVenda Then %>
                                        <span class="badge status-pago" title="Comissões pagas">PAGO</span>
                                     <% Else %>   
                                        <span class="badge status-pendente" title="Comissões pendentes">PENDENTE</span>
                                    <% End If %>
                                </td>    
                                <td>
                                    <%= rsVendas("AnoVenda") & "-" & Right("0"&rsVendas("MesVenda"),2) %>
                                    <br><small class="text-muted"><%= vAno & "T" & rsVendas("Trimestre") %></small>
                                </td>
                                <td><%= FormatDateTime(rsVendas("DataVenda"), 2) %></td>
                                <td>
                                    <strong><%= rsVendas("Empreend_ID") %>-<%= RemoverNumeros(rsVendas("NomeEmpreendimento")) %></strong>
                                    <br><small class="text-muted"><%= RemoverNumeros(rsVendas("Localidade")) %></small>
                                </td>
                                <td><%= rsVendas("Unidade") %></td>
                                <td>
                                    <div class="fw-bold"><%= rsVendas("Diretoria") %></div>
                                    <small class="comissao-info">
                                        <%= rsVendas("ComissaoDiretoria") %>% - R$ <%= FormatNumber(rsVendas("ValorDiretoria"), 2) %>
                                        <% If pagoDiretoriaVenda Then %>
                                            <span class="badge bg-success">PAGO</span>
                                            <% If dataPagamentoDiretoriaVenda <> "" Then %>
                                                <br><span class="data-pagamento">Em: <%= dataPagamentoDiretoriaVenda %></span>
                                            <% End If %>
                                        <% End If %>
                                    </small>
                                </td>
                                <td>
                                    <div class="fw-bold"><%= rsVendas("Gerencia") %></div>
                                    <small class="comissao-info">
                                        <%= rsVendas("ComissaoGerencia") %>% - R$ <%= FormatNumber(rsVendas("ValorGerencia"), 2) %>
                                        <% If pagoGerenciaVenda Then %>
                                            <span class="badge bg-success">PAGO</span>
                                            <% If dataPagamentoGerenciaVenda <> "" Then %>
                                                <br><span class="data-pagamento">Em: <%= dataPagamentoGerenciaVenda %></span>
                                            <% End If %>
                                        <% End If %>
                                    </small>
                                </td>
                                <td>
                                    <div class="fw-bold"><%= rsVendas("Corretor") %></div>
                                    <small class="comissao-info">
                                        <%= rsVendas("ComissaoCorretor") %>% - R$ <%= FormatNumber(rsVendas("ValorCorretor"), 2) %>
                                        <% If pagoCorretorVenda Then %>
                                            <span class="badge bg-success">PAGO</span>
                                            <% If dataPagamentoCorretorVenda <> "" Then %>
                                                <br><span class="data-pagamento">Em: <%= dataPagamentoCorretorVenda %></span>
                                            <% End If %>
                                        <% End If %>
                                    </small>
                                </td>
                                <td class="text-end fw-bold" data-order="<%= rsVendas("ValorUnidade") %>">
                                    <%= FormatNumber(rsVendas("ValorUnidade"), 2) %>
                                </td>
                                <td class="text-end">
                                    <span class="badge badge-comissao"><%= comissaoText %></span>
                                </td>
                                <td>
                                    <small>
                                        <%= FormatDateTime(rsVendas("DataRegistro"),2) %>
                                        <br>por <%= rsVendas("Usuario") %>
                                    </small>
                                </td>
                            </tr>
                            <%
                                    rsVendas.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="12" class="text-center text-muted py-4">
                                    <i class="fas fa-info-circle fa-2x mb-3"></i><br>
                                    Nenhuma venda encontrada para este corretor.
                                </td>
                            </tr>
                            <%
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-light">
                                <th colspan="9" class="text-end">Totais:</th>
                                <th id="totalValor" class="text-end">R$ <%= FormatNumber(totalValor, 2) %></th>
                                <th id="totalComissao" class="text-end">R$ <%= FormatNumber(totalComissaoGeral, 2) %></th>
                                <th colspan="1"></th>
                            </tr>
                        </tfoot>
                    </table>
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
        $('#tabelaVendas').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            pageLength: 25,
            order: [[0, "desc"]],
            responsive: true,
            dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6"f>>rt<"row"<"col-sm-12 col-md-6"i><"col-sm-12 col-md-6"p>>'
        });
    });
    </script>
</body>
</html>

<%
' Fechar recordset e conexão
If Not rsVendas Is Nothing Then
    If rsVendas.State = 1 Then rsVendas.Close
    Set rsVendas = Nothing
End If

If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>