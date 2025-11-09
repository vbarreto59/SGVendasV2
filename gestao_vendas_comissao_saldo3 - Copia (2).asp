<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Obter ano e mês selecionados do filtro
Dim anoSelecionado, mesSelecionado
anoSelecionado = Request.QueryString("ano")
mesSelecionado = Request.QueryString("mes")

If anoSelecionado = "" Then
    anoSelecionado = Year(Date()) ' Ano atual como padrão
End If

If mesSelecionado = "" Then
    mesSelecionado = "0" ' "0" representa "Todos os meses"
End If

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Construir SQL com filtros
Dim sqlWhere
sqlWhere = "WHERE (Excluido <> -1 OR Excluido IS NULL) " & _
           "AND AnoVenda = " & anoSelecionado

If mesSelecionado <> "0" Then
    sqlWhere = sqlWhere & " AND MesVenda = " & mesSelecionado
End If

' Buscar dados resumidos por ano e mês
Set rsResumo = Server.CreateObject("ADODB.Recordset")
sqlResumo = "SELECT " & _
            "AnoVenda, " & _
            "MesVenda, " & _
            "SUM(ValorUnidade) as VGV, " & _
            "SUM(ValorDiretoria + ValorGerencia + ValorCorretor) as ComissaoTotal, " & _
            "SUM(DescontoBruto) as TotalDesconto, " & _
            "SUM(ValorLiqGeral) as ComissaoLiquida " & _
            "FROM Vendas " & _
            sqlWhere & " " & _
            "GROUP BY AnoVenda, MesVenda " & _
            "ORDER BY AnoVenda DESC, MesVenda DESC"

rsResumo.Open sqlResumo, connSales

' Buscar comissões pagas separadamente
Set rsComissoesPagas = Server.CreateObject("ADODB.Recordset")
sqlComissoesPagas = "SELECT " & _
                   "V.AnoVenda, " & _
                   "V.MesVenda, " & _
                   "SUM(PC.ValorPago) as ComissaoPaga " & _
                   "FROM Vendas V " & _
                   "INNER JOIN PAGAMENTOS_COMISSOES PC ON V.ID = PC.ID_Venda " & _
                   "WHERE (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                   "AND V.AnoVenda = " & anoSelecionado & _
                   " AND PC.TipoPagamento = 'Comissão'"

If mesSelecionado <> "0" Then
    sqlComissoesPagas = sqlComissoesPagas & " AND V.MesVenda = " & mesSelecionado
End If

sqlComissoesPagas = sqlComissoesPagas & " GROUP BY V.AnoVenda, V.MesVenda"

rsComissoesPagas.Open sqlComissoesPagas, connSales

' Buscar premiações pagas separadamente
Set rsPremiacoesPagas = Server.CreateObject("ADODB.Recordset")
sqlPremiacoesPagas = "SELECT " & _
                    "V.AnoVenda, " & _
                    "V.MesVenda, " & _
                    "SUM(PC.ValorPago) as PremiacaoPaga " & _
                    "FROM Vendas V " & _
                    "INNER JOIN PAGAMENTOS_COMISSOES PC ON V.ID = PC.ID_Venda " & _
                    "WHERE (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                    "AND V.AnoVenda = " & anoSelecionado & _
                    " AND PC.TipoPagamento = 'Premiação'"

If mesSelecionado <> "0" Then
    sqlPremiacoesPagas = sqlPremiacoesPagas & " AND V.MesVenda = " & mesSelecionado
End If

sqlPremiacoesPagas = sqlPremiacoesPagas & " GROUP BY V.AnoVenda, V.MesVenda"

rsPremiacoesPagas.Open sqlPremiacoesPagas, connSales

' DEBUG: Verificar estrutura da tabela Vendas
Response.Write "<!-- DEBUG: Verificando estrutura da tabela Vendas -->"
On Error Resume Next
Set rsTest = connSales.Execute("SELECT TOP 1 * FROM Vendas")
If Err.Number = 0 Then
    For i = 0 To rsTest.Fields.Count - 1
        Response.Write "<!-- Campo: " & rsTest.Fields(i).Name & " -->"
    Next
Else
    Response.Write "<!-- Erro ao acessar tabela Vendas: " & Err.Description & " -->"
End If
On Error GoTo 0

' Buscar premiação total - CORREÇÃO APLICADA AQUI
Set rsPremiacaoTotal = Server.CreateObject("ADODB.Recordset")

' Verificar quais colunas de premiação existem na tabela
Dim sqlPremiacaoTotal
On Error Resume Next

' Tentativa 1: Verificar se existe coluna Premio
sqlTest = "SELECT TOP 1 Premio FROM Vendas"
Set rsTest = connSales.Execute(sqlTest)
If Err.Number = 0 Then
    sqlPremiacaoTotal = "SELECT " & _
                       "AnoVenda, " & _
                       "MesVenda, " & _
                       "SUM(Premio) as PremiacaoTotal " & _
                       "FROM Vendas " & _
                       sqlWhere & " " & _
                       "GROUP BY AnoVenda, MesVenda"
    Response.Write "<!-- Usando coluna: Premio -->"
Else
    Err.Clear
    ' Tentativa 2: Verificar se existe coluna Premiacao
    sqlTest = "SELECT TOP 1 Premiacao FROM Vendas"
    Set rsTest = connSales.Execute(sqlTest)
    If Err.Number = 0 Then
        sqlPremiacaoTotal = "SELECT " & _
                           "AnoVenda, " & _
                           "MesVenda, " & _
                           "SUM(Premiacao) as PremiacaoTotal " & _
                           "FROM Vendas " & _
                           sqlWhere & " " & _
                           "GROUP BY AnoVenda, MesVenda"
        Response.Write "<!-- Usando coluna: Premiacao -->"
    Else
        Err.Clear
        ' Tentativa 3: Verificar se existe coluna ValorPremiacao
        sqlTest = "SELECT TOP 1 ValorPremiacao FROM Vendas"
        Set rsTest = connSales.Execute(sqlTest)
        If Err.Number = 0 Then
            sqlPremiacaoTotal = "SELECT " & _
                               "AnoVenda, " & _
                               "MesVenda, " & _
                               "SUM(ValorPremiacao) as PremiacaoTotal " & _
                               "FROM Vendas " & _
                               sqlWhere & " " & _
                               "GROUP BY AnoVenda, MesVenda"
            Response.Write "<!-- Usando coluna: ValorPremiacao -->"
        Else
            ' Se nenhuma coluna de premiação for encontrada, usar valor fixo baseado nas premiações pagas
            sqlPremiacaoTotal = "SELECT " & _
                               "V.AnoVenda, " & _
                               "V.MesVenda, " & _
                               "SUM(PC.ValorPago) as PremiacaoTotal " & _
                               "FROM Vendas V " & _
                               "INNER JOIN PAGAMENTOS_COMISSOES PC ON V.ID = PC.ID_Venda " & _
                               "WHERE (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                               "AND V.AnoVenda = " & anoSelecionado & _
                               " AND PC.TipoPagamento = 'Premiação'"
            If mesSelecionado <> "0" Then
                sqlPremiacaoTotal = sqlPremiacaoTotal & " AND V.MesVenda = " & mesSelecionado
            End If
            sqlPremiacaoTotal = sqlPremiacaoTotal & " GROUP BY V.AnoVenda, V.MesVenda"
            Response.Write "<!-- Usando premiações pagas como total (fallback) -->"
        End If
    End If
End If
On Error GoTo 0

Response.Write "<!-- SQL PremiacaoTotal: " & sqlPremiacaoTotal & " -->"

rsPremiacaoTotal.Open sqlPremiacaoTotal, connSales

' Criar arrays para comissões e premiações
Dim comissoesPagas(12), premiacoesPagas(12), premiacaoTotal(12)
For i = 1 To 12
    comissoesPagas(i) = 0
    premiacoesPagas(i) = 0
    premiacaoTotal(i) = 0
Next

Do While Not rsComissoesPagas.EOF
    mes = CInt(rsComissoesPagas("MesVenda"))
    comissoesPagas(mes) = CDbl(rsComissoesPagas("ComissaoPaga"))
    rsComissoesPagas.MoveNext
Loop
rsComissoesPagas.Close
Set rsComissoesPagas = Nothing

Do While Not rsPremiacoesPagas.EOF
    mes = CInt(rsPremiacoesPagas("MesVenda"))
    premiacoesPagas(mes) = CDbl(rsPremiacoesPagas("PremiacaoPaga"))
    rsPremiacoesPagas.MoveNext
Loop
rsPremiacoesPagas.Close
Set rsPremiacoesPagas = Nothing

' DEBUG: Verificar dados da premiação total
Response.Write "<!-- DEBUG: Dados da premiação total -->"
Do While Not rsPremiacaoTotal.EOF
    mes = CInt(rsPremiacaoTotal("MesVenda"))
    premiacaoTotal(mes) = CDbl(rsPremiacaoTotal("PremiacaoTotal"))
    Response.Write "<!-- Mes " & mes & ": " & premiacaoTotal(mes) & " -->"
    rsPremiacaoTotal.MoveNext
Loop
rsPremiacaoTotal.Close
Set rsPremiacaoTotal = Nothing

' Se não encontrou dados de premiação total, usar as premiações pagas como base
If totalPremiacao = 0 Then
    For i = 1 To 12
        If premiacoesPagas(i) > 0 Then
            premiacaoTotal(i) = premiacoesPagas(i)
            Response.Write "<!-- Usando premiação paga como total para mês " & i & ": " & premiacoesPagas(i) & " -->"
        End If
    Next
End If

' Calcular totais
Dim totalVGV, totalComissaoBruta, totalDesconto, totalComissaoLiquida
Dim totalComissaoPaga, totalComissaoAPagar
Dim totalPremiacao, totalPremiacaoPaga, totalPremiacaoAPagar
Dim totalPago, totalAPagar

totalVGV = 0
totalComissaoBruta = 0
totalDesconto = 0
totalComissaoLiquida = 0
totalComissaoPaga = 0
totalComissaoAPagar = 0
totalPremiacao = 0
totalPremiacaoPaga = 0
totalPremiacaoAPagar = 0
totalPago = 0
totalAPagar = 0
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resumo de Comissões | Gestão de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding-top: 60px;
        }
        
        .header-bordo {
            background-color: #800000 !important;
            color: #ffffff !important;
            padding: 5px 20px; 
            margin-bottom: 8px !important; 
            font-size: 20px;
            font-weight: bold;
            text-align: left;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
            width: 100%;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
        }
        
        .card-header {
            background: linear-gradient(to right, #2c3e50, #3498db);
            color: white;
            border-bottom: none;
            padding: 1rem 1.5rem;
            font-weight: 600;
        }
        
        .filter-section {
            background: white;
            border-radius: 12px;
            padding: 1rem;
            margin-bottom: 1.5rem;
        }
        
        .table th {
            background-color: #2c3e50;
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
        
        .valor-positivo {
            color: #28a745;
            font-weight: 600;
        }
        
        .valor-pendente {
            color: #dc3545;
            font-weight: 600;
        }
        
        .valor-desconto {
            color: #fd7e14;
            font-weight: 600;
        }
        
        .valor-liquido {
            color: #17a2b8;
            font-weight: 600;
        }
        
        .valor-premiacao {
            color: #9b59b6;
            font-weight: 600;
        }
        
        .valor-total {
            color: #2c3e50;
            font-weight: 700;
        }
        
        .btn-refresh {
            background-color: #fd7e14;
            border-color: #fd7e14;
            color: white;
        }
        
        .btn-detalhes {
            background-color: #17a2b8;
            border-color: #17a2b8;
            color: white;
            font-size: 0.8rem;
            padding: 0.25rem 0.5rem;
        }
        
        .section-comissao {
            border-left: 4px solid #3498db;
            background: linear-gradient(to right, #3498db, #2980b9);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .section-premiacao {
            border-left: 4px solid #9b59b6;
            background: linear-gradient(to right, #9b59b6, #8e44ad);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .section-total {
            border-left: 4px solid #2c3e50;
            background: linear-gradient(to right, #2c3e50, #34495e);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }

        .info-card {
            background: white;
            border-radius: 8px;
            padding: 1rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #3498db;
            height: 100%;
        }

        .info-card-premiacao {
            border-left: 4px solid #9b59b6;
        }

        .info-card-total {
            border-left: 4px solid #2c3e50;
        }

        .info-card h6 {
            color: #6c757d;
            font-size: 0.9rem;
            margin-bottom: 0.5rem;
        }

        .info-card h4 {
            color: #2c3e50;
            margin-bottom: 0;
            font-weight: 700;
        }
        
        .debug-info {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 10px;
            font-size: 12px;
            color: #856404;
        }
    </style>
</head>
<body>
    <header class="header-bordo">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 style="color: #ffffff !important; margin: 0; font-size: 20px;">
                        <i class="fas fa-money-bill-wave me-2"></i> Resumo de Comissões e Premiações
                    </h1>
                </div>
                <div class="col-md-6 text-end">
                    <a href="gestao_vendas.asp" class="btn btn-light btn-sm" style="color: #333 !important;">
                        <i class="fas fa-arrow-left me-1"></i>Voltar para Vendas
                    </a>
                </div>
            </div>
        </div>
    </header>

    <div class="container-fluid main-content">
        <!-- Filtro de Ano e Mês -->
        <div class="filter-section">
            <div class="row align-items-center filter-row">
                <div class="col-md-6">
                    <h5 class="mb-0"><i class="fas fa-filter me-2"></i>Filtros</h5>
                </div>
                <div class="col-md-6">
                    <form method="GET" action="" class="d-flex gap-2">
                        <select name="ano" class="form-select" onchange="this.form.submit()">
                            <%
                            ' Opções de anos
                            Dim anos
                            anos = Array("2025", "2026")
                            
                            For Each ano In anos
                                If CStr(anoSelecionado) = ano Then
                                    Response.Write "<option value='" & ano & "' selected>" & ano & "</option>"
                                Else
                                    Response.Write "<option value='" & ano & "'>" & ano & "</option>"
                                End If
                            Next
                            %>
                        </select>
                        <select name="mes" class="form-select" onchange="this.form.submit()">
                            <%
                            ' Opções de meses
                            Dim meses
                            Set meses = Server.CreateObject("Scripting.Dictionary")
                            meses.Add "0", "Todos os meses"
                            meses.Add "1", "Janeiro"
                            meses.Add "2", "Fevereiro"
                            meses.Add "3", "Março"
                            meses.Add "4", "Abril"
                            meses.Add "5", "Maio"
                            meses.Add "6", "Junho"
                            meses.Add "7", "Julho"
                            meses.Add "8", "Agosto"
                            meses.Add "9", "Setembro"
                            meses.Add "10", "Outubro"
                            meses.Add "11", "Novembro"
                            meses.Add "12", "Dezembro"
                            
                            For Each key in meses
                                If CStr(mesSelecionado) = key Then
                                    Response.Write "<option value='" & key & "' selected>" & meses(key) & "</option>"
                                Else
                                    Response.Write "<option value='" & key & "'>" & meses(key) & "</option>"
                                End If
                            Next
                            %>
                        </select>
                        <button type="button" class="btn btn-refresh" onclick="location.reload()">
                            <i class="fas fa-sync-alt"></i>
                        </button>
                    </form>
                </div>
            </div>
            <div class="row">
                <div class="col-12">
                    <%
                    Dim tituloFiltro
                    If mesSelecionado = "0" Then
                        tituloFiltro = "Ano " & anoSelecionado & " - Todos os meses"
                    Else
                        tituloFiltro = "Ano " & anoSelecionado & " - " & meses(mesSelecionado)
                    End If
                    %>
                    <small class="text-muted"><i class="fas fa-info-circle me-1"></i>Filtro aplicado: <%= tituloFiltro %></small>
                </div>
            </div>
        </div>

        <!-- Cards de Resumo Rápido -->
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="info-card">
                    <h6><i class="fas fa-chart-line me-2"></i>VGV Total</h6>
                    <h4 class="valor-positivo"><%= FormatNumber(totalVGV, 2) %></h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card">
                    <h6><i class="fas fa-money-bill-wave me-2"></i>Comissão Líquida</h6>
                    <h4 class="valor-liquido"><%= FormatNumber(totalComissaoLiquida, 2) %></h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card info-card-premiacao">
                    <h6><i class="fas fa-trophy me-2"></i>Premiação Total</h6>
                    <h4 class="valor-premiacao"><%= FormatNumber(totalPremiacao, 2) %></h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card info-card-total">
                    <h6><i class="fas fa-hand-holding-usd me-2"></i>Total a Pagar</h6>
                    <h4 class="valor-pendente"><%= FormatNumber(totalAPagar, 2) %></h4>
                </div>
            </div>
        </div>

        <!-- Primeira Tabela - Resumo por Mês -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-table me-2"></i>Resumo por Mês - <%= tituloFiltro %></h5>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table id="tabelaResumo" class="table table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>Ano</th>
                                <th>Mês</th>
                                <th>VGV</th>
                                
                                <!-- Seção Comissão -->
                                <th colspan="5" class="text-center section-comissao">
                                    <i class="fas fa-money-bill-wave me-2"></i>COMISSÃO
                                </th>
                                
                                <!-- Seção Premiação -->
                                <th colspan="3" class="text-center section-premiacao">
                                    <i class="fas fa-trophy me-2"></i>PREMAIAÇÃO
                                </th>
                                
                                <!-- Seção Total -->
                                <th colspan="2" class="text-center section-total">
                                    <i class="fas fa-calculator me-2"></i>TOTAL
                                </th>
                                
                                <th>Ações</th>
                            </tr>
                            <tr>
                                <th></th>
                                <th></th>
                                <th></th>
                                
                                <!-- Subheaders Comissão -->
                                <th>Bruta</th>
                                <th>Desc. Trib.</th>
                                <th>Líquida</th>
                                <th>Paga</th>
                                <th>a Pagar</th>
                                
                                <!-- Subheaders Premiação -->
                                <th>Total</th>
                                <th>Paga</th>
                                <th>a Pagar</th>
                                
                                <!-- Subheaders Total -->
                                <th>Pago</th>
                                <th>a Pagar</th>
                                
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsResumo.EOF Then
                                Do While Not rsResumo.EOF
                                    Dim comissaoPagaMes, premiacaoPagaMes, premiacaoTotalMes
                                    Dim totalPagoMes, comissaoAPagarMes, premiacaoAPagarMes, totalAPagarMes
                                    Dim comissaoLiquidaMes
                                    
                                    mes = CInt(rsResumo("MesVenda"))
                                    
                                    ' Verificar se existe valor pago para este mês
                                    comissaoPagaMes = comissoesPagas(mes)
                                    premiacaoPagaMes = premiacoesPagas(mes)
                                    premiacaoTotalMes = premiacaoTotal(mes)
                                    
                                    ' CORREÇÃO: Se premiação total for 0 mas tem premiação paga, usar premiação paga como total
                                    If premiacaoTotalMes = 0 And premiacaoPagaMes > 0 Then
                                        premiacaoTotalMes = premiacaoPagaMes
                                    End If
                                    
                                    totalPagoMes = comissaoPagaMes + premiacaoPagaMes
                                    comissaoLiquidaMes = CDbl(rsResumo("ComissaoLiquida"))
                                    comissaoAPagarMes = comissaoLiquidaMes - comissaoPagaMes
                                    premiacaoAPagarMes = premiacaoTotalMes - premiacaoPagaMes
                                    totalAPagarMes = comissaoAPagarMes + premiacaoAPagarMes
                                    
                                    ' Garantir que valores não fiquem negativos
                                    If comissaoAPagarMes < 0 Then comissaoAPagarMes = 0
                                    If premiacaoAPagarMes < 0 Then premiacaoAPagarMes = 0
                                    If totalAPagarMes < 0 Then totalAPagarMes = 0
                                    
                                    ' Acumular totais
                                    totalVGV = totalVGV + CDbl(rsResumo("VGV"))
                                    totalComissaoBruta = totalComissaoBruta + CDbl(rsResumo("ComissaoTotal"))
                                    totalDesconto = totalDesconto + CDbl(rsResumo("TotalDesconto"))
                                    totalComissaoLiquida = totalComissaoLiquida + comissaoLiquidaMes
                                    totalComissaoPaga = totalComissaoPaga + comissaoPagaMes
                                    totalComissaoAPagar = totalComissaoAPagar + comissaoAPagarMes
                                    
                                    ' CORREÇÃO: Acumular corretamente a premiação total
                                    totalPremiacao = totalPremiacao + premiacaoTotalMes
                                    totalPremiacaoPaga = totalPremiacaoPaga + premiacaoPagaMes
                                    totalPremiacaoAPagar = totalPremiacaoAPagar + premiacaoAPagarMes
                                    
                                    totalPago = totalPago + totalPagoMes
                                    totalAPagar = totalAPagar + totalAPagarMes
                                    
                                    Dim nomeMes
                                    nomeMes = GetNomeMes(mes)
                            %>
                            <tr>
                                <td><strong><%= rsResumo("AnoVenda") %></strong></td>
                                <td data-order="<%= rsResumo("MesVenda") %>">
                                    <strong><%= nomeMes %></strong>
                                    <br><small class="text-muted">(<%= Right("0" & rsResumo("MesVenda"), 2) %>)</small>
                                </td>
                                <td class="valor-positivo" data-order="<%= rsResumo("VGV") %>"><%= FormatNumber(rsResumo("VGV"), 2) %></td>
                                
                                <!-- Dados Comissão -->
                                <td class="valor-positivo" data-order="<%= rsResumo("ComissaoTotal") %>"><%= FormatNumber(rsResumo("ComissaoTotal"), 2) %></td>
                                <td class="valor-desconto" data-order="<%= rsResumo("TotalDesconto") %>">
                                    <%= FormatNumber(rsResumo("TotalDesconto"), 2) %>
                                </td>
                                <td class="valor-liquido" data-order="<%= comissaoLiquidaMes %>"><%= FormatNumber(comissaoLiquidaMes, 2) %></td>
                                <td class="valor-positivo" data-order="<%= comissaoPagaMes %>"><%= FormatNumber(comissaoPagaMes, 2) %></td>
                                <td class="valor-pendente" data-order="<%= comissaoAPagarMes %>"><%= FormatNumber(comissaoAPagarMes, 2) %></td>
                                
                                <!-- Dados Premiação -->
                                <td class="valor-premiacao" data-order="<%= premiacaoTotalMes %>"><%= FormatNumber(premiacaoTotalMes, 2) %></td>
                                <td class="valor-premiacao" data-order="<%= premiacaoPagaMes %>"><%= FormatNumber(premiacaoPagaMes, 2) %></td>
                                <td class="valor-pendente" data-order="<%= premiacaoAPagarMes %>"><%= FormatNumber(premiacaoAPagarMes, 2) %></td>
                                
                                <!-- Dados Total -->
                                <td class="valor-positivo" data-order="<%= totalPagoMes %>"><%= FormatNumber(totalPagoMes, 2) %></td>
                                <td class="valor-pendente" data-order="<%= totalAPagarMes %>"><%= FormatNumber(totalAPagarMes, 2) %></td>
                                
                                <td>
                                    <a href="gestao_vendas_comissao_detalhes3.asp?ano=<%= rsResumo("AnoVenda") %>&mes=<%= rsResumo("MesVenda") %>" 
                                       class="btn btn-detalhes" 
                                       title="Ver detalhes do mês" target="_blank">
                                        <i class="fas fa-search me-1"></i>Detalhes
                                    </a>
                                </td>
                            </tr>
                            <%
                                    rsResumo.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="15" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhum dado encontrado para o filtro aplicado.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-light">
                                <th colspan="2" class="text-end">Totais:</th>
                                <th class="valor-positivo"><%= FormatNumber(totalVGV, 2) %></th>
                                
                                <!-- Totais Comissão -->
                                <th class="valor-positivo"><%= FormatNumber(totalComissaoBruta, 2) %></th>
                                <th class="valor-desconto"><%= FormatNumber(totalDesconto, 2) %></th>
                                <th class="valor-liquido"><%= FormatNumber(totalComissaoLiquida, 2) %></th>
                                <th class="valor-positivo"><%= FormatNumber(totalComissaoPaga, 2) %></th>
                                <th class="valor-pendente"><%= FormatNumber(totalComissaoAPagar, 2) %></th>
                                
                                <!-- Totais Premiação -->
                                <th class="valor-premiacao"><%= FormatNumber(totalPremiacao, 2) %></th>
                                <th class="valor-premiacao"><%= FormatNumber(totalPremiacaoPaga, 2) %></th>
                                <th class="valor-pendente"><%= FormatNumber(totalPremiacaoAPagar, 2) %></th>
                                
                                <!-- Totais Gerais -->
                                <th class="valor-positivo"><%= FormatNumber(totalPago, 2) %></th>
                                <th class="valor-pendente"><%= FormatNumber(totalAPagar, 2) %></th>
                                
                                <th></th>
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
        $('#tabelaResumo').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            pageLength: 25,
            order: [[0, 'desc'], [1, 'asc']],
            responsive: true,
            dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6"f>>rt<"row"<"col-sm-12 col-md-6"i><"col-sm-12 col-md-6"p>>',
            columnDefs: [
                { 
                    targets: [1],
                    type: 'num'
                },
                { 
                    targets: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
                    type: 'num'
                }
            ]
        });
    });
    </script>
</body>
</html>

<%
' Função para obter nome do mês
Function GetNomeMes(mes)
    Select Case mes
        Case 1: GetNomeMes = "Janeiro"
        Case 2: GetNomeMes = "Fevereiro"
        Case 3: GetNomeMes = "Março"
        Case 4: GetNomeMes = "Abril"
        Case 5: GetNomeMes = "Maio"
        Case 6: GetNomeMes = "Junho"
        Case 7: GetNomeMes = "Julho"
        Case 8: GetNomeMes = "Agosto"
        Case 9: GetNomeMes = "Setembro"
        Case 10: GetNomeMes = "Outubro"
        Case 11: GetNomeMes = "Novembro"
        Case 12: GetNomeMes = "Dezembro"
    End Select
End Function

' Fechar conexões
If IsObject(rsResumo) Then
    If rsResumo.State = 1 Then rsResumo.Close
    Set rsResumo = Nothing
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>