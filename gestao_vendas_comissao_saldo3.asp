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
                   "AND V.AnoVenda = " & anoSelecionado

If mesSelecionado <> "0" Then
    sqlComissoesPagas = sqlComissoesPagas & " AND V.MesVenda = " & mesSelecionado
End If

sqlComissoesPagas = sqlComissoesPagas & " GROUP BY V.AnoVenda, V.MesVenda"

rsComissoesPagas.Open sqlComissoesPagas, connSales

' Criar array para comissões pagas
Dim comissoesPagas(12)
For i = 1 To 12
    comissoesPagas(i) = 0
Next

Do While Not rsComissoesPagas.EOF
    mes = CInt(rsComissoesPagas("MesVenda"))
    comissoesPagas(mes) = CDbl(rsComissoesPagas("ComissaoPaga"))
    rsComissoesPagas.MoveNext
Loop
rsComissoesPagas.Close
Set rsComissoesPagas = Nothing

' Calcular totais
Dim totalVGV, totalComissao, totalDesconto, totalComissaoLiquida, totalPaga, totalAPagar
totalVGV = 0
totalComissao = 0
totalDesconto = 0
totalComissaoLiquida = 0
totalPaga = 0
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
        
        .badge-pago {
            background-color: #28a745;
            color: white;
        }
        
        .badge-pendente {
            background-color: #fd7e14;
            color: white;
        }
        
        .table-warning {
            background-color: #fff3cd !important;
        }
        
        .table-success {
            background-color: #d1e7dd !important;
        }
        
        .mes-header {
            background-color: #e9ecef !important;
            font-weight: bold;
            font-size: 1.1em;
        }
        
        .filter-row {
            margin-bottom: 1rem;
        }
        
        .info-badge {
            font-size: 0.75rem;
            padding: 0.2rem 0.4rem;
        }
        
        .desconto-info {
            font-size: 0.8rem;
            color: #6c757d;
        }
    </style>
</head>
<body>
    <header class="header-bordo">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 style="color: #ffffff !important; margin: 0; font-size: 20px;">
                        <i class="fas fa-money-bill-wave me-2"></i> Resumo de Comissões
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
                                <th>VGV (R$)</th>
                                <th>Comissão Bruta (R$)</th>
                                <th>Desconto Trib. (R$)</th>
                                <th>Comissão Líquida (R$)</th>
                                <th>Comissão Paga (R$)</th>
                                <th>Comissão a Pagar (R$)</th>
                                <th>Status</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsResumo.EOF Then
                                Do While Not rsResumo.EOF
                                    Dim comissaoPagaMes, comissaoAPagarMes, comissaoLiquidaMes
                                    mes = CInt(rsResumo("MesVenda"))
                                    
                                    ' Verificar se existe valor pago para este mês
                                    comissaoPagaMes = comissoesPagas(mes)
                                    comissaoLiquidaMes = CDbl(rsResumo("ComissaoLiquida"))
                                    comissaoAPagarMes = comissaoLiquidaMes - comissaoPagaMes
                                    
                                    ' Garantir que valores não fiquem negativos
                                    If comissaoAPagarMes < 0 Then comissaoAPagarMes = 0
                                    
                                    totalVGV = totalVGV + CDbl(rsResumo("VGV"))
                                    totalComissao = totalComissao + CDbl(rsResumo("ComissaoTotal"))
                                    totalDesconto = totalDesconto + CDbl(rsResumo("TotalDesconto"))
                                    totalComissaoLiquida = totalComissaoLiquida + comissaoLiquidaMes
                                    totalPaga = totalPaga + comissaoPagaMes
                                    totalAPagar = totalAPagar + comissaoAPagarMes
                                    
                                    Dim nomeMes, statusComissao, badgeStatus
                                    nomeMes = GetNomeMes(mes)
                                    
                                    If comissaoAPagarMes <= 0 Then
                                        statusComissao = "Quitado"
                                        badgeStatus = "badge-pago"
                                    Else
                                        statusComissao = "Pendente"
                                        badgeStatus = "badge-pendente"
                                    End If
                            %>
                            <tr>
                                <td><strong><%= rsResumo("AnoVenda") %></strong></td>
                                <td data-order="<%= rsResumo("MesVenda") %>">
                                    <strong><%= nomeMes %></strong>
                                    <br><small class="text-muted">(<%= Right("0" & rsResumo("MesVenda"), 2) %>)</small>
                                </td>
                                <td class="valor-positivo" data-order="<%= rsResumo("VGV") %>">R$ <%= FormatNumber(rsResumo("VGV"), 2) %></td>
                                <td class="valor-positivo" data-order="<%= rsResumo("ComissaoTotal") %>">R$ <%= FormatNumber(rsResumo("ComissaoTotal"), 2) %></td>
                                <td class="valor-desconto" data-order="<%= rsResumo("TotalDesconto") %>">
                                    R$ <%= FormatNumber(rsResumo("TotalDesconto"), 2) %>
                                    <% If CDbl(rsResumo("TotalDesconto")) > 0 Then %>
                                    <br><small class="desconto-info"><i class="fas fa-info-circle"></i> <%= FormatNumber((rsResumo("TotalDesconto")/rsResumo("ComissaoTotal"))*100, 1) %>%</small>
                                    <% End If %>
                                </td>
                                <td class="valor-liquido" data-order="<%= comissaoLiquidaMes %>">R$ <%= FormatNumber(comissaoLiquidaMes, 2) %></td>
                                <td class="valor-positivo" data-order="<%= comissaoPagaMes %>">R$ <%= FormatNumber(comissaoPagaMes, 2) %></td>
                                <td class="valor-pendente" data-order="<%= comissaoAPagarMes %>">R$ <%= FormatNumber(comissaoAPagarMes, 2) %></td>
                                <td data-order="<%= statusComissao %>">
                                    <span class="badge <%= badgeStatus %>">
                                        <%= statusComissao %>
                                    </span>
                                </td>
                                <td>
                                    <a href="gestao_vendas_comissao_detalhes1.asp?ano=<%= rsResumo("AnoVenda") %>&mes=<%= rsResumo("MesVenda") %>" 
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
                                <td colspan="10" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhum dado encontrado para o filtro aplicado.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            rsResumo.Close
                            Set rsResumo = Nothing
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-light">
                                <th colspan="2" class="text-end">Totais:</th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalVGV, 2) %></th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalComissao, 2) %></th>
                                <th class="valor-desconto">R$ <%= FormatNumber(totalDesconto, 2) %></th>
                                <th class="valor-liquido">R$ <%= FormatNumber(totalComissaoLiquida, 2) %></th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalPaga, 2) %></th>
                                <th class="valor-pendente">R$ <%= FormatNumber(totalAPagar, 2) %></th>
                                <th>
                                    <span class="badge <% If totalAPagar = 0 Then Response.Write "badge-pago" Else Response.Write "badge-pendente" End If %>">
                                        <% If totalAPagar = 0 Then Response.Write "Quitado" Else Response.Write "Pendente" End If %>
                                    </span>
                                </th>
                                <th></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>

        <!-- Segunda Tabela - Resumo por Diretoria e Mês -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-building me-2"></i>Resumo por Diretoria e Mês - <%= tituloFiltro %></h5>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table id="tabelaDiretoria" class="table table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>Ano</th>
                                <th>Mês</th>
                                <th>Diretoria</th>
                                <th>VGV (R$)</th>
                                <th>Comissão Bruta (R$)</th>
                                <th>Desconto Trib. (R$)</th>
                                <th>Comissão Líquida (R$)</th>
                                <th>Comissão Paga (R$)</th>
                                <th>Comissão a Pagar (R$)</th>
                                <th>Status</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            ' Buscar dados por mês e diretoria
                            Set rsDiretoria = Server.CreateObject("ADODB.Recordset")
                            
                            sqlDiretoria = "SELECT " & _
                                          "AnoVenda, " & _
                                          "MesVenda, " & _
                                          "Diretoria, " & _
                                          "SUM(ValorUnidade) as VGV, " & _
                                          "SUM(ValorDiretoria + ValorGerencia + ValorCorretor) as ComissaoTotal, " & _
                                          "SUM(DescontoBruto) as TotalDesconto, " & _
                                          "SUM(ValorLiqGeral) as ComissaoLiquida " & _
                                          "FROM Vendas " & _
                                          sqlWhere & " " & _
                                          "GROUP BY AnoVenda, MesVenda, Diretoria " & _
                                          "ORDER BY AnoVenda DESC, MesVenda DESC, Diretoria"
                            
                            rsDiretoria.Open sqlDiretoria, connSales
                            
                            ' Buscar comissões pagas por mês e diretoria
                            Set rsComissoesPagasDiretoria = Server.CreateObject("ADODB.Recordset")
                            sqlComissoesPagasDiretoria = "SELECT " & _
                                                       "V.AnoVenda, " & _
                                                       "V.MesVenda, " & _
                                                       "V.Diretoria, " & _
                                                       "SUM(PC.ValorPago) as ComissaoPaga " & _
                                                       "FROM Vendas V " & _
                                                       "INNER JOIN PAGAMENTOS_COMISSOES PC ON V.ID = PC.ID_Venda " & _
                                                       "WHERE (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                                                       "AND V.AnoVenda = " & anoSelecionado

                            If mesSelecionado <> "0" Then
                                sqlComissoesPagasDiretoria = sqlComissoesPagasDiretoria & " AND V.MesVenda = " & mesSelecionado
                            End If

                            sqlComissoesPagasDiretoria = sqlComissoesPagasDiretoria & " GROUP BY V.AnoVenda, V.MesVenda, V.Diretoria"
                            
                            rsComissoesPagasDiretoria.Open sqlComissoesPagasDiretoria, connSales
                            
                            ' Criar array para comissões pagas por mês e diretoria
                            Dim comissoesPagasDiretoria()
                            ReDim comissoesPagasDiretoria(12, 10) ' meses x diretorias
                            
                            ' Inicializar array
                            For i = 1 To 12
                                For j = 1 To 10
                                    comissoesPagasDiretoria(i, j) = 0
                                Next
                            Next
                            
                            ' Mapear diretorias para índices
                            Dim diretorias, diretoriaIndex
                            Set diretorias = Server.CreateObject("Scripting.Dictionary")
                            diretoriaIndex = 1
                            
                            Do While Not rsComissoesPagasDiretoria.EOF
                                mes = CInt(rsComissoesPagasDiretoria("MesVenda"))
                                diretoria = rsComissoesPagasDiretoria("Diretoria")
                                
                                If Not diretorias.Exists(diretoria) Then
                                    diretorias.Add diretoria, diretoriaIndex
                                    diretoriaIndex = diretoriaIndex + 1
                                End If
                                
                                idx = diretorias(diretoria)
                                comissoesPagasDiretoria(mes, idx) = CDbl(rsComissoesPagasDiretoria("ComissaoPaga"))
                                rsComissoesPagasDiretoria.MoveNext
                            Loop
                            rsComissoesPagasDiretoria.Close
                            Set rsComissoesPagasDiretoria = Nothing
                            
                            ' Variáveis para totais
                            Dim totalGeralVGV, totalGeralComissao, totalGeralDesconto, totalGeralLiquida, totalGeralPaga, totalGeralAPagar
                            Dim totalMesVGV, totalMesComissao, totalMesDesconto, totalMesLiquida, totalMesPaga, totalMesAPagar
                            Dim mesAnterior, primeiroRegistro
                            
                            totalGeralVGV = 0
                            totalGeralComissao = 0
                            totalGeralDesconto = 0
                            totalGeralLiquida = 0
                            totalGeralPaga = 0
                            totalGeralAPagar = 0
                            
                            If Not rsDiretoria.EOF Then
                                mesAnterior = rsDiretoria("MesVenda")
                                primeiroRegistro = True
                                totalMesVGV = 0
                                totalMesComissao = 0
                                totalMesDesconto = 0
                                totalMesLiquida = 0
                                totalMesPaga = 0
                                totalMesAPagar = 0
                                
                                Do While Not rsDiretoria.EOF
                                    Dim comissaoPagaDiretoria, comissaoAPagarDiretoria, comissaoLiquidaDiretoria
                                    mes = CInt(rsDiretoria("MesVenda"))
                                    diretoria = rsDiretoria("Diretoria")
                                    
                                    ' Verificar se é um novo mês
                                    If mes <> mesAnterior And Not primeiroRegistro Then
                                        ' Adicionar linha de total do mês anterior
                            %>
                            <tr class="table-warning">
                                <td colspan="3" class="text-end fw-bold">Total <%= GetNomeMes(mesAnterior) %>:</td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesVGV, 2) %></td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesComissao, 2) %></td>
                                <td class="fw-bold valor-desconto">R$ <%= FormatNumber(totalMesDesconto, 2) %></td>
                                <td class="fw-bold valor-liquido">R$ <%= FormatNumber(totalMesLiquida, 2) %></td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesPaga, 2) %></td>
                                <td class="fw-bold valor-pendente">R$ <%= FormatNumber(totalMesAPagar, 2) %></td>
                                <td>
                                    <span class="badge <% If totalMesAPagar = 0 Then Response.Write "badge-pago" Else Response.Write "badge-pendente" End If %>">
                                        <% If totalMesAPagar = 0 Then Response.Write "Quitado" Else Response.Write "Pendente" End If %>
                                    </span>
                                </td>
                                <td></td>
                            </tr>
                            <%
                                        ' Reiniciar totais do mês
                                        totalMesVGV = 0
                                        totalMesComissao = 0
                                        totalMesDesconto = 0
                                        totalMesLiquida = 0
                                        totalMesPaga = 0
                                        totalMesAPagar = 0
                                        mesAnterior = mes
                                    End If
                                    
                                    ' Calcular comissões para esta linha
                                    If diretorias.Exists(diretoria) Then
                                        idx = diretorias(diretoria)
                                        comissaoPagaDiretoria = comissoesPagasDiretoria(mes, idx)
                                    Else
                                        comissaoPagaDiretoria = 0
                                    End If
                                    
                                    comissaoLiquidaDiretoria = CDbl(rsDiretoria("ComissaoLiquida"))
                                    comissaoAPagarDiretoria = comissaoLiquidaDiretoria - comissaoPagaDiretoria
                                    
                                    ' Garantir que valores não fiquem negativos
                                    If comissaoAPagarDiretoria < 0 Then comissaoAPagarDiretoria = 0
                                    
                                    ' Acumular totais
                                    totalMesVGV = totalMesVGV + CDbl(rsDiretoria("VGV"))
                                    totalMesComissao = totalMesComissao + CDbl(rsDiretoria("ComissaoTotal"))
                                    totalMesDesconto = totalMesDesconto + CDbl(rsDiretoria("TotalDesconto"))
                                    totalMesLiquida = totalMesLiquida + comissaoLiquidaDiretoria
                                    totalMesPaga = totalMesPaga + comissaoPagaDiretoria
                                    totalMesAPagar = totalMesAPagar + comissaoAPagarDiretoria
                                    
                                    totalGeralVGV = totalGeralVGV + CDbl(rsDiretoria("VGV"))
                                    totalGeralComissao = totalGeralComissao + CDbl(rsDiretoria("ComissaoTotal"))
                                    totalGeralDesconto = totalGeralDesconto + CDbl(rsDiretoria("TotalDesconto"))
                                    totalGeralLiquida = totalGeralLiquida + comissaoLiquidaDiretoria
                                    totalGeralPaga = totalGeralPaga + comissaoPagaDiretoria
                                    totalGeralAPagar = totalGeralAPagar + comissaoAPagarDiretoria
                                    
                                    Dim nomeMesDiretoria, statusComissaoDiretoria, badgeStatusDiretoria
                                    nomeMesDiretoria = GetNomeMes(mes)
                                    
                                    If comissaoAPagarDiretoria <= 0 Then
                                        statusComissaoDiretoria = "Quitado"
                                        badgeStatusDiretoria = "badge-pago"
                                    Else
                                        statusComissaoDiretoria = "Pendente"
                                        badgeStatusDiretoria = "badge-pendente"
                                    End If
                            %>
                            <tr>
                                <td><strong><%= rsDiretoria("AnoVenda") %></strong></td>
                                <td data-order="<%= rsDiretoria("MesVenda") %>">
                                    <strong><%= nomeMesDiretoria %></strong>
                                    <br><small class="text-muted">(<%= Right("0" & rsDiretoria("MesVenda"), 2) %>)</small>
                                </td>
                                <td class="fw-bold text-primary"><%= diretoria %></td>
                                <td class="valor-positivo" data-order="<%= rsDiretoria("VGV") %>">R$ <%= FormatNumber(rsDiretoria("VGV"), 2) %></td>
                                <td class="valor-positivo" data-order="<%= rsDiretoria("ComissaoTotal") %>">R$ <%= FormatNumber(rsDiretoria("ComissaoTotal"), 2) %></td>
                                <td class="valor-desconto" data-order="<%= rsDiretoria("TotalDesconto") %>">
                                    R$ <%= FormatNumber(rsDiretoria("TotalDesconto"), 2) %>
                                    <% If CDbl(rsDiretoria("TotalDesconto")) > 0 Then %>
                                    <br><small class="desconto-info"><i class="fas fa-info-circle"></i> <%= FormatNumber((rsDiretoria("TotalDesconto")/rsDiretoria("ComissaoTotal"))*100, 1) %>%</small>
                                    <% End If %>
                                </td>
                                <td class="valor-liquido" data-order="<%= comissaoLiquidaDiretoria %>">R$ <%= FormatNumber(comissaoLiquidaDiretoria, 2) %></td>
                                <td class="valor-positivo" data-order="<%= comissaoPagaDiretoria %>">R$ <%= FormatNumber(comissaoPagaDiretoria, 2) %></td>
                                <td class="valor-pendente" data-order="<%= comissaoAPagarDiretoria %>">R$ <%= FormatNumber(comissaoAPagarDiretoria, 2) %></td>
                                <td data-order="<%= statusComissaoDiretoria %>">
                                    <span class="badge <%= badgeStatusDiretoria %>">
                                        <%= statusComissaoDiretoria %>
                                    </span>
                                </td>
                                <td>
                                    <a href="gestao_vendas_comissao_detalhes1.asp?ano=<%= rsDiretoria("AnoVenda") %>&mes=<%= rsDiretoria("MesVenda") %>&diretoria=<%= Server.URLEncode(diretoria) %>" 
                                       class="btn btn-detalhes" 
                                       title="Ver detalhes da diretoria" target="_blank">
                                        <i class="fas fa-search me-1"></i>Detalhes
                                    </a>
                                </td>
                            </tr>
                            <%
                                    primeiroRegistro = False
                                    rsDiretoria.MoveNext
                                    
                                    ' Se é o último registro, adicionar o total do último mês
                                    If rsDiretoria.EOF Then
                            %>
                            <tr class="table-warning">
                                <td colspan="3" class="text-end fw-bold">Total <%= GetNomeMes(mesAnterior) %>:</td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesVGV, 2) %></td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesComissao, 2) %></td>
                                <td class="fw-bold valor-desconto">R$ <%= FormatNumber(totalMesDesconto, 2) %></td>
                                <td class="fw-bold valor-liquido">R$ <%= FormatNumber(totalMesLiquida, 2) %></td>
                                <td class="fw-bold valor-positivo">R$ <%= FormatNumber(totalMesPaga, 2) %></td>
                                <td class="fw-bold valor-pendente">R$ <%= FormatNumber(totalMesAPagar, 2) %></td>
                                <td>
                                    <span class="badge <% If totalMesAPagar = 0 Then Response.Write "badge-pago" Else Response.Write "badge-pendente" End If %>">
                                        <% If totalMesAPagar = 0 Then Response.Write "Quitado" Else Response.Write "Pendente" End If %>
                                    </span>
                                </td>
                                <td></td>
                            </tr>
                            <%
                                    End If
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="11" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhum dado encontrado para o filtro aplicado.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            
                            ' Fechar recordset da diretoria
                            If IsObject(rsDiretoria) Then
                                If rsDiretoria.State = 1 Then rsDiretoria.Close
                                Set rsDiretoria = Nothing
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-success">
                                <th colspan="3" class="text-end fw-bold">Total Geral:</th>
                                <th class="fw-bold valor-positivo">R$ <%= FormatNumber(totalGeralVGV, 2) %></th>
                                <th class="fw-bold valor-positivo">R$ <%= FormatNumber(totalGeralComissao, 2) %></th>
                                <th class="fw-bold valor-desconto">R$ <%= FormatNumber(totalGeralDesconto, 2) %></th>
                                <th class="fw-bold valor-liquido">R$ <%= FormatNumber(totalGeralLiquida, 2) %></th>
                                <th class="fw-bold valor-positivo">R$ <%= FormatNumber(totalGeralPaga, 2) %></th>
                                <th class="fw-bold valor-pendente">R$ <%= FormatNumber(totalGeralAPagar, 2) %></th>
                                <th>
                                    <span class="badge <% If totalGeralAPagar = 0 Then Response.Write "badge-pago" Else Response.Write "badge-pendente" End If %>">
                                        <% If totalGeralAPagar = 0 Then Response.Write "Quitado" Else Response.Write "Pendente" End If %>
                                    </span>
                                </th>
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
                    targets: [2, 3, 4, 5, 6, 7],
                    type: 'num'
                }
            ]
        });

        $('#tabelaDiretoria').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            pageLength: 50,
            order: [[0, 'desc'], [1, 'asc'], [2, 'asc']],
            responsive: true,
            dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6"f>>rt<"row"<"col-sm-12 col-md-6"i><"col-sm-12 col-md-6"p>>',
            columnDefs: [
                { 
                    targets: [1],
                    type: 'num'
                },
                { 
                    targets: [3, 4, 5, 6, 7, 8],
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
If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>