<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

' Abrir conexão apenas com o banco Sales
Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar ao banco de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno
filtroAno = Request.QueryString("ano")

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

Function GetUniqueValues(tableName, columnName, whereClause)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    sqlQuery = "SELECT DISTINCT " & columnName & " FROM " & tableName & " "
    sqlQuery = sqlQuery & whereClause & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = connSales.Execute(sqlQuery)
    If Err.Number <> 0 Then
        GetUniqueValues = Array()
        Exit Function
    End If
    On Error GoTo 0
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                dict.Add CStr(rs(0)), 1
            End If
            rs.MoveNext
        Loop
    End If
    
    If Not rs Is Nothing Then 
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

' ===============================================
' POPULAR OS SELECTS DO FORMULÁRIO
' ===============================================

Dim uniqueAnos
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")

' ===============================================
' OBTER DADOS DA CONSULTA (APENAS SE ANO ESTIVER PREENCHIDO)
' ===============================================

Dim dadosPessoas, totalGeralVTotal
Set dadosPessoas = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Construir consulta SQL
    Dim sqlConsulta, rsConsulta
    sqlConsulta = "SELECT Vendas.AnoVenda, VENDA_TEMP.Nome, VENDA_TEMP.Cargo, Sum(VENDA_TEMP.VTotal) AS SomaDeVTotal " & _
                  "FROM VENDA_TEMP INNER JOIN Vendas ON VENDA_TEMP.ID_Venda = Vendas.Id " & _
                  "WHERE Vendas.AnoVenda = " & filtroAno & " " & _
                  "GROUP BY Vendas.AnoVenda, VENDA_TEMP.Nome, VENDA_TEMP.Cargo " & _
                  "ORDER BY VENDA_TEMP.Nome, VENDA_TEMP.Cargo, Sum(VENDA_TEMP.VTotal) DESC"

    Set rsConsulta = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsConsulta.Open sqlConsulta, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlConsulta)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados
    totalGeralVTotal = 0
    
    If Not rsConsulta.EOF Then
        Do While Not rsConsulta.EOF
            Dim nomePessoa, cargoPessoa, vTotalPessoa
            nomePessoa = Trim(CStr(rsConsulta("Nome")))
            cargoPessoa = Trim(CStr(rsConsulta("Cargo")))
            vTotalPessoa = CDbl(rsConsulta("SomaDeVTotal"))
            
            ' Verificar se a pessoa já existe no dicionário
            If Not dadosPessoas.Exists(nomePessoa) Then
                Dim infoPessoa
                Set infoPessoa = Server.CreateObject("Scripting.Dictionary")
                infoPessoa.Add "Cargos", Server.CreateObject("Scripting.Dictionary")
                infoPessoa.Add "TotalGeral", 0
                dadosPessoas.Add nomePessoa, infoPessoa
            End If
            
            ' Atualizar dados da pessoa
            Set infoPessoa = dadosPessoas(nomePessoa)
            
            ' Adicionar/atualizar cargo
            infoPessoa("Cargos").Add cargoPessoa, vTotalPessoa
            
            ' Atualizar total geral da pessoa
            infoPessoa("TotalGeral") = infoPessoa("TotalGeral") + vTotalPessoa
            
            ' Atualizar total geral
            totalGeralVTotal = totalGeralVTotal + vTotalPessoa
            
            rsConsulta.MoveNext
        Loop
    End If

    If rsConsulta.State = 1 Then rsConsulta.Close
    Set rsConsulta = Nothing
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório Consolidado por Pessoa</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <!-- DataTables CSS -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/responsive/2.5.0/css/responsive.bootstrap5.min.css">
    <style>
        body {
            background-color: #A5A2A2;
            padding: 20px;
            color: white;
        }
        .card-kpi {
            background-color: #F0ECEC;
            color: black;
            padding: 15px;
            margin-top: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
        }
        .container-fluid {
            max-width: 1800px;
            margin: 0 auto;
        }
        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 20px;
            border-radius: 8px;
            font-size: 1rem;
            margin-bottom: 10px;
            min-height: 120px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .kpi-card h5 {
            font-size: 1rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            line-height: 1.2;
            font-size: 0.9rem;
        }
        .kpi-card i {
            font-size: 1.5rem;
            margin-bottom: 8px;
        }
        .bg-primary-kpi { background-color: #007bff; }
        .bg-success-kpi { background-color: #28a745; }
        .bg-info-kpi { background-color: #17a2b8; }
        .bg-warning-kpi { background-color: #ffc107; color: #000; }
        .bg-danger-kpi { background-color: #dc3545; }
        .bg-secondary-kpi { background-color: #6c757d; }
        .bg-dark-kpi { background-color: #343a40; }
        .bg-maroon-kpi { background-color: #800000; }
        
        .filter-container {
            background-color: #Fff;
            color: black;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .table-responsive {
            background-color: white;
            border-radius: 8px;
            font-size: 0.85rem;
        }
        .table th {
            background-color: #800000;
            color: white;
            position: sticky;
            top: 0;
            font-size: 0.8rem;
        }
        .text-right-v { text-align: right; }
        .text-center-v { text-align: center; }
        .pessoa-header {
            background-color: #e9ecef !important;
            font-weight: bold;
            border-left: 4px solid #800000 !important;
        }
        .cargo-row {
            background-color: #f8f9fa !important;
        }
        .total-pessoa {
            background-color: #d4edda !important;
            font-weight: bold;
            border-top: 2px solid #28a745 !important;
        }
        .total-geral {
            background-color: #800000 !important;
            color: white !important;
            font-weight: bold;
            font-size: 1rem;
        }
        .alert-warning {
            background-color: #fff3cd;
            border-color: #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .table-hover tbody tr:hover {
            background-color: rgba(0,0,0,.075);
        }
        .valor-cell {
            font-weight: bold;
            color: #28a745;
        }
        .cargo-corretor { border-left: 4px solid #28a745 !important; }
        .cargo-gerente { border-left: 4px solid #007bff !important; }
        .cargo-diretor { border-left: 4px solid #6f42c1 !important; }
        .cargo-outros { border-left: 4px solid #fd7e14 !important; }
        .posicao-top {
            background: linear-gradient(135deg, #fff3cd, #ffeaa7) !important;
            font-weight: bold;
        }
        .posicao-1 { background: linear-gradient(135deg, #ffd700, #ffed4e) !important; }
        .posicao-2 { background: linear-gradient(135deg, #c0c0c0, #e0e0e0) !important; }
        .posicao-3 { background: linear-gradient(135deg, #cd7f32, #e3964a) !important; }
        .dataTables_wrapper .dataTables_filter {
            float: right;
            margin-bottom: 10px;
        }
        .dataTables_wrapper .dataTables_length {
            float: left;
            margin-bottom: 10px;
        }
        .dataTables_wrapper .dataTables_paginate {
            float: right;
            margin-top: 10px;
        }
        .dataTables_wrapper .dataTables_info {
            float: left;
            margin-top: 10px;
        }
        .dt-buttons {
            margin-bottom: 10px;
        }
        .dt-button {
            background-color: #6c757d !important;
            border-color: #6c757d !important;
            color: white !important;
            margin-right: 5px;
            margin-bottom: 5px;
        }
        .dt-button:hover {
            background-color: #5a6268 !important;
            border-color: #545b62 !important;
        }
        .badge-bg-purple {
            background-color: #6f42c1 !important;
        }
    </style>
    <style>
        body {
            transform: scale(0.8); 
            transform-origin: 0 0; 
            width: calc(100% / 0.8); 
        }
    </style>    
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-trophy"></i> SGVendas - Ranking por Total Recebido
        </h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-8">
                        <div class="filter-label">Ano</div>
                        <select class="form-select" name="ano" id="anoFilter" required>
                            <option value="">Selecione o ano</option>
                            <%
                            If IsArray(uniqueAnos) Then
                                For Each ano In uniqueAnos
                                    Response.Write "<option value=""" & ano & """"
                                    If CStr(filtroAno) = CStr(ano) Then Response.Write " selected"
                                    Response.Write ">" & ano & "</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-4">
                        <div class="filter-label">&nbsp;</div>
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Ranking
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Then %>
            <div class="alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar o ranking.
            </div>
        <% Else %>
        
        <!-- KPIs Principais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-money-bill-wave"></i>
                    <h5>Total Vendas <%= filtroAno %></h5>
                    <p><%= FormatNumber(totalGeralVTotal, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <i class="fas fa-user-tie"></i>
                    <h5>Total de Pessoas</h5>
                    <p><%= dadosPessoas.Count %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-layer-group"></i>
                    <h5>Total de Cargos</h5>
                    <p>
                        <% 
                        Dim totalCargos
                        totalCargos = 0
                        For Each pessoaKey In dadosPessoas.Keys
                            totalCargos = totalCargos + dadosPessoas(pessoaKey)("Cargos").Count
                        Next
                        Response.Write totalCargos
                        %>
                    </p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-chart-line"></i>
                    <h5>Média por Pessoa</h5>
                    <p>
                        <% 
                        If dadosPessoas.Count > 0 Then 
                            Response.Write "" & FormatNumber(totalGeralVTotal / dadosPessoas.Count, 2)
                        Else 
                            Response.Write "0,00"
                        End If 
                        %>
                    </p>
                </div>
            </div>
        </div>

        <!-- Tabela de Ranking por Total Recebido -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">
                <i class="fas fa-trophy"></i> Ranking por Total Recebido - Ano <%= filtroAno %>
                <small class="text-muted">(Ordenado por maior rendimento acumulado)</small>
            </h3>
            
            <div class="table-responsive" style="overflow-y: auto;">
                <table id="tabelaPessoas" class="table table-striped table-hover table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th class="text-center-v">#</th>
                            <th class="text-center-v">Posição</th>
                            <th class="text-center-v">Nome</th>
                            <th class="text-center-v">Cargo</th>
                            <th class="text-center-v bg-success text-white">Valor por Cargo</th>
                            <th class="text-center-v bg-primary text-white">% do Cargo</th>
                            <th class="text-center-v bg-info text-white">Total da Pessoa</th>
                            <th class="text-center-v bg-warning text-dark">% da Pessoa</th>
                            <th class="text-center-v bg-secondary text-white">% do Total Geral</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        If dadosPessoas.Count > 0 Then
                            Dim arrPessoas, pessoaKey
                            arrPessoas = dadosPessoas.Keys
                            
                            ' Ordenar pessoas por total geral (DECRESCENTE - maior primeiro)
                            For i = 0 To UBound(arrPessoas)
                                For j = i + 1 To UBound(arrPessoas)
                                    If dadosPessoas(arrPessoas(j))("TotalGeral") > dadosPessoas(arrPessoas(i))("TotalGeral") Then
                                        Dim temp
                                        temp = arrPessoas(i)
                                        arrPessoas(i) = arrPessoas(j)
                                        arrPessoas(j) = temp
                                    End If
                                Next
                            Next
                            
                            Dim posicao
                            posicao = 0
                            
                            For Each pessoaKey In arrPessoas
                                posicao = posicao + 1
                                Set infoPessoa = dadosPessoas(pessoaKey)
                                Dim cargosPessoa
                                Set cargosPessoa = infoPessoa("Cargos")
                                Dim arrCargos, cargoKey
                                arrCargos = cargosPessoa.Keys
                                
                                ' Ordenar cargos de forma específica: Corretor, Gerente, Diretor, outros
                                Dim cargosOrdenados
                                cargosOrdenados = Array("Corretor", "Gerente", "Diretor")
                                
                                Dim primeiroCargo
                                primeiroCargo = True
                                Dim linhaCargoCount
                                linhaCargoCount = 0
                                
                                ' Primeiro mostrar cargos na ordem específica
                                For Each cargoEspecifico In cargosOrdenados
                                    If cargosPessoa.Exists(cargoEspecifico) Then
                                        linhaCargoCount = linhaCargoCount + 1
                                        Call ExibirLinhaCargo(pessoaKey, cargoEspecifico, cargosPessoa(cargoEspecifico), infoPessoa("TotalGeral"), primeiroCargo, linhaCargoCount, posicao)
                                        primeiroCargo = False
                                    End If
                                Next
                                
                                ' Depois mostrar outros cargos
                                For Each cargoKey In cargosPessoa.Keys
                                    Dim jaExibido
                                    jaExibido = False
                                    For Each cargoEspecifico In cargosOrdenados
                                        If cargoKey = cargoEspecifico Then
                                            jaExibido = True
                                            Exit For
                                        End If
                                    Next
                                    
                                    If Not jaExibido Then
                                        linhaCargoCount = linhaCargoCount + 1
                                        Call ExibirLinhaCargo(pessoaKey, cargoKey, cargosPessoa(cargoKey), infoPessoa("TotalGeral"), primeiroCargo, linhaCargoCount, posicao)
                                        primeiroCargo = False
                                    End If
                                Next
                                
                                ' Linha de total da pessoa
                                Dim percentualTotalPessoa
                                percentualTotalPessoa = (infoPessoa("TotalGeral") / totalGeralVTotal) * 100
                                
                                ' Determinar classe da posição
                                Dim classePosicao
                                Select Case posicao
                                    Case 1: classePosicao = "posicao-1"
                                    Case 2: classePosicao = "posicao-2"
                                    Case 3: classePosicao = "posicao-3"
                                    Case Else
                                        If posicao <= 10 Then
                                            classePosicao = "posicao-top"
                                        Else
                                            classePosicao = ""
                                        End If
                                End Select
                        %>
                        <tr class="total-pessoa <%= classePosicao %>">
                            <td class="text-center-v">
                                <strong><%= posicao %></strong>
                            </td>
                            <td class="text-center-v">
                                <% If posicao <= 3 Then %>
                                    <span class="badge 
                                    <% 
                                    Select Case posicao
                                        Case 1: Response.Write "bg-warning text-dark"
                                        Case 2: Response.Write "bg-secondary"
                                        Case 3: Response.Write "bg-danger"
                                    End Select 
                                    %>">
                                        <% If posicao = 1 Then %>
                                            <i class="fas fa-trophy"></i> 1º
                                        <% ElseIf posicao = 2 Then %>
                                            <i class="fas fa-medal"></i> 2º
                                        <% ElseIf posicao = 3 Then %>
                                            <i class="fas fa-award"></i> 3º
                                        <% End If %>
                                    </span>
                                <% Else %>
                                    <span class="badge bg-light text-dark"><%= posicao %>º</span>
                                <% End If %>
                            </td>
                            <td colspan="2" class="text-center-v">
                                <strong><%= UCase(pessoaKey) %></strong>
                            </td>
                            <td class="text-right-v bg-success text-white">
                                <strong><%= FormatNumber(infoPessoa("TotalGeral"), 2) %></strong>
                            </td>
                            <td class="text-right-v bg-primary text-white">
                                <strong>100%</strong>
                            </td>
                            <td class="text-right-v bg-info text-white">
                                <strong><%= FormatNumber(infoPessoa("TotalGeral"), 2) %></strong>
                            </td>
                            <td class="text-right-v bg-warning text-dark">
                                <strong>100%</strong>
                            </td>
                            <td class="text-right-v bg-secondary text-white">
                                <strong><%= FormatNumber(percentualTotalPessoa, 2) %>%</strong>
                            </td>
                        </tr>
                        <%
                            Next
                        Else
                        %>
                        <tr>
                            <td colspan="9" class="text-center-v">Nenhum dado encontrado para o ano <%= filtroAno %>.</td>
                        </tr>
                        <%
                        End If
                        %>
                    </tbody>
                    <tfoot>
                        <tr class="total-geral">
                            <td colspan="4" class="text-center-v">
                                <strong>TOTAL GERAL - ANO <%= filtroAno %></strong>
                            </td>
                            <td class="text-right-v">
                                <strong><%= FormatNumber(totalGeralVTotal, 2) %></strong>
                            </td>
                            <td class="text-right-v">
                                <strong>-</strong>
                            </td>
                            <td class="text-right-v">
                                <strong><%= FormatNumber(totalGeralVTotal, 2) %></strong>
                            </td>
                            <td class="text-right-v">
                                <strong>-</strong>
                            </td>
                            <td class="text-right-v">
                                <strong>100%</strong>
                            </td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <!-- Top 10 Ranking -->
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">
                        <i class="fas fa-trophy text-warning"></i> Top 10 - Maiores Rendimentos
                    </h4>
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>Posição</th>
                                    <th>Pessoa</th>
                                    <th class="text-right-v">Total (R$)</th>
                                    <th class="text-right-v">% do Total</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If dadosPessoas.Count > 0 Then
                                    Dim contadorTop
                                    contadorTop = 0
                                    For Each pessoaKey In arrPessoas
                                        If contadorTop < 10 Then
                                            contadorTop = contadorTop + 1
                                            Set infoPessoa = dadosPessoas(pessoaKey)
                                            Dim percentualTop
                                            percentualTop = (infoPessoa("TotalGeral") / totalGeralVTotal) * 100
                                            Dim classeTop
                                            Select Case contadorTop
                                                Case 1: classeTop = "table-warning"
                                                Case 2: classeTop = "table-secondary"
                                                Case 3: classeTop = "table-danger"
                                                Case Else: classeTop = ""
                                            End Select
                                %>
                                <tr class="<%= classeTop %>">
                                    <td>
                                        <strong>
                                        <% If contadorTop = 1 Then %>
                                            <i class="fas fa-trophy text-warning"></i>
                                        <% ElseIf contadorTop = 2 Then %>
                                            <i class="fas fa-medal text-secondary"></i>
                                        <% ElseIf contadorTop = 3 Then %>
                                            <i class="fas fa-award text-danger"></i>
                                        <% Else %>
                                            <i class="fas fa-hashtag text-muted"></i>
                                        <% End If %>
                                        <%= contadorTop %>º
                                        </strong>
                                    </td>
                                    <td><strong><%= pessoaKey %></strong></td>
                                    <td class="text-right-v"><strong><%= FormatNumber(infoPessoa("TotalGeral"), 2) %></strong></td>
                                    <td class="text-right-v"><strong><%= FormatNumber(percentualTop, 2) %>%</strong></td>
                                </tr>
                                <%
                                        Else
                                            Exit For
                                        End If
                                    Next
                                Else
                                %>
                                <tr>
                                    <td colspan="4" class="text-center">Nenhum dado disponível</td>
                                </tr>
                                <%
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">
                        <i class="fas fa-layer-group text-info"></i> Pessoas com Múltiplos Cargos
                    </h4>
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>Pessoa</th>
                                    <th class="text-center-v">Nº de Cargos</th>
                                    <th class="text-right-v">Total (R$)</th>
                                    <th class="text-right-v">Posição</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If dadosPessoas.Count > 0 Then
                                    Dim pessoasMultiplosCargos
                                    Set pessoasMultiplosCargos = Server.CreateObject("Scripting.Dictionary")
                                    
                                    For Each pessoaKey In dadosPessoas.Keys
                                        Set infoPessoa = dadosPessoas(pessoaKey)
                                        If infoPessoa("Cargos").Count > 1 Then
                                            ' Encontrar a posição da pessoa no ranking
                                            Dim posicaoPessoa
                                            posicaoPessoa = 0
                                            For i = 0 To UBound(arrPessoas)
                                                If arrPessoas(i) = pessoaKey Then
                                                    posicaoPessoa = i + 1
                                                    Exit For
                                                End If
                                            Next
                                            pessoasMultiplosCargos.Add pessoaKey, Array(infoPessoa("Cargos").Count, infoPessoa("TotalGeral"), posicaoPessoa)
                                        End If
                                    Next
                                    
                                    If pessoasMultiplosCargos.Count > 0 Then
                                        Dim arrMultiplos, pessoaMultiplo
                                        arrMultiplos = pessoasMultiplosCargos.Keys
                                        
                                        ' Ordenar por número de cargos (decrescente)
                                        For i = 0 To UBound(arrMultiplos)
                                            For j = i + 1 To UBound(arrMultiplos)
                                                If pessoasMultiplosCargos(arrMultiplos(j))(0) > pessoasMultiplosCargos(arrMultiplos(i))(0) Then
                                                    temp = arrMultiplos(i)
                                                    arrMultiplos(i) = arrMultiplos(j)
                                                    arrMultiplos(j) = temp
                                                End If
                                            Next
                                        Next
                                        
                                        For Each pessoaMultiplo In arrMultiplos
                                            Dim infoMultiplo, numCargos, totalMultiplo, posicaoMultiplo
                                            infoMultiplo = pessoasMultiplosCargos(pessoaMultiplo)
                                            numCargos = infoMultiplo(0)
                                            totalMultiplo = infoMultiplo(1)
                                            posicaoMultiplo = infoMultiplo(2)
                                %>
                                <tr>
                                    <td><%= pessoaMultiplo %></td>
                                    <td class="text-center-v">
                                        <span class="badge bg-primary"><%= numCargos %> cargos</span>
                                    </td>
                                    <td class="text-right-v"><strong><%= FormatNumber(totalMultiplo, 2) %></strong></td>
                                    <td class="text-center-v">
                                        <span class="badge 
                                        <% 
                                        If posicaoMultiplo <= 3 Then
                                            Response.Write "bg-warning text-dark"
                                        ElseIf posicaoMultiplo <= 10 Then
                                            Response.Write "bg-info"
                                        Else
                                            Response.Write "bg-secondary"
                                        End If 
                                        %>">
                                            <%= posicaoMultiplo %>º
                                        </span>
                                    </td>
                                </tr>
                                <%
                                        Next
                                    Else
                                %>
                                <tr>
                                    <td colspan="4" class="text-center">Nenhuma pessoa com múltiplos cargos</td>
                                </tr>
                                <%
                                    End If
                                Else
                                %>
                                <tr>
                                    <td colspan="4" class="text-center">Nenhum dado disponível</td>
                                </tr>
                                <%
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <% End If %>
    </div>

    <!-- Scripts -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <!-- DataTables -->
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.print.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.5.0/js/dataTables.responsive.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.5.0/js/responsive.bootstrap5.min.js"></script>

    <script>
    $(document).ready(function() {
        // Inicializar DataTable
        var table = $('#xxxxxtabelaPessoas').DataTable({
            dom: '<"row"<"col-md-6"B><"col-md-6"f>>rt<"row"<"col-md-6"l><"col-md-6"p>>',
            buttons: [
                {
                    extend: 'copy',
                    className: 'btn btn-secondary dt-button',
                    text: '<i class="fas fa-copy"></i> Copiar'
                },
                {
                    extend: 'excel',
                    className: 'btn btn-success dt-button',
                    text: '<i class="fas fa-file-excel"></i> Excel'
                },
                {
                    extend: 'pdf',
                    className: 'btn btn-danger dt-button',
                    text: '<i class="fas fa-file-pdf"></i> PDF'
                },
                {
                    extend: 'print',
                    className: 'btn btn-info dt-button',
                    text: '<i class="fas fa-print"></i> Imprimir'
                }
            ],
            language: {
                url: '//cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json'
            },
            pageLength: 50,
            responsive: true,
            order: [[0, 'asc']], // Ordenar por posição (já ordenado no servidor)
            columnDefs: [
                {
                    targets: '_all',
                    className: 'text-center-v'
                },
                {
                    targets: [4, 5, 6, 7, 8], // Colunas numéricas
                    className: 'text-right-v'
                }
            ]
        });

        // Ajustar altura da tabela
        $('.dataTables_scrollBody').css('max-height', '700px');
    });
    </script>
</body>
</html>

<%
' ===============================================
' FUNÇÃO PARA EXIBIR LINHA DE CARGO
' ===============================================
Sub ExibirLinhaCargo(nomePessoa, cargo, valorCargo, totalPessoa, primeiroCargo, linhaCargoCount, posicaoPessoa)
    Dim percentualTotal, percentualPessoa, classeCargo
    
    percentualTotal = (valorCargo / totalGeralVTotal) * 100
    percentualPessoa = (valorCargo / totalPessoa) * 100
    
    ' Determinar classe CSS baseada no cargo
    Select Case LCase(cargo)
        Case "corretor"
            classeCargo = "cargo-corretor"
        Case "gerente"
            classeCargo = "cargo-gerente"
        Case "diretor"
            classeCargo = "cargo-diretor"
        Case Else
            classeCargo = "cargo-outros"
    End Select
%>
<tr class="cargo-row <%= classeCargo %>">
    <td class="text-center-v">
        <% If primeiroCargo Then %>
            <strong><%= posicaoPessoa %></strong>
        <% End If %>
    </td>
    <td class="text-center-v">
        <% If primeiroCargo Then %>
            <span class="badge 
            <% 
            If posicaoPessoa <= 3 Then
                Select Case posicaoPessoa
                    Case 1: Response.Write "bg-warning text-dark"
                    Case 2: Response.Write "bg-secondary"
                    Case 3: Response.Write "bg-danger"
                End Select
            ElseIf posicaoPessoa <= 10 Then
                Response.Write "bg-info"
            Else
                Response.Write "bg-light text-dark"
            End If 
            %>">
                <%= posicaoPessoa %>º
            </span>
        <% End If %>
    </td>
    <td class="pessoa-header">
        <% If primeiroCargo Then %>
            <strong><%= nomePessoa %></strong>
        <% End If %>
    </td>
    <td class="text-center-v">
        <span class="badge 
        <% 
        Select Case LCase(cargo)
            Case "corretor"
                Response.Write "bg-success"
            Case "gerente"
                Response.Write "bg-primary"
            Case "diretor"
                Response.Write "badge-bg-purple"
            Case Else
                Response.Write "bg-warning text-dark"
        End Select
        %>">
            <%= cargo %>
        </span>
    </td>
    <td class="text-right-v valor-cell"><strong><%= FormatNumber(valorCargo, 2) %></strong></td>
    <td class="text-right-v"><%= FormatNumber(percentualTotal, 2) %>%</td>
    <td class="text-right-v">
        <% If primeiroCargo Then %>
            <strong><%= FormatNumber(totalPessoa, 2) %></strong>
        <% End If %>
    </td>
    <td class="text-right-v"><%= FormatNumber(percentualPessoa, 2) %>%</td>
    <td class="text-right-v">
        <% If primeiroCargo Then %>
            <strong><%= FormatNumber((totalPessoa / totalGeralVTotal) * 100, 2) %>%</strong>
        <% End If %>
    </td>
</tr>
<%
End Sub

' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>