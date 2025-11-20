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

Dim filtroAno, filtroMes
filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")

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

Dim uniqueAnos, uniqueMeses
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueMeses = GetUniqueValues("Vendas", "MesVenda", "WHERE MesVenda IS NOT NULL")

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Março"
arrMesesNome(4) = "Abril"
arrMesesNome(5) = "Maio"
arrMesesNome(6) = "Junho"
arrMesesNome(7) = "Julho"
arrMesesNome(8) = "Agosto"
arrMesesNome(9) = "Setembro"
arrMesesNome(10) = "Outubro"
arrMesesNome(11) = "Novembro"
arrMesesNome(12) = "Dezembro"

' ===============================================
' OBTER DADOS DE COMISSÕES (APENAS SE ANO ESTIVER PREENCHIDO)
' ===============================================

Dim comissoesData, totalGeralComissoes, totalCorretores, totalVendasGeral, totalVGVGeral
Set comissoesData = Server.CreateObject("Scripting.Dictionary")

' Variáveis para as novas tabelas
Dim diretoriaData, gerenciaData
Set diretoriaData = Server.CreateObject("Scripting.Dictionary")
Set gerenciaData = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Construir consulta SQL para corretores
    Dim sqlComissoes, rsComissoes
    sqlComissoes = "SELECT " & _
                   "Corretor, " & _
                   "MesVenda, " & _
                   "SUM(ValorCorretor) as TotalComissao, " & _
                   "COUNT(*) as TotalVendas, " & _
                   "SUM(ValorUnidade) as TotalVGV " & _
                   "FROM Vendas " & _
                   "WHERE Excluido = 0 " & _
                   "AND AnoVenda = " & filtroAno
    
    If filtroMes <> "" Then
        sqlComissoes = sqlComissoes & " AND MesVenda = " & filtroMes
    End If
    
    sqlComissoes = sqlComissoes & " GROUP BY Corretor, MesVenda " & _
                   "ORDER BY Corretor, MesVenda"

    Set rsComissoes = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsComissoes.Open sqlComissoes, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de comissões: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlComissoes)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados de comissões dos corretores
    totalGeralComissoes = 0
    totalCorretores = 0
    totalVendasGeral = 0
    totalVGVGeral = 0
    
    Dim corretoresProcessados
    Set corretoresProcessados = Server.CreateObject("Scripting.Dictionary")

    If Not rsComissoes.EOF Then
        Do While Not rsComissoes.EOF
            Dim corretor, mes, comissaoMes, vendasMes, vgvMes
            corretor = CStr(rsComissoes("Corretor"))
            mes = CStr(rsComissoes("MesVenda"))
            comissaoMes = CDbl(rsComissoes("TotalComissao"))
            vendasMes = CLng(rsComissoes("TotalVendas"))
            vgvMes = CDbl(rsComissoes("TotalVGV"))
            
            ' Adicionar corretor à lista de processados
            If Not corretoresProcessados.Exists(corretor) Then
                corretoresProcessados.Add corretor, 1
                totalCorretores = totalCorretores + 1
            End If
            
            ' *** CORREÇÃO: Verificar se o corretor já existe antes de adicionar ***
            If Not comissoesData.Exists(corretor) Then
                Dim infoCorretor
                Set infoCorretor = Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "TotalComissao", 0
                infoCorretor.Add "TotalVendas", 0
                infoCorretor.Add "TotalVGV", 0
                comissoesData.Add corretor, infoCorretor
            End If
            
            ' Atualizar dados do mês
            Set infoCorretor = comissoesData(corretor)
            
            ' *** CORREÇÃO: Verificar se o mês já existe antes de adicionar ***
            If Not infoCorretor("Meses").Exists(mes) Then
                infoCorretor("Meses").Add mes, Array(comissaoMes, vendasMes, vgvMes)
            Else
                ' Se o mês já existe, somar os valores (caso haja duplicatas na query)
                Dim dadosMesExistentes
                dadosMesExistentes = infoCorretor("Meses")(mes)
                dadosMesExistentes(0) = dadosMesExistentes(0) + comissaoMes
                dadosMesExistentes(1) = dadosMesExistentes(1) + vendasMes
                dadosMesExistentes(2) = dadosMesExistentes(2) + vgvMes
                infoCorretor("Meses")(mes) = dadosMesExistentes
            End If
            
            ' Atualizar totais do corretor
            infoCorretor("TotalComissao") = infoCorretor("TotalComissao") + comissaoMes
            infoCorretor("TotalVendas") = infoCorretor("TotalVendas") + vendasMes
            infoCorretor("TotalVGV") = infoCorretor("TotalVGV") + vgvMes
            
            ' Atualizar totais gerais
            totalGeralComissoes = totalGeralComissoes + comissaoMes
            totalVendasGeral = totalVendasGeral + vendasMes
            totalVGVGeral = totalVGVGeral + vgvMes
            
            rsComissoes.MoveNext
        Loop
    End If

    If rsComissoes.State = 1 Then rsComissoes.Close
    Set rsComissoes = Nothing
    
    ' ===============================================
    ' CONSULTAR DADOS DA DIRETORIA POR NOME
    ' ===============================================
    
    ' Consulta para Diretoria agrupada por nome
    Dim sqlDiretoria, rsDiretoria
    sqlDiretoria = "SELECT " & _
                   "Diretoria, " & _
                   "MesVenda, " & _
                   "SUM(ValorDiretoria) as TotalDiretoria, " & _
                   "COUNT(*) as TotalVendas, " & _
                   "SUM(ValorUnidade) as TotalVGV " & _
                   "FROM Vendas " & _
                   "WHERE Excluido = 0 " & _
                   "AND AnoVenda = " & filtroAno & _
                   " AND Diretoria IS NOT NULL "
    
    If filtroMes <> "" Then
        sqlDiretoria = sqlDiretoria & " AND MesVenda = " & filtroMes
    End If
    
    sqlDiretoria = sqlDiretoria & " GROUP BY Diretoria, MesVenda " & _
                   "ORDER BY Diretoria, MesVenda"

    Set rsDiretoria = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsDiretoria.Open sqlDiretoria, connSales

    If Err.Number = 0 Then
        ' Processar dados da Diretoria por nome
        If Not rsDiretoria.EOF Then
            Do While Not rsDiretoria.EOF
                Dim nomeDiretoria, mesDiretoria, valorDiretoriaMes, vendasDiretoriaMes, vgvDiretoriaMes
                nomeDiretoria = CStr(rsDiretoria("Diretoria"))
                mesDiretoria = CStr(rsDiretoria("MesVenda"))
                valorDiretoriaMes = CDbl(rsDiretoria("TotalDiretoria"))
                vendasDiretoriaMes = CLng(rsDiretoria("TotalVendas"))
                vgvDiretoriaMes = CDbl(rsDiretoria("TotalVGV"))
                
                ' Verificar se o nome da diretoria já existe
                If Not diretoriaData.Exists(nomeDiretoria) Then
                    Dim infoDiretoria
                    Set infoDiretoria = Server.CreateObject("Scripting.Dictionary")
                    infoDiretoria.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                    infoDiretoria.Add "TotalComissao", 0
                    infoDiretoria.Add "TotalVendas", 0
                    infoDiretoria.Add "TotalVGV", 0
                    diretoriaData.Add nomeDiretoria, infoDiretoria
                End If
                
                ' Atualizar dados do mês
                Set infoDiretoria = diretoriaData(nomeDiretoria)
                
                ' Adicionar/atualizar dados do mês
                If Not infoDiretoria("Meses").Exists(mesDiretoria) Then
                    infoDiretoria("Meses").Add mesDiretoria, Array(valorDiretoriaMes, vendasDiretoriaMes, vgvDiretoriaMes)
                Else
                    Dim dadosMesExistentesDiretoria
                    dadosMesExistentesDiretoria = infoDiretoria("Meses")(mesDiretoria)
                    dadosMesExistentesDiretoria(0) = dadosMesExistentesDiretoria(0) + valorDiretoriaMes
                    dadosMesExistentesDiretoria(1) = dadosMesExistentesDiretoria(1) + vendasDiretoriaMes
                    dadosMesExistentesDiretoria(2) = dadosMesExistentesDiretoria(2) + vgvDiretoriaMes
                    infoDiretoria("Meses")(mesDiretoria) = dadosMesExistentesDiretoria
                End If
                
                ' Atualizar totais da diretoria
                infoDiretoria("TotalComissao") = infoDiretoria("TotalComissao") + valorDiretoriaMes
                infoDiretoria("TotalVendas") = infoDiretoria("TotalVendas") + vendasDiretoriaMes
                infoDiretoria("TotalVGV") = infoDiretoria("TotalVGV") + vgvDiretoriaMes
                
                rsDiretoria.MoveNext
            Loop
        End If
    End If
    
    If rsDiretoria.State = 1 Then rsDiretoria.Close
    Set rsDiretoria = Nothing
    
    ' ===============================================
    ' CONSULTAR DADOS DA GERÊNCIA POR NOME
    ' ===============================================
    
    ' Consulta para Gerência agrupada por nome
    Dim sqlGerencia, rsGerencia
    sqlGerencia = "SELECT " & _
                  "Gerencia, " & _
                  "MesVenda, " & _
                  "SUM(ValorGerencia) as TotalGerencia, " & _
                  "COUNT(*) as TotalVendas, " & _
                  "SUM(ValorUnidade) as TotalVGV " & _
                  "FROM Vendas " & _
                  "WHERE Excluido = 0 " & _
                  "AND AnoVenda = " & filtroAno & _
                  " AND Gerencia IS NOT NULL "
    
    If filtroMes <> "" Then
        sqlGerencia = sqlGerencia & " AND MesVenda = " & filtroMes
    End If
    
    sqlGerencia = sqlGerencia & " GROUP BY Gerencia, MesVenda " & _
                  "ORDER BY Gerencia, MesVenda"

    Set rsGerencia = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsGerencia.Open sqlGerencia, connSales

    If Err.Number = 0 Then
        ' Processar dados da Gerência por nome
        If Not rsGerencia.EOF Then
            Do While Not rsGerencia.EOF
                Dim nomeGerencia, mesGerencia, valorGerenciaMes, vendasGerenciaMes, vgvGerenciaMes
                nomeGerencia = CStr(rsGerencia("Gerencia"))
                mesGerencia = CStr(rsGerencia("MesVenda"))
                valorGerenciaMes = CDbl(rsGerencia("TotalGerencia"))
                vendasGerenciaMes = CLng(rsGerencia("TotalVendas"))
                vgvGerenciaMes = CDbl(rsGerencia("TotalVGV"))
                
                ' Verificar se o nome da gerência já existe
                If Not gerenciaData.Exists(nomeGerencia) Then
                    Dim infoGerencia
                    Set infoGerencia = Server.CreateObject("Scripting.Dictionary")
                    infoGerencia.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                    infoGerencia.Add "TotalComissao", 0
                    infoGerencia.Add "TotalVendas", 0
                    infoGerencia.Add "TotalVGV", 0
                    gerenciaData.Add nomeGerencia, infoGerencia
                End If
                
                ' Atualizar dados do mês
                Set infoGerencia = gerenciaData(nomeGerencia)
                
                ' Adicionar/atualizar dados do mês
                If Not infoGerencia("Meses").Exists(mesGerencia) Then
                    infoGerencia("Meses").Add mesGerencia, Array(valorGerenciaMes, vendasGerenciaMes, vgvGerenciaMes)
                Else
                    Dim dadosMesExistentesGerencia
                    dadosMesExistentesGerencia = infoGerencia("Meses")(mesGerencia)
                    dadosMesExistentesGerencia(0) = dadosMesExistentesGerencia(0) + valorGerenciaMes
                    dadosMesExistentesGerencia(1) = dadosMesExistentesGerencia(1) + vendasGerenciaMes
                    dadosMesExistentesGerencia(2) = dadosMesExistentesGerencia(2) + vgvGerenciaMes
                    infoGerencia("Meses")(mesGerencia) = dadosMesExistentesGerencia
                End If
                
                ' Atualizar totais da gerência
                infoGerencia("TotalComissao") = infoGerencia("TotalComissao") + valorGerenciaMes
                infoGerencia("TotalVendas") = infoGerencia("TotalVendas") + vendasGerenciaMes
                infoGerencia("TotalVGV") = infoGerencia("TotalVGV") + vgvGerenciaMes
                
                rsGerencia.MoveNext
            Loop
        End If
    End If
    
    If rsGerencia.State = 1 Then rsGerencia.Close
    Set rsGerencia = Nothing
    
End If ' *** CORREÇÃO: Fechando o If filtroAno <> "" ***
%>


<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Comissões dos Corretores</title>
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
        .corretor-header {
            background-color: #e9ecef !important;
            font-weight: bold;
        }
        .mes-header {
            background-color: #17a2b8;
            color: white;
            font-weight: bold;
        }
        .total-row {
            background-color: #800000;
            color: white;
            font-weight: bold;
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
        .comissao-cell {
            font-weight: bold;
            color: #28a745;
        }
        .media-anual-cell {
            font-weight: bold;
            color: #1D4C7F;
            background-color: #e3f2fd !important;
        }
        .media-real-cell {
            font-weight: bold;
            color: #ff6b00;
            background-color: #fff3e0 !important;
        }
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
        .table-diretoria th {
            background-color: #2c3e50 !important;
        }
        .table-gerencia th {
            background-color: #34495e !important;
        }
        .diretoria-row {
            background-color: #e8f4f8 !important;
        }
        .gerencia-row {
            background-color: #f0e8f8 !important;
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
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-money-bill-wave"></i> SGVendas - Comissões dos Corretores
        </h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-4">
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
                        <div class="filter-label">Mês</div>
                        <select class="form-select" name="mes" id="mesFilter">
                            <option value="">Todos os meses</option>
                            <%
                            If IsArray(uniqueMeses) Then
                                For Each mes In uniqueMeses
                                    If Not IsEmpty(mes) Then
                                        Dim mesNum
                                        mesNum = CInt(mes)
                                        Response.Write "<option value=""" & mes & """"
                                        If CStr(filtroMes) = CStr(mes) Then Response.Write " selected"
                                        Response.Write ">" & arrMesesNome(mesNum) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-4">
                        <div class="filter-label">&nbsp;</div>
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Relatório
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Then %>
            <div class="alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar o relatório de comissões.
            </div>
        <% Else %>
        
        <!-- KPIs Principais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-money-bill-wave"></i>
                    <h5>Total Comissões <%= filtroAno %></h5>
                    <p><%= FormatNumber(totalGeralComissoes, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <i class="fas fa-user-tie"></i>
                    <h5>Corretores com Vendas</h5>
                    <p><%= totalCorretores %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-calendar-alt"></i>
                    <h5>Período</h5>
                    <p>
                        <% 
                        If filtroMes <> "" Then 
                            Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                        Else 
                            Response.Write "Ano " & filtroAno & " (Todos os meses)"
                        End If 
                        %>
                    </p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-chart-line"></i>
                    <h5>Comissão Média</h5>
                    <p>
                        <% 
                        If totalCorretores > 0 Then 
                            Response.Write "" & FormatNumber(totalGeralComissoes / totalCorretores, 2)
                        Else 
                            Response.Write "0,00"
                        End If 
                        %>
                    </p>
                </div>
            </div>
        </div>

        <!-- Tabela de Comissões dos Corretores -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">
                Comissões por Corretor - 
                <% 
                If filtroMes <> "" Then 
                    Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                Else 
                    Response.Write "Ano " & filtroAno
                End If 
                %>
            </h3>
            
            <div class="table-responsive" style="overflow-y: auto;">
                <table id="tabelaComissoes" class="table table-striped table-hover table-bordered" style="width:100%">
                    <thead>
                        <tr>
                            <th class="text-center-v">Corretor</th>
                            <%
                            ' Cabeçalhos dos meses (apenas se não tiver filtro de mês específico)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Response.Write "<th class='text-center-v mes-header'>" & Left(arrMesesNome(i), 3) & "</th>"
                                Next
                            Else
                                Response.Write "<th class='text-center-v mes-header'>Comissão " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>Vendas " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>VGV " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                            End If
                            %>
                            <th class="text-center-v bg-success text-white">Total Comissão</th>
                            <th class="text-center-v bg-primary text-white">Média Anual</th>
                            <th class="text-center-v bg-warning text-dark">Média Real</th>
                            <th class="text-center-v bg-info text-white">Total Vendas</th>
                            <th class="text-center-v bg-secondary text-white">Total VGV</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        If comissoesData.Count > 0 Then
                            Dim arrCorretores, corretorKey
                            arrCorretores = comissoesData.Keys
                            
                            ' Ordenar corretores por total de comissão (decrescente)
                            For i = 0 To UBound(arrCorretores)
                                For j = i + 1 To UBound(arrCorretores)
                                    If comissoesData(arrCorretores(j))("TotalComissao") > comissoesData(arrCorretores(i))("TotalComissao") Then
                                        Dim temp
                                        temp = arrCorretores(i)
                                        arrCorretores(i) = arrCorretores(j)
                                        arrCorretores(j) = temp
                                    End If
                                Next
                            Next
                            
                            For Each corretorKey In arrCorretores
                                Set infoCorretor = comissoesData(corretorKey)
                                Dim mesesCorretor
                                Set mesesCorretor = infoCorretor("Meses")
                                
                                ' Calcular médias
                                Dim mediaAnual, mediaReal, mesesComVenda
                                mediaAnual = 0
                                mediaReal = 0
                                mesesComVenda = 0
                                
                                ' Contar meses com vendas
                                For i = 1 To 12
                                    If mesesCorretor.Exists(CStr(i)) Then
                                        mesesComVenda = mesesComVenda + 1
                                    End If
                                Next
                                
                                ' Média Anual: Total dividido por 12 (sempre)
                                mediaAnual = infoCorretor("TotalComissao") / 12
                                
                                ' Média Real: Total dividido pelos meses com vendas
                                If mesesComVenda > 0 Then
                                    mediaReal = infoCorretor("TotalComissao") / mesesComVenda
                                End If
                        %>
                        <tr>
                            <td class="corretor-header"><%= corretorKey %></td>
                            
                            <%
                            ' Dados dos meses
                            If filtroMes = "" Then
                                ' Mostrar todos os meses
                                For i = 1 To 12
                                    Dim mesKey
                                    mesKey = CStr(i)
                                    If mesesCorretor.Exists(mesKey) Then
                                        Dim dadosMes
                                        dadosMes = mesesCorretor(mesKey)
                                        Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMes(0), 2) & "</td>"
                                    Else
                                        Response.Write "<td class='text-center-v'>-</td>"
                                    End If
                                Next
                            Else
                                ' Mostrar apenas o mês filtrado com detalhes
                                If mesesCorretor.Exists(filtroMes) Then
                                    Dim dadosMesFiltrado
                                    dadosMesFiltrado = mesesCorretor(filtroMes)
                                    Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMesFiltrado(0), 2) & "</td>"
                                    Response.Write "<td class='text-center-v'>" & dadosMesFiltrado(1) & "</td>"
                                    Response.Write "<td class='text-right-v'>" & FormatNumber(dadosMesFiltrado(2), 2) & "</td>"
                                Else
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                End If
                            End If
                            %>
                            
                            <!-- Totais do corretor -->
                            <td class="text-right-v bg-success text-white"><strong><%= FormatNumber(infoCorretor("TotalComissao"), 2) %></strong></td>
                            <td class="text-right-v bg-primary text-black media-anual-cell"><strong><%= FormatNumber(mediaAnual, 2) %></strong></td>
                            <td class="text-right-v bg-warning text-dark media-real-cell"><strong><%= FormatNumber(mediaReal, 2) %></strong></td>
                            <td class="text-center-v bg-info text-white"><strong><%= infoCorretor("TotalVendas") %></strong></td>
                            <td class="text-right-v bg-secondary text-white"><strong><%= FormatNumber(infoCorretor("TotalVGV"), 2) %></strong></td>
                        </tr>
                        <%
                            Next
                        Else
                        %>
                        <tr>
                            <td 
                            <% If filtroMes = "" Then %>
                                colspan="17"
                            <% Else %>
                                colspan="9"
                            <% End If %>
                            class="text-center-v">Nenhum dado encontrado para os filtros selecionados.</td>
                        </tr>
                        <%
                        End If
                        %>
                    </tbody>
                    <tfoot>
                        <tr class="total-row">
                            <td><strong>TOTAIS GERAIS</strong></td>
                            <%
                            ' Totais por mês (apenas se não tiver filtro de mês)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Dim totalMes
                                    totalMes = 0
                                    For Each corretorKey In comissoesData.Keys
                                        Set mesesCorretor = comissoesData(corretorKey)("Meses")
                                        If mesesCorretor.Exists(CStr(i)) Then
                                            totalMes = totalMes + mesesCorretor(CStr(i))(0)
                                        End If
                                    Next
                                    Response.Write "<td class='text-right-v'><strong>" & FormatNumber(totalMes, 2) & "</strong></td>"
                                Next
                            Else
                                Response.Write "<td colspan='3'></td>"
                            End If
                            %>
                            <td class="text-right-v"><strong><%= FormatNumber(totalGeralComissoes, 2) %></strong></td>
                            <td class="text-right-v">
                                <strong>
                                <%
                                ' Média Anual Geral: Total dividido por 12
                                Dim mediaAnualGeral
                                mediaAnualGeral = totalGeralComissoes / 12
                                Response.Write FormatNumber(mediaAnualGeral, 2)
                                %>
                                </strong>
                            </td>
                            <td class="text-right-v">
                                <strong>
                                <%
                                ' Média Real Geral: Total dividido pelos meses com vendas
                                Dim mediaRealGeral, totalMesesComVenda
                                totalMesesComVenda = 0
                                
                                If filtroMes = "" Then
                                    ' Contar meses totais com vendas (considerando todos os corretores)
                                    For i = 1 To 12
                                        Dim mesTemVenda
                                        mesTemVenda = False
                                        For Each corretorKey In comissoesData.Keys
                                            Set mesesCorretor = comissoesData(corretorKey)("Meses")
                                            If mesesCorretor.Exists(CStr(i)) Then
                                                mesTemVenda = True
                                                Exit For
                                            End If
                                        Next
                                        If mesTemVenda Then
                                            totalMesesComVenda = totalMesesComVenda + 1
                                        End If
                                    Next
                                Else
                                    ' Quando há filtro de mês, considera apenas 1 mês
                                    totalMesesComVenda = 1
                                End If
                                
                                If totalMesesComVenda > 0 Then
                                    mediaRealGeral = totalGeralComissoes / totalMesesComVenda
                                Else
                                    mediaRealGeral = 0
                                End If
                                Response.Write FormatNumber(mediaRealGeral, 2)
                                %>
                                </strong>
                            </td>
                            <td class="text-center-v"><strong><%= totalVendasGeral %></strong></td>
                            <td class="text-right-v"><strong><%= FormatNumber(totalVGVGeral, 2) %></strong></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <!-- NOVA TABELA: COMISSÃO DA DIRETORIA -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">
                <i class="fas fa-crown"></i> Comissão da Diretoria - 
                <% 
                If filtroMes <> "" Then 
                    Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                Else 
                    Response.Write "Ano " & filtroAno
                End If 
                %>
            </h3>
            
            <div class="table-responsive" style="overflow-y: auto;">
                <table class="table table-striped table-hover table-bordered table-diretoria" style="width:100%">
                    <thead>
                        <tr>
                            <th class="text-center-v">Diretoria</th>
                            <%
                            ' Cabeçalhos dos meses (apenas se não tiver filtro de mês específico)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Response.Write "<th class='text-center-v mes-header'>" & Left(arrMesesNome(i), 3) & "</th>"
                                Next
                            Else
                                Response.Write "<th class='text-center-v mes-header'>Comissão " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>Vendas " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>VGV " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                            End If
                            %>
                            <th class="text-center-v bg-success text-white">Total Comissão</th>
                            <th class="text-center-v bg-primary text-white">Média Anual</th>
                            <th class="text-center-v bg-warning text-dark">Média Real</th>
                            <th class="text-center-v bg-info text-white">Total Vendas</th>
                            <th class="text-center-v bg-secondary text-white">Total VGV</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        If diretoriaData.Count > 0 Then
                            Dim arrDiretoria, diretoriaKey
                            arrDiretoria = diretoriaData.Keys
                            
                            ' Ordenar diretorias por total de comissão (decrescente)
                            For i = 0 To UBound(arrDiretoria)
                                For j = i + 1 To UBound(arrDiretoria)
                                    If diretoriaData(arrDiretoria(j))("TotalComissao") > diretoriaData(arrDiretoria(i))("TotalComissao") Then
                                        temp = arrDiretoria(i)
                                        arrDiretoria(i) = arrDiretoria(j)
                                        arrDiretoria(j) = temp
                                    End If
                                Next
                            Next
                            
                            For Each diretoriaKey In arrDiretoria
                                Set infoDiretoria = diretoriaData(diretoriaKey)
                                Dim mesesDiretoria
                                Set mesesDiretoria = infoDiretoria("Meses")
                                
                                ' Calcular médias
                                Dim mediaAnualDiretoria, mediaRealDiretoria, mesesComVendaDiretoria
                                mediaAnualDiretoria = 0
                                mediaRealDiretoria = 0
                                mesesComVendaDiretoria = 0
                                
                                ' Contar meses com vendas
                                For i = 1 To 12
                                    If mesesDiretoria.Exists(CStr(i)) Then
                                        mesesComVendaDiretoria = mesesComVendaDiretoria + 1
                                    End If
                                Next
                                
                                ' Média Anual: Total dividido por 12 (sempre)
                                mediaAnualDiretoria = infoDiretoria("TotalComissao") / 12
                                
                                ' Média Real: Total dividido pelos meses com vendas
                                If mesesComVendaDiretoria > 0 Then
                                    mediaRealDiretoria = infoDiretoria("TotalComissao") / mesesComVendaDiretoria
                                End If
                        %>
                        <tr class="diretoria-row">
                            <td class="corretor-header"><%= diretoriaKey %></td>
                            
                            <%
                            ' Dados dos meses
                            If filtroMes = "" Then
                                ' Mostrar todos os meses
                                For i = 1 To 12
                                    mesKey = CStr(i)
                                    If mesesDiretoria.Exists(mesKey) Then
                                        Dim dadosMesDiretoria
                                        dadosMesDiretoria = mesesDiretoria(mesKey)
                                        Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMesDiretoria(0), 2) & "</td>"
                                    Else
                                        Response.Write "<td class='text-center-v'>-</td>"
                                    End If
                                Next
                            Else
                                ' Mostrar apenas o mês filtrado com detalhes
                                If mesesDiretoria.Exists(filtroMes) Then
                                    Dim dadosMesFiltradoDiretoria
                                    dadosMesFiltradoDiretoria = mesesDiretoria(filtroMes)
                                    Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMesFiltradoDiretoria(0), 2) & "</td>"
                                    Response.Write "<td class='text-center-v'>" & dadosMesFiltradoDiretoria(1) & "</td>"
                                    Response.Write "<td class='text-right-v'>" & FormatNumber(dadosMesFiltradoDiretoria(2), 2) & "</td>"
                                Else
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                End If
                            End If
                            %>
                            
                            <!-- Totais da Diretoria -->
                            <td class="text-right-v bg-success text-white"><strong><%= FormatNumber(infoDiretoria("TotalComissao"), 2) %></strong></td>
                            <td class="text-right-v bg-primary text-black media-anual-cell"><strong><%= FormatNumber(mediaAnualDiretoria, 2) %></strong></td>
                            <td class="text-right-v bg-warning text-dark media-real-cell"><strong><%= FormatNumber(mediaRealDiretoria, 2) %></strong></td>
                            <td class="text-center-v bg-info text-white"><strong><%= infoDiretoria("TotalVendas") %></strong></td>
                            <td class="text-right-v bg-secondary text-white"><strong><%= FormatNumber(infoDiretoria("TotalVGV"), 2) %></strong></td>
                        </tr>
                        <%
                            Next
                        Else
                        %>
                        <tr>
                            <td 
                            <% If filtroMes = "" Then %>
                                colspan="17"
                            <% Else %>
                                colspan="9"
                            <% End If %>
                            class="text-center-v">Nenhum dado encontrado para a Diretoria.</td>
                        </tr>
                        <%
                        End If
                        %>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- NOVA TABELA: COMISSÃO DA GERÊNCIA -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">
                <i class="fas fa-user-tie"></i> Comissão da Gerência - 
                <% 
                If filtroMes <> "" Then 
                    Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                Else 
                    Response.Write "Ano " & filtroAno
                End If 
                %>
            </h3>
            
            <div class="table-responsive" style="overflow-y: auto;">
                <table class="table table-striped table-hover table-bordered table-gerencia" style="width:100%">
                    <thead>
                        <tr>
                            <th class="text-center-v">Gerência</th>
                            <%
                            ' Cabeçalhos dos meses (apenas se não tiver filtro de mês específico)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Response.Write "<th class='text-center-v mes-header'>" & Left(arrMesesNome(i), 3) & "</th>"
                                Next
                            Else
                                Response.Write "<th class='text-center-v mes-header'>Comissão " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>Vendas " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>VGV " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                            End If
                            %>
                            <th class="text-center-v bg-success text-white">Total Comissão</th>
                            <th class="text-center-v bg-primary text-white">Média Anual</th>
                            <th class="text-center-v bg-warning text-dark">Média Real</th>
                            <th class="text-center-v bg-info text-white">Total Vendas</th>
                            <th class="text-center-v bg-secondary text-white">Total VGV</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        If gerenciaData.Count > 0 Then
                            Dim arrGerencia, gerenciaKey
                            arrGerencia = gerenciaData.Keys
                            
                            ' Ordenar gerências por total de comissão (decrescente)
                            For i = 0 To UBound(arrGerencia)
                                For j = i + 1 To UBound(arrGerencia)
                                    If gerenciaData(arrGerencia(j))("TotalComissao") > gerenciaData(arrGerencia(i))("TotalComissao") Then
                                        temp = arrGerencia(i)
                                        arrGerencia(i) = arrGerencia(j)
                                        arrGerencia(j) = temp
                                    End If
                                Next
                            Next
                            
                            For Each gerenciaKey In arrGerencia
                                Set infoGerencia = gerenciaData(gerenciaKey)
                                Dim mesesGerencia
                                Set mesesGerencia = infoGerencia("Meses")
                                
                                ' Calcular médias
                                Dim mediaAnualGerencia, mediaRealGerencia, mesesComVendaGerencia
                                mediaAnualGerencia = 0
                                mediaRealGerencia = 0
                                mesesComVendaGerencia = 0
                                
                                ' Contar meses com vendas
                                For i = 1 To 12
                                    If mesesGerencia.Exists(CStr(i)) Then
                                        mesesComVendaGerencia = mesesComVendaGerencia + 1
                                    End If
                                Next
                                
                                ' Média Anual: Total dividido por 12 (sempre)
                                mediaAnualGerencia = infoGerencia("TotalComissao") / 12
                                
                                ' Média Real: Total dividido pelos meses com vendas
                                If mesesComVendaGerencia > 0 Then
                                    mediaRealGerencia = infoGerencia("TotalComissao") / mesesComVendaGerencia
                                End If
                        %>
                        <tr class="gerencia-row">
                            <td class="corretor-header"><%= gerenciaKey %></td>
                            
                            <%
                            ' Dados dos meses
                            If filtroMes = "" Then
                                ' Mostrar todos os meses
                                For i = 1 To 12
                                    mesKey = CStr(i)
                                    If mesesGerencia.Exists(mesKey) Then
                                        Dim dadosMesGerencia
                                        dadosMesGerencia = mesesGerencia(mesKey)
                                        Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMesGerencia(0), 2) & "</td>"
                                    Else
                                        Response.Write "<td class='text-center-v'>-</td>"
                                    End If
                                Next
                            Else
                                ' Mostrar apenas o mês filtrado com detalhes
                                If mesesGerencia.Exists(filtroMes) Then
                                    Dim dadosMesFiltradoGerencia
                                    dadosMesFiltradoGerencia = mesesGerencia(filtroMes)
                                    Response.Write "<td class='text-right-v comissao-cell'>" & FormatNumber(dadosMesFiltradoGerencia(0), 2) & "</td>"
                                    Response.Write "<td class='text-center-v'>" & dadosMesFiltradoGerencia(1) & "</td>"
                                    Response.Write "<td class='text-right-v'>" & FormatNumber(dadosMesFiltradoGerencia(2), 2) & "</td>"
                                Else
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                End If
                            End If
                            %>
                            
                            <!-- Totais da Gerência -->
                            <td class="text-right-v bg-success text-white"><strong><%= FormatNumber(infoGerencia("TotalComissao"), 2) %></strong></td>
                            <td class="text-right-v bg-primary text-black media-anual-cell"><strong><%= FormatNumber(mediaAnualGerencia, 2) %></strong></td>
                            <td class="text-right-v bg-warning text-dark media-real-cell"><strong><%= FormatNumber(mediaRealGerencia, 2) %></strong></td>
                            <td class="text-center-v bg-info text-white"><strong><%= infoGerencia("TotalVendas") %></strong></td>
                            <td class="text-right-v bg-secondary text-white"><strong><%= FormatNumber(infoGerencia("TotalVGV"), 2) %></strong></td>
                        </tr>
                        <%
                            Next
                        Else
                        %>
                        <tr>
                            <td 
                            <% If filtroMes = "" Then %>
                                colspan="17"
                            <% Else %>
                                colspan="9"
                            <% End If %>
                            class="text-center-v">Nenhum dado encontrado para a Gerência.</td>
                        </tr>
                        <%
                        End If
                        %>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Resumo Estatístico -->
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">Top 5 Corretores (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>Posição</th>
                                    <th>Corretor</th>
                                    <th class="text-right-v">Comissão (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If comissoesData.Count > 0 Then
                                    Dim contador
                                    contador = 0
                                    For Each corretorKey In arrCorretores
                                        If contador < 5 Then
                                            Set infoCorretor = comissoesData(corretorKey)
                                %>
                                <tr>
                                    <td><%= contador + 1 %></td>
                                    <td><%= corretorKey %></td>
                                    <td class="text-right-v"><%= FormatNumber(infoCorretor("TotalComissao"), 2) %></td>
                                </tr>
                                <%
                                            contador = contador + 1
                                        Else
                                            Exit For
                                        End If
                                    Next
                                Else
                                %>
                                <tr>
                                    <td colspan="3" class="text-center">Nenhum dado disponível</td>
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
                    <h4 class="text-dark">Estatísticas</h4>
                    <%
                    If comissoesData.Count > 0 Then
                        Dim maiorComissao, menorComissao, corretorMaior, corretorMenor
                        maiorComissao = 0
                        menorComissao = 999999999
                        
                        For Each corretorKey In comissoesData.Keys
                            Set infoCorretor = comissoesData(corretorKey)
                            If infoCorretor("TotalComissao") > maiorComissao Then
                                maiorComissao = infoCorretor("TotalComissao")
                                corretorMaior = corretorKey
                            End If
                            If infoCorretor("TotalComissao") < menorComissao Then
                                menorComissao = infoCorretor("TotalComissao")
                                corretorMenor = corretorKey
                            End If
                        Next
                    %>
                    <p><strong>Maior Comissão:</strong><br>
                    <%= corretorMaior %> - <%= FormatNumber(maiorComissao, 2) %></p>
                    
                    <p><strong>Menor Comissão:</strong><br>
                    <%= corretorMenor %> - <%= FormatNumber(menorComissao, 2) %></p>
                    
                    <p><strong>Média de Comissões:</strong><br>
                    <%= FormatNumber(totalGeralComissoes / comissoesData.Count, 2) %></p>
                    <%
                    Else
                    %>
                    <p class="text-center">Nenhum dado disponível</p>
                    <%
                    End If
                    %>
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
        var table = $('#tabelaComissoes').DataTable({
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
            pageLength: 25,
            responsive: true,
            order: [[ 
                <% 
                ' Ordenar pela coluna Total Comissão (penúltima coluna quando sem filtro de mês)
                If filtroMes = "" Then 
                    Response.Write "16, 'desc'" ' Total Comissão é a 17ª coluna (índice 16)
                Else 
                    Response.Write "7, 'desc'" ' Total Comissão é a 8ª coluna (índice 7)
                End If 
                %> ],
            columnDefs: [
                {
                    targets: '_all',
                    className: 'text-center-v'
                },
                {
                    targets: [ 
                        <% 
                        ' Definir alinhamento à direita para colunas numéricas
                        If filtroMes = "" Then 
                            ' Colunas dos meses (1-12), Total Comissão (16), Médias (17-18), Total VGV (20)
                            Response.Write "1,2,3,4,5,6,7,8,9,10,11,12,16,17,18,20"
                        Else 
                            ' Colunas: Comissão (1), VGV (3), Total Comissão (4), Médias (5-6), Total VGV (8)
                            Response.Write "1,3,4,5,6,8"
                        End If 
                        %> ],
                    className: 'text-right-v'
                }
            ],
            initComplete: function() {
                // Adicionar controles personalizados acima da tabela
                this.api().columns().every(function() {
                    var column = this;
                    var title = $(column.header()).text();
                    
                    // Criar input para filtro em cada coluna
                    if (title !== 'Ações') {
                        $('<input type="text" placeholder="Filtrar ' + title + '" style="width: 100%; margin: 2px;"/>')
                            .appendTo($(column.header()))
                            .on('keyup change', function() {
                                if (column.search() !== this.value) {
                                    column.search(this.value).draw();
                                }
                            });
                    }
                });
            }
        });

        // Ajustar altura da tabela
        $('.dataTables_scrollBody').css('max-height', '600px');
    });
    </script>
</body>
</html>

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>