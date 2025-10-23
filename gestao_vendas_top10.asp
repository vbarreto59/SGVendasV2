<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%

' Função para executar consultas de valores únicos com tratamento de erros
Function GetUniqueValues(fieldName, empreendimentoNome)
    On Error Resume Next
    
    Dim sql, rs, result
    result = ""
    
    ' Corrigindo a consulta para usar JOIN correto ou WHERE apropriado
    sql = "SELECT DISTINCT Vendas." & fieldName & " " & _
          "FROM Vendas " & _
          "INNER JOIN Empreendimento E ON Vendas.Empreend_ID = E.Empreend_ID " & _
          "WHERE Vendas.Excluido = 0 " & _
          "AND E.NomeEmpreendimento = '" & Replace(empreendimentoNome, "'", "''") & "' " & _
          "ORDER BY Vendas." & fieldName
    
    Set rs = connSales.Execute(sql)
    
    If Err.Number <> 0 Then
        result = "Erro ao obter valores únicos: " & Err.Description & "<br>SQL Query: " & Server.HTMLEncode(sql)
        Err.Clear
    Else
        If Not rs.EOF Then
            Do While Not rs.EOF
                If Not IsNull(rs(fieldName)) Then
                    result = result & rs(fieldName) & "|"
                End If
                rs.MoveNext
            Loop
            ' Remove o último separador
            If Len(result) > 0 Then result = Left(result, Len(result)-1)
        Else
            result = "Nenhum dado encontrado"
        End If
    End If
    
    If IsObject(rs) Then rs.Close
    Set rs = Nothing
    
    GetUniqueValues = result
    On Error Goto 0
End Function



' ===============================================
' CONFIGURAÇÃO DE BANCOS DE DADOS E CONEXÕES v1
' ===============================================

' Extrair caminhos dos bancos de dados
Dim dbSunnyPath, dbSunSalesPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunSalesPath = Split(StrConnSales, "Data Source=")(1)

' Verificar se os caminhos foram extraídos corretamente
If dbSunnyPath = "" Or dbSunSalesPath = "" Then
    Response.Write "Erro: Não foi possível extrair caminhos dos bancos de dados"
    Response.End
End If

' Inicializar conexões
Dim connSales, connSunny
Set connSales = Server.CreateObject("ADODB.Connection")
Set connSunny = Server.CreateObject("ADODB.Connection")

On Error Resume Next
connSales.Open StrConnSales ' Conexão com o banco de dados Vendas
connSunny.Open StrConn      ' Conexão com o banco de dados Empreendimento
If Err.Number <> 0 Then
    Response.Write "Erro ao conectar aos bancos de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

' ===============================================
' FUNÇÕES E SUB-ROTINAS GLOBAIS
' ===============================================

' Função personalizada para formatar números para exibição com vírgula
Function FormatCurrencyValue(value, decimalPlaces)
    ' A função nativa FormatNumber retorna um ponto como separador decimal.
    ' Esta função substitui o ponto pela vírgula para o formato brasileiro.
    vValor = Replace(FormatNumber(value, decimalPlaces), ".", "")
    FormatCurrencyValue = FormatCurrency(value, 2)
End Function

' Função para processar e agregar dados de Vendas (VGV) em um dicionário
Sub ProcessaVendas(mainDict, key, valorVenda)
    If Not mainDict.Exists(key) Then
        mainDict.Add key, 0
    End If
    mainDict(key) = mainDict(key) + valorVenda
End Sub

' Função para ordenar as chaves de um dicionário por um valor específico em ordem decrescente
Function SortDictionaryByValue(dict)
    Dim arrKeys, i, j, temp
    
    If dict.Count > 0 Then
        arrKeys = dict.Keys
        For i = 0 To UBound(arrKeys)
            For j = i + 1 To UBound(arrKeys)
                If dict(arrKeys(i)) < dict(arrKeys(j)) Then
                    temp = arrKeys(i)
                    arrKeys(i) = arrKeys(j)
                    arrKeys(j) = temp
                End If
            Next
        Next
    Else
        SortDictionaryByValue = Array()
        Exit Function
    End If
    
    SortDictionaryByValue = arrKeys
End Function

' Função para obter valores únicos de um campo, com base nos filtros
Function GetUniqueValuesWithFilter(conn, fieldName, currentFilterName, joinClause, databasePath)
    Dim dict, rs, sqlWhere, sqlQuery, param
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    Dim baseTable
    baseTable = "Vendas"
    
    sqlWhere = " WHERE Vendas.Excluido = 0 "
    
    If Request.QueryString("ano") <> "" And currentFilterName <> "Vendas.AnoVenda" Then
        sqlWhere = sqlWhere & " AND Vendas.AnoVenda = " & Request.QueryString("ano")
    End If
    If Request.QueryString("mes") <> "" And currentFilterName <> "Vendas.MesVenda" Then
        sqlWhere = sqlWhere & " AND Vendas.MesVenda = " & Request.QueryString("mes")
    End If
    If Request.QueryString("trimestre") <> "" And currentFilterName <> "Vendas.Trimestre" Then
        sqlWhere = sqlWhere & " AND Vendas.Trimestre = " & Request.QueryString("trimestre")
    End If
    If Request.QueryString("semestre") <> "" And currentFilterName <> "Vendas.Semestre" Then
        sqlWhere = sqlWhere & " AND Vendas.Semestre = " & Request.QueryString("semestre")
    End If
    If Request.QueryString("diretoria") <> "" And currentFilterName <> "Vendas.Diretoria" Then
        param = Replace(Request.QueryString("diretoria"), "'", "''")
        sqlWhere = sqlWhere & " AND Vendas.Diretoria = '" & param & "'"
    End If
    If Request.QueryString("gerencia") <> "" And currentFilterName <> "Vendas.Gerencia" Then
        param = Replace(Request.QueryString("gerencia"), "'", "''")
        sqlWhere = sqlWhere & " AND Vendas.Gerencia = '" & param & "'"
    End If
    If Request.QueryString("corretor") <> "" And currentFilterName <> "Vendas.Corretor" Then
        param = Replace(Request.QueryString("corretor"), "'", "''")
        sqlWhere = sqlWhere & " AND Vendas.Corretor = '" & param & "'"
    End If
    If Request.QueryString("empreendimento") <> "" And currentFilterName <> "Empreendimento.NomeEmpreendimento" Then
        param = Replace(Request.QueryString("empreendimento"), "'", "''")
        sqlWhere = sqlWhere & " AND E.NomeEmpreendimento = '" & param & "'"
    End If
    If Request.QueryString("empresa") <> "" And currentFilterName <> "Empresa.NomeEmpresa" Then
        param = Replace(Request.QueryString("empresa"), "'", "''")
        sqlWhere = sqlWhere & " AND Empresa.NomeEmpresa = '" & param & "'"
    End If
    
    sqlQuery = "SELECT DISTINCT " & fieldName & " FROM " & baseTable & joinClause & sqlWhere & " ORDER BY " & fieldName & ";"
    
    On Error Resume Next
    rs.Open sqlQuery, conn
    If Err.Number <> 0 Then
        Response.Write "Erro ao obter valores únicos: " & Err.Description & "<br>"
        Response.Write "SQL Query: " & sqlQuery & "<br>"
        Err.Clear
        GetUniqueValuesWithFilter = Array()
        Exit Function
    End If
    On Error GoTo 0
    
    Dim cleanFieldName
    cleanFieldName = fieldName
    If InStr(fieldName, ".") > 0 Then
        cleanFieldName = Mid(fieldName, InStr(fieldName, ".") + 1)
    End If

    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(cleanFieldName)) Then
                Dim value
                value = CStr(rs(cleanFieldName))
                If Not dict.Exists(value) Then
                    dict.Add value, 1
                End If
            End If
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetUniqueValuesWithFilter = dict.Keys
End Function

' ===============================================
' PROCESSAMENTO DE DADOS COM FILTROS
' ===============================================

' Obter filtros atuais da query string
Dim filtroAno, filtroSemestre, filtroMes, filtroTrimestre, filtroDiretoria, filtroGerencia, filtroCorretor, filtroEmpreendimento, filtroEmpresa
filtroAno = Request.QueryString("ano")
filtroSemestre = Request.QueryString("semestre")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")
filtroEmpresa = Request.QueryString("empresa")

' Definição da cláusula de JOIN para as consultas, já com a sintaxe de linked table
Dim joinClause
joinClause = " INNER JOIN ([;DATABASE=" & dbSunnyPath & "].Empreendimento AS E " & _
             "INNER JOIN [;DATABASE=" & dbSunnyPath & "].Empresa AS M " & _
             "ON E.Empresa_ID = M.Empresa_ID) ON Vendas.Empreend_ID = E.Empreend_ID "

' Popula as listas de filtros dinamicamente
Dim uniqueAnos, uniqueMeses, uniqueSemestres, uniqueTrimestres, uniqueDiretorias, uniqueGerencias, uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas
uniqueAnos = GetUniqueValuesWithFilter(connSales, "Vendas.AnoVenda", "Vendas.AnoVenda", "", "")
uniqueMeses = GetUniqueValuesWithFilter(connSales, "Vendas.MesVenda", "Vendas.MesVenda", "", "")
uniqueSemestres = GetUniqueValuesWithFilter(connSales, "Vendas.Semestre", "Vendas.Semestre", "", "")
uniqueTrimestres = GetUniqueValuesWithFilter(connSales, "Vendas.Trimestre", "Vendas.Trimestre", "", "")
uniqueDiretorias = GetUniqueValuesWithFilter(connSales, "Vendas.Diretoria", "Vendas.Diretoria", "", "")
uniqueGerencias = GetUniqueValuesWithFilter(connSales, "Vendas.Gerencia", "Vendas.Gerencia", "", "")
uniqueCorretores = GetUniqueValuesWithFilter(connSales, "Vendas.Corretor", "Vendas.Corretor", "", "")

' Para Empreendimento e Empresa, a query principal já faz o JOIN
uniqueEmpreendimentos = GetUniqueValuesWithFilter(connSales, "E.NomeEmpreendimento", "Empreendimento.NomeEmpreendimento", joinClause, dbSunnyPath)
uniqueEmpresas = GetUniqueValuesWithFilter(connSales, "M.NomeEmpresa", "Empresa.NomeEmpresa", joinClause, dbSunnyPath)

' Constrói a cláusula WHERE para a query principal
Dim sqlWhere
sqlWhere = " WHERE Vendas.Excluido = 0 "

If filtroAno <> "" Then sqlWhere = sqlWhere & " AND Vendas.AnoVenda = " & filtroAno
If filtroSemestre <> "" Then sqlWhere = sqlWhere & " AND Vendas.Semestre = " & filtroSemestre
If filtroMes <> "" Then sqlWhere = sqlWhere & " AND Vendas.MesVenda = " & filtroMes
If filtroTrimestre <> "" Then sqlWhere = sqlWhere & " AND Vendas.Trimestre = " & filtroTrimestre
If filtroDiretoria <> "" Then sqlWhere = sqlWhere & " AND Vendas.Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
If filtroGerencia <> "" Then sqlWhere = sqlWhere & " AND Vendas.Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
If filtroCorretor <> "" Then sqlWhere = sqlWhere & " AND Vendas.Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
If filtroEmpreendimento <> "" Then sqlWhere = sqlWhere & " AND E.NomeEmpreendimento = '" & Replace(filtroEmpreendimento, "'", "''") & "'"
If filtroEmpresa <> "" Then sqlWhere = sqlWhere & " AND M.NomeEmpresa = '" & Replace(filtroEmpresa, "'", "''") & "'"

' SQL query para recuperar os dados filtrados com os JOINS entre os dois bancos de dados
Dim sql
sql = "SELECT Vendas.*, E.NomeEmpreendimento, E.ComissaoVenda, M.NomeEmpresa " & _
      "FROM (Vendas AS Vendas " & joinClause & ") " & sqlWhere & " ORDER BY Vendas.ID DESC;"

' Cria e abre o Recordset com a nova query
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, connSales

' Dicionários para armazenar os dados agregados de VGV
Dim kpiData
Set kpiData = Server.CreateObject("Scripting.Dictionary")
Set kpiData("VendasMes") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopDiretoriasVGV") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopGerenciasVGV") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopCorretoresVGV") = Server.CreateObject("Scripting.Dictionary")

' Variáveis para KPIs
Dim totalVGV, totalUnidades, ticketMedio
totalVGV = 0
totalUnidades = 0

' Processa os dados do Recordset
If Not rs.EOF Then
    Do While Not rs.EOF
        Dim valorUnidade
        valorUnidade = CDbl(rs("ValorUnidade"))
        
        Dim mes, diretoria, gerencia, corretor
        mes = CStr(rs("MesVenda"))
        diretoria = CStr(rs("Diretoria"))
        gerencia = CStr(rs("Gerencia"))
        corretor = CStr(rs("Corretor"))

        ' Popula os dicionários de KPIs de VGV
        Call ProcessaVendas(kpiData("VendasMes"), mes, valorUnidade)
        Call ProcessaVendas(kpiData("TopCorretoresVGV"), corretor, valorUnidade)
        Call ProcessaVendas(kpiData("TopDiretoriasVGV"), diretoria, valorUnidade)
        Call ProcessaVendas(kpiData("TopGerenciasVGV"), gerencia, valorUnidade)

        ' Soma os totais para os cards de KPI
        totalVGV = totalVGV + valorUnidade
        totalUnidades = totalUnidades + 1

        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing

' Calcular ticket médio
If totalUnidades > 0 Then
    ticketMedio = totalVGV / totalUnidades
Else
    ticketMedio = 0
End If

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

' Preparar dados para o gráfico de VGV por mês
Dim chartLabels(11)
Dim chartData(11)
For i = 1 To 12
    chartLabels(i - 1) = arrMesesNome(i)
    chartData(i - 1) = 0 ' Valor padrão se o mês não tiver dados
    If kpiData("VendasMes").Exists(CStr(i)) Then
        chartData(i - 1) = kpiData("VendasMes")(CStr(i))
    End If
Next

' Preparar os Top 10 para exibição
Dim sortedTopDiretorias, sortedTopGerencias, sortedTopCorretores
sortedTopDiretorias = SortDictionaryByValue(kpiData("TopDiretoriasVGV"))
sortedTopGerencias = SortDictionaryByValue(kpiData("TopGerenciasVGV"))
sortedTopCorretores = SortDictionaryByValue(kpiData("TopCorretoresVGV"))

' Fechar as conexões no final do script
connSales.Close
connSunny.Close
Set connSales = Nothing
Set connSunny = Nothing
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório Top 10 de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            background-color: #A5A2A2;
            padding: 0;
            margin: 0;
            color: white;
            padding-top: 80px; /* Espaço para o header fixo */
        }
        
        .main-content {
            padding: 20px;
        }
        
        .container-fluid {
            max-width: 1400px;
            margin: 0 auto;
        }

        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 20px;
            border-radius: 8px;
            font-size: 1rem;
            margin-bottom: 10px;
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .kpi-card h5 {
            font-size: 1.1rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            line-height: 1.2;
        }
        .kpi-card i {
            font-size: 1.8rem;
            margin-bottom: 10px;
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
        .filter-row {
            margin-bottom: 10px;
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .filter-select {
            width: 100%;
        }
        .filter-btn {
            margin-top: 25px;
        }
        .text-center-v { text-align: center; }
        .text-right-v { text-align: right; }

        .chart-container {
            background-color: #fff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .top-10-container {
            background-color: #fff;
            color: #333;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            height: 100%;
            overflow: auto;
        }
        .top-10-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .top-10-list-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 15px;
            border-bottom: 1px solid #eee;
            transition: background-color 0.3s ease;
        }
        .top-10-list-item:last-child {
            border-bottom: none;
        }
        .top-10-list-item:hover {
            background-color: #f8f9fa;
        }
        .top-10-list-item span {
            font-weight: bold;
        }
        .top-10-list-item .rank-icon {
            font-size: 1.5rem;
            margin-right: 10px;
        }
        .top-10-list-item .top-1 {
            color: #FFD700; /* Gold */
        }
        .top-10-list-item .top-2 {
            color: #C0C0C0; /* Silver */
        }
        .top-10-list-item .top-3 {
            color: #CD7F32; /* Bronze */
        }

        /* Estilo para garantir que o header não sobreponha o conteúdo */
        header {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            z-index: 1000;
            background-color: #fff; /* ou a cor do seu header */
        }
    </style>
</head>
<body>
    <!--#include file="gestao_header.inc"-->
    
    <div class="main-content">
        <div class="container-fluid">
            <h2 class="mt-4 mb-4 text-center">
                <i class="fas fa-handshake"></i> Relatório Top 10 de Vendas (VGV)
            </h2>
            
            <div class="filter-container">
                <form id="filterForm" method="get">
                    <div class="row filter-row">
                        <div class="col-md-2">
                            <div class="filter-label">Ano</div>
                            <select class="form-select filter-select" name="ano" id="anoFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
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

                        <div class="col-md-2">
                            <div class="filter-label">Semestre</div>
                            <select class="form-select filter-select" name="semestre" id="semestreFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueSemestres) Then
                                    For Each semestre In uniqueSemestres
                                        Response.Write "<option value=""" & semestre & """"
                                        If CStr(filtroSemestre) = CStr(semestre) Then Response.Write " selected"
                                        Response.Write ">" & semestre & "º Semestre</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>

                        <div class="col-md-2">
                            <div class="filter-label">Trimestre</div>
                            <select class="form-select filter-select" name="trimestre" id="trimestreFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueTrimestres) Then
                                    For Each trimestre In uniqueTrimestres
                                        Response.Write "<option value=""" & trimestre & """"
                                        If filtroTrimestre = trimestre Then Response.Write " selected"
                                        Response.Write ">" & trimestre & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>

                        <div class="col-md-2">
                            <div class="filter-label">Mês</div>
                            <select class="form-select filter-select" name="mes" id="mesFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
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
                        
                        <div class="col-md-2">
                            <div class="filter-label">Diretoria</div>
                            <select class="form-select filter-select" name="diretoria" id="diretoriaFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueDiretorias) Then
                                    For Each diretoria In uniqueDiretorias
                                        Response.Write "<option value=""" & diretoria & """"
                                        If filtroDiretoria = diretoria Then Response.Write " selected"
                                        Response.Write ">" & diretoria & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-md-2">
                            <div class="filter-label">Gerência</div>
                            <select class="form-select filter-select" name="gerencia" id="gerenciaFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueGerencias) Then
                                    For Each gerencia In uniqueGerencias
                                        Response.Write "<option value=""" & gerencia & """"
                                        If filtroGerencia = gerencia Then Response.Write " selected"
                                        Response.Write ">" & gerencia & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    <div class="row filter-row mt-3">
                        <div class="col-md-2">
                            <div class="filter-label">Corretor</div>
                            <select class="form-select filter-select" name="corretor" id="corretorFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueCorretores) Then
                                    For Each corretor In uniqueCorretores
                                        Response.Write "<option value=""" & corretor & """"
                                        If filtroCorretor = corretor Then Response.Write " selected"
                                        Response.Write ">" & corretor & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-md-2">
                            <div class="filter-label">Empreendimento</div>
                            <select class="form-select filter-select" name="empreendimento" id="empreendimentoFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueEmpreendimentos) Then
                                    For Each empreendimento In uniqueEmpreendimentos
                                        Response.Write "<option value=""" & empreendimento & """"
                                        If filtroEmpreendimento = empreendimento Then Response.Write " selected"
                                        Response.Write ">" & empreendimento & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-md-2">
                            <div class="filter-label">Empresa</div>
                            <select class="form-select filter-select" name="empresa" id="empresaFilter" onchange="this.form.submit()">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueEmpresas) Then
                                    For Each empresa In uniqueEmpresas
                                        Response.Write "<option value=""" & empresa & """"
                                        If filtroEmpresa = empresa Then Response.Write " selected"
                                        Response.Write ">" & empresa & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-md-6 text-end">
                            <button type="button" class="btn btn-secondary filter-btn" onclick="limparFiltros()">
                                <i class="fas fa-times"></i> Limpar Filtros
                            </button>
                        </div>
                    </div>
                </form>
            </div>
            
            <!-- Resto do seu conteúdo (KPIs, gráficos, etc.) -->
            <div class="row mt-4">
                <div class="col-md-4">
                    <div class="kpi-card bg-success-kpi">
                        <i class="fas fa-dollar-sign"></i>
                        <h5>Total de Vendas (VGV)</h5>
                        <p><%= FormatCurrencyValue(totalVGV, 2) %></p>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="kpi-card bg-info-kpi">
                        <i class="fas fa-cubes"></i>
                        <h5>Unidades Vendidas</h5>
                        <p><%= totalUnidades %></p>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="kpi-card bg-warning-kpi">
                        <i class="fas fa-receipt"></i>
                        <h5>Ticket Médio</h5>
                        <p><%= FormatCurrencyValue(ticketMedio, 2) %></p>
                    </div>
                </div>
            </div>

            <!-- Seção de Top 10 -->
            <div class="row mt-4">
                <div class="col-md-4 mb-4">
                    <div class="top-10-container">
                        <h5 class="text-center mb-3">Top 10 Diretorias por VGV</h5>
                        <ul class="top-10-list">
                            <%
                            If IsArray(sortedTopDiretorias) Then
                                For i = 0 To UBound(sortedTopDiretorias)
                                    If i > 9 Then Exit For
                                    Dim diretoriaKey, diretoriaValue, rankIcon
                                    diretoriaKey = sortedTopDiretorias(i)
                                    diretoriaValue = kpiData("TopDiretoriasVGV")(diretoriaKey)
                                    Select Case i
                                        Case 0
                                            rankIcon = "<span class='rank-icon top-1'>🥇</span>"
                                        Case 1
                                            rankIcon = "<span class='rank-icon top-2'>🥈</span>"
                                        Case 2
                                            rankIcon = "<span class='rank-icon top-3'>🥉</span>"
                                        Case Else
                                            rankIcon = "<span class='rank-icon'>" & (i + 1) & ".</span>"
                                    End Select
                                    Response.Write "<li class='top-10-list-item'>" & rankIcon & "<span>" & diretoriaKey & "</span><span> " & FormatCurrencyValue(diretoriaValue, 2) & "</span></li>"
                                Next
                            End If
                            %>
                        </ul>
                    </div>
                </div>
                
                <div class="col-md-4 mb-4">
                    <div class="top-10-container">
                        <h5 class="text-center mb-3">Top 10 Gerências por VGV</h5>
                        <ul class="top-10-list">
                            <%
                            If IsArray(sortedTopGerencias) Then
                                For i = 0 To UBound(sortedTopGerencias)
                                    If i > 9 Then Exit For
                                    Dim gerenciaKey, gerenciaValue
                                    gerenciaKey = sortedTopGerencias(i)
                                    gerenciaValue = kpiData("TopGerenciasVGV")(gerenciaKey)
                                    Select Case i
                                        Case 0
                                            rankIcon = "<span class='rank-icon top-1'>🥇</span>"
                                        Case 1
                                            rankIcon = "<span class='rank-icon top-2'>🥈</span>"
                                        Case 2
                                            rankIcon = "<span class='rank-icon top-3'>🥉</span>"
                                        Case Else
                                            rankIcon = "<span class='rank-icon'>" & (i + 1) & ".</span>"
                                    End Select
                                    Response.Write "<li class='top-10-list-item'>" & rankIcon & "<span>" & gerenciaKey & "</span><span> " & FormatCurrencyValue(gerenciaValue, 2) & "</span></li>"
                                Next
                            End If
                            %>
                        </ul>
                    </div>
                </div>
                
                <div class="col-md-4 mb-4">
                    <div class="top-10-container">
                        <h5 class="text-center mb-3">Top 10 Corretores por VGV</h5>
                        <ul class="top-10-list">
                            <%
                            If IsArray(sortedTopCorretores) Then
                                For i = 0 To UBound(sortedTopCorretores)
                                    If i > 9 Then Exit For
                                    Dim corretorKey, corretorValue
                                    corretorKey = sortedTopCorretores(i)
                                    corretorValue = kpiData("TopCorretoresVGV")(corretorKey)
                                    Select Case i
                                        Case 0
                                            rankIcon = "<span class='rank-icon top-1'>🥇</span>"
                                        Case 1
                                            rankIcon = "<span class='rank-icon top-2'>🥈</span>"
                                        Case 2
                                            rankIcon = "<span class='rank-icon top-3'>🥉</span>"
                                        Case Else
                                            rankIcon = "<span class='rank-icon'>" & (i + 1) & ".</span>"
                                    End Select
                                    Response.Write "<li class='top-10-list-item'>" & rankIcon & "<span>" & corretorKey & "</span><span> " & FormatCurrencyValue(corretorValue, 2) & "</span></li>"
                                Next
                            End If
                            %>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function limparFiltros() {
            window.location.href = window.location.pathname;
        }
    </script>
</body>
</html>