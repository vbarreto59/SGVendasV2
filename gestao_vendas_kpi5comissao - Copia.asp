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

'=============================================

' Atualizar campo Semestre se necessário
sql = "UPDATE Vendas " & _
      "SET Semestre = SWITCH(" & _
      "    Trimestre IN (1, 2), 1, " & _
      "    Trimestre IN (3, 4), 2" & _
      ") " & _
      "WHERE Trimestre IS NOT NULL AND (Semestre IS NULL OR Semestre = 0);"

On Error Resume Next
connSales.Execute sql
On Error GoTo 0

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
    
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

Sub ProcessaComissoes(mainDict, key, comissaoValor)
    If Not mainDict.Exists(key) Then mainDict.Add key, 0
    mainDict(key) = mainDict(key) + comissaoValor
End Sub

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

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno, filtroMes, filtroTrimestre, filtroSemestre
Dim filtroDiretoria, filtroGerencia, filtroCorretor
Dim filtroEmpreendimento, filtroEmpresa

filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroSemestre = Request.QueryString("semestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")
filtroEmpresa = Request.QueryString("empresa")

' Construir cláusula WHERE
Dim sqlWhere
sqlWhere = " WHERE Excluido = 0 "

If filtroAno <> "" Then sqlWhere = sqlWhere & " AND AnoVenda = " & filtroAno
If filtroMes <> "" Then sqlWhere = sqlWhere & " AND MesVenda = " & filtroMes
If filtroTrimestre <> "" Then sqlWhere = sqlWhere & " AND Trimestre = " & filtroTrimestre
If filtroSemestre <> "" Then sqlWhere = sqlWhere & " AND Semestre = " & filtroSemestre
If filtroDiretoria <> "" Then sqlWhere = sqlWhere & " AND Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
If filtroGerencia <> "" Then sqlWhere = sqlWhere & " AND Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
If filtroCorretor <> "" Then sqlWhere = sqlWhere & " AND Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
If filtroEmpreendimento <> "" Then sqlWhere = sqlWhere & " AND NomeEmpreendimento = '" & Replace(filtroEmpreendimento, "'", "''") & "'"
If filtroEmpresa <> "" Then sqlWhere = sqlWhere & " AND NomeEmpresa = '" & Replace(filtroEmpresa, "'", "''") & "'"

' Consulta SQL principal usando apenas a tabela Vendas
Dim sql
sql = "SELECT * FROM Vendas " & sqlWhere & " ORDER BY DataVenda DESC"

' Execução da consulta principal
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rs.Open sql, connSales

If Err.Number <> 0 Then
    Response.Write "Erro na consulta: " & Err.Description & "<br>"
    Response.Write "SQL: " & Server.HTMLEncode(sql)
    Response.End
End If
On Error GoTo 0

' Processar resultados
Dim kpiData, totalVGV, totalComissoes, totalComissoesDiretoria, totalComissoesGerencia, totalComissoesCorretor, totalUnidadesVendidas
Set kpiData = Server.CreateObject("Scripting.Dictionary")
Set kpiData("Mes") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopDiretorias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopGerencias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopCorretores") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpreendimentos") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpresas") = Server.CreateObject("Scripting.Dictionary")

totalVGV = 0
totalComissoes = 0
totalComissoesDiretoria = 0
totalComissoesGerencia = 0
totalComissoesCorretor = 0
totalUnidadesVendidas = 0

If Not rs.EOF Then
    Do While Not rs.EOF
        Dim valorUnidade, comissaoPercentual, comissaoValor
        Dim comissaoDiretor, comissaoGerente, comissaoCorretorTotal
        Dim mes, diretoria, gerencia, corretor, empreendimento, empresa
        
        valorUnidade = CDbl(rs("ValorUnidade"))
        comissaoPercentual = CDbl(rs("ComissaoPercentual"))
        
        ' Calcular comissões baseadas nos valores já existentes na tabela
        comissaoValor = CDbl(rs("ValorComissaoGeral"))
        comissaoDiretor = CDbl(rs("ValorDiretoria"))
        comissaoGerente = CDbl(rs("ValorGerencia"))
        comissaoCorretorTotal = CDbl(rs("ValorCorretor"))
        
        mes = CStr(rs("MesVenda"))
        diretoria = CStr(rs("Diretoria"))
        gerencia = CStr(rs("Gerencia"))
        corretor = CStr(rs("Corretor"))
        empreendimento = CStr(rs("NomeEmpreendimento"))
        empresa = CStr(rs("NomeEmpresa"))

        ' Atualizar KPIs
        If Not IsEmpty(mes) Then Call ProcessaComissoes(kpiData("Mes"), mes, comissaoValor)
        If Not IsEmpty(diretoria) Then Call ProcessaComissoes(kpiData("TopDiretorias"), diretoria, comissaoDiretor)
        If Not IsEmpty(gerencia) Then Call ProcessaComissoes(kpiData("TopGerencias"), gerencia, comissaoGerente)
        If Not IsEmpty(corretor) Then Call ProcessaComissoes(kpiData("TopCorretores"), corretor, comissaoCorretorTotal)
        If Not IsEmpty(empreendimento) Then Call ProcessaComissoes(kpiData("TopEmpreendimentos"), empreendimento, comissaoValor)
        If Not IsEmpty(empresa) Then Call ProcessaComissoes(kpiData("TopEmpresas"), empresa, comissaoValor)

        ' Atualizar totais
        totalVGV = totalVGV + valorUnidade
        totalComissoes = totalComissoes + comissaoValor
        totalComissoesDiretoria = totalComissoesDiretoria + comissaoDiretor
        totalComissoesGerencia = totalComissoesGerencia + comissaoGerente
        totalComissoesCorretor = totalComissoesCorretor + comissaoCorretorTotal
        totalUnidadesVendidas = totalUnidadesVendidas + 1

        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing

' ===============================================
' POPULAR OS SELECTS DO FORMULÁRIO
' ===============================================

Dim uniqueAnos, uniqueMeses, uniqueTrimestres, uniqueSemestres, uniqueDiretorias, uniqueGerencias, uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas

' Popular arrays de valores únicos usando apenas a tabela Vendas
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueMeses = GetUniqueValues("Vendas", "MesVenda", "WHERE MesVenda IS NOT NULL")
uniqueTrimestres = GetUniqueValues("Vendas", "Trimestre", "WHERE Trimestre IS NOT NULL")
uniqueSemestres = GetUniqueValues("Vendas", "Semestre", "WHERE Semestre IS NOT NULL")
uniqueDiretorias = GetUniqueValues("Vendas", "Diretoria", "WHERE Diretoria IS NOT NULL AND Diretoria <> ''")
uniqueGerencias = GetUniqueValues("Vendas", "Gerencia", "WHERE Gerencia IS NOT NULL AND Gerencia <> ''")
uniqueCorretores = GetUniqueValues("Vendas", "Corretor", "WHERE Corretor IS NOT NULL AND Corretor <> ''")
uniqueEmpreendimentos = GetUniqueValues("Vendas", "NomeEmpreendimento", "WHERE NomeEmpreendimento IS NOT NULL AND NomeEmpreendimento <> ''")
uniqueEmpresas = GetUniqueValues("Vendas", "NomeEmpresa", "WHERE NomeEmpresa IS NOT NULL AND NomeEmpresa <> ''")

' Preparar dados para gráficos
Dim arrMesesNome(12), chartLabels(11), chartData(11)
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

For i = 1 To 12
    chartLabels(i-1) = arrMesesNome(i)
    
    If kpiData("Mes").Exists(CStr(i)) Then
        chartData(i-1) = kpiData("Mes")(CStr(i))
    Else
        chartData(i-1) = 0
    End If
Next
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório de Vendas e Comissões</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #A5A2A2;
            padding: 20px;
            color: white;
        }
        .card-kpi {
            background-color: #F0ECEC;
            padding: 15px;
            margin-top: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
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
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;"><i class="fas fa-hand-holding-usd"></i> SGVendas - Relatório de Vendas e Comissões</h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row filter-row">
                    <div class="col-md-2">
                        <div class="filter-label">Ano</div>
                        <select class="form-select filter-select" name="ano" id="anoFilter">
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
                        <select class="form-select filter-select" name="semestre" id="semestreFilter">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueSemestres) Then
                                For Each sem In uniqueSemestres
                                    Response.Write "<option value=""" & sem & """"
                                    If filtroSemestre = sem Then Response.Write " selected"
                                    Response.Write ">" & sem & "º Semestre</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Mês</div>
                        <select class="form-select filter-select" name="mes" id="mesFilter">
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
                        <div class="filter-label">Trimestre</div>
                        <select class="form-select filter-select" name="trimestre" id="trimestreFilter">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueTrimestres) Then
                                For Each trimestre In uniqueTrimestres
                                    Response.Write "<option value=""" & trimestre & """"
                                    If filtroTrimestre = trimestre Then Response.Write " selected"
                                    Response.Write ">" & trimestre & "º Trimestre</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-2">
                        <div class="filter-label">Diretoria</div>
                        <select class="form-select filter-select" name="diretoria" id="diretoriaFilter">
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
                        <select class="form-select filter-select" name="gerencia" id="gerenciaFilter">
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
                        <select class="form-select filter-select" name="corretor" id="corretorFilter">
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
                        <select class="form-select filter-select" name="empreendimento" id="empreendimentoFilter">
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
                        <select class="form-select filter-select" name="empresa" id="empresaFilter">
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
                        <button type="submit" class="btn btn-primary filter-btn">
                            <i class="fas fa-filter"></i> Aplicar Filtros
                        </button>
                        <button type="button" class="btn btn-secondary filter-btn" onclick="limparFiltros()">
                            <i class="fas fa-times"></i> Limpar
                        </button>
                    </div>
                </div>
            </form>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-handshake"></i>
                    <h5>Total de VGV</h5>
                    <p>R$ <%= FormatNumber(totalVGV, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-user-tie"></i>
                    <h5>Comissões Diretoria</h5>
                    <p>R$ <%= FormatNumber(totalComissoesDiretoria, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-users-cog"></i>
                    <h5>Comissões Gerência</h5>
                    <p>R$ <%= FormatNumber(totalComissoesGerencia, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-danger-kpi">
                    <i class="fas fa-user-tag"></i>
                    <h5>Comissões Corretores</h5>
                    <p>R$ <%= FormatNumber(totalComissoesCorretor, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-secondary-kpi">
                    <i class="fas fa-home"></i>
                    <h5>Unidades Vendidas</h5>
                    <p><%= totalUnidadesVendidas %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <i class="fas fa-money-bill-wave"></i>
                    <h5>Total Comissões</h5>
                    <p>R$ <%= FormatNumber(totalComissoes, 2) %></p>
                </div>
            </div>
        </div>
        
        <!-- -------------------------------------------------------------- -->
        <h2 class="text-white mt-5">Relatório Detalhado</h2>
        <div class="card-kpi p-3 rounded">
            <div class="row">
                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Diretorias (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-primary-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Diretoria</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topDiretorias, arrDiretorias
                                Set topDiretorias = kpiData("TopDiretorias")
                                arrDiretorias = SortDictionaryByValue(topDiretorias)
                                If IsArray(arrDiretorias) Then
                                    For i = 0 To UBound(arrDiretorias)
                                        If i < 10 Then ' Mostrar apenas top 10
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrDiretorias(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(topDiretorias(arrDiretorias(i)), 2) %></td>
                                </tr>
                                <%
                                        End If
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Gerências (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-success-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Gerência</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topGerencias, arrGerencias
                                Set topGerencias = kpiData("TopGerencias")
                                arrGerencias = SortDictionaryByValue(topGerencias)
                                If IsArray(arrGerencias) Then
                                    For i = 0 To UBound(arrGerencias)
                                        If i < 10 Then ' Mostrar apenas top 10
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrGerencias(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(topGerencias(arrGerencias(i)), 2) %></td>
                                </tr>
                                <%
                                        End If
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Corretores (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-info-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Corretor</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topCorretores, arrCorretores
                                Set topCorretores = kpiData("TopCorretores")
                                arrCorretores = SortDictionaryByValue(topCorretores)
                                If IsArray(arrCorretores) Then
                                    For i = 0 To UBound(arrCorretores)
                                        If i < 10 Then ' Mostrar apenas top 10
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrCorretores(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(topCorretores(arrCorretores(i)), 2) %></td>
                                </tr>
                                <%
                                        End If
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Empreendimentos</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-warning-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Empreendimento</th>
                                    <th class="text-right-v">VGV (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topEmpreendimentos, arrEmpreendimentos
                                Set topEmpreendimentos = kpiData("TopEmpreendimentos")
                                arrEmpreendimentos = SortDictionaryByValue(topEmpreendimentos)
                                If IsArray(arrEmpreendimentos) Then
                                    For i = 0 To UBound(arrEmpreendimentos)
                                        If i < 10 Then ' Mostrar apenas top 10
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrEmpreendimentos(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(topEmpreendimentos(arrEmpreendimentos(i)), 2) %></td>
                                </tr>
                                <%
                                        End If
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
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

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>