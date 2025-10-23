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
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno, filtroMes, filtroDiretoria, filtroGerencia, filtroLocalidade
filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroLocalidade = Request.QueryString("localidade")

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

Dim uniqueAnos, uniqueMeses, uniqueDiretorias, uniqueGerencias, uniqueLocalidades
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueMeses = GetUniqueValues("Vendas", "MesVenda", "WHERE MesVenda IS NOT NULL")
uniqueDiretorias = GetUniqueValues("Vendas", "Diretoria", "WHERE Diretoria IS NOT NULL AND Diretoria <> ''")
uniqueGerencias = GetUniqueValues("Vendas", "Gerencia", "WHERE Gerencia IS NOT NULL AND Gerencia <> ''")
uniqueLocalidades = GetUniqueValues("Vendas", "Localidade", "WHERE Localidade IS NOT NULL AND Localidade <> ''")

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

' ===============================================
' OBTER DADOS DE VENDAS POR LOCALIDADE
' ===============================================

Dim localidadesData, totalGeralVGV, totalGeralVendas
Set localidadesData = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Primeiro: obter dados básicos das localidades
    Dim sqlVendas, rsVendas
    sqlVendas = "SELECT " & _
                "Vendas.Localidade, " & _
                "SUM(Vendas.ValorUnidade) as TotalVGV, " & _
                "COUNT(*) as TotalVendas, " & _
                "AVG(Vendas.ValorUnidade) as MediaVGV " & _
                "FROM Vendas " & _
                "WHERE Vendas.Excluido = 0 " & _
                "AND Vendas.AnoVenda = " & filtroAno
    
    If filtroMes <> "" Then sqlVendas = sqlVendas & " AND Vendas.MesVenda = " & filtroMes
    If filtroDiretoria <> "" Then sqlVendas = sqlVendas & " AND Vendas.Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
    If filtroGerencia <> "" Then sqlVendas = sqlVendas & " AND Vendas.Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
    If filtroLocalidade <> "" Then sqlVendas = sqlVendas & " AND Vendas.Localidade = '" & Replace(filtroLocalidade, "'", "''") & "'"
    
    sqlVendas = sqlVendas & " GROUP BY Vendas.Localidade " & _
                "ORDER BY SUM(Vendas.ValorUnidade) DESC"
    
    Set rsVendas = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsVendas.Open sqlVendas, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de vendas: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlVendas)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados de vendas por localidade
    totalGeralVGV = 0
    totalGeralVendas = 0

    If Not rsVendas.EOF Then
        Do While Not rsVendas.EOF
            Dim localidade, totalVGV, totalVendas, mediaVGV
            localidade = CStr(rsVendas("Localidade"))
            totalVGV = CDbl(rsVendas("TotalVGV"))
            totalVendas = CLng(rsVendas("TotalVendas"))
            mediaVGV = CDbl(rsVendas("MediaVGV"))
            
            ' Buscar quantidade de empreendimentos distintos para esta localidade
            Dim sqlEmpreendimentos, rsEmpreendimentos, totalEmpreendimentos
            totalEmpreendimentos = 0
            
            sqlEmpreendimentos = "SELECT COUNT(*) as TotalEmp FROM (" & _
                                "SELECT DISTINCT Empreend_Id " & _
                                "FROM Vendas " & _
                                "WHERE Excluido = 0 " & _
                                "AND AnoVenda = " & filtroAno & _
                                " AND Localidade = '" & Replace(localidade, "'", "''") & "'"
            
            If filtroMes <> "" Then sqlEmpreendimentos = sqlEmpreendimentos & " AND MesVenda = " & filtroMes
            If filtroDiretoria <> "" Then sqlEmpreendimentos = sqlEmpreendimentos & " AND Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
            If filtroGerencia <> "" Then sqlEmpreendimentos = sqlEmpreendimentos & " AND Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
            
            sqlEmpreendimentos = sqlEmpreendimentos & ") as Empreendimentos"
            
            Set rsEmpreendimentos = connSales.Execute(sqlEmpreendimentos)
            If Not rsEmpreendimentos.EOF Then
                totalEmpreendimentos = CLng(rsEmpreendimentos("TotalEmp"))
            End If
            
            If rsEmpreendimentos.State = 1 Then rsEmpreendimentos.Close
            Set rsEmpreendimentos = Nothing
            
            ' Adicionar dados da localidade
            localidadesData.Add localidade, Array(totalVGV, totalVendas, mediaVGV, totalEmpreendimentos)
            
            ' Atualizar totais gerais
            totalGeralVGV = totalGeralVGV + totalVGV
            totalGeralVendas = totalGeralVendas + totalVendas
            
            rsVendas.MoveNext
        Loop
    End If

    If rsVendas.State = 1 Then rsVendas.Close
    Set rsVendas = Nothing
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vendas por Localidade</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 10px;
            font-size: 14px;
        }
        .container {
            max-width: 100%;
            margin: 0 auto;
        }
        .header {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 15px;
            margin-bottom: 15px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }
        .page-title {
            color: #2c3e50;
            font-size: 20px;
            font-weight: 700;
            text-align: center;
            margin-bottom: 5px;
        }
        .page-subtitle {
            color: #7f8c8d;
            font-size: 12px;
            text-align: center;
            margin-bottom: 15px;
        }
        .filter-card {
            background: rgba(255, 255, 255, 0.9);
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        .filter-label {
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 6px;
            font-size: 12px;
        }
        .form-select {
            border-radius: 8px;
            border: 1px solid #e9ecef;
            padding: 8px;
            font-size: 12px;
            margin-bottom: 8px;
        }
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            border-radius: 8px;
            padding: 10px;
            font-weight: 600;
            font-size: 12px;
        }
        .btn-secondary {
            background: #6c757d;
            border: none;
            border-radius: 8px;
            padding: 10px;
            font-weight: 600;
            font-size: 12px;
            color: white;
        }
        
        /* ===== ESTILOS MOBILE ===== */
        .mobile-only {
            display: block;
        }
        .desktop-only {
            display: none;
        }
        
        .localidade-card {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 15px;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.3);
            position: relative;
        }
        .posicao-badge {
            position: absolute;
            top: 10px;
            right: 10px;
            background: #3498db;
            color: white;
            width: 30px;
            height: 30px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 14px;
        }
        .localidade-nome {
            font-size: 18px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 10px;
            padding-right: 40px;
        }
        .valor-venda {
            font-size: 24px;
            font-weight: 800;
            color: #27ae60;
            text-align: center;
            margin-bottom: 10px;
        }
        .info-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin: 10px 0;
        }
        .info-item {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 8px;
            text-align: center;
        }
        .info-label {
            font-weight: 600;
            color: #6c757d;
            font-size: 10px;
            display: block;
            margin-bottom: 3px;
        }
        .info-value {
            color: #2c3e50;
            font-weight: 600;
            font-size: 12px;
        }
        .percentual-bar {
            background: #ecf0f1;
            border-radius: 10px;
            height: 20px;
            margin: 8px 0;
            overflow: hidden;
            position: relative;
        }
        .percentual-fill {
            background: linear-gradient(90deg, #27ae60, #2ecc71);
            height: 100%;
            border-radius: 10px;
            transition: width 0.3s ease;
        }
        .percentual-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 10px;
            font-weight: 600;
            color: #2c3e50;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 6px;
            margin: 10px 0;
        }
        .stat-item {
            background: #34495e;
            color: white;
            border-radius: 8px;
            padding: 8px;
            text-align: center;
        }
        .stat-label {
            font-size: 9px;
            opacity: 0.8;
            margin-bottom: 2px;
        }
        .stat-value {
            font-size: 11px;
            font-weight: 700;
        }
        .total-card {
            background: #2c3e50;
            color: white;
            border-radius: 12px;
            padding: 12px;
            margin-top: 10px;
            text-align: center;
        }
        .no-results {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 30px 20px;
            text-align: center;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        .filter-buttons {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin-top: 5px;
        }
        .kpi-mobile {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 10px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        .kpi-value {
            font-size: 18px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 3px;
        }
        .kpi-label {
            font-size: 10px;
            color: #7f8c8d;
            font-weight: 600;
        }
        .kpi-icon {
            font-size: 20px;
            color: #3498db;
            margin-bottom: 5px;
        }
        .top-localidades {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 12px;
            padding: 12px;
            margin-top: 10px;
        }
        .top-title {
            font-size: 14px;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 8px;
            text-align: center;
        }
        .top-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 6px 0;
            border-bottom: 1px solid #f1f2f6;
        }
        .top-item:last-child {
            border-bottom: none;
        }
        .top-pos {
            background: #3498db;
            color: white;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            font-weight: bold;
        }
        .top-name {
            flex: 1;
            margin-left: 8px;
            font-size: 11px;
            font-weight: 600;
        }
        .top-value {
            font-size: 11px;
            font-weight: 700;
            color: #27ae60;
        }
        
        /* ===== ESTILOS DESKTOP ===== */
        @media (min-width: 992px) {
            .mobile-only {
                display: none !important;
            }
            .desktop-only {
                display: block !important;
            }
            
            body {
                padding: 20px;
                font-size: 16px;
            }
            .container {
                max-width: 1400px;
            }
            .header {
                padding: 25px;
                margin-bottom: 20px;
            }
            .page-title {
                font-size: 28px;
                margin-bottom: 10px;
            }
            .page-subtitle {
                font-size: 16px;
                margin-bottom: 20px;
            }
            .filter-card {
                padding: 20px;
                margin-bottom: 20px;
            }
            .filter-label {
                font-size: 14px;
                margin-bottom: 8px;
            }
            .form-select {
                padding: 10px;
                font-size: 14px;
                margin-bottom: 10px;
            }
            .btn-primary, .btn-secondary {
                padding: 12px;
                font-size: 14px;
            }
            
            /* Layout Desktop - Grid */
            .desktop-grid {
                display: grid;
                grid-template-columns: 300px 1fr;
                gap: 20px;
                align-items: start;
            }
            
            .sidebar-desktop {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 15px;
                padding: 20px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
                position: sticky;
                top: 20px;
            }
            
            .main-content-desktop {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 20px;
                align-items: start;
            }
            
            .localidades-grid {
                display: grid;
                grid-template-columns: 1fr;
                gap: 15px;
            }
            
            .localidade-card-desktop {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 15px;
                padding: 20px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
                border: 1px solid rgba(255, 255, 255, 0.3);
            }
            
            .localidade-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 15px;
            }
            
            .localidade-nome-desktop {
                font-size: 20px;
                font-weight: 700;
                color: #2c3e50;
            }
            
            .valor-venda-desktop {
                font-size: 24px;
                font-weight: 800;
                color: #27ae60;
            }
            
            .stats-grid-desktop {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 10px;
                margin: 15px 0;
            }
            
            .stat-item-desktop {
                background: #34495e;
                color: white;
                border-radius: 10px;
                padding: 12px;
                text-align: center;
            }
            
            .stat-label-desktop {
                font-size: 12px;
                opacity: 0.8;
                margin-bottom: 5px;
            }
            
            .stat-value-desktop {
                font-size: 14px;
                font-weight: 700;
            }
            
            .kpi-desktop {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 12px;
                padding: 20px;
                text-align: center;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
                margin-bottom: 20px;
            }
            
            .kpi-value-desktop {
                font-size: 24px;
                font-weight: 700;
                color: #2c3e50;
                margin-bottom: 5px;
            }
            
            .kpi-label-desktop {
                font-size: 14px;
                color: #7f8c8d;
                font-weight: 600;
            }
            
            .table-desktop {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 15px;
                overflow: hidden;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            }
            
            .table-desktop th {
                background: #34495e;
                color: white;
                padding: 15px;
                font-weight: 600;
            }
            
            .table-desktop td {
                padding: 12px 15px;
                border-bottom: 1px solid #f1f2f6;
            }
            
            .table-desktop tr:hover {
                background: #f8f9fa;
            }
            
            .posicao-badge-desktop {
                background: #3498db;
                color: white;
                width: 30px;
                height: 30px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                font-weight: bold;
                font-size: 14px;
            }
        }
        
        /* Breakpoint Tablet */
        @media (min-width: 768px) and (max-width: 991px) {
            .container {
                max-width: 95%;
            }
            .localidade-card {
                padding: 20px;
            }
            .valor-venda {
                font-size: 28px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Cabeçalho -->
        <div class="header">
            <h1 class="page-title">
                <i class="fas fa-map-marker-alt"></i> Vendas por Localidade
            </h1>
            <p class="page-subtitle mobile-only">Versão Responsiva</p>
            <p class="page-subtitle desktop-only">Dashboard Completo</p>
            
            <!-- Filtros -->
            <div class="filter-card">
                <form id="filterForm" method="get">
                    <div class="row g-2">
                        <div class="col-6 col-lg-3">
                            <label class="filter-label">Ano</label>
                            <select class="form-select" name="ano" id="anoFilter" required>
                                <option value="">Selecione</option>
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
                        
                        <div class="col-6 col-lg-3">
                            <label class="filter-label">Mês</label>
                            <select class="form-select" name="mes" id="mesFilter">
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
                        
                        <div class="col-6 col-lg-3">
                            <label class="filter-label">Diretoria</label>
                            <select class="form-select" name="diretoria" id="diretoriaFilter">
                                <option value="">Todas</option>
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
                        
                        <div class="col-6 col-lg-3">
                            <label class="filter-label">Gerência</label>
                            <select class="form-select" name="gerencia" id="gerenciaFilter">
                                <option value="">Todas</option>
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
                        
                        <div class="col-12 col-lg-6">
                            <label class="filter-label">Localidade</label>
                            <select class="form-select" name="localidade" id="localidadeFilter">
                                <option value="">Todas</option>
                                <%
                                If IsArray(uniqueLocalidades) Then
                                    For Each localidade In uniqueLocalidades
                                        Response.Write "<option value=""" & localidade & """"
                                        If filtroLocalidade = localidade Then Response.Write " selected"
                                        Response.Write ">" & localidade & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-12 col-lg-6">
                            <div class="filter-buttons">
                                <button type="submit" class="btn btn-primary">
                                    <i class="fas fa-filter"></i> Aplicar Filtros
                                </button>
                                <button type="button" class="btn btn-secondary" onclick="limparFiltros()">
                                    <i class="fas fa-times"></i> Limpar
                                </button>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
        </div>

        <% If filtroAno = "" Then %>
            <div class="no-results">
                <i class="fas fa-filter"></i>
                <h4>Selecione os filtros</h4>
                <p>Escolha o ano para visualizar o relatório.</p>
            </div>
        <% Else %>
        
        <!-- ===== VERSÃO MOBILE ===== -->
        <div class="mobile-only">
            <!-- KPIs Mobile -->
            <div class="row g-2">
                <div class="col-6">
                    <div class="kpi-mobile">
                        <div class="kpi-icon">
                            <i class="fas fa-handshake"></i>
                        </div>
                        <div class="kpi-value">R$ <%= FormatNumber(totalGeralVGV, 2) %></div>
                        <div class="kpi-label">Total VGV</div>
                    </div>
                </div>
                <div class="col-6">
                    <div class="kpi-mobile">
                        <div class="kpi-icon">
                            <i class="fas fa-home"></i>
                        </div>
                        <div class="kpi-value"><%= totalGeralVendas %></div>
                        <div class="kpi-label">Vendas</div>
                    </div>
                </div>
                <div class="col-6">
                    <div class="kpi-mobile">
                        <div class="kpi-icon">
                            <i class="fas fa-map-marker-alt"></i>
                        </div>
                        <div class="kpi-value"><%= localidadesData.Count %></div>
                        <div class="kpi-label">Localidades</div>
                    </div>
                </div>
                <div class="col-6">
                    <div class="kpi-mobile">
                        <div class="kpi-icon">
                            <i class="fas fa-chart-line"></i>
                        </div>
                        <div class="kpi-value">
                            <% 
                            If totalGeralVendas > 0 Then 
                                Response.Write "R$ " & FormatNumber(totalGeralVGV / totalGeralVendas, 2)
                            Else 
                                Response.Write "R$ 0,00"
                            End If 
                            %>
                        </div>
                        <div class="kpi-label">VGV Médio</div>
                    </div>
                </div>
            </div>

            <!-- Lista de Localidades Mobile -->
            <%
            If localidadesData.Count > 0 Then
                Dim arrLocalidades, localidadeKey, posicao
                arrLocalidades = localidadesData.Keys
                posicao = 0
                
                ' Ordenar localidades por VGV (decrescente)
                For i = 0 To UBound(arrLocalidades)
                    For j = i + 1 To UBound(arrLocalidades)
                        If localidadesData(arrLocalidades(j))(0) > localidadesData(arrLocalidades(i))(0) Then
                            Dim tempLocalidade
                            tempLocalidade = arrLocalidades(i)
                            arrLocalidades(i) = arrLocalidades(j)
                            arrLocalidades(j) = tempLocalidade
                        End If
                    Next
                Next
                
                For Each localidadeKey In arrLocalidades
                    posicao = posicao + 1
                    Dim dadosLocalidade, percentualTotal
                    dadosLocalidade = localidadesData(localidadeKey)
                    
                    If totalGeralVGV > 0 Then
                        percentualTotal = (dadosLocalidade(0) / totalGeralVGV) * 100
                    Else
                        percentualTotal = 0
                    End If
            %>
            <div class="localidade-card">
                <div class="posicao-badge"><%= posicao %></div>
                
                <div class="localidade-nome">
                    <%= localidadeKey %>
                </div>
                
                <div class="valor-venda">
                    R$ <%= FormatNumber(dadosLocalidade(0), 2) %>
                </div>
                
                <!-- Barra de Percentual -->
                <div class="percentual-bar">
                    <div class="percentual-fill" style="width: <%= percentualTotal %>%"></div>
                    <div class="percentual-text"><%= FormatNumber(percentualTotal, 1) %>% do total</div>
                </div>
                
                <!-- Estatísticas -->
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-label">VENDAS</div>
                        <div class="stat-value"><%= dadosLocalidade(1) %></div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-label">VGV MÉDIO</div>
                        <div class="stat-value">R$ <%= FormatNumber(dadosLocalidade(2), 2) %></div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-label">EMPREEND.</div>
                        <div class="stat-value"><%= dadosLocalidade(3) %></div>
                    </div>
                </div>
            </div>
            <%
                Next
            Else
            %>
            <div class="no-results">
                <i class="fas fa-search"></i>
                <h4>Nenhuma venda encontrada</h4>
                <p>Não foram encontradas vendas para os filtros selecionados.</p>
            </div>
            <%
            End If
            %>

            <!-- Top 5 Localidades Mobile -->
            <% If localidadesData.Count > 0 Then %>
            <div class="top-localidades">
                <div class="top-title">
                    <i class="fas fa-trophy"></i> Top 5 Localidades
                </div>
                <%
                Dim contador
                contador = 0
                For Each localidadeKey In arrLocalidades
                    If contador < 5 Then
                        Dim dadosTop
                        dadosTop = localidadesData(localidadeKey)
                %>
                <div class="top-item">
                    <div class="top-pos"><%= contador + 1 %></div>
                    <div class="top-name"><%= localidadeKey %></div>
                    <div class="top-value">R$ <%= FormatNumber(dadosTop(0), 2) %></div>
                </div>
                <%
                        contador = contador + 1
                    Else
                        Exit For
                    End If
                Next
                %>
            </div>
            <% End If %>

            <!-- Totais Mobile -->
            <div class="total-card">
                <div class="row">
                    <div class="col-6">
                        <div class="total-label">Total VGV</div>
                        <div class="total-value">R$ <%= FormatNumber(totalGeralVGV, 2) %></div>
                    </div>
                    <div class="col-6">
                        <div class="total-label">Total Vendas</div>
                        <div class="total-value"><%= totalGeralVendas %></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- ===== VERSÃO DESKTOP ===== -->
        <div class="desktop-only">
            <div class="desktop-grid">
                <!-- Sidebar com KPIs -->
                <div class="sidebar-desktop">
                    <h5 class="mb-4"><i class="fas fa-chart-bar"></i> Resumo Geral</h5>
                    
                    <div class="kpi-desktop">
                        <div class="kpi-value-desktop">R$ <%= FormatNumber(totalGeralVGV, 2) %></div>
                        <div class="kpi-label-desktop">Total VGV</div>
                    </div>
                    
                    <div class="kpi-desktop">
                        <div class="kpi-value-desktop"><%= totalGeralVendas %></div>
                        <div class="kpi-label-desktop">Total de Vendas</div>
                    </div>
                    
                    <div class="kpi-desktop">
                        <div class="kpi-value-desktop"><%= localidadesData.Count %></div>
                        <div class="kpi-label-desktop">Localidades</div>
                    </div>
                    
                    <div class="kpi-desktop">
                        <div class="kpi-value-desktop">
                            <% 
                            If totalGeralVendas > 0 Then 
                                Response.Write "R$ " & FormatNumber(totalGeralVGV / totalGeralVendas, 2)
                            Else 
                                Response.Write "R$ 0,00"
                            End If 
                            %>
                        </div>
                        <div class="kpi-label-desktop">VGV Médio</div>
                    </div>

                    <!-- Top 5 Desktop -->
                    <% If localidadesData.Count > 0 Then %>
                    <div class="mt-4">
                        <h6><i class="fas fa-trophy"></i> Top 5 Localidades</h6>
                        <%
                        contador = 0
                        For Each localidadeKey In arrLocalidades
                            If contador < 5 Then
                                dadosTop = localidadesData(localidadeKey)
                        %>
                        <div class="top-item">
                            <div class="top-pos"><%= contador + 1 %></div>
                            <div class="top-name"><%= localidadeKey %></div>
                            <div class="top-value">R$ <%= FormatNumber(dadosTop(0), 2) %></div>
                        </div>
                        <%
                                contador = contador + 1
                            Else
                                Exit For
                            End If
                        Next
                        %>
                    </div>
                    <% End If %>
                </div>

                <!-- Conteúdo Principal -->
                <div class="main-content-desktop">
                    <!-- Tabela de Dados -->
                    <div class="table-desktop">
                        <table class="table table-hover mb-0">
                            <thead>
                                <tr>
                                    <th>#</th>
                                    <th>Localidade</th>
                                    <th>VGV Total</th>
                                    <th>Vendas</th>
                                    <th>VGV Médio</th>
                                    <th>Empreend.</th>
                                    <th>% do Total</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If localidadesData.Count > 0 Then
                                    posicao = 0
                                    For Each localidadeKey In arrLocalidades
                                        posicao = posicao + 1
                                        dadosLocalidade = localidadesData(localidadeKey)
                                        
                                        If totalGeralVGV > 0 Then
                                            percentualTotal = (dadosLocalidade(0) / totalGeralVGV) * 100
                                        Else
                                            percentualTotal = 0
                                        End If
                                %>
                                <tr>
                                    <td>
                                        <div class="posicao-badge-desktop">
                                            <%= posicao %>
                                        </div>
                                    </td>
                                    <td><strong><%= localidadeKey %></strong></td>
                                    <td class="text-success fw-bold">R$ <%= FormatNumber(dadosLocalidade(0), 2) %></td>
                                    <td><%= dadosLocalidade(1) %></td>
                                    <td>R$ <%= FormatNumber(dadosLocalidade(2), 2) %></td>
                                    <td><%= dadosLocalidade(3) %></td>
                                    <td>
                                        <div class="progress" style="height: 20px;">
                                            <div class="progress-bar bg-success" 
                                                 role="progressbar" 
                                                 style="width: <%= percentualTotal %>%"
                                                 aria-valuenow="<%= percentualTotal %>" 
                                                 aria-valuemin="0" 
                                                 aria-valuemax="100">
                                                <%= FormatNumber(percentualTotal, 1) %>%
                                            </div>
                                        </div>
                                    </td>
                                </tr>
                                <%
                                    Next
                                Else
                                %>
                                <tr>
                                    <td colspan="7" class="text-center py-4">
                                        <i class="fas fa-search fa-2x text-muted mb-2"></i>
                                        <p class="text-muted">Nenhuma venda encontrada para os filtros selecionados.</p>
                                    </td>
                                </tr>
                                <%
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>

                    <!-- Cards Detalhados -->
                    <div class="localidades-grid">
                        <%
                        If localidadesData.Count > 0 Then
                            posicao = 0
                            For Each localidadeKey In arrLocalidades
                                If posicao < 6 Then ' Limita a 6 cards para não sobrecarregar
                                    posicao = posicao + 1
                                    dadosLocalidade = localidadesData(localidadeKey)
                                    
                                    If totalGeralVGV > 0 Then
                                        percentualTotal = (dadosLocalidade(0) / totalGeralVGV) * 100
                                    Else
                                        percentualTotal = 0
                                    End If
                        %>
                        <div class="localidade-card-desktop">
                            <div class="localidade-header">
                                <div class="localidade-nome-desktop">
                                    <i class="fas fa-map-marker-alt text-primary"></i> <%= localidadeKey %>
                                </div>
                                <div class="valor-venda-desktop">
                                    R$ <%= FormatNumber(dadosLocalidade(0), 2) %>
                                </div>
                            </div>
                            
                            <div class="progress mb-3" style="height: 10px;">
                                <div class="progress-bar bg-success" 
                                     role="progressbar" 
                                     style="width: <%= percentualTotal %>%"
                                     aria-valuenow="<%= percentualTotal %>" 
                                     aria-valuemin="0" 
                                     aria-valuemax="100">
                                </div>
                            </div>
                            <div class="text-center text-muted small mb-3">
                                <%= FormatNumber(percentualTotal, 1) %>% do total geral
                            </div>
                            
                            <div class="stats-grid-desktop">
                                <div class="stat-item-desktop">
                                    <div class="stat-label-desktop">VENDAS</div>
                                    <div class="stat-value-desktop"><%= dadosLocalidade(1) %></div>
                                </div>
                                <div class="stat-item-desktop">
                                    <div class="stat-label-desktop">VGV MÉDIO</div>
                                    <div class="stat-value-desktop">R$ <%= FormatNumber(dadosLocalidade(2), 2) %></div>
                                </div>
                                <div class="stat-item-desktop">
                                    <div class="stat-label-desktop">EMPREEND.</div>
                                    <div class="stat-value-desktop"><%= dadosLocalidade(3) %></div>
                                </div>
                                <div class="stat-item-desktop">
                                    <div class="stat-label-desktop">POSIÇÃO</div>
                                    <div class="stat-value-desktop">#<%= posicao %></div>
                                </div>
                            </div>
                        </div>
                        <%
                                End If
                            Next
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>

        <% End If %>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function limparFiltros() {
            window.location.href = window.location.pathname;
        }

        // Animações suaves
        document.addEventListener('DOMContentLoaded', function() {
            const cards = document.querySelectorAll('.localidade-card, .localidade-card-desktop');
            cards.forEach((card, index) => {
                card.style.opacity = '0';
                card.style.transform = 'translateY(20px)';
                
                setTimeout(() => {
                    card.style.transition = 'all 0.5s ease';
                    card.style.opacity = '1';
                    card.style.transform = 'translateY(0)';
                }, index * 100);
            });
        });

        // Melhorar experiência mobile
        if (window.innerWidth < 992) {
            document.addEventListener('touchstart', function() {}, { passive: true });
        }
    </script>
</body>
</html>

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>