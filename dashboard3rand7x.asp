<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 04/12/2025                       -->
<!-- CODIGO_ARQUIVO: SRXGMTAPYU             -->
<!-- OBS: 04 DEZ LABEL NO GRAFICO           -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if 
%>

<%
' FUNÇÃO PARA POPULAR OS SELECTS DE FILTRO
Function GetUniqueValues(conn, fieldName, tableName)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    sqlQuery = "SELECT DISTINCT " & fieldName & " FROM " & tableName & " ORDER BY " & fieldName & ";"
    
    rs.Open sqlQuery, conn
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(fieldName)) Then
                dict.Add CStr(rs(fieldName)), 1
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    GetUniqueValues = dict.Keys
End Function

' FUNÇÃO PARA CONSTRUIR A CLÁUSULA WHERE
Function BuildWhereClause()
    Dim sqlWhere
    sqlWhere = " WHERE 1=1 AND Excluido = 0 AND Excluido IS NOT NULL"

    If Request.QueryString("ano") <> "" Then
        sqlWhere = sqlWhere & " AND AnoVenda = " & Request.QueryString("ano")
    Else
        ' Se não tiver ano filtrado, usar ano atual
        sqlWhere = sqlWhere & " AND AnoVenda = " & Year(Now)
    End If

    If Request.QueryString("mes") <> "" Then
        sqlWhere = sqlWhere & " AND MesVenda = " & Request.QueryString("mes")
    End If
    
    If Request.QueryString("diretoria") <> "" Then
        sqlWhere = sqlWhere & " AND Diretoria = '" & Replace(Request.QueryString("diretoria"), "'", "''") & "'"
    End If

    If Request.QueryString("gerencia") <> "" Then
        sqlWhere = sqlWhere & " AND Gerencia = '" & Replace(Request.QueryString("gerencia"), "'", "''") & "'"
    End If

    If Request.QueryString("corretor") <> "" Then
        sqlWhere = sqlWhere & " AND Corretor = '" & Replace(Request.QueryString("corretor"), "'", "''") & "'"
    End If

    If Request.QueryString("empreendimento") <> "" Then
        sqlWhere = sqlWhere & " AND NomeEmpreendimento = '" & Replace(Request.QueryString("empreendimento"), "'", "''") & "'"
    End If
    
    BuildWhereClause = sqlWhere
End Function

' =======================================================
' INÍCIO DO PROCESSAMENTO
' =======================================================

' Inicializar array de meses ANTES de qualquer uso
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

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConnSales

Dim whereClause
whereClause = BuildWhereClause()

Dim uniqueAnos, uniqueMeses, uniqueDiretorias, uniqueGerencias, uniqueCorretores, uniqueEmpreendimentos
uniqueAnos = GetUniqueValues(conn, "AnoVenda", "Vendas")
uniqueMeses = GetUniqueValues(conn, "MesVenda", "Vendas")
uniqueDiretorias = GetUniqueValues(conn, "Diretoria", "Vendas")
uniqueGerencias = GetUniqueValues(conn, "Gerencia", "Vendas")
uniqueCorretores = GetUniqueValues(conn, "Corretor", "Vendas")
uniqueEmpreendimentos = GetUniqueValues(conn, "NomeEmpreendimento", "Vendas")

' CALCULAR TICKET MÉDIO E QUANTIDADE DE UNIDADES
Dim ticketMedio, quantidadeUnidades, totalVendas, metaValor, metaRealizada, metaPercentual
quantidadeUnidades = 0
totalVendas = 0

SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalUnidades")) Then
        quantidadeUnidades = rs("TotalUnidades")
    End If
    If Not IsNull(rs("TotalVendas")) Then
        totalVendas = rs("TotalVendas")
    End If
End If
rs.Close

If quantidadeUnidades > 0 And totalVendas > 0 Then
    ticketMedio = totalVendas / quantidadeUnidades
Else
    ticketMedio = 0
End If

' BUSCAR METAS CONFORME FILTRO
Dim anoFiltro, mesFiltro
anoFiltro = Request.QueryString("ano")
mesFiltro = Request.QueryString("mes")

' Se não tiver filtro de ano, usar ano atual
If anoFiltro = "" Then
    anoFiltro = Year(Now)
End If

metaValor = 0
metaRealizada = 0
metaPercentual = 0

' Buscar metas de acordo com o filtro
If mesFiltro <> "" Then
    ' Se tiver filtro de mês, buscar meta específica do mês
    SQL = "SELECT Meta FROM MetaEmpresa WHERE Ano = " & anoFiltro & " AND Mes = " & mesFiltro
Else
    ' Se não tiver filtro de mês, somar todas as metas do ano
    SQL = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoFiltro
End If

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If mesFiltro <> "" Then
        ' Meta específica do mês
        If Not IsNull(rs("Meta")) Then
            metaValor = rs("Meta")
        End If
    Else
        ' Soma das metas do ano
        If Not IsNull(rs("MetaTotal")) Then
            metaValor = rs("MetaTotal")
        End If
    End If
End If
rs.Close

' Calcular percentual realizado
If metaValor > 0 Then
    metaRealizada = totalVendas
    metaPercentual = (metaRealizada / metaValor) * 100
    If metaPercentual > 100 Then metaPercentual = 100
End If

' =======================================================
' CALCULAR VGV POR DIRETORIA E GERÊNCIA COM % DA META
' =======================================================

' Estruturas para armazenar os dados
Dim dictDiretoriaVGV, dictDiretoriaPercent
Dim dictGerenciaVGV, dictGerenciaPercent

Set dictDiretoriaVGV = Server.CreateObject("Scripting.Dictionary")
Set dictDiretoriaPercent = Server.CreateObject("Scripting.Dictionary")
Set dictGerenciaVGV = Server.CreateObject("Scripting.Dictionary")
Set dictGerenciaPercent = Server.CreateObject("Scripting.Dictionary")

' 1. Buscar VGV por Diretoria
SQL = "SELECT Diretoria, SUM(ValorUnidade) AS VGV FROM Vendas " & whereClause & " AND Diretoria IS NOT NULL AND Diretoria <> '' GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

Do While Not rs.EOF
    Dim dirNome, dirVGV
    dirNome = Trim(rs("Diretoria"))
    dirVGV = 0
    If Not IsNull(rs("VGV")) Then
        dirVGV = CDbl(rs("VGV"))
    End If
    
    If dirNome <> "" Then
        dictDiretoriaVGV.Add dirNome, dirVGV
        
        ' Calcular % da meta
        If metaValor > 0 And dirVGV > 0 Then
            Dim dirPercent
            dirPercent = Round((dirVGV / metaValor) * 100, 1)
            dictDiretoriaPercent.Add dirNome, dirPercent
        Else
            dictDiretoriaPercent.Add dirNome, 0
        End If
    End If
    rs.MoveNext
Loop
rs.Close

' 2. Buscar VGV por Gerência
SQL = "SELECT Gerencia, SUM(ValorUnidade) AS VGV FROM Vendas " & whereClause & " AND Gerencia IS NOT NULL AND Gerencia <> '' GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

Do While Not rs.EOF
    Dim gerNome, gerVGV
    gerNome = Trim(rs("Gerencia"))
    gerVGV = 0
    If Not IsNull(rs("VGV")) Then
        gerVGV = CDbl(rs("VGV"))
    End If
    
    If gerNome <> "" Then
        dictGerenciaVGV.Add gerNome, gerVGV
        
        ' Calcular % da meta
        If metaValor > 0 And gerVGV > 0 Then
            Dim gerPercent
            gerPercent = Round((gerVGV / metaValor) * 100, 1)
            dictGerenciaPercent.Add gerNome, gerPercent
        Else
            dictGerenciaPercent.Add gerNome, 0
        End If
    End If
    rs.MoveNext
Loop
rs.Close

' Determinar texto do card de metas
Dim metaTitulo, metaSubtitulo
If mesFiltro <> "" Then
    metaTitulo = "Meta do Mês"
    If IsNumeric(mesFiltro) Then
        Dim mesNum
        mesNum = CInt(mesFiltro)
        If mesNum >= 1 And mesNum <= 12 Then
            metaSubtitulo = arrMesesNome(mesNum) & "/" & anoFiltro
        Else
            metaSubtitulo = "Mês " & mesFiltro & "/" & anoFiltro
        End If
    Else
        metaSubtitulo = "Mês " & mesFiltro & "/" & anoFiltro
    End If
ElseIf anoFiltro <> "" Then
    metaTitulo = ""
    metaSubtitulo = "Ano " & anoFiltro
Else
    metaTitulo = "Sem meta definida"
    metaSubtitulo = ""
End If

Dim autoTime
autoTime = Request.QueryString("autotime")
If autoTime = "" Then autoTime = 5
%>

<%
' =======================================================
' DADOS PARA O GRÁFICO DE QUANTIDADES VENDIDAS - VERSÃO SIMPLIFICADA
' =======================================================

Dim datasetsJSONQuantidades
datasetsJSONQuantidades = ""
Dim colorIndexQuant
colorIndexQuant = 0

' Cores diferentes para diferenciar dos valores
Dim colorsQuant
colorsQuant = Array("rgba(65, 105, 225, 1)", "rgba(50, 205, 50, 1)", "rgba(255, 140, 0, 1)", _
                    "rgba(148, 0, 211, 1)", "rgba(220, 20, 60, 1)", "rgba(30, 144, 255, 1)")

' Buscar anos
SQL_Anos_Quant = "SELECT DISTINCT AnoVenda FROM Vendas " & whereClause & " ORDER BY AnoVenda"
Set rsAnosQuant = Server.CreateObject("ADODB.Recordset")
rsAnosQuant.Open SQL_Anos_Quant, conn

Do Until rsAnosQuant.EOF
    ano = rsAnosQuant("AnoVenda")
    
    ' Inicializar array para 12 meses
    Dim monthlyData(12)
    For i = 1 to 12
        monthlyData(i) = 0
    Next
    
    ' Buscar quantidades por mês
    SQL_Dados_Quant = "SELECT MesVenda, COUNT(*) AS Quantidade FROM Vendas " & whereClause & _
                      " AND AnoVenda = " & ano & " GROUP BY MesVenda ORDER BY MesVenda"
    Set rsDadosQuant = Server.CreateObject("ADODB.Recordset")
    rsDadosQuant.Open SQL_Dados_Quant, conn
    
    Do Until rsDadosQuant.EOF
        If Not IsNull(rsDadosQuant("MesVenda")) Then
            mesNum = CInt(rsDadosQuant("MesVenda"))
            If mesNum >= 1 And mesNum <= 12 Then
                If Not IsNull(rsDadosQuant("Quantidade")) Then
                    monthlyData(mesNum) = rsDadosQuant("Quantidade")
                End If
            End If
        End If
        rsDadosQuant.MoveNext
    Loop
    rsDadosQuant.Close
    Set rsDadosQuant = Nothing
    
    ' Construir string de dados
    Dim dataString
    dataString = ""
    For i = 1 to 12
        dataString = dataString & monthlyData(i) & ","
    Next
    dataString = Left(dataString, Len(dataString) - 1) ' Remover última vírgula
    
    ' Adicionar ao dataset
    datasetsJSONQuantidades = datasetsJSONQuantidades & "{"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "label: 'Qtd " & ano & "',"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "data: [" & dataString & "],"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "borderColor: '" & colorsQuant(colorIndexQuant Mod 6) & "',"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "backgroundColor: '" & Replace(colorsQuant(colorIndexQuant Mod 6), "1)", "0.5)") & "',"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "borderWidth: 2,"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "fill: true"
    datasetsJSONQuantidades = datasetsJSONQuantidades & "},"
    
    colorIndexQuant = colorIndexQuant + 1
    rsAnosQuant.MoveNext
Loop

If Right(datasetsJSONQuantidades, 1) = "," Then 
    datasetsJSONQuantidades = Left(datasetsJSONQuantidades, Len(datasetsJSONQuantidades) - 1)
End If

If Not rsAnosQuant Is Nothing Then
    rsAnosQuant.Close
    Set rsAnosQuant = Nothing
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard de Vendas - Modo Autônomo</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <style>
        body {
            background-color: #f8f9fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            transform: scale(0.8);
            transform-origin: 0 0;
            width: calc(100% / 0.8);
            position: absolute;
            min-height: 100vh;
        }
        
        h1 {
            color: #343a40;
            text-align: center;
            margin-bottom: 30px !important;
            font-weight: 700;
        }
        
        h5 {
            color: #ffffff;
            margin-bottom: 15px;
            font-weight: 600;
        }
        
        .card {
            border: none;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
            transition: transform 0.3s ease;
        }
        
        .card:hover {
            transform: translateY(-5px);
        }
        
        .card-header {
            background: linear-gradient(135deg, #800000 0%, #B22222 100%);
            color: white;
            border-radius: 10px 10px 0 0 !important;
            font-weight: 600;
        }
        
        /* Layout 2 por linha para cards métricos */
        .two-per-row-metrics {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            width: 100%;
        }
        
        .two-per-row-metrics > .card {
            width: calc(50% - 7.5px);
            min-height: 150px; 
        }
        
        /* Cores dos cards */
        .bg-green { background-color: #28a745 !important; }
        .bg-orange { background-color: #f3722c !important; }
        .bg-purple { background-color: #7209b7 !important; }
        .bg-info { background-color: #4361ee !important; }
        .bg-primary { background-color: #4361ee !important; }
        .bg-success { background-color: #4cc9f0 !important; }
        .bg-warning { background-color: #f72585 !important; }
        .bg-pink { background: linear-gradient(135deg, #ff4d8d 0%, #ff6b9d 100%) !important; }
        .bg-teal { background-color: #2a9d8f !important; }
        
        .metric-card-small {
            text-align: center;
            padding: 15px !important;
            height: 120px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        
        .metric-value-small {
            font-size: 1.2rem !important;
            font-weight: bold;
            margin: 8px 0;
        }
        
        .metric-label {
            font-size: 0.9rem;
            color: rgba(255, 255, 255, 0.9);
            margin-bottom: 0;
        }
        
        .meta-subtitle {
            font-size: 0.8rem;
            color: rgba(255, 255, 255, 0.8);
            margin-top: -5px;
            margin-bottom: 10px;
        }
        
        .progress-container {
            margin-top: 10px;
            background: rgba(255, 255, 255, 0.3);
            border-radius: 10px;
            overflow: hidden;
        }
        
        .progress-bar {
            height: 10px;
            background: linear-gradient(90deg, #4cd964 0%, #5ac8fa 100%);
            border-radius: 10px;
            transition: width 1s ease-in-out;
        }
        
        /* Layout principal */
        .main-container {
            display: grid;
            grid-template-columns: 250px 1fr;
            gap: 20px;
            padding: 20px;
        }
        
        .filter-column {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
        
        /* ESTILOS PARA LISTAGEM DE VGV */
        .vgv-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            font-size: 0.9rem;
        }
        
        .vgv-table th {
            background-color: #f8f9fa;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            color: #495057;
            border-bottom: 2px solid #dee2e6;
        }
        
        .vgv-table td {
            padding: 10px 12px;
            border-bottom: 1px solid #e9ecef;
            vertical-align: middle;
        }
        
        .vgv-value {
            font-weight: bold;
            color: #2c3e50;
        }
        
        .vgv-percent {
            font-weight: bold;
            text-align: center;
            min-width: 80px;
        }
        
        .percent-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 0.8rem;
            font-weight: bold;
            min-width: 60px;
            text-align: center;
        }
        
        .percent-excelente { background-color: #28a745; color: white; }
        .percent-bom { background-color: #17a2b8; color: white; }
        .percent-medio { background-color: #ffc107; color: black; }
        .percent-baixo { background-color: #fd7e14; color: white; }
        .percent-critico { background-color: #dc3545; color: white; }
        
        .progress-small {
            height: 6px;
            background-color: #e9ecef;
            border-radius: 3px;
            overflow: hidden;
            margin-top: 5px;
        }
        
        .total-row {
            background-color: #e8f4fd !important;
            font-weight: bold;
        }
        
        .ranking {
            width: 30px;
            text-align: center;
            font-weight: bold;
            color: #6c757d;
        }
        
        .table-container {
            max-height: 600px;
            overflow-y: auto;
            border: 1px solid #dee2e6;
            border-radius: 5px;
        }
        
        /* Spinner e timer */
        .spinner-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.8);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }
        
        #countdown-timer {
            position: fixed;
            top: 20px;
            left: 20px;
            background-color: rgba(0, 0, 0, 0.7);
            color: white;
            padding: 5px 10px;
            border-radius: 5px;
            font-size: 1rem;
            font-weight: bold;
            display: none;
            z-index: 999;
            width: auto;
        }
        
        /* Estilo para aumentar a altura das barras dos gráficos */
        .grafico-container {
            height: 350px !important; /* Aumentado de 250px para 350px */
            min-height: 350px;
        }
        
        /* Responsividade */
        @media (max-width: 992px) {
            .main-container {
                grid-template-columns: 1fr;
            }
            .sidebar {
                grid-column: 1 / -1;
            }
        }
        
        @media (max-width: 768px) {
            .two-per-row-metrics {
                flex-direction: column;
            }
            
            .two-per-row-metrics > .card {
                width: 100%;
            }
            
            .grafico-container {
                height: 300px !important;
                min-height: 300px;
            }
        }
        
        /* Estilos para a nova seção de piores gerências 05/12/2025 */
        .bg-excelente { background-color: #28a745 !important; color: white; }
        .bg-bom { background-color: #17a2b8 !important; color: white; }
        .bg-medio { background-color: #ffc107 !important; color: black; }
        .bg-baixo { background-color: #fd7e14 !important; color: white; }
        .bg-critico { background-color: #dc3545 !important; color: white; }
        
        /* Ajuste no layout para acomodar a nova seção */
        .row.mb-4 .col-md-6 {
            margin-bottom: 20px;
        }
        
        /* Estilo para a linha de média das piores */
        .total-row[style*="background-color: #f8d7da"] td {
            font-weight: bold;
            color: #721c24;
        }
    </style>
</head>
<body>

<div class="spinner-overlay" id="loadingSpinner">
    <div class="spinner-border text-primary" role="status">
        <span class="visually-hidden">Loading...</span>
    </div>
</div>

<div class="container-fluid">
    <h1 class="mb-1 text-center">Dashboard de Vendas</h1>
    <div class="main-container">
        <div class="sidebar">
            <div class="text-center mt-4">
                <a href="dashb_comp_metas3.asp" class="btn btn-primary btn-sm" target="_blank">
                    <i class="fas fa-arrow-right"></i> Dashboard Metas
                </a>
            </div>
            <div class="filter-column">
                <h5 class="text-center">Filtros</h5>
                <form method="get" id="filterForm">
                    <div class="mb-3">
                        <label for="anoFilter" class="form-label">Ano</label>
                        <select class="form-select" id="anoFilter" name="ano">
                            <option value="">Todos</option>
                            <% 
                            Dim anoAtual, anoSelecionado
                            anoAtual = Year(Now)
                            anoSelecionado = Request.QueryString("ano")
                            
                            For Each ano In uniqueAnos 
                                Response.Write "<option value='" & ano & "'"
                                If (anoSelecionado = "" And CStr(ano) = CStr(anoAtual)) Or anoSelecionado = CStr(ano) Then 
                                    Response.Write " selected"
                                End If
                                Response.Write ">" & ano & "</option>"
                            Next 
                            %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="mesFilter" class="form-label">Mês</label>
                        <select class="form-select" id="mesFilter" name="mes">
                            <option value="">Todos</option>
                            <% For Each mes In uniqueMeses %>
                                <option value="<%=mes%>" <% If Request.QueryString("mes") = CStr(mes) Then Response.Write "selected" %>><%=arrMesesNome(CInt(mes))%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="diretoriaFilter" class="form-label">Diretoria</label>
                        <select class="form-select" id="diretoriaFilter" name="diretoria">
                            <option value="">Todas</option>
                            <% For Each dir In uniqueDiretorias %>
                                <option value="<%=dir%>" <% If Request.QueryString("diretoria") = dir Then Response.Write "selected" %>><%=dir%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="gerenciaFilter" class="form-label">Gerência</label>
                        <select class="form-select" id="gerenciaFilter" name="gerencia">
                            <option value="">Todas</option>
                            <% For Each ger In uniqueGerencias %>
                                <option value="<%=ger%>" <% If Request.QueryString("gerencia") = ger Then Response.Write "selected" %>><%=ger%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="corretorFilter" class="form-label">Corretor</label>
                        <select class="form-select" id="corretorFilter" name="corretor">
                            <option value="">Todos</option>
                            <% For Each corr In uniqueCorretores %>
                                <option value="<%=corr%>" <% If Request.QueryString("corretor") = corr Then Response.Write "selected" %>><%=corr%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="empreendimentoFilter" class="form-label">Empreendimento</label>
                        <select class="form-select" id="empreendimentoFilter" name="empreendimento">
                            <option value="">Todos</option>
                            <% For Each emp In uniqueEmpreendimentos %>
                                <option value="<%=emp%>" <% If Request.QueryString("empreendimento") = emp Then Response.Write "selected" %>><%=emp%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label for="autoTimeFilter" class="form-label">Tempo de Atualização</label>
                        <select class="form-select" id="autoTimeFilter" name="autotime">
                            <option value="5" <% If CStr(autoTime) = "5" Then Response.Write "selected" %>>5s</option>
                            <option value="10" <% If CStr(autoTime) = "10" Then Response.Write "selected" %>>10s</option>
                            <option value="15" <% If CStr(autoTime) = "15" Then Response.Write "selected" %>>15s</option>
                            <option value="20" <% If CStr(autoTime) = "20" Then Response.Write "selected" %>>20s</option>
                            <option value="25" <% If CStr(autoTime) = "25" Then Response.Write "selected" %>>25s</option>
                            <option value="30" <% If CStr(autoTime) = "30" Then Response.Write "selected" %>>30s</option>
                        </select>
                    </div>
                    <div class="d-grid gap-2">
                        <button type="submit" class="btn btn-primary">
                            <i class="fas fa-search"></i> Filtrar
                        </button>
                        <button type="button" class="btn btn-secondary" onclick="window.location.href='<%= Request.ServerVariables("SCRIPT_NAME") %>'">
                            <i class="fas fa-times"></i> Limpar Filtros
                        </button>
                        <button type="button" class="btn btn-info mt-3" id="autoModeBtn">
                            <i class="fas fa-play-circle"></i> Iniciar Modo Autônomo
                        </button>
                    </div>
                </form>
            </div>
        </div>

        <div class="content-center">
            <!-- Cards Métricos Superiores -->
            <div class="two-per-row-metrics">
                <div class="card bg-orange text-white">
                    <div class="card-body metric-card-small">
                        <i class="fas fa-chart-line fa-2x mb-2"></i>
                        <div class="metric-value-small">R$ <%=FormatNumber(totalVendas, 0)%></div>
                        <p class="metric-label">VGV Total</p>
                    </div>
                </div>
                
                <div class="card bg-purple text-white">
                    <div class="card-body metric-card-small">
                        <i class="fas fa-ticket-alt fa-2x mb-2"></i>
                        <div class="metric-value-small">R$ <%=FormatNumber(ticketMedio, 2)%></div>
                        <p class="metric-label">Ticket Médio</p>
                    </div>
                </div>
                
                <div class="card bg-green text-white">
                    <div class="card-body metric-card-small">
                        <i class="fas fa-cube fa-2x mb-2"></i>
                        <div class="metric-value-small"><%=FormatNumber(quantidadeUnidades, 0)%></div>
                        <p class="metric-label">Unidades Vendidas</p>
                    </div>
                </div>
                
                <div class="card bg-info text-white">
                    <div class="card-body metric-card-small">
                        <h5 class="metric-label-small"><%=metaTitulo%></h5>
                        <% If metaSubtitulo <> "" Then %>
                            <p class="meta-subtitle"><%=metaSubtitulo%></p>
                        <% End If %>
                        
                        <% If metaValor > 0 Then %>
                            <div class="metric-value-normal"><%=FormatNumber(metaPercentual, 1)%>%</div>
                            <div class="progress-container">
                                <div class="progress-bar" style="width: <%=metaPercentual%>%"></div>
                            </div>
                            <div class="row mt-2">
                                <div class="col-6">
                                    <small>Meta: R$ <%=FormatNumber(metaValor, 0)%></small>
                                </div>
                                <div class="col-6">
                                    <small>Realizado: R$ <%=FormatNumber(metaRealizada, 0)%></small>
                                </div>
                            </div>
                        <% Else %>
                            <div class="metric-value-normal">--</div>
                            <p class="small mt-2"><i class="fas fa-info-circle"></i> Defina metas no sistema</p>
                        <% End If %>
                    </div>
                </div>
            </div>
<!-- @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  -->
                       <!-- Seção VGV por Diretoria com Quantidade -->
            <div class="row mb-4">
                <div class="col-md-12">
                    <div class="card">
                        <div class="card-header text-white" style="background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%);">
                            <h5 class="mb-0"><i class="fas fa-building"></i> VGV por Diretoria</h5>
                        </div>
                        <div class="card-body">
                            <%
                            ' Buscar dados de diretorias com quantidade de unidades
                            Dim dictDiretoriaQuantidade
                            Set dictDiretoriaQuantidade = Server.CreateObject("Scripting.Dictionary")
                            
                            ' Buscar quantidade por diretoria
                            SQL = "SELECT Diretoria, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " AND Diretoria IS NOT NULL AND Diretoria <> '' GROUP BY Diretoria"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn
                            
                            Do While Not rs.EOF
                                Dim dirNomeQtd
                                dirNomeQtd = Trim(rs("Diretoria"))
                                If dirNomeQtd <> "" Then
                                    If Not IsNull(rs("Quantidade")) Then
                                        dictDiretoriaQuantidade(dirNomeQtd) = rs("Quantidade")
                                    Else
                                        dictDiretoriaQuantidade(dirNomeQtd) = 0
                                    End If
                                End If
                                rs.MoveNext
                            Loop
                            rs.Close
                            
                            If dictDiretoriaVGV.Count > 0 Then 
                                Dim dirTotalQuantidade
                                dirTotalQuantidade = 0
                                
                                ' Calcular total de quantidade - CORREÇÃO AQUI
                                For Each dirNome In dictDiretoriaQuantidade
                                    dirTotalQuantidade = dirTotalQuantidade + dictDiretoriaQuantidade(dirNome)
                                Next
                            %>
                                <div class="table-container" style="max-height: 800px;">
                                    <table class="vgv-table">
                                        <thead>
                                            <tr>
                                                <th style="width: 40px;">#</th>
                                                <th>Diretoria</th>
                                                <th style="text-align: right;">VGV</th>
                                                <th style="text-align: center;">Qtd Unidades</th>
                                                <th style="text-align: center;">% da Meta</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <%
                                            Dim dirRank, dirTotalVGV
                                            dirRank = 1
                                            dirTotalVGV = 0
                                            
                                            ' Calcular total VGV
                                            For Each dirNome In dictDiretoriaVGV
                                                dirTotalVGV = dirTotalVGV + dictDiretoriaVGV(dirNome)
                                            Next
                                            
                                            Dim dirKeys
                                            dirKeys = dictDiretoriaVGV.Keys
                                            
                                            Dim tempDirArray()
                                            ReDim tempDirArray(dictDiretoriaVGV.Count - 1, 1)
                                            
                                            i = 0
                                            For Each dirNome In dirKeys
                                                tempDirArray(i, 0) = dirNome
                                                tempDirArray(i, 1) = dictDiretoriaVGV(dirNome)
                                                i = i + 1
                                            Next
                                            
                                            Dim j, tempNome, tempValor
                                            For i = 0 To UBound(tempDirArray, 1) - 1
                                                For j = i + 1 To UBound(tempDirArray, 1)
                                                    If tempDirArray(i, 1) < tempDirArray(j, 1) Then
                                                        tempNome = tempDirArray(i, 0)
                                                        tempValor = tempDirArray(i, 1)
                                                        tempDirArray(i, 0) = tempDirArray(j, 0)
                                                        tempDirArray(i, 1) = tempDirArray(j, 1)
                                                        tempDirArray(j, 0) = tempNome
                                                        tempDirArray(j, 1) = tempValor
                                                    End If
                                                Next
                                            Next
                                            
                                            For i = 0 To UBound(tempDirArray, 1)
                                                dirNome = tempDirArray(i, 0)
                                                Dim dirVGVDisplay, dirPercentDisplay, dirPercentClass, dirQuantidade
                                                dirVGVDisplay = dictDiretoriaVGV(dirNome)
                                                
                                                ' Obter quantidade - com verificação
                                                dirQuantidade = 0
                                                If dictDiretoriaQuantidade.Exists(dirNome) Then
                                                    dirQuantidade = dictDiretoriaQuantidade(dirNome)
                                                End If
                                                
                                                If dictDiretoriaPercent.Exists(dirNome) Then
                                                    dirPercentDisplay = dictDiretoriaPercent(dirNome)
                                                Else
                                                    dirPercentDisplay = 0
                                                End If
                                                
                                                If dirPercentDisplay >= 100 Then
                                                    dirPercentClass = "percent-excelente"
                                                ElseIf dirPercentDisplay >= 75 Then
                                                    dirPercentClass = "percent-bom"
                                                ElseIf dirPercentDisplay >= 50 Then
                                                    dirPercentClass = "percent-medio"
                                                ElseIf dirPercentDisplay >= 25 Then
                                                    dirPercentClass = "percent-baixo"
                                                Else
                                                    dirPercentClass = "percent-critico"
                                                End If
                                                
                                                Dim progressWidthDir
                                                If dirTotalVGV > 0 Then
                                                    progressWidthDir = (dirVGVDisplay / dirTotalVGV) * 100
                                                Else
                                                    progressWidthDir = 0
                                                End If
                                                
                                                Dim progressWidthQtd
                                                If dirTotalQuantidade > 0 Then
                                                    progressWidthQtd = (dirQuantidade / dirTotalQuantidade) * 100
                                                Else
                                                    progressWidthQtd = 0
                                                End If
                                                
                                                Dim percentTotalDir
                                                If dirTotalVGV > 0 Then
                                                    percentTotalDir = FormatNumber((dirVGVDisplay / dirTotalVGV) * 100, 1)
                                                Else
                                                    percentTotalDir = "0.0"
                                                End If
                                                %>
                                                <tr>
                                                    <td class="ranking"><%=dirRank%></td>
                                                    <td><strong><%=dirNome%></strong></td>
                                                    <td style="text-align: right;" class="vgv-value">
                                                        R$ <%=FormatNumber(dirVGVDisplay, 0)%>
                                                        <div class="progress-small">
                                                            <div class="progress-bar-small bg-info" style="width: <%=progressWidthDir%>%"></div>
                                                        </div>
                                                        <small class="text-muted">
                                                            <%=percentTotalDir%>% do total
                                                        </small>
                                                    </td>
                                                    <td style="text-align: center;">
                                                        <span class="badge bg-primary">
                                                            <%=dirQuantidade%> un
                                                        </span>
                                                        <div class="progress-small mt-1">
                                                            <div class="progress-bar-small bg-primary" style="width: <%=progressWidthQtd%>%"></div>
                                                        </div>
                                                        <small class="text-muted">
                                                            <%
                                                            Dim percentTotalQtd
                                                            If dirTotalQuantidade > 0 Then
                                                                percentTotalQtd = FormatNumber((dirQuantidade / dirTotalQuantidade) * 100, 1)
                                                            Else
                                                                percentTotalQtd = "0.0"
                                                            End If
                                                            %>
                                                            <%=percentTotalQtd%>% das unidades
                                                        </small>
                                                    </td>
                                                    <td style="text-align: center;">
                                                        <span class="percent-badge <%=dirPercentClass%>">
                                                            <%=FormatNumber(dirPercentDisplay, 1)%>%
                                                        </span>
                                                        <div class="progress-small">
                                                            <%
                                                            Dim progressWidthMetaDir
                                                            If dirPercentDisplay > 100 Then
                                                                progressWidthMetaDir = 100
                                                            Else
                                                                progressWidthMetaDir = dirPercentDisplay
                                                            End If
                                                            %>
                                                            <div class="progress-bar-small <%=dirPercentClass%>" style="width: <%=progressWidthMetaDir%>%"></div>
                                                        </div>
                                                    </td>
                                                </tr>
                                                <%
                                                dirRank = dirRank + 1
                                            Next
                                            %>
                                            <tr class="total-row">
                                                <td colspan="2"><strong>TOTAL DIRETORIAS</strong></td>
                                                <td style="text-align: right;">
                                                    <strong>R$ <%=FormatNumber(dirTotalVGV, 0)%></strong>
                                                </td>
                                                <td style="text-align: center;">
                                                    <strong><%=dirTotalQuantidade%> un</strong>
                                                </td>
                                                <td style="text-align: center;">
                                                    <%
                                                    Dim dirTotalPercent
                                                    If metaValor > 0 Then
                                                        dirTotalPercent = Round((dirTotalVGV / metaValor) * 100, 1)
                                                    Else
                                                        dirTotalPercent = 0
                                                    End If
                                                    %>
                                                    <strong><%=FormatNumber(dirTotalPercent, 1)%>%</strong>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            <% Else %>
                                <p class="text-center text-muted mt-3">
                                    <i class="fas fa-info-circle"></i> Nenhuma diretoria com vendas no período
                                </p>
                            <% End If %>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Seção VGV por Gerência -->
<!-- Seção VGV por Gerência -->
<div class="row mb-4">
    <div class="col-md-12">
        <div class="card">
            <div class="card-header text-white" style="background: linear-gradient(135deg, #4cc9f0 0%, #3a0ca3 100%);">
                <h5 class="mb-0"><i class="fas fa-user-tie"></i> VGV por Gerência</h5>
            </div>
            <div class="card-body">
                <% 
                ' Verificar se o dicionário existe e tem dados
                If Not dictGerenciaVGV Is Nothing Then
                    If dictGerenciaVGV.Count > 0 Then 
                %>
                    <div class="table-container" style="max-height: 800px;">
                        <table class="vgv-table">
                            <thead>
                                <tr>
                                    <th style="width: 40px;">#</th>
                                    <th>Gerência</th>
                                    <th style="text-align: right;">VGV</th>
                                    <th style="text-align: center;">% da Meta</th>
                                    <th style="text-align: center;">Qtd</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim gerRank, gerTotalVGV, gerCount, gerTotalQtd
                                gerRank = 1
                                gerTotalVGV = 0
                                gerCount = 0
                                gerTotalQtd = 0
                                
                                ' Criar dicionário para quantidades
                                Dim dictGerenciaQtd
                                Set dictGerenciaQtd = Server.CreateObject("Scripting.Dictionary")
                                
                                ' Buscar quantidade por gerência
                                SQL = "SELECT Gerencia, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " AND Gerencia IS NOT NULL AND Gerencia <> '' GROUP BY Gerencia"
                                Set rs = Server.CreateObject("ADODB.Recordset")
                                rs.Open SQL, conn
                                
                                Do While Not rs.EOF
                                    Dim gerNomeQtd
                                    gerNomeQtd = Trim(rs("Gerencia"))
                                    If gerNomeQtd <> "" Then
                                        If Not IsNull(rs("Quantidade")) Then
                                            dictGerenciaQtd.Add gerNomeQtd, CLng(rs("Quantidade"))
                                        Else
                                            dictGerenciaQtd.Add gerNomeQtd, 0
                                        End If
                                    End If
                                    rs.MoveNext
                                Loop
                                rs.Close
                                
                                ' Calcular total VGV e total quantidade
                                For Each gerNome In dictGerenciaVGV
                                    gerTotalVGV = gerTotalVGV + dictGerenciaVGV(gerNome)
                                    If dictGerenciaQtd.Exists(gerNome) Then
                                        gerTotalQtd = gerTotalQtd + dictGerenciaQtd(gerNome)
                                    End If
                                Next
                                
                                Dim gerKeys
                                gerKeys = dictGerenciaVGV.Keys
                                
                                Dim tempGerArray()
                                ReDim tempGerArray(dictGerenciaVGV.Count - 1, 2) ' 0=Nome, 1=VGV, 2=Qtd
                                
                                i = 0
                                For Each gerNome In gerKeys
                                    tempGerArray(i, 0) = gerNome
                                    tempGerArray(i, 1) = dictGerenciaVGV(gerNome)
                                    If dictGerenciaQtd.Exists(gerNome) Then
                                        tempGerArray(i, 2) = dictGerenciaQtd(gerNome)
                                    Else
                                        tempGerArray(i, 2) = 0
                                    End If
                                    i = i + 1
                                Next
                                
                                ' Ordenar por VGV (decrescente)
                                For i = 0 To UBound(tempGerArray, 1) - 1
                                    For j = i + 1 To UBound(tempGerArray, 1)
                                        If tempGerArray(i, 1) < tempGerArray(j, 1) Then
                                            tempNome = tempGerArray(i, 0)
                                            tempValor = tempGerArray(i, 1)
                                            tempQtd = tempGerArray(i, 2)
                                            
                                            tempGerArray(i, 0) = tempGerArray(j, 0)
                                            tempGerArray(i, 1) = tempGerArray(j, 1)
                                            tempGerArray(i, 2) = tempGerArray(j, 2)
                                            
                                            tempGerArray(j, 0) = tempNome
                                            tempGerArray(j, 1) = tempValor
                                            tempGerArray(j, 2) = tempQtd
                                        End If
                                    Next
                                Next
                                
                                For i = 0 To UBound(tempGerArray, 1)
                                    If gerCount >= 10 Then Exit For
                                    
                                    gerNome = tempGerArray(i, 0)
                                    Dim gerVGVDisplay, gerPercentDisplay, gerPercentClass, gerQtdDisplay
                                    gerVGVDisplay = tempGerArray(i, 1)
                                    gerQtdDisplay = tempGerArray(i, 2)
                                    
                                    If dictGerenciaPercent.Exists(gerNome) Then
                                        gerPercentDisplay = dictGerenciaPercent(gerNome)
                                    Else
                                        gerPercentDisplay = 0
                                    End If
                                    
                                    If gerPercentDisplay >= 100 Then
                                        gerPercentClass = "percent-excelente"
                                    ElseIf gerPercentDisplay >= 75 Then
                                        gerPercentClass = "percent-bom"
                                    ElseIf gerPercentDisplay >= 50 Then
                                        gerPercentClass = "percent-medio"
                                    ElseIf gerPercentDisplay >= 25 Then
                                        gerPercentClass = "percent-baixo"
                                    Else
                                        gerPercentClass = "percent-critico"
                                    End If
                                    
                                    Dim progressWidthGer
                                    If gerTotalVGV > 0 Then
                                        progressWidthGer = (gerVGVDisplay / gerTotalVGV) * 100
                                    Else
                                        progressWidthGer = 0
                                    End If
                                    
                                    Dim percentTotalGer
                                    If gerTotalVGV > 0 Then
                                        percentTotalGer = FormatNumber((gerVGVDisplay / gerTotalVGV) * 100, 1)
                                    Else
                                        percentTotalGer = "0.0"
                                    End If
                                    
                                    ' Calcular % da quantidade total
                                    Dim progressWidthQtdGer
                                    If gerTotalQtd > 0 Then
                                        progressWidthQtdGer = (gerQtdDisplay / gerTotalQtd) * 100
                                    Else
                                        progressWidthQtdGer = 0
                                    End If
                                    %>
                                    <tr>
                                        <td class="ranking"><%=gerRank%></td>
                                        <td><strong><%=gerNome%></strong></td>
                                        <td style="text-align: right;" class="vgv-value">
                                            R$ <%=FormatNumber(gerVGVDisplay, 0)%>
                                            <div class="progress-small">
                                                <div class="progress-bar-small bg-success" style="width: <%=progressWidthGer%>%"></div>
                                            </div>
                                            <small class="text-muted">
                                                <%=percentTotalGer%>% do total VGV
                                            </small>
                                        </td>
                                        <td style="text-align: center;">
                                            <span class="percent-badge <%=gerPercentClass%>">
                                                <%=FormatNumber(gerPercentDisplay, 1)%>%
                                            </span>
                                            <div class="progress-small">
                                                <%
                                                Dim progressWidthMetaGer
                                                If gerPercentDisplay > 100 Then
                                                    progressWidthMetaGer = 100
                                                Else
                                                    progressWidthMetaGer = gerPercentDisplay
                                                End If
                                                %>
                                                <div class="progress-bar-small <%=gerPercentClass%>" style="width: <%=progressWidthMetaGer%>%"></div>
                                            </div>
                                        </td>
                                        <td style="text-align: center;">
                                            <span class="badge bg-info">
                                                <%=FormatNumber(gerQtdDisplay, 0)%>
                                            </span>
                                            <div class="progress-small mt-1">
                                                <div class="progress-bar-small bg-info" style="width: <%=progressWidthQtdGer%>%"></div>
                                            </div>
                                            <small class="text-muted">
                                                <%
                                                Dim percentTotalQtdGer
                                                If gerTotalQtd > 0 Then
                                                    percentTotalQtdGer = FormatNumber((gerQtdDisplay / gerTotalQtd) * 100, 1)
                                                Else
                                                    percentTotalQtdGer = "0.0"
                                                End If
                                                %>
                                                <%=percentTotalQtdGer%>% das unidades
                                            </small>
                                        </td>
                                    </tr>
                                    <%
                                    gerRank = gerRank + 1
                                    gerCount = gerCount + 1
                                Next
                                %>
                                <tr class="total-row">
                                    <td colspan="2"><strong>TOTAL GERÊNCIAS (Top 10)</strong></td>
                                    <td style="text-align: right;">
                                        <strong>R$ <%=FormatNumber(gerTotalVGV, 0)%></strong>
                                    </td>
                                    <td style="text-align: center;">
                                        <%
                                        Dim gerTotalPercent
                                        If metaValor > 0 Then
                                            gerTotalPercent = Round((gerTotalVGV / metaValor) * 100, 1)
                                        Else
                                            gerTotalPercent = 0
                                        End If
                                        %>
                                        <strong><%=FormatNumber(gerTotalPercent, 1)%>%</strong>
                                    </td>
                                    <td style="text-align: center;">
                                        <strong><%=FormatNumber(gerTotalQtd, 0)%></strong>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                <% Else %>
                    <p class="text-center text-muted mt-3">
                        <i class="fas fa-info-circle"></i> Nenhuma gerência com vendas no período
                    </p>
                <% End If %>
                <% Else %>
                    <p class="text-center text-muted mt-3">
                        <i class="fas fa-info-circle"></i> Dados de gerência não disponíveis
                    </p>
                <% End If %>
            </div>
        </div>
    </div>
</div>




            <!-- NOVA SEÇÃO: TOP 5 PIORES GERÊNCIAS EM RELAÇÃO À META -->
<div class="row mb-4">
    <div class="col-md-12">
        <div class="card">
            <div class="card-header text-white" style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);">
                <h5 class="mb-0"><i class="fas fa-exclamation-triangle"></i> Top 5 Gerências Abaixo das Metas</h5>
            </div>
            <div class="card-body">
                <% 
                ' =======================================================
                ' CÁLCULO DAS PIORES GERÊNCIAS EM RELAÇÃO À META
                ' =======================================================
                
                If dictGerenciaVGV.Count > 0 Then 
                    ' Calcular o total de VGV de todas as gerências
                    Dim gerTotalVGVForWorst
                    gerTotalVGVForWorst = 0
                    
                    For Each gerNome In dictGerenciaVGV
                        gerTotalVGVForWorst = gerTotalVGVForWorst + dictGerenciaVGV(gerNome)
                    Next
                    
                    ' Criar array para ordenar por % da meta (do menor para o maior)
                    Dim worstGerArray()
                    ReDim worstGerArray(dictGerenciaVGV.Count - 1, 3) ' 0=Nome, 1=VGV, 2=%, 3=Class
                    
                    i = 0
                    For Each gerNome In dictGerenciaVGV
                        worstGerArray(i, 0) = gerNome
                        worstGerArray(i, 1) = dictGerenciaVGV(gerNome)
                        
                        If dictGerenciaPercent.Exists(gerNome) Then
                            worstGerArray(i, 2) = dictGerenciaPercent(gerNome)
                        Else
                            worstGerArray(i, 2) = 0
                        End If
                        
                        ' Classificar por performance
                        If worstGerArray(i, 2) >= 100 Then
                            worstGerArray(i, 3) = "excelente"
                        ElseIf worstGerArray(i, 2) >= 75 Then
                            worstGerArray(i, 3) = "bom"
                        ElseIf worstGerArray(i, 2) >= 50 Then
                            worstGerArray(i, 3) = "medio"
                        ElseIf worstGerArray(i, 2) >= 25 Then
                            worstGerArray(i, 3) = "baixo"
                        Else
                            worstGerArray(i, 3) = "critico"
                        End If
                        
                        i = i + 1
                    Next
                    
                    ' Ordenar por % da meta (do menor para o maior - piores primeiro)
                    For i = 0 To UBound(worstGerArray, 1) - 1
                        For j = i + 1 To UBound(worstGerArray, 1)
                            If worstGerArray(i, 2) > worstGerArray(j, 2) Then
                                tempNome = worstGerArray(i, 0)
                                tempValor = worstGerArray(i, 1)
                                tempPercent = worstGerArray(i, 2)
                                tempClass = worstGerArray(i, 3)
                                
                                worstGerArray(i, 0) = worstGerArray(j, 0)
                                worstGerArray(i, 1) = worstGerArray(j, 1)
                                worstGerArray(i, 2) = worstGerArray(j, 2)
                                worstGerArray(i, 3) = worstGerArray(j, 3)
                                
                                worstGerArray(j, 0) = tempNome
                                worstGerArray(j, 1) = tempValor
                                worstGerArray(j, 2) = tempPercent
                                worstGerArray(j, 3) = tempClass
                            End If
                        Next
                    Next
                    
                    ' Calcular média das piores
                    Dim worstGerTotalVGV, worstGerAvgVGV, worstGerAvgPercent, worstCount
                    worstGerTotalVGV = 0
                    worstGerAvgPercent = 0
                    worstCount = 0
                    
                    ' Pegar apenas as 5 piores
                    Dim showWorstCount
                    showWorstCount = 5
                    If UBound(worstGerArray, 1) + 1 < showWorstCount Then
                        showWorstCount = UBound(worstGerArray, 1) + 1
                    End If
                %>
                
                <div class="table-container" style="max-height: 400px;">
                    <table class="vgv-table">
                        <thead>
                            <tr>
                                <th style="width: 30px;">#</th>
                                <th>Gerência</th>
                                <th style="text-align: right;">VGV</th>
                                <th style="text-align: center;">% da Meta</th>
                                <th style="text-align: center;">Qtd</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            ' Buscar quantidades para as piores gerências
                            Dim dictWorstGerenciaQtd
                            Set dictWorstGerenciaQtd = Server.CreateObject("Scripting.Dictionary")
                            
                            ' Buscar quantidade por gerência (apenas para as piores)
                            SQL = "SELECT Gerencia, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " AND Gerencia IS NOT NULL AND Gerencia <> '' GROUP BY Gerencia"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn
                            
                            Do While Not rs.EOF
                              ' ' Dim gerNomeQtd
                                gerNomeQtd = Trim(rs("Gerencia"))
                                If gerNomeQtd <> "" Then
                                    If Not IsNull(rs("Quantidade")) Then
                                        dictWorstGerenciaQtd.Add gerNomeQtd, CLng(rs("Quantidade"))
                                    Else
                                        dictWorstGerenciaQtd.Add gerNomeQtd, 0
                                    End If
                                End If
                                rs.MoveNext
                            Loop
                            rs.Close
                            
                            For i = 0 To showWorstCount - 1
                                If worstCount >= 5 Then Exit For
                                
                                Dim worstGerNome, worstGerVGV, worstGerPercent, worstGerClass, worstGerQtd
                                worstGerNome = worstGerArray(i, 0)
                                worstGerVGV = worstGerArray(i, 1)
                                worstGerPercent = worstGerArray(i, 2)
                                worstGerClass = worstGerArray(i, 3)
                                
                                ' Obter quantidade
                                worstGerQtd = 0
                                If dictWorstGerenciaQtd.Exists(worstGerNome) Then
                                    worstGerQtd = dictWorstGerenciaQtd(worstGerNome)
                                End If
                                
                                ' Acumular para média
                                worstGerTotalVGV = worstGerTotalVGV + worstGerVGV
                                worstGerAvgPercent = worstGerAvgPercent + worstGerPercent
                                worstCount = worstCount + 1
                                
                                ' Determinar classe CSS
                                Dim worstPercentClass
                                Select Case worstGerClass
                                    Case "excelente"
                                        worstPercentClass = "percent-excelente"
                                    Case "bom"
                                        worstPercentClass = "percent-bom"
                                    Case "medio"
                                        worstPercentClass = "percent-medio"
                                    Case "baixo"
                                        worstPercentClass = "percent-baixo"
                                    Case "critico"
                                        worstPercentClass = "percent-critico"
                                End Select
                                
                                ' Calcular % do total de VGV
                                Dim worstProgressWidth
                                If gerTotalVGVForWorst > 0 Then
                                    worstProgressWidth = (worstGerVGV / gerTotalVGVForWorst) * 100
                                Else
                                    worstProgressWidth = 0
                                End If
                                %>
                                <tr>
                                    <td class="ranking"><%=i+1%></td>
                                    <td><strong><%=worstGerNome%></strong></td>
                                    <td style="text-align: right;" class="vgv-value">
                                        R$ <%=FormatNumber(worstGerVGV, 0)%>
                                        <div class="progress-small">
                                            <div class="progress-bar-small" style="width: <%=worstProgressWidth%>%; background-color: #dc3545;"></div>
                                        </div>
                                    </td>
                                    <td style="text-align: center;">
                                        <span class="percent-badge <%=worstPercentClass%>">
                                            <%=FormatNumber(worstGerPercent, 1)%>%
                                        </span>
                                        <div class="progress-small">
                                            <%
                                            Dim worstProgressWidthMeta
                                            If worstGerPercent > 100 Then
                                                worstProgressWidthMeta = 100
                                            Else
                                                worstProgressWidthMeta = worstGerPercent
                                            End If
                                            %>
                                            <div class="progress-bar-small <%=worstPercentClass%>" style="width: <%=worstProgressWidthMeta%>%"></div>
                                        </div>
                                    </td>
                                    <td style="text-align: center;">
                                        <span class="badge bg-danger">
                                            <%=worstGerQtd%>
                                        </span>
                                    </td>
                                </tr>
                                <%
                            Next
                            
                            ' Calcular médias
                            If worstCount > 0 Then
                                worstGerAvgVGV = worstGerTotalVGV / worstCount
                                worstGerAvgPercent = worstGerAvgPercent / worstCount
                            Else
                                worstGerAvgVGV = 0
                                worstGerAvgPercent = 0
                            End If
                            %>
                            
                            <!-- LINHA DE MÉDIA -->
                            <tr style="background-color: #f8d7da; font-weight: bold;">
                                <td colspan="2">MÉDIA (Top <%=worstCount%>)</td>
                                <td style="text-align: right;">
                                    R$ <%=FormatNumber(worstGerAvgVGV, 0)%>
                                </td>
                                <td style="text-align: center;">
                                    <%
                                    Dim worstAvgPercentClass
                                    If worstGerAvgPercent >= 100 Then
                                        worstAvgPercentClass = "percent-excelente"
                                    ElseIf worstGerAvgPercent >= 75 Then
                                        worstAvgPercentClass = "percent-bom"
                                    ElseIf worstGerAvgPercent >= 50 Then
                                        worstAvgPercentClass = "percent-medio"
                                    ElseIf worstGerAvgPercent >= 25 Then
                                        worstAvgPercentClass = "percent-baixo"
                                    Else
                                        worstAvgPercentClass = "percent-critico"
                                    End If
                                    %>
                                    <span class="<%=worstAvgPercentClass%>" style="padding: 3px 8px; border-radius: 12px;">
                                        <%=FormatNumber(worstGerAvgPercent, 1)%>%
                                    </span>
                                </td>
                                <td style="text-align: center;">
                                    --
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                
                <!-- INDICADOR DE ALERTA -->
                <div class="alert alert-danger mt-3 p-2" role="alert">
                    <div class="d-flex align-items-center">
                        <i class="fas fa-exclamation-circle me-2"></i>
                        <small><strong>Atenção:</strong> Estas são as <%=worstCount%> gerências com menor % de atingimento da meta.</small>
                    </div>
                </div>
                
                <% Else %>
                    <p class="text-center text-muted mt-3">
                        <i class="fas fa-info-circle"></i> Nenhuma gerência com vendas no período
                    </p>
                <% End If %>
            </div>
        </div>
    </div>
</div>

            <!-- NOVA SEÇÃO: TOP 5 EMPRESAS -->
            <div class="row mb-4">
                <div class="col-md-12">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0"><i class="fas fa-building"></i> Top 5 Empresas</h5>
                        </div>
                        <div class="card-body">
                            <%
                            ' Buscar top 5 empresas pelo campo NomeEmpresa
                            Dim rsEmpresas, SQL_Empresas, contadorEmpresas
                            contadorEmpresas = 0
                            
                            SQL_Empresas = "SELECT TOP 5 NomeEmpresa, SUM(ValorUnidade) AS Total, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " AND NomeEmpresa IS NOT NULL AND NomeEmpresa <> '' GROUP BY NomeEmpresa ORDER BY SUM(ValorUnidade) DESC"
                            
                            Set rsEmpresas = Server.CreateObject("ADODB.Recordset")
                            rsEmpresas.Open SQL_Empresas, conn
                            
                            If Not rsEmpresas.EOF Then
                                Response.Write "<div class='table-container' style='max-height: 300px;'>"
                                Response.Write "<table class='vgv-table'>"
                                Response.Write "<thead>"
                                Response.Write "<tr>"
                                Response.Write "<th style='width: 30px;'>#</th>"
                                Response.Write "<th>Empresa</th>"
                                Response.Write "<th style='text-align: right;'>VGV</th>"
                                Response.Write "<th style='text-align: center;'>Qtd</th>"
                                Response.Write "</tr>"
                                Response.Write "</thead>"
                                Response.Write "<tbody>"
                                
                                Do While Not rsEmpresas.EOF And contadorEmpresas < 5
                                    contadorEmpresas = contadorEmpresas + 1
                                    Response.Write "<tr>"
                                    Response.Write "<td class='ranking'>" & contadorEmpresas & "</td>"
                                    Response.Write "<td><strong>" & rsEmpresas("NomeEmpresa") & "</strong></td>"
                                    Response.Write "<td style='text-align: right;' class='vgv-value'>R$ " & FormatNumber(rsEmpresas("Total"), 0) & "</td>"
                                    Response.Write "<td style='text-align: center;'>"
                                    Response.Write "<span class='badge bg-info'>" & rsEmpresas("Quantidade") & "</span>"
                                    Response.Write "</td>"
                                    Response.Write "</tr>"
                                    rsEmpresas.MoveNext
                                Loop
                                
                                Response.Write "</tbody>"
                                Response.Write "</table>"
                                Response.Write "</div>"
                            Else
                                Response.Write "<p class='text-center text-muted mt-3'><i class='fas fa-info-circle'></i> Nenhuma empresa com vendas no período</p>"
                            End If
                            
                            ' Fechar recordset
                            rsEmpresas.Close
                            Set rsEmpresas = Nothing
                            %>
                            
                            <!-- Nota informativa -->
                            <div class="alert alert-info mt-3 p-2" role="alert">
                                <div class="d-flex align-items-center">
                                    <i class="fas fa-info-circle me-2"></i>
                                    <small><strong>Nota:</strong> Empresas com maior valor de venda (VGV)</small>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

<!-- @@@@@@@@@@@@@@@@@@@@@@@@@@@@ -->
                        <!-- Tops (em 2 colunas) -->
            <div class="row mb-4">
                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏆 Top 10 Corretores</h5>
                        </div>
                        <div class="card-body">
                            <%
                            SQL = "SELECT TOP 10 Corretor, SUM(ValorUnidade) AS Total, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " GROUP BY Corretor ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            If Not rs.EOF Then
                                Response.Write "<div class='table-container' style='max-height: 300px;'>"
                                Response.Write "<table class='vgv-table'>"
                                Response.Write "<thead>"
                                Response.Write "<tr>"
                                Response.Write "<th style='width: 30px;'>#</th>"
                                Response.Write "<th>Corretor</th>"
                                Response.Write "<th style='text-align: right;'>VGV</th>"
                                Response.Write "<th style='text-align: center;'>Qtd</th>"
                                Response.Write "</tr>"
                                Response.Write "</thead>"
                                Response.Write "<tbody>"
                                
                                contador = 0
                                Do While Not rs.EOF And contador < 10
                                    contador = contador + 1
                                    Response.Write "<tr>"
                                    Response.Write "<td class='ranking'>" & contador & "</td>"
                                    Response.Write "<td><strong>" & rs("Corretor") & "</strong></td>"
                                    Response.Write "<td style='text-align: right;' class='vgv-value'>R$ " & FormatNumber(rs("Total"), 0) & "</td>"
                                    Response.Write "<td style='text-align: center;'>"
                                    Response.Write "<span class='badge bg-primary'>" & rs("Quantidade") & " un</span>"
                                    Response.Write "</td>"
                                    Response.Write "</tr>"
                                    rs.MoveNext
                                Loop
                                
                                Response.Write "</tbody>"
                                Response.Write "</table>"
                                Response.Write "</div>"
                            Else
                                Response.Write "<p class='text-center text-muted mt-3'><i class='fas fa-info-circle'></i> Nenhum corretor com vendas no período</p>"
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </div>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">👔 Top 10 Gerentes</h5>
                        </div>
                        <div class="card-body">
                            <%
                            SQL = "SELECT TOP 10 Gerencia, SUM(ValorUnidade) AS Total, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            If Not rs.EOF Then
                                Response.Write "<div class='table-container' style='max-height: 300px;'>"
                                Response.Write "<table class='vgv-table'>"
                                Response.Write "<thead>"
                                Response.Write "<tr>"
                                Response.Write "<th style='width: 30px;'>#</th>"
                                Response.Write "<th>Gerência</th>"
                                Response.Write "<th style='text-align: right;'>VGV</th>"
                                Response.Write "<th style='text-align: center;'>Qtd</th>"
                                Response.Write "</tr>"
                                Response.Write "</thead>"
                                Response.Write "<tbody>"
                                
                                contador = 0
                                Do While Not rs.EOF And contador < 10
                                    contador = contador + 1
                                    Response.Write "<tr>"
                                    Response.Write "<td class='ranking'>" & contador & "</td>"
                                    Response.Write "<td><strong>" & rs("Gerencia") & "</strong></td>"
                                    Response.Write "<td style='text-align: right;' class='vgv-value'>R$ " & FormatNumber(rs("Total"), 0) & "</td>"
                                    Response.Write "<td style='text-align: center;'>"
                                    Response.Write "<span class='badge bg-success'>" & rs("Quantidade") & " un</span>"
                                    Response.Write "</td>"
                                    Response.Write "</tr>"
                                    rs.MoveNext
                                Loop
                                
                                Response.Write "</tbody>"
                                Response.Write "</table>"
                                Response.Write "</div>"
                            Else
                                Response.Write "<p class='text-center text-muted mt-3'><i class='fas fa-info-circle'></i> Nenhuma gerência com vendas no período</p>"
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </div>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏢 Top 5 Diretorias</h5>
                        </div>
                        <div class="card-body">
                            <%
                            SQL = "SELECT TOP 5 Diretoria, SUM(ValorUnidade) AS Total, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            If Not rs.EOF Then
                                Response.Write "<div class='table-container' style='max-height: 300px;'>"
                                Response.Write "<table class='vgv-table'>"
                                Response.Write "<thead>"
                                Response.Write "<tr>"
                                Response.Write "<th style='width: 30px;'>#</th>"
                                Response.Write "<th>Diretoria</th>"
                                Response.Write "<th style='text-align: right;'>VGV</th>"
                                Response.Write "<th style='text-align: center;'>Qtd</th>"
                                Response.Write "</tr>"
                                Response.Write "</thead>"
                                Response.Write "<tbody>"
                                
                                contador = 0
                                Do While Not rs.EOF And contador < 5
                                    contador = contador + 1
                                    Response.Write "<tr>"
                                    Response.Write "<td class='ranking'>" & contador & "</td>"
                                    Response.Write "<td><strong>" & rs("Diretoria") & "</strong></td>"
                                    Response.Write "<td style='text-align: right;' class='vgv-value'>R$ " & FormatNumber(rs("Total"), 0) & "</td>"
                                    Response.Write "<td style='text-align: center;'>"
                                    Response.Write "<span class='badge bg-info'>" & rs("Quantidade") & " un</span>"
                                    Response.Write "</td>"
                                    Response.Write "</tr>"
                                    rs.MoveNext
                                Loop
                                
                                Response.Write "</tbody>"
                                Response.Write "</table>"
                                Response.Write "</div>"
                            Else
                                Response.Write "<p class='text-center text-muted mt-3'><i class='fas fa-info-circle'></i> Nenhuma diretoria com vendas no período</p>"
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </div>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏗️ Top 5 Empreendimentos</h5>
                        </div>
                        <div class="card-body">
                            <%
                            SQL = "SELECT TOP 5 NomeEmpreendimento, SUM(ValorUnidade) AS Total, COUNT(*) AS Quantidade FROM Vendas " & whereClause & " GROUP BY NomeEmpreendimento ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            If Not rs.EOF Then
                                Response.Write "<div class='table-container' style='max-height: 300px;'>"
                                Response.Write "<table class='vgv-table'>"
                                Response.Write "<thead>"
                                Response.Write "<tr>"
                                Response.Write "<th style='width: 30px;'>#</th>"
                                Response.Write "<th>Empreendimento</th>"
                                Response.Write "<th style='text-align: right;'>VGV</th>"
                                Response.Write "<th style='text-align: center;'>Qtd</th>"
                                Response.Write "</tr>"
                                Response.Write "</thead>"
                                Response.Write "<tbody>"
                                
                                contador = 0
                                Do While Not rs.EOF And contador < 5
                                    contador = contador + 1
                                    Response.Write "<tr>"
                                    Response.Write "<td class='ranking'>" & contador & "</td>"
                                    Response.Write "<td><strong>" & rs("NomeEmpreendimento") & "</strong></td>"
                                    Response.Write "<td style='text-align: right;' class='vgv-value'>R$ " & FormatNumber(rs("Total"), 0) & "</td>"
                                    Response.Write "<td style='text-align: center;'>"
                                    Response.Write "<span class='badge bg-warning'>" & rs("Quantidade") & " un</span>"
                                    Response.Write "</td>"
                                    Response.Write "</tr>"
                                    rs.MoveNext
                                Loop
                                
                                Response.Write "</tbody>"
                                Response.Write "</table>"
                                Response.Write "</div>"
                            Else
                                Response.Write "<p class='text-center text-muted mt-3'><i class='fas fa-info-circle'></i> Nenhum empreendimento com vendas no período</p>"
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </div>
                    </div>
                </div>
            </div>

<!-- @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ -->

            <!-- Gráficos -->
            <div class="row">
                <div class="col-md-12 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">📈 Gráfico de Vendas por Ano-Mês</h5>
                        </div>
                        <div class="card-body grafico-container">
                            <canvas id="graficoVendas"></canvas>
                        </div>
                    </div>
                </div>

                <div class="col-md-12 mb-4">
                    <div class="card">
                        <div class="card-header text-white" style="background: linear-gradient(135deg, #7209b7 0%, #4361ee 100%);">
                            <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Quantidades Vendidas por Ano-Mês</h5>
                        </div>
                        <div class="card-body grafico-container">
                            <canvas id="graficoQuantidades"></canvas>
                        </div>
                    </div>
                </div>
            </div>



            </div>
        </div>
    </div>
    
    <!-- Countdown Timer -->
    <div id="countdown-timer">Atualizando em: <span id="seconds-left">0</span>s</div>
</div>

<!-- Scripts -->
<script>
// =======================================================
// FUNÇÕES GERAIS
// =======================================================

/**
 * Formata valores monetários
 * @param {number} valor - Valor a ser formatado
 * @returns {string} Valor formatado
 */
function formatarMoeda(valor) {
    if (valor >= 1000000) {
        return 'R$ ' + (valor / 1000000).toFixed(1).replace('.', ',') + 'M';
    } else if (valor >= 1000) {
        return 'R$ ' + Math.round(valor / 1000) + 'K';
    }
    return 'R$ ' + Math.round(valor).toLocaleString('pt-BR');
}

/**
 * Formata números com separadores
 * @param {number} valor - Valor a ser formatado
 * @returns {string} Valor formatado
 */
function formatarNumero(valor) {
    return valor.toLocaleString('pt-BR');
}

// =======================================================
// GERENCIAMENTO DO TIMER AUTOMÁTICO
// =======================================================

const filterNames = ['ano', 'mes', 'diretoria', 'gerencia', 'corretor', 'empreendimento'];
let timerInterval;

const urlParams = new URLSearchParams(window.location.search);
const timerDuration = parseInt(urlParams.get('autotime')) || 10;

const autoModeBtn = document.getElementById('autoModeBtn');
const loadingSpinner = document.getElementById('loadingSpinner');
const countdownTimer = document.getElementById('countdown-timer');
const secondsLeftSpan = document.getElementById('seconds-left');

/**
 * Inicia o timer de atualização automática
 */
function startTimer() {
    let secondsLeft = timerDuration;
    secondsLeftSpan.textContent = secondsLeft;
    countdownTimer.style.display = 'block';

    timerInterval = setInterval(() => {
        secondsLeft--;
        secondsLeftSpan.textContent = secondsLeft;
        if (secondsLeft <= 0) {
            clearInterval(timerInterval);
            const nextState = getNextFilterState();
            window.location.href = window.location.pathname + '?' + nextState;
        }
    }, 1000);
}

/**
 * Para o timer de atualização automática
 */
function stopTimer() {
    clearInterval(timerInterval);
    countdownTimer.style.display = 'none';
}

/**
 * Obtém o próximo estado do filtro para o modo automático
 * @returns {string} Query string com os próximos parâmetros
 */
function getNextFilterState() {
    const currentParams = new URLSearchParams(window.location.search);
    
    let filterName = currentParams.get('auto_filter') || filterNames[0];
    let filterIndex = filterNames.indexOf(filterName);

    let selectElement = document.getElementById(filterName + 'Filter');
    let currentOptionIndex = selectElement.selectedIndex;
    let nextOptionIndex = (currentOptionIndex + 1) % selectElement.options.length;
    
    let nextFilterName = filterName;

    if (nextOptionIndex === 0) {
        filterIndex = (filterIndex + 1) % filterNames.length;
        nextFilterName = filterNames[filterIndex];
        selectElement = document.getElementById(nextFilterName + 'Filter');
        nextOptionIndex = 0;
    }

    const nextParams = new URLSearchParams();
    nextParams.set('auto_mode', 'on');
    nextParams.set('auto_filter', nextFilterName);
    nextParams.set('autotime', timerDuration);
    nextParams.set(nextFilterName, selectElement.options[nextOptionIndex].value);

    return nextParams.toString();
}

// Verificar se o modo automático está ativo
const isAutoModeActive = urlParams.get('auto_mode') === 'on';

// Configurar botão do modo automático
if (isAutoModeActive) {
    autoModeBtn.innerHTML = '<i class="fas fa-pause-circle"></i> Parar Modo Autônomo';
    autoModeBtn.classList.remove('btn-info');
    autoModeBtn.classList.add('btn-danger');
    startTimer();
}

// Event listener para o botão de modo automático
autoModeBtn.addEventListener('click', function() {
    if (isAutoModeActive) {
        const currentParams = new URLSearchParams(window.location.search);
        currentParams.delete('auto_mode');
        currentParams.delete('auto_filter');
        currentParams.delete('autotime');
        window.location.href = window.location.pathname + '?' + currentParams.toString();
    } else {
        const currentParams = new URLSearchParams(window.location.search);
        const selectedTime = document.getElementById('autoTimeFilter').value;
        currentParams.set('auto_mode', 'on');
        currentParams.set('autotime', selectedTime);
        window.location.href = window.location.pathname + '?' + currentParams.toString();
    }
});

// Event listener para o formulário de filtros
document.getElementById('filterForm').addEventListener('submit', function() {
    document.getElementById('loadingSpinner').style.display = 'flex';
});
</script>

<%
' =======================================================
' DADOS PARA OS GRÁFICOS
' =======================================================

Dim datasetsJSON, colors(5), colorIndex, ano, SQL_Anos, rsAnos, SQL_Dados, rsDados
Dim dadosAno, dadosPorMes(12), mesAtual, i, vTotal

datasetsJSON = ""
colorIndex = 0

colors(0) = "rgba(255, 99, 132, 1)"
colors(1) = "rgba(54, 162, 235, 1)"
colors(2) = "rgba(255, 206, 86, 1)"
colors(3) = "rgba(75, 192, 192, 1)"
colors(4) = "rgba(153, 102, 255, 1)"
colors(5) = "rgba(255, 159, 64, 1)"

SQL_Anos = "SELECT DISTINCT AnoVenda FROM Vendas " & whereClause & " ORDER BY AnoVenda"
Set rsAnos = Server.CreateObject("ADODB.Recordset")
rsAnos.Open SQL_Anos, conn

Do Until rsAnos.EOF
    ano = rsAnos("AnoVenda")
    SQL_Dados = "SELECT MesVenda, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & ano & " GROUP BY MesVenda ORDER BY MesVenda"
    Set rsDados = Server.CreateObject("ADODB.Recordset")
    rsDados.Open SQL_Dados, conn

    For i = 1 to 12
        dadosPorMes(i) = "0"
    Next

    Do Until rsDados.EOF
        mesAtual = CInt(rsDados("MesVenda"))
        If Not IsNull(rsDados("Total")) Then
            vTotal = Replace(rsDados("Total"),",",".")
            dadosPorMes(mesAtual) = vTotal
        End If
        rsDados.MoveNext
    Loop
    rsDados.Close
    Set rsDados = Nothing

    dadosAno = ""
    For i = 1 to 12
        dadosAno = dadosAno & dadosPorMes(i) & ","
    Next
    If Right(dadosAno, 1) = "," Then dadosAno = Left(dadosAno, Len(dadosAno) - 1)

    datasetsJSON = datasetsJSON & "{ "
    datasetsJSON = datasetsJSON & "label: 'Vendas " & ano & "', "
    datasetsJSON = datasetsJSON & "data: [" & dadosAno & "], "
    datasetsJSON = datasetsJSON & "borderColor: '" & colors(colorIndex Mod 6) & "', "
    datasetsJSON = datasetsJSON & "backgroundColor: '" & Replace(colors(colorIndex Mod 6), "1)", "0.7)") & "', "
    datasetsJSON = datasetsJSON & "borderWidth: 2, "
    datasetsJSON = datasetsJSON & "borderRadius: 4, "
    datasetsJSON = datasetsJSON & "fill: false, "
    datasetsJSON = datasetsJSON & "tension: 0.3 "
    datasetsJSON = datasetsJSON & "},"

    colorIndex = colorIndex + 1
    rsAnos.MoveNext
Loop

If Right(datasetsJSON, 1) = "," Then datasetsJSON = Left(datasetsJSON, Len(datasetsJSON) - 1)

If Not rsAnos Is Nothing Then
    If Not rsAnos.EOF Then rsAnos.Close
    Set rsAnos = Nothing
End If

conn.Close
Set conn = Nothing
%>

<script>
// =======================================================
// SCRIPT PARA O GRÁFICO DE VENDAS (VALORES) COM LABELS VERTICAIS
// =======================================================

// Aguardar o DOM estar completamente carregado
document.addEventListener('DOMContentLoaded', function() {
    inicializarGraficoVendas();
});

/**
 * Inicializa o gráfico de vendas (valores monetários) com labels verticais
 */
function inicializarGraficoVendas() {
    const ctx = document.getElementById('graficoVendas').getContext('2d');
    
    if (!ctx) {
        console.error('Elemento graficoVendas não encontrado');
        return;
    }
    
    try {
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                datasets: [<%=datasetsJSON%>]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleFont: {
                            size: 16,
                            weight: 'bold'
                        },
                        bodyFont: {
                            size: 14
                        },
                        padding: 12,
                        displayColors: true,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += new Intl.NumberFormat('pt-BR', {
                                        style: 'currency',
                                        currency: 'BRL',
                                        minimumFractionDigits: 0
                                    }).format(context.parsed.y);
                                }
                                return label;
                            }
                        }
                    },
                    // CONFIGURAÇÃO PARA LABELS VERTICAIS (ROTACIONADOS)
                    datalabels: {
                        display: function(context) {
                            // Mostrar apenas valores maiores que 0
                            return context.dataset.data[context.dataIndex] > 0;
                        },
                        color: '#FFFFFF',
                        font: {
                            size: 11,
                            weight: 'bold',
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        formatter: function(value) {
                            if (value >= 1000000) {
                                return 'R$ ' + (value / 1000000).toFixed(1).replace('.', ',') + 'M';
                            } else if (value >= 1000) {
                                return 'R$ ' + Math.round(value / 1000) + 'K';
                            }
                            return 'R$ ' + Math.round(value).toLocaleString('pt-BR');
                        },
                        anchor: 'center',      // Âncora no final da barra
                        align: 'center',       // Alinhamento no topo
                        clamp: true,
                        rotation: -90,      // ROTAÇÃO DE -90 GRAUS PARA FICAR VERTICAL
                        padding: {
                            top: 4,
                            right: 4,
                            bottom: 4,
                            left: 4
                        },
                        backgroundColor: 'rgba(0, 0, 0, 1.6)',
                        borderColor: 'rgba(255, 255, 255, 0.3)',
                        borderWidth: 1,
                        borderRadius: 4,
                        textAlign: 'center',
                        offset: 8           // Offset para afastar um pouco da barra
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                size: 12,
                                weight: 'bold'
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        },
                        ticks: {
                            font: {
                                size: 12,
                                weight: 'bold'
                            },
                            callback: function(value) {
                                return new Intl.NumberFormat('pt-BR', {
                                    style: 'currency',
                                    currency: 'BRL',
                                    minimumFractionDigits: 0
                                }).format(value);
                            }
                        },
                        // Aumentar o espaço para os labels verticais
                        afterFit: function(scaleInstance) {
                            scaleInstance.width = 80; // Aumentar largura do eixo Y
                        }
                    }
                },
                // Aumentar o espaço entre as barras
                layout: {
                    padding: {
                        left: 20,
                        right: 20,
                        top: 40,
                        bottom: 20
                    }
                }
            },
            plugins: [ChartDataLabels]
        });
        
        // CONFIGURAÇÃO PARA AUMENTAR A ALTURA DAS BARRAS
        if (chart.data.datasets[0]) {
            // Aumentar a espessura das barras significativamente
            chart.data.datasets[0].barThickness = 50;      // Aumentado de 45 para 50
            chart.data.datasets[0].maxBarThickness = 80;   // Aumentado de 60 para 80
            chart.data.datasets[0].borderRadius = 8;
            
            // Se houver mais datasets, configurar também
            for (let i = 1; i < chart.data.datasets.length; i++) {
                chart.data.datasets[i].barThickness = 50;
                chart.data.datasets[i].maxBarThickness = 80;
                chart.data.datasets[i].borderRadius = 8;
            }
        }
        
        // Ajustar a escala do gráfico para dar mais altura às barras
        chart.options.scales.y.ticks.maxTicksLimit = 8; // Reduzir número de ticks para dar mais espaço
        
        chart.update();
        
        console.log('Gráfico de vendas inicializado com sucesso - Labels Verticais');
        console.log('Número de datasets:', chart.data.datasets.length);
        console.log('Configuração de barras:', {
            barThickness: chart.data.datasets[0]?.barThickness,
            maxBarThickness: chart.data.datasets[0]?.maxBarThickness
        });
        
    } catch (error) {
        console.error('Erro ao inicializar gráfico de vendas:', error);
    }
}
</script>

<script>
// =======================================================
// SCRIPT PARA O GRÁFICO DE QUANTIDADES VENDIDAS COM LABELS HORIZONTAIS
// =======================================================

// Aguardar o DOM estar completamente carregado
document.addEventListener('DOMContentLoaded', function() {
    inicializarGraficoQuantidades();
});

/**
 * Inicializa o gráfico de quantidades vendidas com labels horizontais
 */
function inicializarGraficoQuantidades() {
    const ctx = document.getElementById('graficoQuantidades').getContext('2d');
    
    if (!ctx) {
        console.error('Elemento graficoQuantidades não encontrado');
        return;
    }
    
    try {
        const chartQuantidades = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                datasets: [<%=datasetsJSONQuantidades%>]
            },
            options: {
                animation: {
                    duration: 1000,
                    easing: 'easeInOutQuad'
                },
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleFont: {
                            size: 16,
                            weight: 'bold'
                        },
                        bodyFont: {
                            size: 14
                        },
                        padding: 12,
                        displayColors: true,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += context.parsed.y.toLocaleString('pt-BR') + ' unidades';
                                }
                                return label;
                            }
                        }
                    },
                    // CONFIGURAÇÃO PARA LABELS HORIZONTAIS
                    datalabels: {
                        display: function(context) {
                            // Mostrar apenas valores maiores que 0
                            return context.dataset.data[context.dataIndex] > 0;
                        },
                        color: '#FFFFFF',
                        font: {
                            size: 11,
                            weight: 'bold',
                            family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
                        },
                        formatter: function(value) {
                            return value.toLocaleString('pt-BR');
                        },
                        anchor: 'center',
                        align: 'center',
                        clamp: true,
                        rotation: 0,           // LABELS HORIZONTAIS
                        padding: {
                            top: 4,
                            right: 4,
                            bottom: 4,
                            left: 4
                        },
                        backgroundColor: 'rgba(0, 0, 0, 0.6)',
                        borderColor: 'rgba(255, 255, 255, 0.3)',
                        borderWidth: 1,
                        borderRadius: 4,
                        textAlign: 'center',
                        offset: 0
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                size: 12,
                                weight: 'bold'
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        },
                        ticks: {
                            font: {
                                size: 12,
                                weight: 'bold'
                            },
                            callback: function(value) {
                                return value.toLocaleString('pt-BR') + ' un';
                            }
                        }
                    }
                },
                // Aumentar o espaço entre as barras
                layout: {
                    padding: {
                        left: 20,
                        right: 20,
                        top: 40,
                        bottom: 20
                    }
                }
            },
            plugins: [ChartDataLabels]
        });
        
        // CONFIGURAÇÃO PARA AUMENTAR A ALTURA DAS BARRAS (IGUAL AO PRIMEIRO GRÁFICO)
        if (chartQuantidades.data.datasets[0]) {
            // Aumentar a espessura das barras significativamente
            chartQuantidades.data.datasets[0].barThickness = 50;      // Aumentado
            chartQuantidades.data.datasets[0].maxBarThickness = 80;   // Aumentado
            chartQuantidades.data.datasets[0].borderRadius = 8;
            
            // Se houver mais datasets, configurar também
            for (let i = 1; i < chartQuantidades.data.datasets.length; i++) {
                chartQuantidades.data.datasets[i].barThickness = 50;
                chartQuantidades.data.datasets[i].maxBarThickness = 80;
                chartQuantidades.data.datasets[i].borderRadius = 8;
            }
        }
        
        chartQuantidades.update();
        
        console.log('Gráfico de quantidades inicializado com sucesso - Labels Horizontais');
        console.log('Número de datasets:', chartQuantidades.data.datasets.length);
        console.log('Configuração de barras:', {
            barThickness: chartQuantidades.data.datasets[0]?.barThickness,
            maxBarThickness: chartQuantidades.data.datasets[0]?.maxBarThickness
        });
        
    } catch (error) {
        console.error('Erro ao inicializar gráfico de quantidades:', error);
    }
}
</script>

</body>
</html>