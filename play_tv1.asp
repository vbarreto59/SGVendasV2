<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conSunSales.asp"-->

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
    Set rs = Nothing
    
    GetUniqueValues = dict.Keys
End Function

' FUNÇÃO PARA CONSTRUIR A CLÁUSULA WHERE
Function BuildWhereClause()
    Dim sqlWhere
    sqlWhere = " WHERE 1=1 AND Excluido = 0 AND Excluido IS NOT NULL"

    If Request.QueryString("ano") <> "" Then
        sqlWhere = sqlWhere & " AND AnoVenda = " & Request.QueryString("ano")
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

' FUNÇÃO PARA SERIALIZAR ARRAY EM JSON
Function JSON_Serialize(arr)
    Dim i, result
    result = "["
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            result = result & arr(i)
        Else
            result = result & """" & Replace(arr(i), """", "\""") & """"
        End If
        If i < UBound(arr) Then result = result & ","
    Next
    result = result & "]"
    JSON_Serialize = result
End Function

' =======================================================
' INÍCIO DO PROCESSAMENTO
' =======================================================

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConnSales

Dim whereClause
whereClause = BuildWhereClause()

' Dados do mês atual
Dim anoAtual, mesAtual
anoAtual = Year(Date())
mesAtual = Month(Date())

' CALCULAR TICKET MÉDIO E QUANTIDADE DE UNIDADES
Dim ticketMedio, quantidadeUnidades, totalVendas, ticketMedioAno, quantidadeUnidadesAno, totalVendasAno
quantidadeUnidades = 0
totalVendas = 0

' Dados do mês atual
SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAtual
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
Set rs = Nothing

If quantidadeUnidades > 0 And totalVendas > 0 Then
    ticketMedio = totalVendas / quantidadeUnidades
Else
    ticketMedio = 0
End If

' Dados do ano atual
SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalUnidades")) Then
        quantidadeUnidadesAno = rs("TotalUnidades")
    End If
    If Not IsNull(rs("TotalVendas")) Then
        totalVendasAno = rs("TotalVendas")
    End If
End If
rs.Close
Set rs = Nothing

If quantidadeUnidadesAno > 0 And totalVendasAno > 0 Then
    ticketMedioAno = totalVendasAno / quantidadeUnidadesAno
Else
    ticketMedioAno = 0
End If

' Calcular variação em relação ao mês anterior
Dim mesAnterior, totalVendasMesAnterior, variacaoMensal
mesAnterior = mesAtual - 1
If mesAnterior = 0 Then
    mesAnterior = 12
End If

SQL = "SELECT SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAnterior
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalVendas")) Then
        totalVendasMesAnterior = rs("TotalVendas")
    Else
        totalVendasMesAnterior = 0
    End If
Else
    totalVendasMesAnterior = 0
End If
rs.Close
Set rs = Nothing

If totalVendasMesAnterior > 0 Then
    variacaoMensal = ((totalVendas - totalVendasMesAnterior) / totalVendasMesAnterior) * 100
Else
    variacaoMensal = 100
End If

' Calcular performance vs meta (exemplo com meta fictícia de 1.000.000)
Dim metaAnual, performanceAnual
metaAnual = 1000000 ' Valor de exemplo
performanceAnual = (totalVendasAno / metaAnual) * 100

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

Dim autoTime
autoTime = Request.QueryString("autotime")
If autoTime = "" Then autoTime = 10

' Preparar dados para os gráficos

' Gráfico de vendas anual por mês
Dim dadosVendasAnual(12), mesesAno(12)
For i = 1 to 12
    mesesAno(i-1) = Left(arrMesesNome(i), 3)
    
    SQL = "SELECT SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & i
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn
    
    If Not rs.EOF Then
        If Not IsNull(rs("Total")) Then
            dadosVendasAnual(i-1) = rs("Total")
        Else
            dadosVendasAnual(i-1) = 0
        End If
    Else
        dadosVendasAnual(i-1) = 0
    End If
    rs.Close
    Set rs = Nothing
Next

' Gráfico de diretorias
Dim diretorias, totaisDiretorias
SQL = "SELECT Diretoria, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

diretorias = ""
totaisDiretorias = ""
Do Until rs.EOF
    diretorias = diretorias & """" & Replace(rs("Diretoria"), """", "\""") & ""","
    totaisDiretorias = totaisDiretorias & rs("Total") & ","
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Right(diretorias, 1) = "," Then diretorias = Left(diretorias, Len(diretorias) - 1)
If Right(totaisDiretorias, 1) = "," Then totaisDiretorias = Left(totaisDiretorias, Len(totaisDiretorias) - 1)

' Gráfico de gerências (top 10)
Dim gerencias, totaisGerencias
SQL = "SELECT TOP 10 Gerencia, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

gerencias = ""
totaisGerencias = ""
Do Until rs.EOF
    gerencias = gerencias & """" & Replace(rs("Gerencia"), """", "\""") & ""","
    totaisGerencias = totaisGerencias & rs("Total") & ","
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Right(gerencias, 1) = "," Then gerencias = Left(gerencias, Len(gerencias) - 1)
If Right(totaisGerencias, 1) = "," Then totaisGerencias = Left(totaisGerencias, Len(totaisGerencias) - 1)

' Gráfico de ticket médio mensal
Dim ticketMedioMensal(12)
For i = 1 to 12
    SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & i
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn
    
    If Not rs.EOF Then
        If Not IsNull(rs("TotalUnidades")) And Not IsNull(rs("TotalVendas")) Then
            If rs("TotalUnidades") > 0 And rs("TotalVendas") > 0 Then
                ticketMedioMensal(i-1) = rs("TotalVendas") / rs("TotalUnidades")
            Else
                ticketMedioMensal(i-1) = 0
            End If
        Else
            ticketMedioMensal(i-1) = 0
        End If
    Else
        ticketMedioMensal(i-1) = 0
    End If
    rs.Close
    Set rs = Nothing
Next

' Gráfico de tipo de unidade (exemplo com dados fictícios)
Dim tiposUnidade, vendasTiposUnidade
tiposUnidade = "'Apartamento','Casa','Sobrado','Terreno','Comercial'"
vendasTiposUnidade = "450000,320000,280000,150000,80000" ' Valores de exemplo

' NÃO FECHAR A CONEXÃO AQUI - vamos mantê-la aberta para as consultas dentro dos slides
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard de Vendas - Sala de Vendas</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            background-color: #0a1929;
            color: #ffffff;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            overflow-x: hidden;
        }
        .dashboard-container {
            padding: 20px;
        }
        .header {
            text-align: center;
            margin-bottom: 20px;
            padding: 10px;
            background: linear-gradient(135deg, #1a3a5f 0%, #0a1929 100%);
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }
        .header h1 {
            color: #ffffff;
            font-weight: 700;
            margin: 0;
            font-size: 2.5rem;
        }
        .header h2 {
            color: #a0c4ff;
            font-weight: 400;
            margin: 0;
            font-size: 1.5rem;
        }
        .card {
            background-color: #1a3a5f;
            border: none;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            margin-bottom: 20px;
            color: #ffffff;
            transition: transform 0.3s ease;
        }
        .card:hover {
            transform: translateY(-5px);
        }
        .card-header {
            background: linear-gradient(135deg, #2a4a7a 0%, #1a3a5f 100%);
            color: white;
            border-radius: 10px 10px 0 0 !important;
            font-weight: 600;
            padding: 15px 20px;
            border-bottom: 1px solid #2a4a7a;
        }
        .metric-card {
            text-align: center;
            padding: 25px 15px;
            height: 100%;
        }
        .metric-value {
            font-size: 2.5rem;
            font-weight: bold;
            margin: 10px 0;
            color: #a0c4ff;
        }
        .metric-label {
            font-size: 1rem;
            color: #c0d6ff;
            margin-bottom: 0;
        }
        .metric-icon {
            font-size: 2.5rem;
            margin-bottom: 15px;
            color: #a0c4ff;
        }
        .list-group-item {
            background-color: #1a3a5f;
            border: 1px solid #2a4a7a;
            color: #ffffff;
            padding: 15px 20px;
        }
        .badge {
            font-weight: 600;
            padding: 8px 12px;
            border-radius: 10px;
        }
        .controls {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 1000;
            background-color: rgba(26, 58, 95, 0.9);
            border-radius: 10px;
            padding: 15px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }
        .slide {
            display: none;
        }
        .slide.active {
            display: block;
        }
        .comparison-chart {
            height: 300px;
        }
        .venda-item {
            border-left: 4px solid #a0c4ff;
            padding-left: 15px;
            margin-bottom: 10px;
        }
        .venda-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .venda-details {
            font-size: 0.9rem;
            color: #c0d6ff;
        }
        .chart-container {
            position: relative;
            height: 100%;
            min-height: 300px;
        }
        .progress {
            height: 10px;
            margin-bottom: 10px;
        }
        .progress-bar {
            background-color: #a0c4ff;
        }
        .countdown-timer {
            position: fixed;
            top: 20px;
            left: 20px;
            background-color: rgba(26, 58, 95, 0.9);
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            font-size: 1rem;
            font-weight: bold;
            z-index: 999;
            display: none;
        }
        .info-box {
            background-color: #1a3a5f;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 15px;
            border-left: 4px solid #a0c4ff;
        }
        .info-title {
            font-size: 1rem;
            color: #c0d6ff;
            margin-bottom: 5px;
        }
        .info-value {
            font-size: 1.5rem;
            font-weight: bold;
            color: #ffffff;
        }
        .trend-up {
            color: #4ade80;
        }
        .trend-down {
            color: #f87171;
        }
        .slide-indicator {
            position: fixed;
            bottom: 20px;
            left: 20px;
            background-color: rgba(26, 58, 95, 0.9);
            border-radius: 10px;
            padding: 10px 15px;
            z-index: 999;
        }
        .slide-dot {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background-color: rgba(255, 255, 255, 0.3);
            margin: 0 5px;
            cursor: pointer;
        }
        .slide-dot.active {
            background-color: #a0c4ff;
        }
    </style>
</head>
<body>

<div class="countdown-timer" id="countdown-timer">
    <i class="fas fa-clock"></i> Próxima visualização em: <span id="seconds-left">0</span>s
</div>

<div class="slide-indicator" id="slide-indicator">
    <span class="slide-dot active" data-slide="1"></span>
    <span class="slide-dot" data-slide="2"></span>
    <span class="slide-dot" data-slide="3"></span>
    <span class="slide-dot" data-slide="4"></span>
    <span class="slide-dot" data-slide="5"></span>
</div>

<div class="dashboard-container">
    <div class="header">
        <h1>Dashboard de Vendas - Sala de Vendas</h1>
        <h2><%=arrMesesNome(mesAtual)%> de <%=anoAtual%></h2>
    </div>

    <!-- Slide 1: Visão Geral do Mês -->
    <div class="slide active" id="slide1">
        <div class="row">
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-money-bill-wave metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(totalVendas, 2)%></div>
                    <p class="metric-label">Vendas do Mês</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-cube metric-icon"></i>
                    <div class="metric-value"><%=FormatNumber(quantidadeUnidades, 0)%></div>
                    <p class="metric-label">Unidades Vendidas</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-ticket-alt metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(ticketMedio, 2)%></div>
                    <p class="metric-label">Ticket Médio Mensal</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-chart-line metric-icon"></i>
                    <div class="metric-value">
                        <% If variacaoMensal >= 0 Then %>
                            <span class="trend-up">+<%=FormatNumber(variacaoMensal, 1)%>%</span>
                        <% Else %>
                            <span class="trend-down"><%=FormatNumber(variacaoMensal, 1)%>%</span>
                        <% End If %>
                    </div>
                    <p class="metric-label">Variação vs Mês Anterior</p>
                </div>
            </div>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Vendas do Ano por Mês</h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoVendasAnual"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-trophy"></i> Top 5 Corretores do Mês</h5>
                    </div>
                    <div class="card-body">
                        <%
                        ' Usar a conexão que já está aberta
                        SQL = "SELECT TOP 5 Corretor, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAtual & " GROUP BY Corretor ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide1 = Server.CreateObject("ADODB.Recordset")
                        rsSlide1.Open SQL, conn
                        
                        Do Until rsSlide1.EOF
                            Response.Write "<div class='info-box'>"
                            Response.Write "<div class='info-title'>" & rsSlide1("Corretor") & "</div>"
                            Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide1("Total"), 2) & "</div>"
                            Response.Write "</div>"
                            rsSlide1.MoveNext
                        Loop
                        rsSlide1.Close
                        Set rsSlide1 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 2: Comparativo de Diretorias -->
    <div class="slide" id="slide2">
        <div class="row">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-pie"></i> Comparativo de Diretorias - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoDiretorias"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-list-ol"></i> Ranking de Diretorias</h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT Diretoria, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide2 = Server.CreateObject("ADODB.Recordset")
                        rsSlide2.Open SQL, conn
                        
                        contador = 1
                        Do Until rsSlide2.EOF
                            Response.Write "<div class='info-box'>"
                            Response.Write "<div class='info-title'>#" & contador & " " & rsSlide2("Diretoria") & "</div>"
                            Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide2("Total"), 2) & "</div>"
                            Response.Write "</div>"
                            contador = contador + 1
                            rsSlide2.MoveNext
                        Loop
                        rsSlide2.Close
                        Set rsSlide2 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 3: Comparativo de Gerências -->
    <div class="slide" id="slide3">
        <div class="row">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Comparativo de Gerências - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoGerencias"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-list-ol"></i> Top 10 Gerências</h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 10 Gerencia, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide3 = Server.CreateObject("ADODB.Recordset")
                        rsSlide3.Open SQL, conn
                        
                        contador = 1
                        Do Until rsSlide3.EOF
                            Response.Write "<div class='info-box'>"
                            Response.Write "<div class='info-title'>#" & contador & " " & rsSlide3("Gerencia") & "</div>"
                            Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide3("Total"), 2) & "</div>"
                            Response.Write "</div>"
                            contador = contador + 1
                            rsSlide3.MoveNext
                        Loop
                        rsSlide3.Close
                        Set rsSlide3 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 4: Ticket Médio e Eficiência -->
    <div class="slide" id="slide4">
        <div class="row">
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-ticket-alt metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(ticketMedioAno, 2)%></div>
                    <p class="metric-label">Ticket Médio Anual</p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-money-bill-wave metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(totalVendasAno, 2)%></div>
                    <p class="metric-label">Vendas do Ano</p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-cube metric-icon"></i>
                    <div class="metric-value"><%=FormatNumber(quantidadeUnidadesAno, 0)%></div>
                    <p class="metric-label">Unidades Vendidas no Ano</p>
                </div>
            </div>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-line"></i> Evolução do Ticket Médio Mensal</h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoTicketMedio"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-building"></i> Top 5 Empreendimentos</h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 5 NomeEmpreendimento, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY NomeEmpreendimento ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide4 = Server.CreateObject("ADODB.Recordset")
                        rsSlide4.Open SQL, conn
                        
                        Do Until rsSlide4.EOF
                            Response.Write "<div class='info-box'>"
                            Response.Write "<div class='info-title'>" & rsSlide4("NomeEmpreendimento") & "</div>"
                            Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide4("Total"), 2) & "</div>"
                            Response.Write "</div>"
                            rsSlide4.MoveNext
                        Loop
                        rsSlide4.Close
                        Set rsSlide4 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 5: Últimas Vendas e Performance -->
    <div class="slide" id="slide5">
        <div class="row">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-clock"></i> Últimas 10 Vendas</h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 10 Corretor, ValorUnidade, NomeEmpreendimento, Gerencia, DiaVenda, MesVenda, AnoVenda FROM Vendas " & whereClause & " ORDER BY AnoVenda DESC, MesVenda DESC, DiaVenda DESC, ID DESC"
                        Set rsSlide5 = Server.CreateObject("ADODB.Recordset")
                        rsSlide5.Open SQL, conn
                        
                        If Not rsSlide5.EOF Then
                            Do While Not rsSlide5.EOF
                                %>
                                <div class="venda-item">
                                    <div class="venda-info">
                                        <strong><%=rsSlide5("Corretor")%></strong>
                                        <span class="badge bg-primary">R$ <%=FormatNumber(rsSlide5("ValorUnidade"), 2)%></span>
                                    </div>
                                    <div class="venda-details">
                                        <small>
                                            <i class="fas fa-building"></i> <%=rsSlide5("NomeEmpreendimento")%> | 
                                            <i class="fas fa-user-tie"></i> <%=rsSlide5("Gerencia")%> | 
                                            <i class="fas fa-calendar"></i> <%=rsSlide5("DiaVenda")%>/<%=rsSlide5("MesVenda")%>/<%=rsSlide5("AnoVenda")%>
                                        </small>
                                    </div>
                                </div>
                                <%
                                rsSlide5.MoveNext
                            Loop
                        Else
                            Response.Write "<p class='text-center text-muted'>Nenhuma venda encontrada</p>"
                        End If
                        
                        rsSlide5.Close
                        Set rsSlide5 = Nothing
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-tachometer-alt"></i> Metas e Performance</h5>
                    </div>
                    <div class="card-body">
                        <div class="info-box">
                            <div class="info-title">Meta Anual</div>
                            <div class="info-value">R$ <%=FormatNumber(metaAnual, 2)%></div>
                        </div>
                        <div class="info-box">
                            <div class="info-title">Vendas Realizadas</div>
                            <div class="info-value">R$ <%=FormatNumber(totalVendasAno, 2)%></div>
                        </div>
                        <div class="info-box">
                            <div class="info-title">Performance</div>
                            <div class="info-value"><%=FormatNumber(performanceAnual, 1)%>%</div>
                        </div>
                        <div class="progress mt-3">
                            <div class="progress-bar" role="progressbar" style="width: <%=performanceAnual%>%;" aria-valuenow="<%=performanceAnual%>" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                    </div>
                </div>
                
                <div class="card mt-4">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-area"></i> Vendas por Tipo de Unidade</h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoTipoUnidade"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<div class="controls">
    <div class="btn-group" role="group">
        <button type="button" class="btn btn-primary" id="prevSlide">
            <i class="fas fa-chevron-left"></i>
        </button>
        <button type="button" class="btn btn-success" id="playPause">
            <i class="fas fa-play" id="playIcon"></i>
        </button>
        <button type="button" class="btn btn-primary" id="nextSlide">
            <i class="fas fa-chevron-right"></i>
        </button>
    </div>
    <div class="mt-2">
        <select class="form-select form-select-sm" id="autoTimeSelect">
            <option value="5" <% If CStr(autoTime) = "5" Then Response.Write "selected" %>>5s</option>
            <option value="10" <% If CStr(autoTime) = "10" Then Response.Write "selected" %>>10s</option>
            <option value="15" <% If CStr(autoTime) = "15" Then Response.Write "selected" %>>15s</option>
            <option value="20" <% If CStr(autoTime) = "20" Then Response.Write "selected" %>>20s</option>
            <option value="25" <% If CStr(autoTime) = "25" Then Response.Write "selected" %>>25s</option>
            <option value="30" <% If CStr(autoTime) = "30" Then Response.Write "selected" %>>30s</option>
        </select>
    </div>
</div>

<script>
    // Configuração do slideshow
    let currentSlide = 1;
    const totalSlides = 5;
    let autoPlay = true;
    let slideInterval;
    const countdownTimer = document.getElementById('countdown-timer');
    const secondsLeftSpan = document.getElementById('seconds-left');
    const playPauseBtn = document.getElementById('playPause');
    const playIcon = document.getElementById('playIcon');
    const autoTimeSelect = document.getElementById('autoTimeSelect');
    let slideDuration = <%=autoTime%>;

    function showSlide(slideNumber) {
        // Esconder todos os slides
        document.querySelectorAll('.slide').forEach(slide => {
            slide.classList.remove('active');
        });
        
        // Mostrar o slide atual
        document.getElementById('slide' + slideNumber).classList.add('active');
        currentSlide = slideNumber;
        
        // Atualizar indicadores
        document.querySelectorAll('.slide-dot').forEach((dot, index) => {
            if (index + 1 === slideNumber) {
                dot.classList.add('active');
            } else {
                dot.classList.remove('active');
            }
        });
        
        // Reiniciar o contador
        resetCountdown();
    }

    function nextSlide() {
        let next = currentSlide + 1;
        if (next > totalSlides) next = 1;
        showSlide(next);
    }

    function prevSlide() {
        let prev = currentSlide - 1;
        if (prev < 1) prev = totalSlides;
        showSlide(prev);
    }

    function startAutoPlay() {
        if (autoPlay) {
            slideInterval = setInterval(nextSlide, slideDuration * 1000);
            countdownTimer.style.display = 'block';
            playIcon.classList.remove('fa-play');
            playIcon.classList.add('fa-pause');
        }
    }

    function stopAutoPlay() {
        clearInterval(slideInterval);
        countdownTimer.style.display = 'none';
        playIcon.classList.remove('fa-pause');
        playIcon.classList.add('fa-play');
    }

    function toggleAutoPlay() {
        autoPlay = !autoPlay;
        if (autoPlay) {
            startAutoPlay();
        } else {
            stopAutoPlay();
        }
    }

    function resetCountdown() {
        if (autoPlay) {
            let secondsLeft = slideDuration;
            secondsLeftSpan.textContent = secondsLeft;
            
            // Atualizar a cada segundo
            clearInterval(countdownTimer.interval);
            countdownTimer.interval = setInterval(() => {
                secondsLeft--;
                secondsLeftSpan.textContent = secondsLeft;
                
                if (secondsLeft <= 0) {
                    clearInterval(countdownTimer.interval);
                }
            }, 1000);
        }
    }

    function updateSlideDuration() {
        slideDuration = parseInt(autoTimeSelect.value);
        if (autoPlay) {
            stopAutoPlay();
            startAutoPlay();
        }
    }

    // Event listeners
    document.getElementById('nextSlide').addEventListener('click', nextSlide);
    document.getElementById('prevSlide').addEventListener('click', prevSlide);
    document.getElementById('playPause').addEventListener('click', toggleAutoPlay);
    autoTimeSelect.addEventListener('change', updateSlideDuration);

    // Event listeners para os indicadores de slide
    document.querySelectorAll('.slide-dot').forEach(dot => {
        dot.addEventListener('click', function() {
            const slideNumber = parseInt(this.getAttribute('data-slide'));
            showSlide(slideNumber);
        });
    });

    // Iniciar o slideshow
    startAutoPlay();

    // Gráfico de vendas anual
    const ctxVendasAnual = document.getElementById('graficoVendasAnual').getContext('2d');
    new Chart(ctxVendasAnual, {
        type: 'bar',
        data: {
            labels: <%=JSON_Serialize(mesesAno)%>,
            datasets: [{
                label: 'Vendas Mensais',
                data: <%=JSON_Serialize(dadosVendasAnual)%>,
                backgroundColor: 'rgba(160, 196, 255, 0.7)',
                borderColor: 'rgba(160, 196, 255, 1)',
                borderWidth: 2,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: {
                        color: '#ffffff',
                        font: {
                            size: 14
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    callbacks: {
                        label: function(context) {
                            return 'R$ ' + context.parsed.y.toLocaleString('pt-BR');
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        }
                    }
                },
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        },
                        callback: function(value) {
                            return 'R$ ' + value.toLocaleString('pt-BR');
                        }
                    }
                }
            }
        }
    });

    // Gráfico de diretorias
    const ctxDiretorias = document.getElementById('graficoDiretorias').getContext('2d');
    new Chart(ctxDiretorias, {
        type: 'pie',
        data: {
            labels: [<%=diretorias%>],
            datasets: [{
                data: [<%=totaisDiretorias%>],
                backgroundColor: [
                    'rgba(255, 99, 132, 0.7)',
                    'rgba(54, 162, 235, 0.7)',
                    'rgba(255, 206, 86, 0.7)',
                    'rgba(75, 192, 192, 0.7)',
                    'rgba(153, 102, 255, 0.7)',
                    'rgba(255, 159, 64, 0.7)',
                    'rgba(199, 199, 199, 0.7)',
                    'rgba(83, 102, 255, 0.7)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)',
                    'rgba(199, 199, 199, 1)',
                    'rgba(83, 102, 255, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = Math.round((value / total) * 100);
                            return `${label}: R$ ${value.toLocaleString('pt-BR')} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });

    // Gráfico de gerências
    const ctxGerencias = document.getElementById('graficoGerencias').getContext('2d');
    new Chart(ctxGerencias, {
        type: 'bar',
        data: {
            labels: [<%=gerencias%>],
            datasets: [{
                label: 'Vendas por Gerência',
                data: [<%=totaisGerencias%>],
                backgroundColor: 'rgba(160, 196, 255, 0.7)',
                borderColor: 'rgba(160, 196, 255, 1)',
                borderWidth: 2,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    callbacks: {
                        label: function(context) {
                            return 'R$ ' + context.parsed.x.toLocaleString('pt-BR');
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        },
                        callback: function(value) {
                            return 'R$ ' + value.toLocaleString('pt-BR');
                        }
                    }
                },
                y: {
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        }
                    }
                }
            }
        }
    });

    // Gráfico de ticket médio
    const ctxTicketMedio = document.getElementById('graficoTicketMedio').getContext('2d');
    new Chart(ctxTicketMedio, {
        type: 'line',
        data: {
            labels: <%=JSON_Serialize(mesesAno)%>,
            datasets: [{
                label: 'Ticket Médio Mensal',
                data: <%=JSON_Serialize(ticketMedioMensal)%>,
                backgroundColor: 'rgba(160, 196, 255, 0.2)',
                borderColor: 'rgba(160, 196, 255, 1)',
                borderWidth: 3,
                tension: 0.3,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: {
                        color: '#ffffff',
                        font: {
                            size: 14
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    callbacks: {
                        label: function(context) {
                            return 'R$ ' + context.parsed.y.toLocaleString('pt-BR');
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        }
                    }
                },
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(255,255,255,0.1)'
                    },
                    ticks: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        },
                        callback: function(value) {
                            return 'R$ ' + value.toLocaleString('pt-BR');
                        }
                    }
                }
            }
        }
    });

    // Gráfico de tipo de unidade
    const ctxTipoUnidade = document.getElementById('graficoTipoUnidade').getContext('2d');
    new Chart(ctxTipoUnidade, {
        type: 'doughnut',
        data: {
            labels: [<%=tiposUnidade%>],
            datasets: [{
                data: [<%=vendasTiposUnidade%>],
                backgroundColor: [
                    'rgba(255, 99, 132, 0.7)',
                    'rgba(54, 162, 235, 0.7)',
                    'rgba(255, 206, 86, 0.7)',
                    'rgba(75, 192, 192, 0.7)',
                    'rgba(153, 102, 255, 0.7)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#ffffff',
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = Math.round((value / total) * 100);
                            return `${label}: R$ ${value.toLocaleString('pt-BR')} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
</script>
</body>
</html>

<%
' FECHAR A CONEXÃO APENAS NO FINAL DO ARQUIVO
If IsObject(conn) Then
    If conn.State = 1 Then ' Se a conexão está aberta
        conn.Close
    End If
    Set conn = Nothing
End If
%>