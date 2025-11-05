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

' =======================================================
' INÍCIO DO PROCESSAMENTO
' =======================================================

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
If autoTime = "" Then autoTime = 5
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard de Vendas - Modo Autônomo</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            background-color: #f8f9fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
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
        .list-group-item {
            border-left: none;
            border-right: none;
            font-weight: 500;
        }
        .list-group-item:first-child {
            border-top: none;
        }
        .list-group-item:last-child {
            border-bottom: none;
        }
        .bg-primary {
            background-color: #4361ee !important;
        }
        .bg-success {
            background-color: #4cc9f0 !important;
        }
        .bg-info {
            background-color: #3a0ca3 !important;
        }
        .bg-warning {
            background-color: #f72585 !important;
        }
        .badge {
            font-weight: 600;
            padding: 5px 10px;
            border-radius: 10px;
        }
        .filter-column {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
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
        .spinner-border {
            width: 5rem;
            height: 5rem;
            border-width: .5em;
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
        /* Layout de três colunas */
        .main-container {
            display: grid;
            grid-template-columns: 250px 1fr 1fr; /* Coluna 1 fixa, Coluna 2 e 3 flexíveis */
            gap: 20px;
            padding: 20px;
        }
        @media (max-width: 992px) {
            .main-container {
                grid-template-columns: 1fr; /* Em telas menores, o layout volta a ser de uma coluna */
            }
            .sidebar {
                grid-column: 1 / -1; /* Ocupa a largura total */
            }
            .content-center {
                grid-column: 1 / -1; /* Ocupa a largura total */
            }
            .content-right {
                grid-column: 1 / -1; /* Ocupa a largura total */
            }
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
    <h1 class="mb-1 text-center">SGVendas - Dashboard de Vendas</h1>
    <div class="main-container">
        <div class="sidebar">
            <div class="text-center mt-4">
                <a href="gestao_painel2.asp" class="btn btn-primary btn-sm" target="_blank">
                    <i class="fas fa-arrow-right"></i> Gerenciar Vendas
                </a>
            </div>
            <div class="filter-column">
                <h5 class="text-center">Filtros</h5>
                <form method="get" id="filterForm">
                    <div class="mb-3">
                        <label for="anoFilter" class="form-label">Ano</label>
                        <select class="form-select" id="anoFilter" name="ano">
                            <option value="">Todos</option>
                            <% For Each ano In uniqueAnos %>
                                <option value="<%=ano%>" <% If Request.QueryString("ano") = CStr(ano) Then Response.Write "selected" %>><%=ano%></option>
                            <% Next %>
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
            <div class="row">
                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏆 Top 10 Corretores</h5>
                        </div>
                        <ul class="list-group list-group-flush">
                            <%
                            SQL = "SELECT Vendas.Corretor, Sum(Vendas.ValorUnidade) AS Total FROM Vendas " & whereClause & " GROUP BY Vendas.Corretor ORDER BY Sum(Vendas.ValorUnidade) DESC;"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            contador = 0
                            Do Until rs.EOF Or contador = 10
                                Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'><span>" & rs("Corretor") & "</span><span class='badge bg-primary'>R$ " & FormatNumber(rs("Total"), 2) & "</span></li>"
                                contador = contador + 1
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </ul>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">👔 Top 10 Gerentes</h5>
                        </div>
                        <ul class="list-group list-group-flush">
                            <%
                            SQL = "SELECT Gerencia, SUM(ValorUnidade) AS Total FROM vendas " & whereClause & " GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            contador = 0
                            Do Until rs.EOF Or contador = 10
                                Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'><span>" & rs("Gerencia") & "</span><span class='badge bg-success'>R$ " & FormatNumber(rs("Total"), 2) & "</span></li>"
                                contador = contador + 1
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </ul>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏢 Top 10 Diretorias</h5>
                        </div>
                        <ul class="list-group list-group-flush">
                            <%
                            SQL = "SELECT Diretoria, SUM(ValorUnidade) AS Total FROM vendas " & whereClause & " GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            contador = 0
                            Do Until rs.EOF Or contador = 10
                                Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'><span>" & rs("Diretoria") & "</span><span class='badge bg-info'>R$ " & FormatNumber(rs("Total"), 2) & "</span></li>"
                                contador = contador + 1
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </ul>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card">
                        <div class="card-header text-white">
                            <h5 class="mb-0">🏗️ Top 10 Empreendimentos</h5>
                        </div>
                        <ul class="list-group list-group-flush">
                            <%
                            SQL = "SELECT TOP 10 NomeEmpreendimento, SUM(ValorUnidade) AS Total FROM vendas " & whereClause & " GROUP BY NomeEmpreendimento ORDER BY SUM(ValorUnidade) DESC"
                            Set rs = Server.CreateObject("ADODB.Recordset")
                            rs.Open SQL, conn

                            If Not rs.EOF Then
                                Do While Not rs.EOF
                                    Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'><span>" & rs("NomeEmpreendimento") & "</span><span class='badge bg-warning'>R$ " & FormatNumber(rs("Total"), 2) & "</span></li>"
                                    rs.MoveNext
                                Loop
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </ul>
                    </div>
                </div>
            </div>
        </div>

        <div class="content-right">
            <div class="card">
                <div class="card-header text-white">
                    <h5 class="mb-0">📈 Gráfico de Vendas por Ano-Mês</h5>
                </div>
                <div class="card-body">
                    <canvas id="graficoVendas" height="250"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <%
    Dim datasetsJSON, colors(5), colorIndex, ano, SQL_Anos, rsAnos, SQL_Dados, rsDados
    Dim dadosAno, dadosPorMes(12), mesAtual, i
    
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
                dadosPorMes(mesAtual) = CStr(rsDados("Total"))
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
    <div id="countdown-timer">Atualizando em: <span id="seconds-left">0</span>s</div>
    
    <script>
        const filterNames = ['ano', 'mes', 'diretoria', 'gerencia', 'corretor', 'empreendimento'];
        let timerInterval;
        
        const urlParams = new URLSearchParams(window.location.search);
        const timerDuration = parseInt(urlParams.get('autotime')) || 10;
        
        const autoModeBtn = document.getElementById('autoModeBtn');
        const loadingSpinner = document.getElementById('loadingSpinner');
        const countdownTimer = document.getElementById('countdown-timer');
        const secondsLeftSpan = document.getElementById('seconds-left');

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

        function stopTimer() {
            clearInterval(timerInterval);
            countdownTimer.style.display = 'none';
        }

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
        
        const isAutoModeActive = urlParams.get('auto_mode') === 'on';

        if (isAutoModeActive) {
            autoModeBtn.innerHTML = '<i class="fas fa-pause-circle"></i> Parar Modo Autônomo';
            autoModeBtn.classList.remove('btn-info');
            autoModeBtn.classList.add('btn-danger');
            startTimer();
        }

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
        
        document.getElementById('filterForm').addEventListener('submit', function() {
            document.getElementById('loadingSpinner').style.display = 'flex';
        });

        const ctx = document.getElementById('graficoVendas').getContext('2d');
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                datasets: [<%=datasetsJSON%>]
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
                                    label += new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(context.parsed.y);
                                }
                                return label;
                            }
                        }
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
                                return 'R$ ' + value.toLocaleString('pt-BR');
                            }
                        }
                    }
                }
            }
        });
    </script>
    </div>
</div>
<!--#include file="footer.inc"-->
</body>
</html>