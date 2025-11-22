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
' VARIÁVEL PARA CONTROLE DA LARGURA DO GRÁFICO
' ===============================================
Dim vLargura
vLargura = 1000 ' ALTERE ESTE VALOR PARA AJUSTAR A LARGURA: 1200, 1400, 1600, 1800, 2000

' ===============================================
' OBTER DADOS DE VENDAS POR SEMANA
' ===============================================

Dim sqlVendas, rsVendas, anoFiltro
anoFiltro = Request.QueryString("ano")
If anoFiltro = "" Then
    anoFiltro = Year(Date())
End If

' Função para calcular semana do ano no Access
Function GetWeekNumber(dtDate)
    Dim dtFirstDay, iWeekday, iOffset, dtFirstSunday
    
    ' Primeiro dia do ano
    dtFirstDay = DateSerial(Year(dtDate), 1, 1)
    
    ' Dia da semana do primeiro dia (1=Domingo, 2=Segunda, ..., 7=Sábado)
    iWeekday = Weekday(dtFirstDay)
    
    ' Calcular o primeiro domingo do ano
    If iWeekday = 1 Then
        dtFirstSunday = dtFirstDay
    Else
        dtFirstSunday = DateAdd("d", 8 - iWeekday, dtFirstDay)
    End If
    
    ' Calcular número da semana
    If dtDate < dtFirstSunday Then
        GetWeekNumber = 1
    Else
        GetWeekNumber = Int((dtDate - dtFirstSunday) / 7) + 2
    End If
End Function

' Função para criar data a partir de DiaVenda, MesVenda, AnoVenda
Function CreateDateFromFields(dia, mes, ano)
    If IsNumeric(dia) And IsNumeric(mes) And IsNumeric(ano) Then
        If dia >= 1 And dia <= 31 And mes >= 1 And mes <= 12 And ano >= 2000 Then
            CreateDateFromFields = DateSerial(ano, mes, dia)
        Else
            CreateDateFromFields = Null
        End If
    Else
        CreateDateFromFields = Null
    End If
End Function

' Consulta para obter todas as vendas do ano
sqlVendas = "SELECT " & _
            "DiaVenda, MesVenda, AnoVenda " & _
            "FROM Vendas " & _
            "WHERE (Excluido <> -1 OR Excluido IS NULL) " & _
            "AND AnoVenda = " & anoFiltro & " " & _
            "AND DiaVenda IS NOT NULL AND MesVenda IS NOT NULL AND AnoVenda IS NOT NULL " & _
            "ORDER BY AnoVenda, MesVenda, DiaVenda"

Set rsVendas = Server.CreateObject("ADODB.Recordset")
rsVendas.Open sqlVendas, connSales

' Criar array para armazenar as quantidades por semana
Dim quantidades(53)
Dim maxQuantidade
maxQuantidade = 0

' Inicializar array
For i = 1 To 53
    quantidades(i) = 0
Next

' Processar vendas e agrupar por semana
If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        If Not IsNull(rsVendas("DiaVenda")) And Not IsNull(rsVendas("MesVenda")) And Not IsNull(rsVendas("AnoVenda")) Then
            Dim dataVenda
            dataVenda = CreateDateFromFields(rsVendas("DiaVenda"), rsVendas("MesVenda"), rsVendas("AnoVenda"))
            
            If Not IsNull(dataVenda) Then
                semana = GetWeekNumber(dataVenda)
                If semana >= 1 And semana <= 53 Then
                    quantidades(semana) = quantidades(semana) + 1
                    
                    If quantidades(semana) > maxQuantidade Then
                        maxQuantidade = quantidades(semana)
                    End If
                End If
            End If
        End If
        rsVendas.MoveNext
    Loop
End If

' Obter anos disponíveis para o filtro
Dim rsAnos, uniqueAnos
Set rsAnos = connSales.Execute("SELECT DISTINCT AnoVenda as Ano FROM Vendas WHERE AnoVenda IS NOT NULL ORDER BY AnoVenda DESC")
uniqueAnos = ""
If Not rsAnos.EOF Then
    Do While Not rsAnos.EOF
        uniqueAnos = uniqueAnos & "<option value=""" & rsAnos("Ano") & """"
        If CStr(rsAnos("Ano")) = CStr(anoFiltro) Then
            uniqueAnos = uniqueAnos & " selected"
        End If
        uniqueAnos = uniqueAnos & ">" & rsAnos("Ano") & "</option>"
        rsAnos.MoveNext
    Loop
End If
rsAnos.Close
Set rsAnos = Nothing

rsVendas.Close
Set rsVendas = Nothing
connSales.Close
Set connSales = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vendas por Semana</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
            padding: 15px;
        }
        .container {
            max-width: <%= vLargura %>px;
            margin: 0 auto;
        }
        .header {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }
        .page-title {
            color: #2c3e50;
            font-size: 24px;
            font-weight: 700;
            text-align: center;
            margin-bottom: 5px;
        }
        .page-subtitle {
            color: #7f8c8d;
            font-size: 14px;
            text-align: center;
            margin-bottom: 15px;
        }
        .filter-card {
            background: rgba(255, 255, 255, 0.9);
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 20px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        .chart-container {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
            min-height: 500px;
            width: 100%;
            overflow-x: auto;
        }
        .chart-wrapper {
            min-width: <%= vLargura - 50 %>px;
            height: 400px;
            position: relative;
        }
        .stats-card {
            background: rgba(255, 255, 255, 0.9);
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 15px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        .stat-number {
            font-size: 24px;
            font-weight: 700;
            color: #2c3e50;
            text-align: center;
        }
        .stat-label {
            font-size: 12px;
            color: #7f8c8d;
            text-align: center;
        }
        .form-select {
            border-radius: 8px;
            border: 1px solid #e9ecef;
            padding: 10px;
            font-size: 14px;
        }
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: 600;
        }
        
        /* NOVO LAYOUT PARA SEMANAS */
        .weeks-container {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
        .weeks-section {
            margin-bottom: 25px;
        }
        .weeks-section:last-child {
            margin-bottom: 0;
        }
        .section-title {
            color: #2c3e50;
            font-size: 16px;
            font-weight: 700;
            margin-bottom: 15px;
            text-align: center;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 8px;
        }
        .weeks-grid {
            display: grid;
            grid-template-columns: repeat(14, 1fr);
            gap: 8px;
            margin-bottom: 15px;
        }
        .week-item-compact {
            background: #f8f9fa;
            border-radius: 6px;
            padding: 8px 4px;
            text-align: center;
            border: 1px solid #e9ecef;
            transition: all 0.3s ease;
        }
        .week-item-compact:hover {
            transform: translateY(-2px);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        }
        .week-number-compact {
            font-size: 10px;
            font-weight: 600;
            color: #6c757d;
            margin-bottom: 4px;
        }
        .week-bar-compact {
            border-radius: 3px;
            min-height: 4px;
            margin-bottom: 4px;
        }
        .week-quantity-compact {
            font-size: 11px;
            font-weight: 700;
        }
        .week-high {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }
        .week-medium {
            background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);
        }
        .week-low {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        .week-zero {
            background: #e9ecef;
        }
        .quantity-high {
            color: #28a745;
        }
        .quantity-medium {
            color: #fd7e14;
        }
        .quantity-low {
            color: #667eea;
        }
        .quantity-zero {
            color: #6c757d;
        }
        .chart-legend {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin: 20px 0;
            flex-wrap: wrap;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 12px;
            font-weight: 500;
        }
        .legend-color {
            width: 14px;
            height: 14px;
            border-radius: 3px;
        }
        .data-info {
            background: #e7f3ff;
            border-radius: 8px;
            padding: 10px;
            margin-top: 15px;
            font-size: 12px;
            color: #2c3e50;
            text-align: center;
        }
        .config-info {
            background: #fff3cd;
            border-radius: 8px;
            padding: 8px 12px;
            margin-bottom: 15px;
            font-size: 11px;
            color: #856404;
            text-align: center;
            border: 1px solid #ffeaa7;
        }
        @media (max-width: 768px) {
            .weeks-grid {
                grid-template-columns: repeat(7, 1fr);
            }
            .container {
                max-width: 100%;
                padding: 10px;
            }
            .chart-wrapper {
                min-width: 100%;
            }
        }
        @media (max-width: 480px) {
            .weeks-grid {
                grid-template-columns: repeat(4, 1fr);
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Cabeçalho -->
        <div class="header">
            <h1 class="page-title">
                <i class="fas fa-chart-bar"></i> Vendas por Semana - <%= anoFiltro %>
            </h1>
            <p class="page-subtitle">Quantidade de unidades vendidas por semana do ano</p>
            
            <!-- Informação de Configuração -->
            <div class="config-info">
                <i class="fas fa-cog"></i> 
                Largura do gráfico: <strong><%= vLargura %>px</strong> 
                | Alterar variável <code>vLargura</code> no código
            </div>
            
            <!-- Filtros -->
            <div class="filter-card">
                <form id="filterForm" method="get" class="row g-3 align-items-center">
                    <div class="col-md-6">
                        <label class="form-label fw-bold">Ano:</label>
                        <select class="form-select" name="ano" id="anoFilter">
                            <%= uniqueAnos %>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <button type="submit" class="btn btn-primary mt-4">
                            <i class="fas fa-filter"></i> Aplicar Filtro
                        </button>
                    </div>
                </form>
            </div>

            <!-- Estatísticas -->
            <%
            Dim totalVendas, semanasComVenda, semanaMaiorVenda, maiorQuantidade
            totalVendas = 0
            semanasComVenda = 0
            semanaMaiorVenda = 0
            maiorQuantidade = 0
            
            For i = 1 To 53
                totalVendas = totalVendas + quantidades(i)
                If quantidades(i) > 0 Then
                    semanasComVenda = semanasComVenda + 1
                End If
                If quantidades(i) > maiorQuantidade Then
                    maiorQuantidade = quantidades(i)
                    semanaMaiorVenda = i
                End If
            Next
            %>
            
            <div class="row">
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number"><%= totalVendas %></div>
                        <div class="stat-label">Total de Vendas</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number"><%= semanasComVenda %></div>
                        <div class="stat-label">Semanas com Vendas</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number">
                            <% If semanasComVenda > 0 Then %>
                                <%= FormatNumber(totalVendas / semanasComVenda, 1) %>
                            <% Else %>
                                0
                            <% End If %>
                        </div>
                        <div class="stat-label">Média por Semana</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number">
                            <% If semanaMaiorVenda > 0 Then %>
                                <%= maiorQuantidade %>
                            <% Else %>
                                0
                            <% End If %>
                        </div>
                        <div class="stat-label">Melhor Semana (Sem <%= semanaMaiorVenda %>)</div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Gráfico -->
        <div class="chart-container">
            <div class="chart-wrapper">
                <canvas id="vendasChart"></canvas>
            </div>
        </div>

        <!-- VISUALIZAÇÃO DETALHADA POR SEMANA -->
        <div class="weeks-container">
            <h5 class="text-center mb-4" style="color: #2c3e50;">
                <i class="fas fa-th-list"></i> Visualização Detalhada por Semana
            </h5>
            
            <div class="chart-legend">
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%);"></div>
                    <span>Alta (≥ 5 vendas)</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);"></div>
                    <span>Média (2-4 vendas)</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);"></div>
                    <span>Baixa (1 venda)</span>
                </div>
            </div>

            <!-- Primeira seção: Semanas 1-14 -->
            <div class="weeks-section">
                <div class="section-title">Semanas 1-14</div>
                <div class="weeks-grid">
                    <%
                    For semana = 1 To 14
                        quantidade = quantidades(semana)
                        Call RenderWeekItem(semana, quantidade)
                    Next
                    %>
                </div>
            </div>

            <!-- Segunda seção: Semanas 15-28 -->
            <div class="weeks-section">
                <div class="section-title">Semanas 15-28</div>
                <div class="weeks-grid">
                    <%
                    For semana = 15 To 28
                        quantidade = quantidades(semana)
                        Call RenderWeekItem(semana, quantidade)
                    Next
                    %>
                </div>
            </div>

            <!-- Terceira seção: Semanas 29-42 -->
            <div class="weeks-section">
                <div class="section-title">Semanas 29-42</div>
                <div class="weeks-grid">
                    <%
                    For semana = 29 To 42
                        quantidade = quantidades(semana)
                        Call RenderWeekItem(semana, quantidade)
                    Next
                    %>
                </div>
            </div>

            <!-- Quarta seção: Semanas 43-53 -->
            <div class="weeks-section">
                <div class="section-title">Semanas 43-53</div>
                <div class="weeks-grid">
                    <%
                    For semana = 43 To 53
                        quantidade = quantidades(semana)
                        Call RenderWeekItem(semana, quantidade)
                    Next
                    ' Preencher com espaços vazios para manter o layout
                    For semana = 54 To 56
                    %>
                    <div class="week-item-compact" style="visibility: hidden;">
                        <div class="week-number-compact">Sem <%= semana %></div>
                        <div class="week-bar-compact week-zero"></div>
                        <div class="week-quantity-compact quantity-zero">0</div>
                    </div>
                    <%
                    Next
                    %>
                </div>
            </div>

            <div class="data-info">
                <i class="fas fa-info-circle"></i> 
                Contagem de Vendas por Semana - Ano <%= anoFiltro %>
            </div>
        </div>
    </div>

    <script>
        // Dados para o gráfico
        const semanas = [
            <% For i = 1 To 53: Response.Write "'" & i & "'" & IIf(i < 53, ",", ""): Next %>
        ];
        
        const quantidades = [
            <% For i = 1 To 53: Response.Write quantidades(i) & IIf(i < 53, ",", ""): Next %>
        ];

        // Configuração da largura do gráfico
        const LARGURA_GRAFICO = <%= vLargura - 100 %>;

        // Calcular o valor máximo para ajustar o eixo Y
        const maxQuantidade = Math.max(...quantidades);
        const suggestedMax = Math.max(maxQuantidade + 2, 5);

        // Configuração do gráfico
        const ctx = document.getElementById('vendasChart').getContext('2d');
        const vendasChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: semanas,
                datasets: [{
                    label: 'Unidades Vendidas',
                    data: quantidades,
                    backgroundColor: function(context) {
                        const value = context.dataset.data[context.dataIndex];
                        if (value >= 5) {
                            return 'rgba(40, 167, 69, 0.8)';
                        } else if (value >= 2) {
                            return 'rgba(255, 193, 7, 0.8)';
                        } else if (value >= 1) {
                            return 'rgba(102, 126, 234, 0.8)';
                        } else {
                            return 'rgba(222, 226, 230, 0.8)';
                        }
                    },
                    borderColor: function(context) {
                        const value = context.dataset.data[context.dataIndex];
                        if (value >= 5) {
                            return 'rgba(40, 167, 69, 1)';
                        } else if (value >= 2) {
                            return 'rgba(255, 193, 7, 1)';
                        } else if (value >= 1) {
                            return 'rgba(102, 126, 234, 1)';
                        } else {
                            return 'rgba(222, 226, 230, 1)';
                        }
                    },
                    borderWidth: 1,
                    borderRadius: 4,
                    barPercentage: 0.7,
                    categoryPercentage: 0.8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 40,
                        bottom: 20,
                        left: 10,
                        right: 10
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const value = context.parsed.y;
                                return 'Semana ' + (context.dataIndex + 1) + ': ' + value + ' unidade(s)';
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Quantidade de Unidades Vendidas'
                        },
                        ticks: {
                            stepSize: 1,
                            precision: 0
                        },
                        suggestedMax: suggestedMax,
                        grace: '15%'
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Número da Semana'
                        },
                        ticks: {
                            maxTicksLimit: 53,
                            callback: function(value, index, values) {
                                return index % 4 === 0 ? this.getLabelForValue(value) : '';
                            }
                        },
                        afterFit: function(scale) {
                            scale.width = LARGURA_GRAFICO;
                        }
                    }
                }
            },
            plugins: [{
                id: 'datalabels',
                afterDatasetsDraw: function(chart, args, options) {
                    const { ctx, chartArea: { top, bottom, left, right } } = chart;
                    
                    ctx.save();
                    
                    chart.data.datasets.forEach(function(dataset, i) {
                        const meta = chart.getDatasetMeta(i);
                        if (!meta.hidden) {
                            meta.data.forEach(function(element, index) {
                                const value = dataset.data[index];
                                
                                if (value > 0) {
                                    const barHeight = bottom - element.y;
                                    const isTallBar = barHeight > 25;
                                    
                                    if (isTallBar) {
                                        ctx.fillStyle = '#2c3e50';
                                        ctx.font = 'bold 11px Arial';
                                        ctx.textAlign = 'center';
                                        ctx.textBaseline = 'bottom';
                                        
                                        const yPosition = element.y - 8;
                                        
                                        if (yPosition > top + 15) {
                                            ctx.fillText(value, element.x, yPosition);
                                        }
                                    } else {
                                        ctx.fillStyle = '#ffffff';
                                        ctx.font = 'bold 10px Arial';
                                        ctx.textAlign = 'center';
                                        ctx.textBaseline = 'middle';
                                        
                                        const yPosition = element.y + (barHeight / 2);
                                        ctx.fillText(value, element.x, yPosition);
                                    }
                                }
                            });
                        }
                    });
                    
                    ctx.restore();
                }
            }]
        });

        // Ajustar o gráfico após o carregamento
        setTimeout(function() {
            vendasChart.resize();
        }, 100);

        window.addEventListener('resize', function() {
            vendasChart.resize();
        });
    </script>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<%
' Função para renderizar cada item de semana
Sub RenderWeekItem(semana, quantidade)
    ' Definir classe de cor baseada na quantidade
    Dim barClass, quantityClass
    If quantidade >= 5 Then
        barClass = "week-high"
        quantityClass = "quantity-high"
    ElseIf quantidade >= 2 Then
        barClass = "week-medium"
        quantityClass = "quantity-medium"
    ElseIf quantidade >= 1 Then
        barClass = "week-low"
        quantityClass = "quantity-low"
    Else
        barClass = "week-zero"
        quantityClass = "quantity-zero"
    End If
    %>
    <div class="week-item-compact" title="Semana <%= semana %>: <%= quantidade %> venda(s)">
        <div class="week-number-compact">Sem<%= semana %></div>
        <div class="week-bar-compact <%= barClass %>"></div>
        <div class="week-quantity-compact <%= quantityClass %>"><%= quantidade %></div>
    </div>
    <%
End Sub

' Função IIF para VBScript
Function IIf(expr, trueval, falseval)
    If expr Then
        IIf = trueval
    Else
        IIf = falseval
    End If
End Function
%>