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
' OBTER DADOS DE VALORES VENDIDOS POR MÊS
' ===============================================

Dim sqlVendas, rsVendas, anoFiltro
anoFiltro = Request.QueryString("ano")
If anoFiltro = "" Then
    anoFiltro = Year(Date())
End If

' Consulta para obter valores vendidos agrupados por mês
sqlVendas = "SELECT " & _
            "MesVenda, " & _
            "SUM(ValorUnidade) as TotalValor " & _
            "FROM Vendas " & _
            "WHERE (Excluido <> -1 OR Excluido IS NULL) " & _
            "AND AnoVenda = " & anoFiltro & " " & _
            "AND MesVenda IS NOT NULL " & _
            "AND ValorUnidade IS NOT NULL " & _
            "GROUP BY MesVenda " & _
            "ORDER BY MesVenda"
            'response.Write sqlVendas
           ' Response.end 

Set rsVendas = Server.CreateObject("ADODB.Recordset")
rsVendas.Open sqlVendas, connSales

' Criar array para armazenar os valores por mês
Dim valores(12)
Dim acumulado(12)
Dim maxValor
maxValor = 0

' Inicializar arrays
For i = 1 To 12
    valores(i) = 0
    acumulado(i) = 0
Next

' Processar vendas e preencher arrays
If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        If Not IsNull(rsVendas("MesVenda")) And Not IsNull(rsVendas("TotalValor")) Then
            mes = CInt(rsVendas("MesVenda"))
            If mes >= 1 And mes <= 12 Then
                ' Converter o valor para número (remover formatação se houver)
                Dim valorVenda
                valorVenda = rsVendas("TotalValor")
                ' Se for string, converter para número
                If IsNumeric(valorVenda) Then
                    valores(mes) = CDbl(valorVenda)
                Else
                    ' Remover pontos e substituir vírgula por ponto
                    Dim valorLimpo
                    valorLimpo = Replace(valorVenda, ".", "")
                    valorLimpo = Replace(valorLimpo, ",", ".")
                    If IsNumeric(valorLimpo) Then
                        valores(mes) = CDbl(valorLimpo)
                    Else
                        valores(mes) = 0
                    End If
                End If
                
                If valores(mes) > maxValor Then
                    maxValor = valores(mes)
                End If
            End If
        End If
        rsVendas.MoveNext
    Loop
End If

' Calcular valores acumulados
Dim totalAcumulado
totalAcumulado = 0
For i = 1 To 12
    totalAcumulado = totalAcumulado + valores(i)
    acumulado(i) = totalAcumulado
Next

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

' Nomes dos meses
Dim meses
meses = Array("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", _
              "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")

' Função para formatar valores em Real
Function FormatReal(valor)
    If IsNumeric(valor) Then
        ' Formatar como moeda brasileira
        FormatReal = "R$ " & FormatNumber(valor, 2, -1, -1, -1)
    Else
        FormatReal = "R$ 0,00"
    End If
End Function

' Função para formatar valores sem símbolo (apenas números)
Function FormatValor(valor)
    If IsNumeric(valor) Then
        FormatValor = FormatNumber(valor, 2, -1, -1, -1)
    Else
        FormatValor = "0,00"
    End If
End Function

' Função para converter string monetária para número
Function ParseValorMonetario(valorString)
    If IsNumeric(valorString) Then
        ParseValorMonetario = CDbl(valorString)
    Else
        ' Remover R$, espaços e outros caracteres não numéricos
        Dim valorLimpo
        valorLimpo = Replace(valorString, "R$", "")
        valorLimpo = Replace(valorLimpo, " ", "")
        valorLimpo = Replace(valorLimpo, ".", "")
        valorLimpo = Replace(valorLimpo, ",", ".")
        If IsNumeric(valorLimpo) Then
            ParseValorMonetario = CDbl(valorLimpo)
        Else
            ParseValorMonetario = 0
        End If
    End If
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Valores Vendidos por Mês</title>
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
        
        /* LAYOUT PARA MESES */
        .months-container {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
        .months-grid {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 12px;
            margin-bottom: 15px;
        }
        .month-item {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 12px 8px;
            text-align: center;
            border: 1px solid #e9ecef;
            transition: all 0.3s ease;
        }
        .month-item:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }
        .month-name {
            font-size: 12px;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 6px;
        }
        .month-bar {
            border-radius: 4px;
            min-height: 6px;
            margin-bottom: 6px;
        }
        .month-value {
            font-size: 11px;
            font-weight: 700;
            margin-bottom: 4px;
            line-height: 1.2;
        }
        .month-acumulado {
            font-size: 10px;
            color: #6c757d;
            font-weight: 500;
            line-height: 1.2;
        }
        .month-high {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }
        .month-medium {
            background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);
        }
        .month-low {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        .month-zero {
            background: #e9ecef;
        }
        .value-high {
            color: #28a745;
        }
        .value-medium {
            color: #fd7e14;
        }
        .value-low {
            color: #667eea;
        }
        .value-zero {
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
            .months-grid {
                grid-template-columns: repeat(3, 1fr);
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
            .months-grid {
                grid-template-columns: repeat(2, 1fr);
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Cabeçalho -->
        <div class="header">
            <h1 class="page-title">
                <i class="fas fa-money-bill-wave"></i> Valores Vendidos por Mês - <%= anoFiltro %>
            </h1>
            <p class="page-subtitle">Valores em Reais (R$) vendidos por mês com acumulado anual</p>
            
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
            Dim totalValor, mesesComVenda, mesMaiorValor, maiorValorMes
            totalValor = 0
            mesesComVenda = 0
            mesMaiorValor = 0
            maiorValorMes = 0
            
            For i = 1 To 12
                totalValor = totalValor + valores(i)
                If valores(i) > 0 Then
                    mesesComVenda = mesesComVenda + 1
                End If
                If valores(i) > maiorValorMes Then
                    maiorValorMes = valores(i)
                    mesMaiorValor = i
                End If
            Next
            
            ' Calcular média mensal
            Dim mediaMensal
            If mesesComVenda > 0 Then
                mediaMensal = totalValor / mesesComVenda
            Else
                mediaMensal = 0
            End If
            %>
            
            <div class="row">
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number"><%= FormatReal(totalValor) %></div>
                        <div class="stat-label">Total Vendido</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number"><%= mesesComVenda %></div>
                        <div class="stat-label">Meses com Vendas</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number"><%= FormatReal(mediaMensal) %></div>
                        <div class="stat-label">Média por Mês</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="stats-card">
                        <div class="stat-number">
                            <% If mesMaiorValor > 0 Then %>
                                <%= FormatReal(maiorValorMes) %>
                            <% Else %>
                                R$ 0,00
                            <% End If %>
                        </div>
                        <div class="stat-label">Melhor Mês (<%= meses(mesMaiorValor) %>)</div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Gráfico com Barras e Linhas -->
        <div class="chart-container">
            <div class="chart-wrapper">
                <canvas id="vendasChart"></canvas>
            </div>
        </div>

        <!-- VISUALIZAÇÃO DETALHADA POR MÊS -->
        <div class="months-container">
            <h5 class="text-center mb-4" style="color: #2c3e50;">
                <i class="fas fa-th-list"></i> Visualização Detalhada por Mês
            </h5>
            
            <div class="chart-legend">
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%);"></div>
                    <span>Alto (≥ R$ 1.000)</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);"></div>
                    <span>Médio (R$ 500 - R$ 999)</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);"></div>
                    <span>Baixo (≤ R$ 499)</span>
                </div>
            </div>

            <div class="months-grid">
                <%
                For mes = 1 To 12
                    valor = valores(mes)
                    acumuladoMes = acumulado(mes)
                    Call RenderMonthItem(mes, valor, acumuladoMes)
                Next
                %>
            </div>

            <div class="data-info">
                <i class="fas fa-info-circle"></i> 
                Valores Vendidos por Mês - Ano <%= anoFiltro %> | 
                Barras: Valor do mês | 
                Números abaixo: Acumulado anual
            </div>
        </div>
    </div>
<!-- ==================================================== -->
<script>
        // Dados para o gráfico
        const meses = [
            <% For i = 1 To 12: Response.Write "'" & meses(i) & "'" & IIf(i < 12, ",", ""): Next %>
        ];
        
        const valores = [
            <% 
            For i = 1 To 12 
                ' Garantir que o valor use ponto como separador decimal
                If IsNumeric(valores(i)) Then
                    Response.Write Replace(CStr(valores(i)), ",", ".")
                Else
                    Response.Write "0"
                End If
                If i < 12 Then Response.Write ","
            Next 
            %>
        ];

        const acumulado = [
            <% 
            For i = 1 To 12 
                ' Garantir que o valor use ponto como separador decimal
                If IsNumeric(acumulado(i)) Then
                    Response.Write Replace(CStr(acumulado(i)), ",", ".")
                Else
                    Response.Write "0"
                End If
                If i < 12 Then Response.Write ","
            Next 
            %>
        ];

        // Debug: verificar os valores no console
        console.log('Valores:', valores);
        console.log('Acumulado:', acumulado);

        // Configuração da largura do gráfico
        const LARGURA_GRAFICO = <%= vLargura - 100 %>;

        // Calcular o valor máximo para ajustar o eixo Y
        const maxValor = Math.max(...valores);
        const maxAcumulado = Math.max(...acumulado);
        const suggestedMax = Math.max(maxValor * 1.2, 100);

        // Função para formatar valores em Real com abreviação
        function formatReal(valor) {
            if (valor >= 1000000) {
                // Milhões
                return 'R$ ' + (valor / 1000000).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
            } else if (valor >= 1000) {
                // Milhares
                return 'R$ ' + (valor / 1000).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'K';
            } else {
                // Valores normais
                return 'R$ ' + valor.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }
        }

        // Função para formatar valores curtos (para labels no gráfico)
        function formatShort(valor) {
            if (valor >= 1000000) {
                return (valor / 1000000).toFixed(1) + 'M';
            } else if (valor >= 1000) {
                return (valor / 1000).toFixed(1) + 'K';
            } else {
                return valor.toFixed(0);
            }
        }

        // Configuração do gráfico
        const ctx = document.getElementById('vendasChart').getContext('2d');
        const vendasChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: meses,
                datasets: [
                    {
                        label: 'Valor do Mês',
                        data: valores,
                        backgroundColor: function(context) {
                            const value = context.dataset.data[context.dataIndex];
                            if (value >= 1000) {
                                return 'rgba(40, 167, 69, 0.8)';
                            } else if (value >= 500) {
                                return 'rgba(255, 193, 7, 0.8)';
                            } else if (value >= 1) {
                                return 'rgba(102, 126, 234, 0.8)';
                            } else {
                                return 'rgba(222, 226, 230, 0.8)';
                            }
                        },
                        borderColor: function(context) {
                            const value = context.dataset.data[context.dataIndex];
                            if (value >= 1000) {
                                return 'rgba(40, 167, 69, 1)';
                            } else if (value >= 500) {
                                return 'rgba(255, 193, 7, 1)';
                            } else if (value >= 1) {
                                return 'rgba(102, 126, 234, 1)';
                            } else {
                                return 'rgba(222, 226, 230, 1)';
                            }
                        },
                        borderWidth: 1,
                        borderRadius: 4,
                        barPercentage: 0.6,
                        categoryPercentage: 0.7,
                        order: 2
                    },
                    {
                        label: 'Acumulado Anual',
                        data: acumulado,
                        type: 'line',
                        borderColor: 'rgba(220, 53, 69, 0.8)',
                        backgroundColor: 'rgba(220, 53, 69, 0.1)',
                        borderWidth: 3,
                        pointBackgroundColor: 'rgba(220, 53, 69, 1)',
                        pointBorderColor: '#ffffff',
                        pointBorderWidth: 2,
                        pointRadius: 5,
                        pointHoverRadius: 7,
                        fill: true,
                        tension: 0.3,
                        order: 1,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 60,
                        bottom: 20,
                        left: 10,
                        right: 10
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 15
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += formatReal(context.parsed.y);
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Valor do Mês (R$)'
                        },
                        ticks: {
                            callback: function(value) {
                                return formatReal(value);
                            }
                        },
                        suggestedMax: suggestedMax,
                        grace: '15%'
                    },
                    y1: {
                        beginAtZero: true,
                        position: 'right',
                        title: {
                            display: true,
                            text: 'Acumulado Anual (R$)'
                        },
                        ticks: {
                            callback: function(value) {
                                return formatReal(value);
                            }
                        },
                        grid: {
                            drawOnChartArea: false
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Meses'
                        },
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45
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
                    
                    // Adicionar valores nas barras
                    chart.data.datasets[0].data.forEach(function(value, index) {
                        const meta = chart.getDatasetMeta(0);
                        const element = meta.data[index];
                        
                        if (value > 0) {
                            const barHeight = bottom - element.y;
                            const isTallBar = barHeight > 25;
                            
                            ctx.fillStyle = isTallBar ? '#2c3e50' : '#ffffff';
                            ctx.font = 'bold 10px Arial';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = isTallBar ? 'bottom' : 'middle';
                            
                            const yPosition = isTallBar ? element.y - 8 : element.y + (barHeight / 2);
                            
                            if (yPosition > top + 15) {
                                // Usar formatação curta para os labels no gráfico
                                ctx.fillText(formatShort(value), element.x, yPosition);
                            }
                        }
                    });
                    
                    // Adicionar valores na linha de acumulado
                    chart.data.datasets[1].data.forEach(function(value, index) {
                        const meta = chart.getDatasetMeta(1);
                        const element = meta.data[index];
                        
                        // Verificar se o ponto está muito próximo do topo
                        const isNearTop = element.y < top + 30;
                        
                        ctx.fillStyle = '#dc3545';
                        ctx.font = 'bold 10px Arial';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = isNearTop ? 'top' : 'bottom';
                        
                        const yPosition = isNearTop ? element.y + 12 : element.y - 10;
                        // Usar formatação curta para os labels no gráfico
                        ctx.fillText(formatShort(value), element.x, yPosition);
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
    <!-- ==================================================== -->

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<%
' Função para renderizar cada item de mês
Sub RenderMonthItem(mes, valor, acumulado)
    ' Definir classe de cor baseada no valor
    Dim barClass, valueClass
    If valor >= 1000 Then
        barClass = "month-high"
        valueClass = "value-high"
    ElseIf valor >= 500 Then
        barClass = "month-medium"
        valueClass = "value-medium"
    ElseIf valor >= 1 Then
        barClass = "month-low"
        valueClass = "value-low"
    Else
        barClass = "month-zero"
        valueClass = "value-zero"
    End If
    %>
    <div class="month-item" title="<%= meses(mes) %>: <%= FormatReal(valor) %> | Acumulado: <%= FormatReal(acumulado) %>">
        <div class="month-name"><%= Left(meses(mes), 3) %></div>
        <div class="month-bar <%= barClass %>"></div>
        <div class="month-value <%= valueClass %>"><%= FormatValor(valor) %></div>
        <div class="month-acumulado">Acum: <%= FormatValor(acumulado) %></div>
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