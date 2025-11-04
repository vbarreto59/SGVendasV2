<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% ' funcional'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<!--#include file="gestao_header.inc"-->

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Obter ano selecionado do filtro
Dim anoSelecionado
anoSelecionado = Request.QueryString("ano")
If anoSelecionado = "" Then
    anoSelecionado = Year(Date()) ' Ano atual como padrão
End If

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Array com nomes dos meses
Dim meses(12)
meses(1) = "Jan"
meses(2) = "Fev" 
meses(3) = "Mar"
meses(4) = "Abr"
meses(5) = "Mai"
meses(6) = "Jun"
meses(7) = "Jul"
meses(8) = "Ago"
meses(9) = "Set"
meses(10) = "Out"
meses(11) = "Nov"
meses(12) = "Dez"

' Arrays para armazenar totais
Dim vendasMensais(12), metasMensais(12), diferencasMensais(12), coresMensais(12)

' Inicializar arrays
For i = 1 To 12
    vendasMensais(i) = 0
    metasMensais(i) = 0
    diferencasMensais(i) = 0
    coresMensais(i) = "#28a745" ' Verde padrão
Next

' Buscar vendas por mês do ano selecionado
On Error Resume Next
Set rsVendasMensais = Server.CreateObject("ADODB.Recordset")
sqlVendas = "SELECT MesVenda, SUM(ValorUnidade) as TotalVendas " & _
            "FROM Vendas " & _
            "WHERE AnoVenda = " & anoSelecionado & " AND (Excluido <> -1 OR Excluido IS NULL) " & _
            "GROUP BY MesVenda " & _
            "ORDER BY MesVenda"
rsVendasMensais.Open sqlVendas, connSales

If Err.Number = 0 Then
    Do While Not rsVendasMensais.EOF
        mes = CInt(rsVendasMensais("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            vendasMensais(mes) = CDbl(rsVendasMensais("TotalVendas"))
        End If
        rsVendasMensais.MoveNext
    Loop
    rsVendasMensais.Close
    Set rsVendasMensais = Nothing
End If
On Error GoTo 0

' Buscar metas da empresa
On Error Resume Next
Set rsMetas = Server.CreateObject("ADODB.Recordset")
sqlMetas = "SELECT Mes, Meta FROM MetaEmpresa WHERE Ano = " & anoSelecionado & " ORDER BY Mes"
rsMetas.Open sqlMetas, connSales

If Err.Number = 0 Then
    Do While Not rsMetas.EOF
        mes = CInt(rsMetas("Mes"))
        If mes >= 1 And mes <= 12 Then
            metasMensais(mes) = CDbl(rsMetas("Meta"))
            diferencasMensais(mes) = vendasMensais(mes) - metasMensais(mes)
            ' Definir cores baseadas no desempenho
            If vendasMensais(mes) >= metasMensais(mes) Then
                coresMensais(mes) = "#007bff" ' Azul claro para meta atingida/superada
            Else
                coresMensais(mes) = "#dc3545" ' Vermelho para meta não atingida
            End If
        End If
        rsMetas.MoveNext
    Loop
    rsMetas.Close
    Set rsMetas = Nothing
End If
On Error GoTo 0

' -----------
' Buscar quantidade total de unidades vendidas no ano
Dim totalUnidades
totalUnidades = 0

Set rsUnidades = Server.CreateObject("ADODB.Recordset")
sqlUnidades = "SELECT COUNT(*) as TotalUnidades FROM Vendas WHERE AnoVenda = " & anoSelecionado & " AND (Excluido <> -1 OR Excluido IS NULL)"
rsUnidades.Open sqlUnidades, connSales

If Not rsUnidades.EOF Then
    totalUnidades = rsUnidades("TotalUnidades")
End If
rsUnidades.Close
Set rsUnidades = Nothing

' Calcular Ticket Médio
Dim ticketMedio
If totalUnidades > 0 Then
    ticketMedio = totalVendasAno / totalUnidades
Else
    ticketMedio = 0
End If

' ----------









' Buscar últimas 3 vendas com mais informações
Set rsUltimasVendas = Server.CreateObject("ADODB.Recordset")
sqlUltimasVendas = "SELECT TOP 3 " & _
                   "V.ID, " & _
                   "V.Empreend_ID, " & _
                   "V.NomeEmpreendimento, " & _
                   "V.Unidade, " & _
                   "V.ValorUnidade, " & _
                   "V.DataVenda, " & _
                   "V.Corretor, " & _
                   "V.Localidade, " & _
                   "V.MesVenda, " & _
                   "V.AnoVenda, " & _
                   "V.ComissaoPercentual, " & _
                   "V.ValorComissaoGeral, " & _
                   "V.Diretoria, " & _
                   "V.Gerencia " & _
                   "FROM Vendas V " & _
                   "WHERE V.AnoVenda = " & anoSelecionado & " AND (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                   "ORDER BY V.DataVenda DESC, V.ID DESC"

rsUltimasVendas.Open sqlUltimasVendas, connSales

' Calcular totais gerais
Dim totalVendasAno, totalMetaAno, totalDiferencaAno
totalVendasAno = 0
totalMetaAno = 0
totalDiferencaAno = 0

For i = 1 To 12
    totalVendasAno = totalVendasAno + vendasMensais(i)
    totalMetaAno = totalMetaAno + metasMensais(i)
Next
totalDiferencaAno = totalVendasAno - totalMetaAno

' Determinar a classe CSS para o card de diferença
Dim cardDiferencaClass
If totalDiferencaAno >= 0 Then
    cardDiferencaClass = "success"
Else
    cardDiferencaClass = "danger"
End If

' Determinar o ícone para a diferença
Dim iconeDiferenca
If totalDiferencaAno >= 0 Then
    iconeDiferenca = "fa-arrow-up"
Else
    iconeDiferenca = "fa-arrow-down"
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KPIs e Metas | Gestão de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --accent: #e74c3c;
            --success: #28a745;
            --warning: #fd7e14;
            --light-bg: #f8f9fa;
        }
        
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding-top: 80px;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1rem 0;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1rem 1.5rem;
            font-weight: 600;
        }
        
        .mes-card {
            transition: transform 0.2s;
            height: 100%;
        }
        
        .mes-card:hover {
            transform: translateY(-2px);
        }
        
        .filter-section {
            background: white;
            border-radius: 12px;
            padding: 1rem;
            margin-bottom: 1.5rem;
        }
        
        .ultimas-vendas {
            background: white;
            padding: 1.5rem;
            margin-top: 2rem;
        }
        
        .venda-item {
            border-left: 4px solid var(--secondary);
            padding: 1rem;
            margin-bottom: 1rem;
            background: #f8f9fa;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
        }
        
        .venda-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.5rem;
            border-bottom: 1px solid #dee2e6;
            padding-bottom: 0.5rem;
        }
        
        .venda-id {
            font-weight: bold;
            font-size: 1.1rem;
        }
        
        .venda-data {
            color: #6c757d;
            font-size: 0.85rem;
        }
        
        .venda-corretor {
            font-weight: bold;
            margin-top: 0.3rem;
        }
        
        .venda-detalhes {
            display: grid;
            grid-template-columns: 2fr 1fr 1fr;
            gap: 1rem;
            margin-top: 0.5rem;
        }
        
        .venda-empreendimento {
            font-weight: bold;
        }
        
        .venda-valor {
            text-align: right;
            font-weight: bold;
            color: #28a745;
        }
        
        .venda-comissao {
            text-align: right;
            color: #6c757d;
            font-size: 0.85rem;
        }
        
        .btn-refresh {
            background-color: var(--warning);
            border-color: var(--warning);
            color: white;
        }
        
        .chart-container {
            position: relative;
            height: 400px;
            width: 100%;
        }
        
        .venda-periodo {
            background: #e9ecef;
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.8rem;
            display: inline-block;
            margin-top: 0.3rem;
        }
    </style>


<style>
.mes-card {
    transition: transform 0.2s;
    height: 100%;
}

.mes-card.atingiu-meta {
    background-color: #e3f2fd; /* Azul claro para metas atingidas */
    border-left: 4px solid #007bff;
}

.mes-card.nao-atingiu-meta {
    background-color: #ffebee; /* Vermelho claro para metas não atingidas */
    border-left: 4px solid #dc3545;
}

.mes-card:hover {
    transform: translateY(-2px);
}
</style>    
</head>
<body>
    <header class="app-header">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 class="app-title"><i class="fas fa-chart-line me-2"></i> KPIs e Metas</h1>
                </div>
                <div class="col-md-6 text-end">
                    <a href="gestao_vendas.asp" class="btn btn-light btn-sm">
                        <i class="fas fa-arrow-left me-1"></i>Voltar para Vendas
                    </a>
                </div>
            </div>
        </div>
    </header>

    <div class="container-fluid main-content">
        <!-- Filtro de Ano -->
        <div class="filter-section">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h5 class="mb-0"><i class="fas fa-filter me-2"></i>Filtros</h5>
                </div>
                <div class="col-md-6">
                    <form method="GET" action="" class="d-flex gap-2">
                        <select name="ano" class="form-select" onchange="this.form.submit()">
                            <%
                            ' Opção para 2025
                            If anoSelecionado = "2025" Then
                                Response.Write "<option value='2025' selected>2025</option>"
                            Else
                                Response.Write "<option value='2025'>2025</option>"
                            End If
                            
                            ' Opção para 2026
                            If anoSelecionado = "2026" Then
                                Response.Write "<option value='2026' selected>2026</option>"
                            Else
                                Response.Write "<option value='2026'>2026</option>"
                            End If
                            %>
                        </select>
                        <button type="button" class="btn btn-refresh" onclick="location.reload()">
                            <i class="fas fa-sync-alt"></i>
                        </button>
                    </form>
                </div>
            </div>
        </div>

        <!-- Gráfico de Metas vs Vendas -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Metas vs Vendas - <%= anoSelecionado %></h5>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="graficoMetasVendas"></canvas>
                </div>
            </div>
        </div>

        <!-- Cards dos Meses ###################### -->

<!-- Cards dos Meses -->
<div class="card">
    <div class="card-header">
        <h5 class="mb-0"><i class="fas fa-calendar-alt me-2"></i>Desempenho Mensal - <%= anoSelecionado %></h5>
    </div>
    <div class="card-body">
        <div class="row">
            <%
            For i = 1 To 12
                Dim badgeClass, iconeMeta, borderClass, classeFundo
                
                If diferencasMensais(i) >= 0 Then
                    badgeClass = "bg-success"
                    iconeMeta = "fa-arrow-up"
                    borderClass = "success"
                    classeFundo = "atingiu-meta"
                Else
                    badgeClass = "bg-danger"
                    iconeMeta = "fa-arrow-down"
                    borderClass = "danger"
                    classeFundo = "nao-atingiu-meta"
                End If
            %>
            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6 mb-3">
                <div class="card mes-card <%= classeFundo %> border-<%= borderClass %>">
                    <div class="card-body text-center p-2">
                        <h6 class="card-title fw-bold"><%= meses(i) %></h6>
                        <div class="mb-1">
                            <small class="text-muted">Vendas:</small>
                            <div class="fw-bold">R$ <%= FormatNumber(vendasMensais(i), 2) %></div>
                        </div>
                        <div class="mb-1">
                            <small class="text-muted">Meta:</small>
                            <div>R$ <%= FormatNumber(metasMensais(i), 2) %></div>
                        </div>
                        <div class="mt-2">
                            <span class="badge <%= badgeClass %>">
                                <i class="fas <%= iconeMeta %> me-1"></i>
                                R$ <%= FormatNumber(Abs(diferencasMensais(i)), 2) %>
                            </span>
                        </div>
                    </div>
                </div>
            </div>
            <% Next %>
        </div>
    </div>
</div>
        <!-- ######################################### -->

<!-- Totais do Ano -->
<div class="row">
    <div class="col-md-2">
        <div class="card bg-primary text-white">
            <div class="card-body text-center">
                <h6 class="card-title">Total Vendido</h6>
                <h4>R$ <%= FormatNumber(totalVendasAno, 2) %></h4>
            </div>
        </div>
    </div>
    <div class="col-md-2">
        <div class="card bg-warning text-white">
            <div class="card-body text-center">
                <h6 class="card-title">Meta do Ano</h6>
                <h4>R$ <%= FormatNumber(totalMetaAno, 2) %></h4>
            </div>
        </div>
    </div>
    <div class="col-md-2">
        <div class="card bg-<%= cardDiferencaClass %> text-white">
            <div class="card-body text-center">
                <h6 class="card-title">Diferença</h6>
                <h4>
                    R$ <%= FormatNumber(Abs(totalDiferencaAno), 2) %>
                    <i class="fas <%= iconeDiferenca %>"></i>
                </h4>
            </div>
        </div>
    </div>
    <div class="col-md-2">
        <div class="card bg-info text-white">
            <div class="card-body text-center">
                <h6 class="card-title">Unidades Vendidas</h6>
                <h4><%= totalUnidades %></h4>
                <small>unidades</small>
            </div>
        </div>
    </div>
    <div class="col-md-2">
        <div class="card bg-success text-white">
            <div class="card-body text-center">
                <%
                  ticketMedio = totalVendasAno/totalUnidades    
                %>
                <h6 class="card-title">Ticket Médio</h6>
                <h4>R$ <%= FormatNumber(ticketMedio, 2) %></h4>
                <small>por unidade</small>
            </div>
        </div>
    </div>
    <div class="col-md-2">
        <div class="card bg-secondary text-white">
            <div class="card-body text-center">
                <h6 class="card-title">Desempenho</h6>
                <h4>
                    <%
                    Dim percentualDesempenho
                    If totalMetaAno > 0 Then
                        percentualDesempenho = (totalVendasAno / totalMetaAno) * 100
                    Else
                        percentualDesempenho = 0
                    End If
                    %>
                    <%= FormatNumber(percentualDesempenho, 1) %>%
                </h4>
                <small>da meta</small>
            </div>
        </div>
    </div>
</div>

<!-- Últimas 3 Vendas -->
<div class="ultimas-vendas border border-2 border-dark rounded-3 p-4 bg-white">
    <h5 class="mb-4"><i class="fas fa-clock me-2"></i>Últimas 3 Vendas Realizadas</h5>
    <div class="row">
        <%
        Dim contadorVenda
        contadorVenda = 0
        
        If Not rsUltimasVendas.EOF Then
            Do While Not rsUltimasVendas.EOF
                contadorVenda = contadorVenda + 1
                ' Formatar período (Ano-Mês)
                periodo = rsUltimasVendas("AnoVenda") & "-" & Right("0" & rsUltimasVendas("MesVenda"), 2)
                
                ' Definir cores diferentes para cada card
                Dim cardColor, borderColor
                Select Case contadorVenda
                    Case 1
                        cardColor = "bg-primary"
                        borderColor = "border-primary"
                    Case 2
                        cardColor = "bg-success" 
                        borderColor = "border-success"
                    Case 3
                        cardColor = "bg-warning"
                        borderColor = "border-warning"
                End Select
        %>
        <div class="col-md-4 mb-3 border border-2 border-dark rounded">

            <div class="card h-100 <%= borderColor %> shadow-sm">
                <div class="card-header <%= cardColor %> text-white py-2">
                    <div class="d-flex justify-content-between align-items-center">
                        <strong class="small">Venda #<%= contadorVenda %></strong>
                        <span class="badge bg-light text-dark"><%= periodo %></span>
                    </div>
                </div>
                <div class="card-body p-3">
                    <!-- ID e Empreendimento -->
                    <div class="mb-2">
                        <h6 class="card-title text-dark mb-1">
                            <strong><%= rsUltimasVendas("ID") %>-<%= UCase(rsUltimasVendas("NomeEmpreendimento")) %></strong>
                        </h6>
                      <div class="small text-dark">
                            <i class="fas fa-map-marker-alt me-1"></i>
                            <strong><%= UCase(rsUltimasVendas("Localidade")) %></strong>
                        </div>
                        <small class="text-muted">
                            <i class="fas fa-calendar me-1"></i>
                            <%= FormatDateTime(rsUltimasVendas("DataVenda"), 2) %>
                        </small>
                    </div>
                    
                    <!-- Diretoria e Gerência -->
                    <div class="mb-2">
                        <div class="row small">
                            <div class="col-6">
                                <div class="text-center p-1 bg-info bg-opacity-10 rounded">
                                    <i class="fas fa-building me-1"></i>
                                    <strong>Diretoria</strong>
                                    <div class="text-dark"><%= rsUltimasVendas("Diretoria") %></div>
                                </div>
                            </div>
                            <div class="col-6">
                                <div class="text-center p-1 bg-warning bg-opacity-10 rounded">
                                    <i class="fas fa-users me-1"></i>
                                    <strong>Gerência</strong>
                                    <div class="text-dark"><%= rsUltimasVendas("Gerencia") %></div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Localidade e Corretor -->
                    <div class="mb-3 p-2 bg-light rounded">

                        <div class="small text-muted">
                            <i class="fas fa-user me-1"></i>
                            <%= UCase(rsUltimasVendas("Corretor")) %>
                        </div>
                    </div>
                    
                    <!-- Valores -->
                    <div class="row text-center mb-2">
                        <div class="col-6">
                            <div class="border-end">
                                <div class="text-success fw-bold">R$ <%= FormatNumber(rsUltimasVendas("ValorUnidade"), 2) %></div>
                                <small class="text-muted">Valor</small>
                            </div>
                        </div>
                        <div class="col-6">
                            <div class="text-info fw-bold">R$ <%= FormatNumber(rsUltimasVendas("ValorComissaoGeral"), 2) %></div>
                            <small class="text-muted">Comissão</small>
                        </div>
                    </div>
                    
                    <!-- Percentual -->
                    <div class="text-center">
                        <span class="badge <%= cardColor %>">
                            <i class="fas fa-percentage me-1"></i>
                            <%= rsUltimasVendas("ComissaoPercentual") %>% Comissão
                        </span>
                    </div>
                </div>
            </div>
        </div>

        <%
                rsUltimasVendas.MoveNext
            Loop
        Else
        %>
        <div class="col-12">
            <div class="alert alert-info text-center">
                <i class="fas fa-info-circle me-2"></i>Nenhuma venda encontrada para o ano de <%= anoSelecionado %>.
            </div>
        </div>
        <%
        End If
        rsUltimasVendas.Close
        Set rsUltimasVendas = Nothing
        %>
    </div>
</div>
<br>



    </div>

    <script>
    // Dados para o gráfico
    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    const vendas = [<%= Join(vendasMensais, ",") %>];
    const metas = [<%= Join(metasMensais, ",") %>];
    const coresVendas = ['<%= Join(coresMensais, "','") %>'];

    // Configuração do gráfico
    const ctx = document.getElementById('graficoMetasVendas').getContext('2d');
    const graficoMetasVendas = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: meses,
            datasets: [
                {
                    label: 'Vendas',
                    data: vendas,
                    backgroundColor: coresVendas,
                    borderColor: coresVendas,
                    borderWidth: 1,
                    barPercentage: 0.6,
                },
                {
                    label: 'Meta',
                    data: metas,
                    type: 'bar',
                    backgroundColor: 'rgba(253, 126, 20, 0.7)',
                    borderColor: 'rgba(253, 126, 20, 1)',
                    borderWidth: 1,
                    barPercentage: 0.6,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Valor (R$)'
                    },
                    ticks: {
                        callback: function(value) {
                            return 'R$ ' + value.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                        }
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Meses'
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            label += 'R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                            return label;
                        }
                    }
                }
            }
        }
    });

    // Função para atualizar a página mantendo o filtro
    function atualizarPagina() {
        const anoSelecionado = document.querySelector('select[name="ano"]').value;
        window.location.href = 'gestao_vendas_metas.asp?ano=' + anoSelecionado;
    }

    // Atualizar a página a cada 30 segundos
    setInterval(atualizarPagina, 60000);
    </script>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<%
' Fechar conexões
If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>