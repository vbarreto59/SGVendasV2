<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% ' funcional cores com problemas'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<!--#include file="gestao_header.inc"-->


<%
' ===========================================================
' Função substituta para Nz() do Access
' ===========================================================
Function Nz(valor, opcional)
    If IsNull(valor) Or IsEmpty(valor) Or valor = "" Then
        If IsMissing(opcional) Then
            Nz = 0
        Else
            Nz = opcional
        End If
    Else
        Nz = valor
    End If
End Function
%>



<%
' ==============================================================================
' INCLUSÕES E CONFIGURAÇÕES
' ==============================================================================

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

' Array com nomes dos meses (utilizado no VBScript e no JavaScript)
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

' Arrays para armazenar totais e cores
Dim vendasMensais(12), metasMensais(12), diferencasMensais(12), coresMensais(12)

' Inicializar arrays (Jan a Dez = 1 a 12)
For i = 1 To 12
    vendasMensais(i) = 0
    metasMensais(i) = 0
    diferencasMensais(i) = 0
    coresMensais(i) = "#6c757d" ' Cinza padrão (sem vendas/meta)
Next

' ==============================================================================
' BUSCAR VENDAS
' ==============================================================================
Set rsVendasMensais = Server.CreateObject("ADODB.Recordset")
sqlVendas = "SELECT MesVenda, SUM(ValorUnidade) as TotalVendas " & _
            "FROM Vendas " & _
            "WHERE AnoVenda = " & anoSelecionado & " AND (Excluido <> -1 OR Excluido IS NULL) " & _
            "GROUP BY MesVenda " & _
            "ORDER BY MesVenda"

rsVendasMensais.Open sqlVendas, connSales

If Not rsVendasMensais.EOF Then
    Do While Not rsVendasMensais.EOF
        mes = CInt(rsVendasMensais("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            If Not IsNull(rsVendasMensais("TotalVendas")) Then
                vendasMensais(mes) = CDbl(rsVendasMensais("TotalVendas"))
            Else
                vendasMensais(mes) = 0
            End If
        End If
        rsVendasMensais.MoveNext
    Loop
End If
rsVendasMensais.Close
Set rsVendasMensais = Nothing

' ===========================================================
' Definir cores com base nas metas e vendas
' ===========================================================
Set rsMetas = Server.CreateObject("ADODB.Recordset")
sqlMetas = "SELECT Mes, Meta FROM MetaEmpresa WHERE Ano = " & anoSelecionado & " ORDER BY Mes"
rsMetas.Open sqlMetas, connSales

If Not rsMetas.EOF Then
    Do While Not rsMetas.EOF
        mes = CInt(rsMetas("Mes"))
        If mes >= 1 And mes <= 12 Then
            metasMensais(mes) = CDbl(Nz(rsMetas("Meta"), 0))
            diferencasMensais(mes) = vendasMensais(mes) - metasMensais(mes)
            
            If vendasMensais(mes) > metasMensais(mes) Then
                coresMensais(mes) = "#007bff" ' Azul
            ElseIf vendasMensais(mes) < metasMensais(mes) Then
                coresMensais(mes) = "#dc3545" ' Vermelho
            Else
                coresMensais(mes) = "#007bff" ' Azul
            End If
        End If
        rsMetas.MoveNext
    Loop
End If
rsMetas.Close
Set rsMetas = Nothing

' ===========================================================
' AQUI entra o novo código (antes de fechar o ASP)
' ===========================================================
Dim strCoresJS
strCoresJS = ""
For i = 1 To 12
    If i > 1 Then strCoresJS = strCoresJS & ","
    strCoresJS = strCoresJS & "'" & coresMensais(i) & "'"
Next

' ==============================================================================
' CÁLCULOS GERAIS (Total Unidades, Últimas Vendas, Totais Anuais)
' ==============================================================================

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

' Buscar últimas 3 vendas com mais informações
Set rsUltimasVendas = Server.CreateObject("ADODB.Recordset")
sqlUltimasVendas = "SELECT TOP 3 V.ID, V.NomeEmpreendimento, V.Unidade, V.ValorUnidade, V.DataVenda, V.Corretor, V.Localidade, V.MesVenda, V.AnoVenda, V.ComissaoPercentual, V.ValorComissaoGeral, V.Diretoria, V.Gerencia " & _
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

' Calcular Ticket Médio
Dim ticketMedio
If totalUnidades > 0 Then
    ticketMedio = totalVendasAno / totalUnidades
Else
    ticketMedio = 0
End If

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

' ==============================================================================
' CORREÇÃO CRÍTICA PARA O GRÁFICO (VBScript -> JavaScript)
' Forçar uso de PONTO (.) como separador decimal para o Chart.js
' ==============================================================================
Dim strVendasJS, strMetasJS, valorVenda, valorMeta
strVendasJS = ""
strMetasJS = ""

For i = 1 To 12
    ' Formata o número com 2 casas decimais e remove agrupamento de milhares
    ' Em seguida, substitui a vírgula (decimal VBScript) por ponto (decimal JS)
    valorVenda = Replace(FormatNumber(vendasMensais(i), 2, , , False), ",", ".")
    valorMeta = Replace(FormatNumber(metasMensais(i), 2, , , False), ",", ".")
    
    If i > 1 Then
        strVendasJS = strVendasJS & ","
        strMetasJS = strMetasJS & ","
    End If
    
    strVendasJS = strVendasJS & valorVenda
    strMetasJS = strMetasJS & valorMeta
Next
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
        
        /* CORREÇÃO DEFINITIVA: Cores INTUITIVAS para os cards */
        .mes-card.atingiu-meta {
            background-color: #e8f5e8; /* Verde claro para metas ATINGIDAS/SUPERADAS */
            border-left: 4px solid #28a745;
        }
        
        .mes-card.nao-atingiu-meta {
            background-color: #ffebee; /* Vermelho claro para metas NÃO ATINGIDAS */
            border-left: 4px solid #dc3545;
        }
        
        .mes-card.sem-meta {
            background-color: #f8f9fa; /* Cinza claro para meses sem meta */
            border-left: 4px solid #6c757d;
        }
        
        .mes-card.com-vendas {
            background-color: #e3f2fd; /* Azul claro para meses com vendas sem meta */
            border-left: 4px solid #17a2b8;
        }

        /* Legenda do gráfico */
        .legenda-grafico {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 15px;
            flex-wrap: wrap;
        }
        
        .item-legenda {
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 0.9rem;
        }
        
        .cor-legenda {
            width: 15px;
            height: 15px;
            border-radius: 3px;
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
        <div class="filter-section">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h5 class="mb-0"><i class="fas fa-filter me-2"></i>Filtros</h5>
                </div>
                <div class="col-md-6">
                    <form method="GET" action="" class="d-flex gap-2">
                        <select name="ano" class="form-select" onchange="this.form.submit()">
                            <%
                            ' Opções de ano
                            Dim ano
                            For ano = 2025 To Year(Date()) + 1 ' Exemplo: 2024 até 2 anos à frente
                                Response.Write "<option value='" & ano & "'"
                                If CStr(ano) = anoSelecionado Then
                                    Response.Write " selected"
                                End If
                                Response.Write ">" & ano & "</option>"
                            Next
                            %>
                        </select>
                        <button type="button" class="btn btn-refresh" onclick="location.reload()">
                            <i class="fas fa-sync-alt"></i>
                        </button>
                    </form>
                </div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Metas vs Vendas - <%= anoSelecionado %></h5>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="graficoMetasVendas"></canvas>
                </div>
                <!-- Legenda do gráfico -->
                <div class="legenda-grafico">
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #28a745;"></div>
                        <span>Meta Atingida/Superada</span>
                    </div>
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #dc3545;"></div>
                        <span>Meta Não Atingida</span>
                    </div>
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #17a2b8;"></div>
                        <span>Vendas sem Meta</span>
                    </div>
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #6c757d;"></div>
                        <span>Sem Vendas/Meta</span>
                    </div>
                </div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-calendar-alt me-2"></i>Desempenho Mensal - <%= anoSelecionado %></h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <%
                    For i = 1 To 12
                        Dim badgeClass, iconeMeta, borderClass, classeFundo, textoDiferenca
                       
                        ' CORREÇÃO DEFINITIVA: Lógica INTUITIVA para os CARDS
                        If metasMensais(i) > 0 Then
                            If diferencasMensais(i) >= 0 Then
                                ' Meta ATINGIDA/SUPERADA (BOM)
                                badgeClass = "bg-success" ' Verde para badge
                                iconeMeta = "fa-check"
                                borderClass = "success"
                                classeFundo = "atingiu-meta"
                                textoDiferenca = "R$ " & FormatNumber(Abs(diferencasMensais(i)), 2)
                            Else
                                ' Meta NÃO ATINGIDA (RUIM)
                                badgeClass = "bg-danger" ' Vermelho para badge
                                iconeMeta = "fa-times"
                                borderClass = "danger"
                                classeFundo = "nao-atingiu-meta"
                                textoDiferenca = "R$ " & FormatNumber(Abs(diferencasMensais(i)), 2)
                            End If
                        Else
                            ' SEM META
                            If vendasMensais(i) > 0 Then
                                ' Com vendas mas sem meta
                                badgeClass = "bg-info" ' Azul para badge
                                iconeMeta = "fa-chart-line"
                                borderClass = "info"
                                classeFundo = "com-vendas"
                                textoDiferenca = "Com Vendas"
                            Else
                                ' Sem vendas e sem meta
                                badgeClass = "bg-secondary" ' Cinza para badge
                                iconeMeta = "fa-minus"
                                borderClass = "secondary"
                                classeFundo = "sem-meta"
                                textoDiferenca = "Sem Meta"
                            End If
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
                                        <%= textoDiferenca %>
                                    </span>
                                </div>
                            </div>
                        </div>
                    </div>
                    <% Next %>
                </div>
            </div>
        </div>

        <!-- Cards de totais -->
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

        <!-- Resto do código das últimas vendas permanece igual -->
    </div>

    <script>
    // Dados para o gráfico (agora com o formato decimal corrigido)
    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    
    // Utilizando as strings com ponto decimal geradas no VBScript
    const vendas = [<%= strVendasJS %>];
    const metas = [<%= strMetasJS %>];
    
   const coresVendas = [<%= strCoresJS %>];

    // Debug no console para verificar os valores injetados
    console.log('Vendas (JS):', vendas);
    console.log('Metas (JS):', metas);
    console.log('Cores (JS):', coresVendas);

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
                    backgroundColor: 'rgba(157, 209, 169, 0.7)',
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

    // Atualizar a página a cada 60 segundos
    setInterval(atualizarPagina, 60000);
    </script>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<%
' Fechar conexões
If Not rsUltimasVendas Is Nothing Then
    If Not rsUltimasVendas.State = 0 Then rsUltimasVendas.Close
End If

If Not rsMetas Is Nothing Then
    If Not rsMetas.State = 0 Then rsMetas.Close
End If

If Not rsVendasMensais Is Nothing Then
    If Not rsVendasMensais.State = 0 Then rsVendasMensais.Close
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>