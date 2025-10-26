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
' OBTER DADOS DOS CORRETORES
' ===============================================

Dim sqlCorretores, rsCorretores

' Consulta modificada para incluir CorretorId numérico
sqlCorretores = "SELECT " & _
                "Corretor as NomeCorretor, " & _
                "CorretorId, " & _
                "COUNT(*) as QtdVendas, " & _
                "SUM(ValorUnidade) as VGVPeriodo, " & _
                "SUM(ValorCorretor) as ComissaoPeriodo, " & _
                "MAX(MesVenda) as UltimoMes, " & _
                "MAX(AnoVenda) as UltimoAno " & _
                "FROM Vendas " & _
                "WHERE Excluido = 0 " & _
                "AND CorretorId IS NOT NULL " & _
                "GROUP BY Corretor, CorretorId " & _
                "ORDER BY SUM(ValorUnidade) DESC"

Set rsCorretores = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rsCorretores.Open sqlCorretores, connSales

If Err.Number <> 0 Then
    Response.Write "Erro na consulta de corretores: " & Err.Description & "<br>"
    Response.Write "SQL: " & Server.HTMLEncode(sqlCorretores)
    Response.End
End If
On Error GoTo 0

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

' Verificar se há dados
Dim temDados
temDados = Not rsCorretores.EOF
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Listagem de Corretores</title>
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
            color: black;
            padding: 15px;
            margin-top: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
        }
        .container-fluid {
            max-width: 1800px;
            margin: 0 auto;
        }
        .table-responsive {
            background-color: white;
            border-radius: 8px;
        }
        .table th {
            background-color: #800000;
            color: white;
            position: sticky;
            top: 0;
        }
        .text-right-v { text-align: right; }
        .text-center-v { text-align: center; }
        .text-left-v { text-align: left; }
        .corretor-header {
            background-color: #800000;
            color: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .badge-vendas {
            background-color: #28a745;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8rem;
        }
        .badge-dias {
            background-color: #17a2b8;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8rem;
        }
        .table tbody tr:hover {
            background-color: #f5f5f5;
        }
        .alert-info {
            background-color: #d1ecf1;
            border-color: #bee5eb;
            color: #0c5460;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .btn-detalhes {
            background-color: #800000;
            color: white;
            border: none;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8rem;
            transition: background-color 0.3s;
        }
        .btn-detalhes:hover {
            background-color: #600000;
            color: white;
        }
        .btn-action {
            min-width: 100px;
        }
        .user-id {
            font-size: 0.7rem;
            color: #6c757d;
            display: block;
            margin-top: 2px;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-users"></i> SGVendas - Listagem de Corretores
        </h2>
        
        <div class="corretor-header">
            <div class="row">
                <div class="col-md-12">
                    <h4><i class="fas fa-chart-bar"></i> Desempenho dos Corretores</h4>
                    <p class="mb-0">Período Completo - Agrupado por Corretor</p>
                </div>
            </div>
        </div>

        <% If Not temDados Then %>
        <div class="alert-info">
            <h4><i class="fas fa-info-circle"></i> Nenhum Corretor Encontrado</h4>
            <p>Não foram encontrados corretores com vendas na base de dados.</p>
            <p><strong>SQL executada:</strong> <%= Server.HTMLEncode(sqlCorretores) %></p>
        </div>
        <% Else %>

        <!-- Resumo Geral -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="card text-white bg-primary mb-3">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-users"></i> Total Corretores</h5>
                        <%
                        Dim totalCorretores
                        totalCorretores = 0
                        If Not rsCorretores.EOF Then
                            rsCorretores.MoveFirst
                            Do While Not rsCorretores.EOF
                                totalCorretores = totalCorretores + 1
                                rsCorretores.MoveNext
                            Loop
                            rsCorretores.MoveFirst
                        End If
                        %>
                        <p class="card-text display-6"><%= totalCorretores %></p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-white bg-success mb-3">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-handshake"></i> VGV Total</h5>
                        <%
                        Dim vgvTotal
                        vgvTotal = 0
                        If Not rsCorretores.EOF Then
                            Do While Not rsCorretores.EOF
                                vgvTotal = vgvTotal + CDbl(rsCorretores("VGVPeriodo"))
                                rsCorretores.MoveNext
                            Loop
                            rsCorretores.MoveFirst
                        End If
                        %>
                        <p class="card-text">R$ <%= FormatNumber(vgvTotal, 2) %></p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-white bg-warning mb-3">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-money-bill-wave"></i> Comissão Total</h5>
                        <%
                        Dim comissaoTotal
                        comissaoTotal = 0
                        If Not rsCorretores.EOF Then
                            Do While Not rsCorretores.EOF
                                comissaoTotal = comissaoTotal + CDbl(rsCorretores("ComissaoPeriodo"))
                                rsCorretores.MoveNext
                            Loop
                            rsCorretores.MoveFirst
                        End If
                        %>
                        <p class="card-text">R$ <%= FormatNumber(comissaoTotal, 2) %></p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card text-white bg-info mb-3">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-home"></i> Vendas Total</h5>
                        <%
                        Dim vendasTotal
                        vendasTotal = 0
                        If Not rsCorretores.EOF Then
                            Do While Not rsCorretores.EOF
                                vendasTotal = vendasTotal + CLng(rsCorretores("QtdVendas"))
                                rsCorretores.MoveNext
                            Loop
                            rsCorretores.MoveFirst
                        End If
                        %>
                        <p class="card-text"><%= vendasTotal %></p>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tabela de Corretores -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">Listagem de Corretores - Consolidado</h3>
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th class="text-left-v">Nome do Corretor</th>
                            <th class="text-center-v">QTD Vendas</th>
                            <th class="text-right-v">VGV (R$)</th>
                            <th class="text-right-v">Comissão (R$)</th>
                            <th class="text-center-v">Último Mês</th>
                            <th class="text-center-v">1 Venda a Cada</th>
                            <th class="text-right-v">Ticket Médio (R$)</th>
                            <th class="text-center-v">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Do While Not rsCorretores.EOF
                            Dim nomeCorretor, corretorId, qtdVendas, vgvPeriodo, comissaoPeriodo
                            Dim ultimoMes, ultimoAno, mediaDiasVenda, mesComissaoFormatado, ticketMedio
                            
                            nomeCorretor = rsCorretores("NomeCorretor")
                            corretorId = rsCorretores("CorretorId")
                            qtdVendas = rsCorretores("QtdVendas")
                            vgvPeriodo = rsCorretores("VGVPeriodo")
                            comissaoPeriodo = rsCorretores("ComissaoPeriodo")
                            ultimoMes = rsCorretores("UltimoMes")
                            ultimoAno = rsCorretores("UltimoAno")
                            
                            ' Calcular média de dias por venda
                            If qtdVendas > 0 Then
                                mediaDiasVenda = FormatNumber(365 / qtdVendas, 1)
                                ticketMedio = vgvPeriodo / qtdVendas
                            Else
                                mediaDiasVenda = "N/A"
                                ticketMedio = 0
                            End If
                            
                            ' Formatar mês de comissão
                            If Not IsNull(ultimoMes) And Not IsNull(ultimoAno) Then
                                mesComissaoFormatado = arrMesesNome(ultimoMes) & "/" & ultimoAno
                            Else
                                mesComissaoFormatado = "N/A"
                            End If
                    %>
                    <tr>
                        <td class="text-left-v">
                            <strong><%= nomeCorretor %></strong>
                            <% If Not IsNull(corretorId) Then %>
                                <span class="user-id">ID: <%= corretorId %></span>
                            <% End If %>
                        </td>
                        <td class="text-center-v">
                            <span class="badge-vendas"><%= qtdVendas %></span>
                        </td>
                        <td class="text-right-v">
                            <strong>R$ <%= FormatNumber(vgvPeriodo, 2) %></strong>
                        </td>
                        <td class="text-right-v">
                            <strong>R$ <%= FormatNumber(comissaoPeriodo, 2) %></strong>
                        </td>
                        <td class="text-center-v">
                            <%= mesComissaoFormatado %>
                        </td>
                        <td class="text-center-v">
                            <% If mediaDiasVenda <> "N/A" Then %>
                                <span class="badge-dias"><%= mediaDiasVenda %> dias</span>
                            <% Else %>
                                <span class="badge bg-secondary">N/A</span>
                            <% End If %>
                        </td>
                        <td class="text-right-v">
                            R$ <%= FormatNumber(ticketMedio, 2) %>
                        </td>
                        <td class="text-center-v">
                            <button class="btn btn-detalhes btn-action" 
                                    onclick="verHistoricoPorId(<%= corretorId %>, '<%= Server.URLEncode(nomeCorretor) %>')"
                                    title="Ver histórico detalhado de vendas">
                                <i class="fas fa-search"></i> Detalhes
                            </button>
                        </td>
                    </tr>
                    <%
                            rsCorretores.MoveNext
                        Loop
                    %>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Legenda e Informações -->
        <div class="row mt-4">
            <div class="col-md-12">
                <div class="card-kpi">
                    <h5 class="text-dark"><i class="fas fa-info-circle"></i> Informações da Listagem</h5>
                    <div class="row">
                        <div class="col-md-3">
                            <p><strong>Período:</strong> Todos os dados disponíveis</p>
                            <p><strong>QTD Vendas:</strong> Número total de vendas</p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>VGV:</strong> Valor Geral de Vendas (somatório)</p>
                            <p><strong>Comissão:</strong> Valor total de comissões</p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>Último Mês:</strong> Mês mais recente com atividade</p>
                            <p><strong>1 Venda a Cada:</strong> Média de dias entre vendas</p>
                        </div>
                        <div class="col-md-3">
                            <p><strong>Ticket Médio:</strong> Valor médio por venda (VGV/Qtd)</p>
                            <p><strong>Ações:</strong> Clique em "Detalhes" para ver o histórico completo</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <% End If %>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Ordenação simples da tabela
        document.addEventListener('DOMContentLoaded', function() {
            const table = document.querySelector('table');
            if (table) {
                const headers = table.querySelectorAll('th');
                
                headers.forEach((header, index) => {
                    // Não aplicar ordenação na coluna de Ações (última coluna)
                    if (index < headers.length - 1) {
                        header.style.cursor = 'pointer';
                        header.addEventListener('click', () => {
                            sortTable(index);
                        });
                    }
                });
            }
        });

        function sortTable(columnIndex) {
            const table = document.querySelector('table');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            
            rows.sort((a, b) => {
                const aText = a.cells[columnIndex].textContent.trim();
                const bText = b.cells[columnIndex].textContent.trim();
                
                // Verificar se é numérico (remover R$ e pontos)
                const aNum = parseFloat(aText.replace('R$', '').replace('.', '').replace(',', '.'));
                const bNum = parseFloat(bText.replace('R$', '').replace('.', '').replace(',', '.'));
                
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return bNum - aNum; // Ordenar do maior para o menor
                } else {
                    return aText.localeCompare(bText);
                }
            });
            
            // Reordenar as linhas
            rows.forEach(row => tbody.appendChild(row));
        }

        // Função para ver histórico detalhado usando CorretorId numérico
        function verHistoricoPorId(corretorId, corretorNome) {
            // Decodificar o nome do corretor (pode conter caracteres especiais)
            const corretorDecoded = decodeURIComponent(corretorNome);
            
            // Obter o ano atual para o filtro
            const anoAtual = new Date().getFullYear();
            
            // Redirecionar para a página de relatório do corretor usando CorretorId numérico
// Abre a página de relatório do corretor em uma NOVA ABA/JANELA
window.open('gestao_corretores_relat1.asp?corretorid=' + corretorId + '&corretor=' + encodeURIComponent(corretorDecoded) + '&ano=' + anoAtual, '_blank');
        }

        // Função alternativa para abrir em nova janela/popup
        function verHistoricoPopupPorId(corretorId, corretorNome) {
            const corretorDecoded = decodeURIComponent(corretorNome);
            const anoAtual = new Date().getFullYear();
            const url = 'gestao_corretores_relat1.asp?corretorid=' + corretorId + '&corretor=' + encodeURIComponent(corretorDecoded) + '&ano=' + anoAtual;
            
            window.open(url, '_blank', 'width=1200,height=800,scrollbars=yes');
        }
    </script>
</body>
</html>

<%
' Fechar recordset e conexão
If Not rsCorretores Is Nothing Then
    If rsCorretores.State = 1 Then rsCorretores.Close
    Set rsCorretores = Nothing
End If

If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>