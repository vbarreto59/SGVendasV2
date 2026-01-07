<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' Função auxiliar
Function SafeCDbl(value)
    If IsNull(value) Or IsEmpty(value) Then
        SafeCDbl = 0
    ElseIf IsNumeric(value) Then
        SafeCDbl = CDbl(value)
    Else
        SafeCDbl = 0 
    End If
End Function

' Obter diretoria da sessão do usuário
Dim diretoriaSessao
diretoriaSessao = Session("Dir_NomeDiretoria")

' Verificar se a sessão tem diretoria
If diretoriaSessao = "" Or IsNull(diretoriaSessao) Then
    Response.Write "<div class='alert alert-danger'>ERRO: Diretoria não encontrada na sessão. Faça login novamente.</div>"
    Response.End
End If

' 1. Conexão com o Banco de Dados Access MDB
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open StrConnSales
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>ERRO DE CONEXÃO COM O BANCO ACCESS: " & Err.Description & "</div>"
    Response.End
End If
On Error GoTo 0

' 2. Consulta SQL principal adaptada para Access MDB - FILTRANDO PELA DIRETORIA DA SESSÃO
Dim sql, rs, hasRecords
sql = "SELECT " & _
      "VT.ID_Venda, " & _
      "VT.UserID, " & _
      "VT.Nome, " & _
      "VT.Diretoria, " & _
      "VT.Diretoria, " & _
      "VT.Gerencia, " & _
      "SUM(VT.VTotal) AS TotalComissaoDevida, " & _
      "SUM(VT.VBruto) AS VBrutoConsolidado, " & _
      "(SELECT SUM(ValorPago) FROM PAGAMENTOS_COMISSOES " & _
      " WHERE UsuariosUserId = VT.UserID AND ID_Venda = VT.ID_Venda) AS ValorPagoPorVenda " & _
      "FROM VENDA_TEMP AS VT " & _
      "WHERE VT.Diretoria = '" & Replace(diretoriaSessao, "'", "''") & "' " & _
      "GROUP BY VT.ID_Venda, VT.UserID, VT.Nome, VT.Diretoria, VT.Gerencia " & _
      "ORDER BY VT.ID_Venda DESC, VT.Nome"

Set rs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rs.Open sql, conn
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>ERRO NA CONSULTA SQL (ACCESS MDB): " & Err.Description & "</div>"
    Response.Write "<div class='alert alert-info'>SQL: " & Server.HTMLEncode(sql) & "</div>"
    hasRecords = False
Else
    hasRecords = Not rs.EOF
End If
On Error GoTo 0

' Variáveis para totais
Dim totalComissoesGeral, totalPagoGeral, totalSaldoGeral
totalComissoesGeral = 0
totalPagoGeral = 0
Dim totalVBrutoGeral
totalVBrutoGeral = 0

If hasRecords Then
    rs.MoveFirst
    Do While Not rs.EOF
        Dim vComissao, vPago, vBruto
        vComissao = SafeCDbl(rs("TotalComissaoDevida"))
        vPago = SafeCDbl(rs("ValorPagoPorVenda"))
        vBruto = SafeCDbl(rs("VBrutoConsolidado"))
        
        totalComissoesGeral = totalComissoesGeral + vComissao
        totalPagoGeral = totalPagoGeral + vPago
        totalVBrutoGeral = totalVBrutoGeral + vBruto
        
        rs.MoveNext
    Loop
    rs.MoveFirst
    totalSaldoGeral = totalComissoesGeral - totalPagoGeral
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Comissões - <%=diretoriaSessao%></title>
    <!-- Bootstrap local ou CDN -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            font-size: 14px;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
        }
        .container-fluid {
            padding: 15px;
            max-width: 1400px;
            margin: 0 auto;
        }
        .table th { 
            background-color: #008080; 
            color: white; 
            font-size: 12px;
            padding: 10px 8px;
            border-bottom: 2px solid #006666;
        }
        .table td {
            padding: 8px;
            font-size: 12px;
            vertical-align: middle;
        }
        .bg-pago { background-color: #e8f5e9 !important; }
        .bg-pendente { background-color: #ffebee !important; }
        .bg-parcial { background-color: #fffde7 !important; }
        .valor-numero { 
            font-family: 'Consolas', 'Courier New', monospace; 
            text-align: right;
            white-space: nowrap;
        }
        .header-section {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            color: white;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .metric-card {
            padding: 15px;
            border-radius: 8px;
            color: white;
            text-align: center;
            margin-bottom: 15px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            transition: transform 0.2s;
        }
        .metric-card:hover {
            transform: translateY(-3px);
        }
        .metric-value {
            font-size: 1.6rem;
            font-weight: bold;
            margin: 5px 0;
        }
        .metric-label {
            font-size: 0.85rem;
            opacity: 0.9;
        }
        .badge-custom {
            font-size: 0.75rem;
            padding: 4px 8px;
            font-weight: 500;
        }
        .table-responsive {
            overflow-x: auto;
            margin-bottom: 20px;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        .dataTables_wrapper {
            font-size: 12px;
            padding: 10px;
        }
        h1 {
            font-size: 1.8rem;
            margin-bottom: 5px;
        }
        .btn-sm {
            padding: 5px 10px;
            font-size: 12px;
        }
        .info-badge {
            background-color: #e3f2fd;
            color: #1565c0;
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 15px;
            border-left: 4px solid #2196f3;
        }
        .status-indicator {
            display: inline-block;
            width: 10px;
            height: 10px;
            border-radius: 50%;
            margin-right: 5px;
        }
        .status-pago { background-color: #4caf50; }
        .status-pendente { background-color: #f44336; }
        .status-parcial { background-color: #ff9800; }
        .card {
            border: none;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        .card-header {
            border-radius: 8px 8px 0 0 !important;
            font-weight: 600;
            padding: 12px 15px;
        }
        .footer-info {
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 6px;
            margin-top: 20px;
            font-size: 11px;
            color: #666;
            border-top: 1px solid #dee2e6;
        }
        .highlight-row:hover {
            background-color: #f8f9fa !important;
        }
    </style>
</head>
<body>

<div class="container-fluid">
    <!-- HEADER COM INFORMAÇÕES DA DIRETORIA -->
    <div class="header-section">
        <div class="row align-items-center">
            <div class="col-md-8">
                <h1 class="mb-2"><i class="fas fa-money-bill-wave me-2"></i>Relatório de Comissões</h1>
                <p class="mb-0">
                    <i class="fas fa-building me-1"></i><strong>Diretoria:</strong> <%=diretoriaSessao%>
                </p>
                <p class="mb-0 mt-1 small opacity-75">
                    <i class="fas fa-info-circle me-1"></i>Filtrado automaticamente pela sua diretoria
                </p>
            </div>
            <div class="col-md-4 text-end">
                <button class="btn btn-light btn-sm me-2" onclick="window.print()" title="Imprimir relatório">
                    <i class="fas fa-print me-1"></i>Imprimir
                </button>
                <button class="btn btn-light btn-sm me-2" onclick="exportarParaExcel()" title="Exportar para Excel">
                    <i class="fas fa-file-excel me-1"></i>Excel
                </button>
                <button class="btn btn-light btn-sm" onclick="recarregarPagina()" title="Atualizar dados">
                    <i class="fas fa-sync-alt me-1"></i>Atualizar
                </button>
            </div>
        </div>
    </div>
    
    <!-- INFORMAÇÃO DE FILTRO ATIVO -->
    <div class="info-badge">
        <div class="row align-items-center">
            <div class="col-md-8">
                <i class="fas fa-filter text-primary me-2"></i>
                <strong>Filtro ativo:</strong> Diretoria <span class="badge bg-primary"><%=diretoriaSessao%></span>
                <small class="text-muted ms-2">(<%=Year(Now())%>/<%=Month(Now())%>/<%=Day(Now())%> - <%=Hour(Now())%>:<%=Minute(Now())%>)</small>
            </div>
            <div class="col-md-4 text-end">
                <small class="text-muted">
                    <i class="fas fa-database me-1"></i>Sistema Access MDB
                </small>
            </div>
        </div>
    </div>
    
    <!-- MÉTRICAS RÁPIDAS -->
    <div class="row mb-4">
        <div class="col-md-3">
            <%
            Dim corMetrica1
            If totalSaldoGeral > 0 Then
                corMetrica1 = "#dc3545"
            ElseIf totalSaldoGeral < 0 Then
                corMetrica1 = "#c82333"
            Else
                corMetrica1 = "#6c757d"
            End If
            %>
            <div class="metric-card" style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%);">
                <div class="metric-label"><i class="fas fa-hand-holding-usd me-1"></i> Comissões Devidas</div>
                <div class="metric-value">R$ <%= FormatNumber(totalComissoesGeral, 2) %></div>
                <small class="opacity-75">
                    Total a pagar
                </small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="metric-card" style="background: linear-gradient(135deg, #007bff 0%, #00bfff 100%);">
                <div class="metric-label"><i class="fas fa-money-check-alt me-1"></i> Valor Pago</div>
                <div class="metric-value">R$ <%= FormatNumber(totalPagoGeral, 2) %></div>
                <small class="opacity-75">Já liquidado</small>
            </div>
        </div>
        <div class="col-md-3">
            <%
            Dim corCardSaldo, textoSaldo
            If totalSaldoGeral > 0 Then
                corCardSaldo = "linear-gradient(135deg, #dc3545 0%, #c82333 100%)"
                textoSaldo = "A PAGAR"
            ElseIf totalSaldoGeral < 0 Then
                corCardSaldo = "linear-gradient(135deg, #6c757d 0%, #5a6268 100%)"
                textoSaldo = "EXCEDENTE"
            Else
                corCardSaldo = "linear-gradient(135deg, #6c757d 0%, #5a6268 100%)"
                textoSaldo = "QUITADO"
            End If
            %>
            <div class="metric-card" style="background: <%=corCardSaldo%>;">
                <div class="metric-label"><i class="fas fa-balance-scale me-1"></i> Saldo Pendente</div>
                <div class="metric-value">R$ <%= FormatNumber(totalSaldoGeral, 2) %></div>
                <small class="opacity-75">
                    <span class="badge bg-danger badge-custom"><%=textoSaldo%></span>
                </small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="metric-card" style="background: linear-gradient(135deg, #6f42c1 0%, #9b59b6 100%);">
                <div class="metric-label"><i class="fas fa-percentage me-1"></i> % Pago</div>
                <div class="metric-value">
                    <%
                    Dim percentualPago
                    If totalComissoesGeral > 0 Then
                        percentualPago = FormatNumber((totalPagoGeral / totalComissoesGeral) * 100, 1)
                    Else
                        percentualPago = 0
                    End If
                    Response.Write percentualPago & "%"
                    %>
                </div>
                <small class="opacity-75">
                    <div class="progress mt-1" style="height: 5px;">
                        <div class="progress-bar bg-white" style="width: <%=percentualPago%>%;"></div>
                    </div>
                </small>
            </div>
        </div>
    </div>
    
    <!-- TABELA PRINCIPAL -->
    <% If hasRecords Then %>
    <div class="table-responsive">
        <table id="relatorioTable" class="table table-sm table-striped table-hover" style="width:100%; font-size: 12px;">
            <thead class="table-dark">
                <tr>
                    <th width="8%" class="text-center">ID Venda</th>
                    <th width="22%">Colaborador</th>
                    <th width="15%">Gerência</th>
                    <th width="12%" class="text-end">V. Bruto</th>
                    <th width="13%" class="text-end">Comissão</th>
                    <th width="13%" class="text-end">Pago</th>
                    <th width="13%" class="text-end">Saldo</th>
                    <th width="8%" class="text-center">Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                'Dim contadorVendas, vComissao, vPago, vSaldo, statusPagamento, statusClass, vBruto
                'Dim totalVBrutoFiltro, totalComissaoFiltro, totalPagoFiltro
                contadorVendas = 0
                totalVBrutoFiltro = 0
                totalComissaoFiltro = 0
                totalPagoFiltro = 0
                
                Do While Not rs.EOF
                    contadorVendas = contadorVendas + 1
                    
                    vBruto = SafeCDbl(rs("VBrutoConsolidado"))
                    vComissao = SafeCDbl(rs("TotalComissaoDevida"))
                    vPago = SafeCDbl(rs("ValorPagoPorVenda"))
                    vSaldo = vComissao - vPago
                    
                    totalVBrutoFiltro = totalVBrutoFiltro + vBruto
                    totalComissaoFiltro = totalComissaoFiltro + vComissao
                    totalPagoFiltro = totalPagoFiltro + vPago
                    
                    If vComissao > 0 And vPago = 0 Then
                        statusPagamento = "Pendente"
                        statusClass = "bg-pendente"
                    ElseIf vSaldo <= 0 Then
                        If vSaldo < 0 Then
                            statusPagamento = "Pago (Exc.)"
                        Else
                            statusPagamento = "Pago"
                        End If
                        statusClass = "bg-pago"
                    Else
                        statusPagamento = "Parcial"
                        statusClass = "bg-parcial"
                    End If
                %>
                <tr class="<%= statusClass %> highlight-row">
                    <td class="text-center">
                        <span class="badge bg-dark badge-custom">
                            <i class="fas fa-hashtag me-1"></i><%= rs("ID_Venda") %>
                        </span>
                    </td>
                    <td>
                        <div class="fw-bold"><%= rs("Nome") %></div>
                        <small class="text-muted">
                            <i class="fas fa-user me-1"></i>ID: <%= rs("UserID") %>
                        </small>
                    </td>
                    <td>
                        <span class="badge bg-info badge-custom">
                            <i class="fas fa-sitemap me-1"></i><%= rs("Gerencia") %>
                        </span>
                    </td>
                    <td class="valor-numero">
                        <strong>R$ <%= FormatNumber(vBruto, 2) %></strong>
                    </td>
                    <td class="valor-numero">
                        <span class="text-primary fw-bold">
                            R$ <%= FormatNumber(vComissao, 2) %>
                        </span>
                    </td>
                    <td class="valor-numero">
                        <%
                        If vPago > 0 Then
                            Response.Write "<span class='text-success fw-bold'>"
                        Else
                            Response.Write "<span class='text-muted'>"
                        End If
                        %>
                        R$ <%= FormatNumber(vPago, 2) %>
                        </span>
                    </td>
                    <td class="valor-numero">
                        <%
                        If vSaldo > 0 Then
                            Response.Write "<span class='text-danger fw-bold'>"
                        ElseIf vSaldo < 0 Then
                            Response.Write "<span class='text-success fw-bold'>"
                        Else
                            Response.Write "<span class='text-muted'>"
                        End If
                        %>
                        R$ <%= FormatNumber(vSaldo, 2) %>
                        </span>
                    </td>
                    <td class="text-center">
                        <%
                        Dim badgeClass, indicatorClass
                        If vSaldo > 0 Then
                            badgeClass = "bg-danger"
                            indicatorClass = "status-pendente"
                        ElseIf vSaldo < 0 Then
                            badgeClass = "bg-success"
                            indicatorClass = "status-pago"
                        Else
                            badgeClass = "bg-secondary"
                            indicatorClass = "status-pago"
                        End If
                        %>
                        <span class="badge <%=badgeClass%> badge-custom">
                            <span class="status-indicator <%=indicatorClass%>"></span>
                            <%= statusPagamento %>
                        </span>
                    </td>
                </tr>
                <%
                    rs.MoveNext
                Loop
                %>
            </tbody>
            <tfoot class="table-dark">
                <tr>
                    <td colspan="3" class="text-end"><strong>TOTAIS (<%=contadorVendas%> vendas):</strong></td>
                    <td class="valor-numero"><strong>R$ <%= FormatNumber(totalVBrutoFiltro, 2) %></strong></td>
                    <td class="valor-numero"><strong>R$ <%= FormatNumber(totalComissaoFiltro, 2) %></strong></td>
                    <td class="valor-numero"><strong>R$ <%= FormatNumber(totalPagoFiltro, 2) %></strong></td>
                    <td class="valor-numero">
                        <%
                        Dim totalSaldoClass
                        If totalSaldoGeral > 0 Then
                            totalSaldoClass = "text-warning"
                        ElseIf totalSaldoGeral < 0 Then
                            totalSaldoClass = "text-success"
                        Else
                            totalSaldoClass = "text-white"
                        End If
                        %>
                        <strong class="<%=totalSaldoClass%>">
                            R$ <%= FormatNumber(totalSaldoGeral, 2) %>
                        </strong>
                    </td>
                    <td class="text-center">
                        <%
                        Dim totalStatusBadge, totalStatusText
                        If totalSaldoGeral > 0 Then
                            totalStatusBadge = "bg-warning text-dark"
                            totalStatusText = "PENDENTE"
                        ElseIf totalSaldoGeral < 0 Then
                            totalStatusBadge = "bg-success"
                            totalStatusText = "SUPERAVIT"
                        Else
                            totalStatusBadge = "bg-secondary"
                            totalStatusText = "QUITADO"
                        End If
                        %>
                        <span class="badge <%=totalStatusBadge%> badge-custom">
                            <%= totalStatusText %>
                        </span>
                    </td>
                </tr>
            </tfoot>
        </table>
    </div>
    
    
    <% Else %>
    <!-- SEM REGISTROS -->
    <div class="alert alert-info text-center py-5">
        <i class="fas fa-database fa-3x mb-4 text-info"></i>
        <h4 class="fw-bold text-info">Nenhuma comissão encontrada</h4>
        <p class="mb-3">
            Não há registros de comissão para a diretoria <strong><%=diretoriaSessao%></strong> no momento.
        </p>
        <div class="mt-3">
            <button class="btn btn-info me-2" onclick="recarregarPagina()">
                <i class="fas fa-sync-alt me-1"></i>Verificar Novamente
            </button>
            <a href="gestao_vendas_diretoria2.asp" class="btn btn-outline-info">
                <i class="fas fa-arrow-left me-1"></i>Voltar ao Dashboard
            </a>
        </div>
    </div>
    <% End If %>
    
    <!-- RODAPÉ -->
    <div class="footer-info">
        <div class="row">
            <div class="col-md-6">
                <i class="fas fa-copyright me-1"></i> Relatório gerado automaticamente pelo sistema
            </div>
            <div class="col-md-6 text-end">
                <i class="fas fa-building me-1"></i> Diretoria: <%=diretoriaSessao%> | 
                <i class="fas fa-clock me-1 ms-2"></i> <%=FormatDateTime(Now(), 4)%>
            </div>
        </div>
    </div>
</div>

<% 
' Fechar conexões
If hasRecords Then
    rs.Close
End If
Set rs = Nothing
conn.Close
Set conn = Nothing
%>

<!-- Scripts -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>

<!-- DataTables apenas se houver muitos registros -->
<% If hasRecords And contadorVendas > 10 Then %>
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
<% End If %>

<script>
    // Configuração DataTables apenas se necessário
    <% If hasRecords And contadorVendas > 10 Then %>
    $(document).ready(function() {
        $('#relatorioTable').DataTable({
            "order": [[0, "desc"]],
            "pageLength": 25,
            "lengthMenu": [[10, 25, 50, 100], [10, 25, 50, 100]],
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            "dom": '<"row"<"col-md-6"l><"col-md-6"f>><"row"<"col-md-12"t>><"row"<"col-md-6"i><"col-md-6"p>>',
            "initComplete": function() {
                console.log('Tabela carregada com ' + <%=contadorVendas%> + ' registros');
            }
        });
    });
    <% End If %>
    
    // Função para exportar para Excel
    function exportarParaExcel() {
        var html = '';
        html += '<html xmlns:x="urn:schemas-microsoft-com:office:excel">';
        html += '<head>';
        html += '<meta charset="UTF-8">';
        html += '<style>';
        html += 'td { mso-number-format:\@; padding: 4px; border: 1px solid #ccc; }';
        html += 'th { background-color: #008080; color: white; font-weight: bold; padding: 8px; border: 1px solid #006666; }';
        html += '.valor { mso-number-format:"R\\$ #,##0.00"; }';
        html += '</style>';
        html += '</head>';
        html += '<body>';
        
        html += '<table border="1">';
        html += '<tr><th colspan="8" style="background:#2c3e50;color:white;padding:12px;font-size:14px;">Relatório de Comissões - Diretoria: <%=diretoriaSessao%></th></tr>';
        html += '<tr><td colspan="8" style="padding:8px;font-size:11px;">Gerado em: <%=FormatDateTime(Now(), 2)%> às <%=FormatDateTime(Now(), 4)%></td></tr>';
        
        // Cabeçalho
        html += '<tr>';
        html += '<th>ID Venda</th>';
        html += '<th>Colaborador</th>';
        html += '<th>Gerência</th>';
        html += '<th>V. Bruto</th>';
        html += '<th>Comissão Devida</th>';
        html += '<th>Valor Pago</th>';
        html += '<th>Saldo</th>';
        html += '<th>Status</th>';
        html += '</tr>';
        
        // Dados da tabela
        $('#relatorioTable tbody tr').each(function() {
            var cells = $(this).find('td');
            html += '<tr>';
            cells.each(function(index) {
                var text = $(this).text().trim();
                if (index >= 3 && index <= 6) { // Colunas numéricas
                    html += '<td class="valor">' + text + '</td>';
                } else {
                    html += '<td>' + text + '</td>';
                }
            });
            html += '</tr>';
        });
        
        // Totais
        html += '<tr style="font-weight:bold;background:#343a40;color:white;">';
        html += '<td colspan="3" align="right">TOTAIS (<%=contadorVendas%> registros):</td>';
        html += '<td class="valor">R$ <%= FormatNumber(totalVBrutoFiltro, 2) %></td>';
        html += '<td class="valor">R$ <%= FormatNumber(totalComissaoFiltro, 2) %></td>';
        html += '<td class="valor">R$ <%= FormatNumber(totalPagoFiltro, 2) %></td>';
        html += '<td class="valor">R$ <%= FormatNumber(totalSaldoGeral, 2) %></td>';
        <%
        Dim statusFinal
        If totalSaldoGeral > 0 Then
            statusFinal = "PENDENTE"
        ElseIf totalSaldoGeral < 0 Then
            statusFinal = "SUPERAVIT"
        Else
            statusFinal = "QUITADO"
        End If
        statusFinal = ""
        %>
        html += '<td><%=statusFinal%></td>';
        html += '</tr>';
        
        html += '</table>';
        html += '</body></html>';
        
        // Criar arquivo Excel
        var blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'Comissoes_<%=Replace(diretoriaSessao, " ", "_")%>_<%=Year(Now())%><%=Right("0" & Month(Now()), 2)%><%=Right("0" & Day(Now()), 2)%>.xls';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
    
    // Função para recarregar a página
    function recarregarPagina() {
        location.reload();
    }
    
    // Adicionar classe para impressão
    window.onbeforeprint = function() {
        $('.metric-card').addClass('no-shadow');
        $('.header-section').css('box-shadow', 'none');
    };
    
    // Auto-atualização a cada 5 minutos (300000 ms)
    setTimeout(function() {
        console.log('Auto-atualizando dados...');
        recarregarPagina();
    }, 300000);
</script>
</body>
</html>