<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="AtualizarVendas.asp"-->

<%
' ====================================================================
' Conexão e Variáveis
' ====================================================================
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Função para buscar total pago (definida ANTES do loop)
Function GetTotalPagoVenda(idVenda, tipoRecebedor, tipoPagamento)
    Dim sql, rs, total
    total = 0
    
    sql = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
          "WHERE ID_Venda = " & idVenda & " AND TipoRecebedor = '" & tipoRecebedor & "' AND TipoPagamento = '" & tipoPagamento & "'"
    
    Set rs = connSales.Execute(sql)
    If Not rs.EOF And Not IsNull(rs("TotalPago")) Then
        total = CDbl(rs("TotalPago"))
    End If
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    GetTotalPagoVenda = total
End Function

' ====================================================================
' Consulta principal para todas as vendas
' ====================================================================
Dim sqlVendas, rsVendas
sqlVendas = "SELECT " & _
           "v.ID, v.NomeEmpreendimento, v.Unidade, v.DataVenda, v.ValorComissaoGeral, " & _
           "v.Diretoria, v.Gerencia, v.Corretor, " & _
           "v.ValorDiretoria, v.PremioDiretoria, " & _
           "v.ValorGerencia, v.PremioGerencia, " & _
           "v.ValorCorretor, v.PremioCorretor, " & _
           "c.ID_Comissoes, c.StatusPagamento " & _
           "FROM Vendas AS v " & _
           "LEFT JOIN COMISSOES_A_PAGAR AS c ON v.ID = c.ID_Venda " & _
           "WHERE v.excluido = 0 " & _
           "ORDER BY v.DataVenda DESC, v.ID DESC"

Set rsVendas = connSales.Execute(sqlVendas)
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Comissões a Pagar - 2</title>
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <!-- DataTables CSS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/responsive/2.2.9/css/responsive.bootstrap5.min.css">
    <style>
        body {
            background-color: #f8f9fa;
            color: #333;
            padding: 20px;
        }
        .container-fluid {
            background-color: #fff;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 0 15px rgba(0,0,0,0.1);
        }
        .header-title {
            color: #800000;
            border-bottom: 2px solid #800000;
            padding-bottom: 10px;
            margin-bottom: 25px;
        }
        .table {
            background-color: #fff;
        }
        .table thead th {
            background-color: #800000;
            color: #fff;
            border: none;
        }
        .status-badge {
            font-size: 0.75rem;
            padding: 0.35em 0.65em;
            border-radius: 0.25rem;
        }
        .status-pago {
            background-color: #28a745;
            color: white;
        }
        .status-pendente {
            background-color: #dc3545;
            color: white;
        }
        .status-parcial {
            background-color: #ffc107;
            color: #212529;
        }
        .btn-pagar-tudo {
            background: linear-gradient(135deg, #28a745, #20c997);
            border: none;
            color: white;
            font-weight: 600;
        }
        .btn-pagar-tudo:hover {
            background: linear-gradient(135deg, #218838, #1e9e8a);
            color: white;
        }
        .valor-destaque {
            font-weight: bold;
            color: #2c5aa0;
        }
        .valor-premio {
            font-weight: bold;
            color: #e83e8c;
        }
        .row-paga {
            background-color: #e8f5e8 !important;
        }
        .row-pendente {
            background-color: #ffe6e6 !important;
        }
        .row-parcial {
            background-color: #fff3cd !important;
        }
        .card-valor {
            border-left: 4px solid #28a745;
            background-color: #f8f9fa;
        }
        .card-premio {
            border-left: 4px solid #e83e8c;
            background-color: #f8f9fa;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="row mb-4">
            <div class="col-12">
                <h1 class="text-center header-title">
                    <i class="fas fa-list-alt me-2"></i>Comissões a Pagar - 2
                </h1>
            </div>
        </div>

        <div class="row mb-4">
            <div class="col-md-6">
                <a href="gestao_vendas_gerenc_comissoes.asp" class="btn btn-primary">
                    <i class="fas fa-arrow-left me-2"></i>Voltar para Comissões
                </a>
            </div>
            <div class="col-md-6 text-end">
                <div class="btn-group">
                    <button class="btn btn-info" onclick="location.reload()">
                        <i class="fas fa-sync-alt me-2"></i>Atualizar
                    </button>
                </div>
            </div>
        </div>

        <!-- Cards de Resumo -->
        <div class="row mb-4">
            <div class="col-md-4">
                <div class="card card-valor">
                    <div class="card-body">
                        <h5 class="card-title">Total de Vendas</h5>
                        <h3 class="card-text valor-destaque" id="totalVendas">0</h3>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body">
                        <h5 class="card-title">Comissão Total</h5>
                        <h3 class="card-text text-success" id="totalComissao">R$ 0,00</h3>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card card-premio">
                    <div class="card-body">
                        <h5 class="card-title">Prêmio Total</h5>
                        <h3 class="card-text valor-premio" id="totalPremio">R$ 0,00</h3>
                    </div>
                </div>
            </div>
        </div>

        <div class="table-responsive">
            <table id="vendasTable" class="table table-striped table-bordered table-hover align-middle" style="width:100%">
                <thead>
                    <tr>
                        <th class="text-center">ID Venda</th>
                        <th class="text-center">Status</th>
                        <th class="text-center">Empreendimento</th>
                        <th class="text-center">Data</th>
                        <th class="text-center">Diretoria</th>
                        <th class="text-center">Gerência</th>
                        <th class="text-center">Corretor</th>
                        <th class="text-center">Valores</th>
                        <th class="text-center">Ações</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim totalVendas, totalComissao, totalPremio
                    totalVendas = 0
                    totalComissao = 0
                    totalPremio = 0
                    
                    If Not rsVendas.EOF Then
                        Do While Not rsVendas.EOF
                            totalVendas = totalVendas + 1
                            
                            ' Calcular totais
                            If Not IsNull(rsVendas("ValorComissaoGeral")) Then
                                totalComissao = totalComissao + CDbl(rsVendas("ValorComissaoGeral"))
                            End If
                            
                            If Not IsNull(rsVendas("PremioDiretoria")) Then
                                totalPremio = totalPremio + CDbl(rsVendas("PremioDiretoria"))
                            End If
                            If Not IsNull(rsVendas("PremioGerencia")) Then
                                totalPremio = totalPremio + CDbl(rsVendas("PremioGerencia"))
                            End If
                            If Not IsNull(rsVendas("PremioCorretor")) Then
                                totalPremio = totalPremio + CDbl(rsVendas("PremioCorretor"))
                            End If
                            
                            ' Determinar status e classe da linha
                            Dim status, statusClass, rowClass
                            status = "PENDENTE"
                            If Not IsNull(rsVendas("StatusPagamento")) Then
                                status = rsVendas("StatusPagamento")
                            End If
                            
                            Select Case UCase(status)
                                Case "PAGA": 
                                    statusClass = "status-pago"
                                    rowClass = "row-paga"
                                Case "PAGA PARCIALMENTE": 
                                    statusClass = "status-parcial"
                                    rowClass = "row-parcial"
                                Case "PENDENTE": 
                                    statusClass = "status-pendente"
                                    rowClass = "row-pendente"
                                Case Else: 
                                    statusClass = "bg-secondary"
                                    rowClass = ""
                            End Select
                            
                            ' Buscar pagamentos realizados para calcular saldos
                            Dim saldoDiretoria, saldoGerencia, saldoCorretor
                            Dim saldoPremioDiretoria, saldoPremioGerencia, saldoPremioCorretor
                            
                            ' Inicializar valores
                            saldoDiretoria = 0
                            saldoGerencia = 0
                            saldoCorretor = 0
                            saldoPremioDiretoria = 0
                            saldoPremioGerencia = 0
                            saldoPremioCorretor = 0
                            
                            ' Calcular valores totais
                            If Not IsNull(rsVendas("ValorDiretoria")) Then
                                saldoDiretoria = CDbl(rsVendas("ValorDiretoria"))
                            End If
                            If Not IsNull(rsVendas("ValorGerencia")) Then
                                saldoGerencia = CDbl(rsVendas("ValorGerencia"))
                            End If
                            If Not IsNull(rsVendas("ValorCorretor")) Then
                                saldoCorretor = CDbl(rsVendas("ValorCorretor"))
                            End If
                            If Not IsNull(rsVendas("PremioDiretoria")) Then
                                saldoPremioDiretoria = CDbl(rsVendas("PremioDiretoria"))
                            End If
                            If Not IsNull(rsVendas("PremioGerencia")) Then
                                saldoPremioGerencia = CDbl(rsVendas("PremioGerencia"))
                            End If
                            If Not IsNull(rsVendas("PremioCorretor")) Then
                                saldoPremioCorretor = CDbl(rsVendas("PremioCorretor"))
                            End If
                            
                            ' Calcular saldos subtraindo pagamentos já realizados
                            saldoDiretoria = saldoDiretoria - GetTotalPagoVenda(rsVendas("ID"), "diretoria", "Comissão")
                            saldoGerencia = saldoGerencia - GetTotalPagoVenda(rsVendas("ID"), "gerencia", "Comissão")
                            saldoCorretor = saldoCorretor - GetTotalPagoVenda(rsVendas("ID"), "corretor", "Comissão")
                            saldoPremioDiretoria = saldoPremioDiretoria - GetTotalPagoVenda(rsVendas("ID"), "diretoria", "Premiação")
                            saldoPremioGerencia = saldoPremioGerencia - GetTotalPagoVenda(rsVendas("ID"), "gerencia", "Premiação")
                            saldoPremioCorretor = saldoPremioCorretor - GetTotalPagoVenda(rsVendas("ID"), "corretor", "Premiação")
                            
                            ' Verificar se há saldos pendentes
                            Dim temSaldoPendente
                            temSaldoPendente = (saldoDiretoria > 0 Or saldoGerencia > 0 Or saldoCorretor > 0 Or _
                                              saldoPremioDiretoria > 0 Or saldoPremioGerencia > 0 Or saldoPremioCorretor > 0)
                    %>
                    <tr class="<%= rowClass %>">
                        <td class="text-center">
                            <%= Year(rsVendas("DataVenda")) & "-" & Right("0" & Month(rsVendas("DataVenda")),2) & "-" & Right("0" & Day(rsVendas("DataVenda")),2) %><br>
                            <strong>V<%= rsVendas("ID") %></strong>
                            <% If Not IsNull(rsVendas("ID_Comissoes")) Then %>
                                <small class="text-muted">C<%= rsVendas("ID_Comissoes") %></small>
                            <% End If %>
                        </td>
                        <td class="text-center">
                            <span class="status-badge <%= statusClass %>"><%= status %></span>
                        </td>
                        <td>
                            <strong><%= rsVendas("NomeEmpreendimento") %></strong><br>
                            <small class="text-muted"><%= rsVendas("Unidade") %></small>
                        </td>
                        <td class="text-center">
                            <%= FormatDateTime(rsVendas("DataVenda"), 2) %>
                        </td>
                        <td>
                            <strong><%= rsVendas("Diretoria") %></strong><br>
                            <small class="text-muted">
                                Comissão: R$ <%= FormatNumber(rsVendas("ValorDiretoria"), 2) %><br>
                                <% If CDbl(rsVendas("PremioDiretoria")) > 0 Then %>
                                Prêmio: R$ <%= FormatNumber(rsVendas("PremioDiretoria"), 2) %>
                                <% End If %>
                            </small>
                        </td>
                        <td>
                            <strong><%= rsVendas("Gerencia") %></strong><br>
                            <small class="text-muted">
                                Comissão: R$ <%= FormatNumber(rsVendas("ValorGerencia"), 2) %><br>
                                <% If CDbl(rsVendas("PremioGerencia")) > 0 Then %>
                                Prêmio: R$ <%= FormatNumber(rsVendas("PremioGerencia"), 2) %>
                                <% End If %>
                            </small>
                        </td>
                        <td>
                            <strong><%= rsVendas("Corretor") %></strong><br>
                            <small class="text-muted">
                                Comissão: R$ <%= FormatNumber(rsVendas("ValorCorretor"), 2) %><br>
                                <% If CDbl(rsVendas("PremioCorretor")) > 0 Then %>
                                Prêmio: R$ <%= FormatNumber(rsVendas("PremioCorretor"), 2) %>
                                <% End If %>
                            </small>
                        </td>
                        <td class="text-center">
                            <div class="valor-destaque">
                                R$ <%= FormatNumber(rsVendas("ValorComissaoGeral"), 2) %>
                            </div>
                            <% 
                            Dim totalPremioVenda
                            totalPremioVenda = 0
                            If Not IsNull(rsVendas("PremioDiretoria")) Then totalPremioVenda = totalPremioVenda + CDbl(rsVendas("PremioDiretoria"))
                            If Not IsNull(rsVendas("PremioGerencia")) Then totalPremioVenda = totalPremioVenda + CDbl(rsVendas("PremioGerencia"))
                            If Not IsNull(rsVendas("PremioCorretor")) Then totalPremioVenda = totalPremioVenda + CDbl(rsVendas("PremioCorretor"))
                            
                            If totalPremioVenda > 0 Then 
                            %>
                            <div class="valor-premio">
                                <i class="fas fa-trophy"></i> R$ <%= FormatNumber(totalPremioVenda, 2) %>
                            </div>
                            <% End If %>
                        </td>
                        <td class="text-center">
                            <div class="btn-group-vertical" role="group">
                                <% If temSaldoPendente Then %>
                                <button class="btn btn-pagar-tudo btn-sm mb-1" 
                                    data-bs-toggle="modal" 
                                    data-bs-target="#pagarTodosModal"
                                    data-id-venda="<%= rsVendas("ID") %>"
                                    data-id-comissao="<%= rsVendas("ID_Comissoes") %>"
                                    data-empreendimento="<%= Server.HTMLEncode(rsVendas("NomeEmpreendimento")) %>"
                                    data-unidade="<%= Server.HTMLEncode(rsVendas("Unidade")) %>"
                                    data-comissao-total="<%= FormatNumber(rsVendas("ValorComissaoGeral"), 2) %>">
                                    <i class="fas fa-money-bill-wave me-1"></i> Pagar Tudo
                                </button>
                                <% Else %>
                                <span class="badge bg-success">Tudo Pago</span>
                                <% End If %>
                                
                                
                                <button class="btn btn-warning btn-sm ver-pagamentos-btn"
                                    data-bs-toggle="modal" 
                                    data-bs-target="#viewPaymentsModal"
                                    data-id-venda="<%= rsVendas("ID") %>">
                                    <i class="fas fa-eye me-1"></i> Pagamentos
                                </button>
                            </div>
                        </td>
                    </tr>
                    <%
                            rsVendas.MoveNext
                        Loop
                    Else
                    %>
                    <tr>
                        <td colspan="9" class="text-center">Nenhuma venda encontrada.</td>
                    </tr>
                    <%
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Modal para Pagar Todos -->
    <div class="modal fade" id="pagarTodosModal" tabindex="-1" aria-labelledby="pagarTodosModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header bg-success text-white">
                    <h5 class="modal-title" id="pagarTodosModalLabel">
                        <i class="fas fa-money-bill-wave me-2"></i>Pagar Todas as Comissões e Prêmios
                    </h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <form id="pagarTodosForm" action="gestao_vendas_salvar_pag_todos1.asp" method="post">
                    <div class="modal-body">
                        <input type="hidden" id="pagarTodosVendaId" name="ID_Venda">
                        <input type="hidden" id="pagarTodosComissaoId" name="ID_Comissao">
                        
                        <div class="row mb-3">
                            <div class="col-12">
                                <div class="alert alert-info">
                                    <h6><i class="fas fa-info-circle me-2"></i>Resumo da Venda</h6>
                                    <div id="resumoVenda">
                                        <!-- Preenchido via JavaScript -->
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="pagarTodosDataPagamento" class="form-label">Data do Pagamento *</label>
                                    <input type="date" class="form-control" id="pagarTodosDataPagamento" name="DataPagamento" required>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="mb-3">
                                    <label for="pagarTodosStatusPagamento" class="form-label">Status do Pagamento *</label>
                                    <select class="form-select" id="pagarTodosStatusPagamento" name="Status" required>
                                        <option value="">Selecione...</option>
                                        <option value="Em processamento">Em processamento</option>
                                        <option value="Agendado">Agendado</option>
                                        <option value="Realizado">Realizado</option>
                                    </select>
                                </div>
                            </div>
                        </div>

                        <div class="mb-3">
                            <label for="pagarTodosObs" class="form-label">Observações</label>
                            <textarea class="form-control" id="pagarTodosObs" name="Obs" rows="3" placeholder="Observações para todos os pagamentos"></textarea>
                        </div>

                        <div class="alert alert-warning">
                            <h6><i class="fas fa-exclamation-triangle me-2"></i>Atenção</h6>
                            <p class="mb-0">Esta ação irá pagar <strong>TODOS os valores pendentes</strong> de comissões e prêmios para esta venda. Verifique os valores antes de confirmar.</p>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                            <i class="fas fa-times me-2"></i>Cancelar
                        </button>
                        <button type="submit" class="btn btn-success">
                            <i class="fas fa-check me-2"></i>Confirmar Todos os Pagamentos
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Modal para Visualizar Pagamentos (reutilizado) -->
    <div class="modal fade" id="viewPaymentsModal" tabindex="-1" aria-labelledby="viewPaymentsModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header bg-primary text-white">
                    <h5 class="modal-title" id="viewPaymentsModalLabel">Histórico de Pagamentos</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="table-responsive">
                        <table class="table table-striped table-hover" id="paymentsTable">
                            <thead class="table-dark">
                                <tr>
                                    <th>ID</th>
                                    <th>Data</th>
                                    <th>Tipo</th>
                                    <th class="text-end">Valor</th>
                                    <th>Destinatário</th>
                                    <th>Cargo</th>
                                    <th>Status</th>
                                    <th>Observações</th>
                                </tr>
                            </thead>
                            <tbody id="paymentsTableBody">
                                <!-- Os dados serão preenchidos via JavaScript -->
                            </tbody>
                        </table>
                    </div>
                    <div id="noPaymentsMessage" class="alert alert-info mt-3" style="display: none;">
                        Nenhum pagamento encontrado para esta venda.
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <!-- DataTables JS -->
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/responsive/2.2.9/js/dataTables.responsive.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/responsive/2.2.9/js/responsive.bootstrap5.min.js"></script>

    <script>
    // Atualizar cards de resumo
    document.getElementById('totalVendas').textContent = '<%= totalVendas %>';
    document.getElementById('totalComissao').textContent = 'R$ <%= FormatNumber(totalComissao, 2) %>';
    document.getElementById('totalPremio').textContent = 'R$ <%= FormatNumber(totalPremio, 2) %>';

    // Inicializar DataTable
    $(document).ready(function() {
        $('#vendasTable').DataTable({
            responsive: true,
            order: [[0, "desc"]],
            pageLength: 25,
            lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "Todos"]],
            language: {
                url: 'https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json'
            },
            dom: '<"top"lf>rt<"bottom"ip>',
            initComplete: function() {
                this.api().columns.adjust().responsive.recalc();
            }
        });

        // Evento para o modal Pagar Todos
        $('#pagarTodosModal').on('show.bs.modal', function(event) {
            const button = $(event.relatedTarget);
            const idVenda = button.data('id-venda');
            const idComissao = button.data('id-comissao');
            const empreendimento = button.data('empreendimento');
            const unidade = button.data('unidade');
            const comissaoTotal = button.data('comissao-total');

            $('#pagarTodosVendaId').val(idVenda);
            $('#pagarTodosComissaoId').val(idComissao);

            // Preencher data atual
            const today = new Date();
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const day = String(today.getDate()).padStart(2, '0');
            const todayStr = `${year}-${month}-${day}`;
            $('#pagarTodosDataPagamento').val(todayStr);

            // Preencher resumo
            $('#resumoVenda').html(`
                <p><strong>Venda:</strong> V${idVenda}</p>
                <p><strong>Empreendimento:</strong> ${empreendimento} - ${unidade}</p>
                <p><strong>Comissão Total:</strong> ${comissaoTotal}</p>
            `);
        });

        // Validação do formulário Pagar Todos
        $('#pagarTodosForm').submit(function(e) {
            if ($('#pagarTodosDataPagamento').val() === '') {
                alert('Por favor, selecione a data do pagamento.');
                e.preventDefault();
                return;
            }
            
            if ($('#pagarTodosStatusPagamento').val() === '') {
                alert('Por favor, selecione o status do pagamento.');
                e.preventDefault();
                return;
            }
            
            if (!confirm('Tem certeza que deseja realizar TODOS os pagamentos pendentes para esta venda?')) {
                e.preventDefault();
                return;
            }
        });

        // Função para carregar pagamentos (reutilizada)
        function loadPayments(idVenda) {
            $('#paymentsTableBody').html('<tr><td colspan="8" class="text-center"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Carregando...</span></div></td></tr>');
            $('#noPaymentsMessage').hide();

            $.ajax({
                url: 'get_pagamentos_por_comissao.asp',
                type: 'GET',
                dataType: 'json',
                data: { idVenda: idVenda },
                success: function(response) {
                    if (response && response.success && response.data && Array.isArray(response.data) && response.data.length > 0) {
                        let html = '';
                        response.data.forEach(function(payment) {
                            const valorPago = (typeof payment.ValorPago === 'string') ? 
                                parseFloat(payment.ValorPago.replace(',', '.')) : (payment.ValorPago || 0);
                            
                            html += `
                                <tr>
                                    <td>#${payment.ID_Pagamento || 'N/A'}</td>
                                    <td>${payment.DataPagamento}</td>
                                    <td>${payment.TipoPagamento}</td>
                                    <td class="text-end">R$ ${valorPago.toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                                    <td>${payment.UsuariosNome || 'N/A'}</td>
                                    <td>${(payment.TipoRecebedor || 'N/A').toUpperCase()}</td>
                                    <td><span class="badge bg-success">${payment.Status || 'N/A'}</span></td>
                                    <td>${payment.Obs || '-'}</td>
                                </tr>`;
                        });
                        $('#paymentsTableBody').html(html);
                        $('#noPaymentsMessage').hide();
                    } else {
                        $('#paymentsTableBody').html('<tr><td colspan="8" class="text-center">Nenhum pagamento encontrado.</td></tr>');
                        $('#noPaymentsMessage').show();
                    }
                },
                error: function(xhr, status, error) {
                    $('#paymentsTableBody').html(`
                        <tr>
                            <td colspan="8" class="text-center text-danger">
                                Erro ao carregar pagamentos.
                            </td>
                        </tr>
                    `);
                    $('#noPaymentsMessage').hide();
                }
            });
        }

        // Evento para o modal de pagamentos
        $('#viewPaymentsModal').on('show.bs.modal', function(event) {
            const button = $(event.relatedTarget);
            const idVenda = button.data('id-venda');
            loadPayments(idVenda);
        });
    });
    </script>
</body>
</html>

<%
' ====================================================================
' Fechar conexões
' ====================================================================
If Not rsVendas Is Nothing Then rsVendas.Close
Set rsVendas = Nothing

If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close
If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
Set connSales = Nothing
Set conn = Nothing
%>