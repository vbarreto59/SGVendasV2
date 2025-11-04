<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="AtualizarVendas.asp"-->


<%
' Função para formatar datas
Function FormatDateForDisplay(dateValue)
    If Not IsNull(dateValue) And IsDate(dateValue) Then
        FormatDateForDisplay = FormatDateTime(dateValue, 2)
    Else
        FormatDateForDisplay = "N/A"
    End If
End Function

' Função para formatar valores monetários
Function FormatCurrencyForDisplay(value)
    If Not IsNull(value) And IsNumeric(value) Then
        FormatCurrencyForDisplay = "R$ " & FormatNumber(value, 2)
    Else
        FormatCurrencyForDisplay = "R$ 0,00"
    End If
End Function
%>

<%
' ====================================================================
' Conexão e Variáveis - Otimizado para apenas uma conexão
' ====================================================================


' As variáveis relacionadas à conexão principal (StrConn, conn) não são necessárias para este bloco.
Dim rsComissoes
Dim sqlComissoes, sqlCheckStatus, sqlUpdateStatus
Dim comissaoId, vendaId
Dim userIdDiretoria, userIdGerencia, userIdCorretor
Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
Dim dbSalesPath

dbSalesPath = Split(StrConnSales, "Data Source=")(1)

' ====================================================================
' Sua consulta principal para as comissões a pagar (AJUSTADA e SIMPLIFICADA)
' ====================================================================

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn


Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales


sqlComissoes = "SELECT c.ID_Comissoes, c.ID_Venda, v.NomeEmpreendimento, v.Unidade, v.DataVenda, v.ValorComissaoGeral, " & _
               "c.UserIdDiretoria, c.NomeDiretor, v.ComissaoDIretoria, v.ValorDiretoria, " & _
               "c.UserIdGerencia, c.NomeGerente, v.ComissaoGerencia, v.ValorGerencia, " & _
               "c.UserIdCorretor, c.NomeCorretor, v.ComissaoCorretor, v.ValorCorretor, v.ID, v.Diretoria, v.Gerencia," & _
               "c.StatusPagamento " & _
               "FROM COMISSOES_A_PAGAR AS c INNER JOIN Vendas AS v ON c.ID_Venda = v.ID " & _
               "WHERE v.excluido = 0 ORDER BY c.ID_Comissoes DESC;"

Set rsComissoes = connSales.Execute(sqlComissoes)


' ====================================================================
' Script para Verificar e Atualizar Status de Comissões (PAGA/PENDENTE) - Otimizado
' ====================================================================
Response.Buffer = True
Response.Expires = -1
On Error GoTo 0 ' Habilita tratamento de erro explícito para o bloco todo

Dim rsCheckStatus
sqlCheckStatus = "SELECT c.ID_Comissoes, c.ID_Venda, c.StatusPagamento, " & _
                 "v.ValorDiretoria, v.ValorGerencia, v.ValorCorretor " & _
                 "FROM COMISSOES_A_PAGAR c INNER JOIN Vendas v ON c.ID_Venda = v.ID ORDER by c.ID_Comissoes"

Set rsCheckStatus = connSales.Execute(sqlCheckStatus)

Do While Not rsCheckStatus.EOF
    Dim comissaoIdCheck, vendaIdCheck, currentStatusComissao
    Dim valorDirCheck, valorGerCheck, valorCorCheck
    Dim totalDirPaid, totalGerPaid, totalCorPaid
    Dim newStatusComissao

    comissaoIdCheck = rsCheckStatus("ID_Comissoes")
    vendaIdCheck = rsCheckStatus("ID_Venda")
    currentStatusComissao = rsCheckStatus("StatusPagamento")
    valorDirCheck = rsCheckStatus("ValorDiretoria")
    valorGerCheck = rsCheckStatus("ValorGerencia")
    valorCorCheck = rsCheckStatus("ValorCorretor")

    totalDirPaid = 0
    totalGerPaid = 0
    totalCorPaid = 0

    Dim sqlGetPaid, rsGetPaid
    ' --- Verificar pagamentos para Diretoria (agora na conexão 'connSales') ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'diretoria'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalDirPaid = rsGetPaid("TotalPago")
    If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing

    ' --- Verificar pagamentos para Gerência ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'gerencia'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalGerPaid = rsGetPaid("TotalPago")
    If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing

    ' --- Verificar pagamentos para Corretor ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'corretor'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalCorPaid = rsGetPaid("TotalPago")
    If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing

    newStatusComissao = "PAGA"
    If CDbl(valorDirCheck) > 0 And CDbl(totalDirPaid) < CDbl(valorDirCheck) Then newStatusComissao = "PENDENTE"
    If CDbl(valorGerCheck) > 0 And CDbl(totalGerPaid) < CDbl(valorGerCheck) Then newStatusComissao = "PENDENTE"
    If CDbl(valorCorCheck) > 0 And CDbl(totalCorPaid) < CDbl(valorCorCheck) Then newStatusComissao = "PENDENTE"

    If newStatusComissao <> currentStatusComissao Then
        sqlUpdateStatus = "UPDATE COMISSOES_A_PAGAR SET StatusPagamento = '" & newStatusComissao & "' WHERE ID_Comissoes = " & comissaoIdCheck
        connSales.Execute(sqlUpdateStatus)
    End If

    rsCheckStatus.MoveNext
Loop

If Not rsCheckStatus Is Nothing Then rsCheckStatus.Close
Set rsCheckStatus = Nothing

' ====================================================================
' Bloco de Limpeza (IMPORTANTE)
' ====================================================================
If Not connSales Is Nothing Then If connSales.State = adStateOpen Then connSales.Close : Set connSales = Nothing  
%>    
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lista de Comissões a Pagar</title>
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <!-- DataTables CSS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/responsive/2.2.9/css/responsive.bootstrap5.min.css">
    <!-- jQuery e jQuery Mask (para o modal) -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.16/jquery.mask.min.js"></script>
    <style>
        body {
            background-color: #807777;
            color: #fff;
            padding: 20px;
        }
        .container-fluid {
            background-color: #fff;
            color: #000;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0,0,0,0.5);
        }
        .table {
            background-color: #f8f9fa;
        }
        .table thead th {
            background-color: #800000;
            color: #fff;
        }
        .table-striped > tbody > tr:nth-of-type(odd) {
            background-color: #f1f1f1;
        }
        .status-badge {
            font-size: 0.8rem;
            padding: 0.25em 0.5em;
            border-radius: 0.25rem;
        }
        .status-pago {
            background-color: #28a745;
            color: white;
        }
        .status-pendente {
            background-color: #ffc107;
            color: #212529;
        }
        .status-parcial {
            background-color: #17a2b8;
            color: white;
        }
        .header-title {
            color: #800000;
        }
        .total-row {
            font-weight: bold;
            background-color: #e9ecef;
        }
        .dataTables_wrapper .dataTables_length, 
        .dataTables_wrapper .dataTables_filter, 
        .dataTables_wrapper .dataTables_info, 
        .dataTables_wrapper .dataTables_paginate {
            color: #000 !important;
        }
        .dataTables_wrapper .dataTables_filter input {
            color: #000 !important;
            background-color: #fff !important;
        }
        .dataTables_wrapper .dataTables_length select {
            color: #000 !important;
            background-color: #fff !important;
        }

        /* Estilos para o modal */
        .modal-content {
            color: #000;
        }
        .modal-header, .modal-body, .modal-footer {
            color: #000;
        }
        .modal-title {
            color: #000;
        }
        .form-label {
            color: #333;
        }
        .form-control-plaintext {
            color: #555;
        }
        .modal-body input[type="text"],
        .modal-body input[type="date"],
        .modal-body textarea,
        .modal-body select {
            color: #000;
            background-color: #fff;
        }
    </style>
</head>
<body>
    <div class="container-fluid mt-5">
        <h2 class="text-center mb-4 header-title"><i class="fas fa-coins me-2"></i>Comissões a Pagar</h2>
        <a href="gestao_vendas_comissao_saldo1.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Saldos</a>
        <a href="inserirVendasTeste.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Inserir Testes</a>
        <br><br>
        <div class="table-responsive">
            <table id="comissoesTable" class="table table-striped table-bordered align-middle nowrap" style="width:100%">
                <thead>
                    <tr>
                        <th class="text-center">ID Comissão</th>
                        <th class="text-center">Status</th>
                        <th class="text-center">Venda</th>
                        <th class="text-center">Data Venda</th>
                        <th class="text-center">Total Comissão</th>
                        <th class="text-center">Diretoria</th>
                        <th class="text-center">Gerência</th>
                        <th class="text-center">Corretor</th>
                        
                        <th class="text-center">Ações</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If Not rsComissoes.EOF Then
                        Do While Not rsComissoes.EOF
                            comissaoId = rsComissoes("ID_Comissoes")
                            vendaId = rsComissoes("ID_Venda")
                            userIdDiretoria = rsComissoes("UserIdDiretoria")
                            userIdGerencia = rsComissoes("UserIdGerencia")
                            userIdCorretor = rsComissoes("UserIdCorretor")

                            ' Inicializa valores pagos
                            totalPagoDiretoria = 0
                            totalPagoGerencia = 0
                            totalPagoCorretor = 0


                            ' ====================================================================
                            ' Consulta para obter os pagamentos já realizados para esta venda e tipo de recebedor
                            ' ====================================================================
                            ' #### Pagamentos para Diretoria
                           '' sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago, MAX(DataPagamento) as DataPagamento FROM PAGAMENTOS_COMISSOES " & _
                                            '"WHERE ID_Venda = " & vendaId & " AND UsuariosUserId = " & userIdDiretoria & " AND TipoRecebedor = 'diretoria'"

                             sqlPagamentos = "SELECT Sum(ValorPago) AS ValorTotalPago, MAX(DataPagamento) as DataPagamento  " & _
                                            "FROM PAGAMENTOS_COMISSOES " & _
                                            "WHERE PAGAMENTOS_COMISSOES.ID_Venda=" & vendaId & " " & _
                                            "AND PAGAMENTOS_COMISSOES.UsuariosUserId=" & userIdDiretoria & " " & _
                                            "AND PAGAMENTOS_COMISSOES.TipoRecebedor='diretoria';"
                                   


                            Set rsPagamentos = connSales.Execute(sqlPagamentos)
                            
                            Dim dataPagamentoDiretoria
                            If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                                totalPagoDiretoria = rsPagamentos("ValorTotalPago")
                                If Not IsNull(rsPagamentos("DataPagamento")) Then
                                    dataPagamentoDiretoria = FormatDateTime(rsPagamentos("DataPagamento"), 2) ' Formata a data para exibir
                                End If
                            End If
                            If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing



                            sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago, MAX(DataPagamento) as DataPag " & _
                                            "FROM PAGAMENTOS_COMISSOES " & _
                                            "WHERE ID_Venda = " & vendaId & " " & _
                                            "AND UsuariosUserId = " & userIdGerencia & " " & _
                                            "AND TipoRecebedor = 'gerencia'"
                            Set rsPagamentos = connSales.Execute(sqlPagamentos)
                            Dim dataPagamentoGerencia
                            If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                                totalPagoGerencia = rsPagamentos("ValorTotalPago")
                                If Not IsNull(rsPagamentos("DataPag")) Then
                                    dataPagamentoGerencia = FormatDateTime(rsPagamentos("DataPag"), 2)
                                End If
                            End If
                            If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing


                            ' ------------- #### Pagamentos para Corretor
                            sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago, MAX(DataPagamento) as DataPag " & _
                                            "FROM PAGAMENTOS_COMISSOES " & _
                                            "WHERE ID_Venda = " & vendaId & " " & _
                                            "AND UsuariosUserId = " & userIdCorretor & " " & _
                                            "AND TipoRecebedor = 'corretor'"
                            Set rsPagamentos = connSales.Execute(sqlPagamentos)
                            Dim dataPagamentoCorretor
                            If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                                totalPagoCorretor = rsPagamentos("ValorTotalPago")
                                If Not IsNull(rsPagamentos("DataPag")) Then
                                    dataPagamentoCorretor = FormatDateTime(rsPagamentos("DataPag"), 2)
                                End If
                            End If
                            If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing
                            
                            ' ====================================================================
                            ' Determina o status da comissão com base no campo StatusPagamento
                            ' ====================================================================
                            Dim status, statusClass
                            status = rsComissoes("StatusPagamento")
                            Select Case UCase(status)
                                Case "PAGA"
                                    statusClass = "status-pago"
                                Case "PAGA PARCIALMENTE"
                                    statusClass = "status-parcial"
                                Case "PENDENTE"
                                    statusClass = "status-pendente"
                                Case Else
                                    statusClass = "bg-secondary text-white" ' Default
                            End Select
                    %>
                    <tr>
                        <td class="text-center"><%="C"& rsComissoes("ID_Comissoes") %>-<%="V"&vendaID%></td>
                        <td class="text-center"><span class="status-badge <%= statusClass %>"><%= UCase(status) %></span></td>
                        <td class="text-center"><small class="text-muted"><b><%= rsComissoes("NomeEmpreendimento") %></b></small><br><%= rsComissoes("Unidade") %><br><small class="text-muted">ID Venda: <%= rsComissoes("ID_Venda") %></small></td>
                        <td class="text-center"><%= FormatDateTime(rsComissoes("DataVenda"), 2) %></td>
                        <td class="text-center">R$ <%= FormatNumber(rsComissoes("ValorComissaoGeral"), 2) %></td>

                        <td class="text-center">
                            <div><b><%= rsComissoes("Diretoria") %><br><%= userIdDiretoria&"-"&rsComissoes("NomeDiretor") %></b></div>
                            <small class="text-muted">A pagar: R$ <%= FormatNumber(rsComissoes("ValorDiretoria"), 2) %></small><br>
                            <small class="text-success">Pago: R$ <%= FormatNumber(totalPagoDiretoria, 2) %></small>
                            <% If Not IsEmpty(dataPagamentoDiretoria) Then %>
                                <br>
                            <% End If %>
                            <% If totalPagoDiretoria >= rsComissoes("ValorDiretoria") And rsComissoes("ValorDiretoria") > 0 Then %>
                                <i class="fas fa-check-circle text-success ms-1"></i>
                            <% End If %>
                        </td>
                        <td class="text-center">
                            <div><b><%= rsComissoes("Gerencia") %><br><%= userIdGerencia&"-"& rsComissoes("NomeGerente") %></b></div>
                            <small class="text-muted">A pagar: R$ <%= FormatNumber(rsComissoes("ValorGerencia"), 2) %></small><br>
                            <small class="text-success">Pago: R$ <%= FormatNumber(totalPagoGerencia, 2) %></small>
                            <% If Not IsEmpty(dataPagamentoGerencia) Then %>
                                <br>
                            <% End If %>
                            <% If totalPagoGerencia >= rsComissoes("ValorGerencia") And rsComissoes("ValorGerencia") > 0 Then %>
                                <i class="fas fa-check-circle text-success ms-1"></i>
                            <% End If %>
                        </td>

                        <td class="text-center">
                            <div><b><%= userIdCorretor &"-"&rsComissoes("NomeCorretor") %></b></div>
                            <small class="text-muted">A pagar: R$ <%= FormatNumber(rsComissoes("ValorCorretor"), 2) %></small><br>
                            <small class="text-success">Pago: R$ <%= FormatNumber(totalPagoCorretor, 2) %></small>
                            <% If Not IsEmpty(dataPagamentoCorretor) Then %>
                                <br>
                            <% End If %>
                            <% If totalPagoCorretor >= rsComissoes("ValorCorretor") And rsComissoes("ValorCorretor") > 0 Then %>
                                <i class="fas fa-check-circle text-success ms-1"></i>
                            <% End If %>
                        </td>

                        <td class="text-center">
                            <button class="btn btn-primary btn-sm mb-1" 
                                data-bs-toggle="modal" data-bs-target="#paymentModal"
                                data-id-comissao="<%= rsComissoes("ID_Comissoes") %>"
                                data-id-venda="<%= rsComissoes("ID_Venda") %>"
                                data-diretoria-id="<%= userIdDiretoria %>"
                                data-diretoria-nome="<%= rsComissoes("NomeDiretor") %>"
                                data-diretoria-apagar="<%= FormatNumber(rsComissoes("ValorDiretoria"), 2) %>"
                                data-diretoria-pago="<%= FormatNumber(totalPagoDiretoria, 2) %>"
                                data-gerencia-id="<%= userIdGerencia %>"
                                data-gerencia-nome="<%= rsComissoes("NomeGerente") %>"
                                data-gerencia-apagar="<%= FormatNumber(rsComissoes("ValorGerencia"), 2) %>"
                                data-gerencia-pago="<%= FormatNumber(totalPagoGerencia, 2) %>"
                                data-corretor-id="<%= userIdCorretor %>"
                                data-corretor-nome="<%= rsComissoes("NomeCorretor") %>"
                                data-corretor-apagar="<%= FormatNumber(rsComissoes("ValorCorretor"), 2) %>"
                                data-corretor-pago="<%= FormatNumber(totalPagoCorretor, 2) %>"
                            >
                                <i class="fas fa-hand-holding-usd"></i> Pagar Comiss.
                            </button>

                            <button class="btn btn-info btn-sm mb-1 view-payments-btn"
                                data-bs-toggle="modal" 
                                data-bs-target="#viewPaymentsModal"
                                data-id-venda="<%= rsComissoes("ID_Venda") %>">
                                <i class="fas fa-eye"></i> Ver Pagamentos
                            </button>
                         
                            <button class="btn btn-danger btn-sm" onclick="confirmDelete(<%= rsComissoes("ID_Comissoes") %>)"><i class="fas fa-trash-alt"></i> Excluir</button>
                        </td>
                    </tr>
                    <%
                            rsComissoes.MoveNext
                        Loop
                    Else
                    %>
                    <tr>
                        <td colspan="9" class="text-center">Nenhuma comissão a pagar encontrada.</td>
                    </tr>
                    <%
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Modal de Pagamento -->
    <div class="modal fade" id="paymentModal" tabindex="-1" aria-labelledby="paymentModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="paymentModalLabel">Realizar Pagamento de Comissão</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <form id="paymentForm" action="gestao_vendas_salvar_pagamento.asp" method="post">
                    <div class="modal-body">
                        <input type="hidden" id="modalComissaoId" name="ID_Comissao">
                        <input type="hidden" id="modalVendaId" name="ID_Venda">
                        <input type="hidden" id="modalUserId" name="UserId">

                        <div class="mb-3">
                            <label for="modalRecipient" class="form-label">Para quem será o pagamento?</label>
                            <select class="form-select" id="modalRecipient" name="RecipientType" required>
                                <option value="">Selecione...</option>
                                <!-- Opções preenchidas via JS -->
                            </select>
                        </div>

                        <div class="mb-3">
                            <label class="form-label">Valor Total a Pagar:</label>
                            <p class="form-control-plaintext" id="modalValorAPagarTotal">R$ 0,00</p>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Valor Já Pago:</label>
                            <p class="form-control-plaintext" id="modalValorJaPago">R$ 0,00</p>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Saldo a Pagar:</label>
                            <p class="form-control-plaintext" id="modalSaldoAPagar">R$ 0,00</p>
                        </div>

                        <div class="mb-3">
                            <label for="modalValorAPagarInput" class="form-label">Valor a Pagar (nesta transação) *</label>
                            <input type="text" class="form-control" id="modalValorAPagarInput" name="ValorPago" required>
                        </div>
                        <div class="mb-3">
                            <label for="modalDataPagamento" class="form-label">Data do Pagamento *</label>
                            <input type="date" class="form-control" id="modalDataPagamento" name="DataPagamento" required>
                        </div>
                        <div class="mb-3">
                            <label for="modalStatusPagamento" class="form-label">Status do Pagamento *</label>
                            <select class="form-select" id="modalStatusPagamento" name="Status" required>
                                <option value="">Selecione...</option>
                                <option value="Em processamento">Em processamento</option>
                                <option value="Agendado">Agendado</option>
                                <option value="Realizado">Realizado</option>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="modalObs" class="form-label">Observações</label>
                            <textarea class="form-control" id="modalObs" name="Obs" rows="3"></textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="submit" class="btn btn-primary">Salvar Pagamento</button>
                    </div>
                </form>
            </div>
        </div>
    </div>


<!-- Modal para Visualizar Pagamentos -->
<div class="modal fade" id="viewPaymentsModal" tabindex="-1" aria-labelledby="viewPaymentsModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
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
                                <th class="text-end">Valor</th>
                                <th>Destinatário</th>
                                <th>Tipo</th>
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
<!-- --------------------------------------------------------------------------------------------------------------------- -->
    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- DataTables JS -->
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/responsive/2.2.9/js/dataTables.responsive.min.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/responsive/2.2.9/js/responsive.bootstrap5.min.js"></script>

    <!-- Script principal -->
   <script>
// Função para confirmar a exclusão de uma comissão
    function confirmDelete(id) {
        if (window.confirm("Tem certeza que deseja excluir esta comissão?")) {
            window.location.href = "gestao_comissao_delete.asp?id=" + id;
        }
    }

    // Função para formatar números para exibição em moeda brasileira
    function formatCurrency(value) {
        if (!value && value !== 0) return '0,00';
        return parseFloat(value).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Função para parsear números de moeda brasileira para float
    function parseCurrency(value) {
        if (!value && value !== 0) return 0;
        if (typeof value === 'number') return value;
        return parseFloat(value.replace('R$', '').replace(/\./g, '').replace(',', '.'));
    }

    // Função auxiliar para classes de status
    function getStatusBadgeClass(status) {
        if (!status) return 'bg-secondary';
        status = status.toLowerCase();
        if (status.includes('pago') || status.includes('realizado')) return 'bg-success';
        if (status.includes('pendente')) return 'bg-warning text-dark';
        if (status.includes('processamento') || status.includes('agendado')) return 'bg-primary';
        return 'bg-secondary';
    }

    // Função para formatar datas
    function FormatDateForDisplay(dateString) {
        if (!dateString) return 'N/A';
        const date = new Date(dateString);
        return isNaN(date) ? 'N/A' : date.toLocaleDateString('pt-BR');
    }

    // Função para excluir um pagamento específico
    function deletePayment(paymentId, idVenda) {
        if (confirm(`Tem certeza que deseja excluir o pagamento ID ${paymentId}?`)) {
            $.ajax({
                url: 'gestao_vendas_excluir_pagamento.asp',
                method: 'POST',
                data: { ID_Pagamento: paymentId },
                dataType: 'json',
                success: function(response) {
                    console.log('Resposta do servidor (exclusão):', response);
                    if (response.success) {
                        alert('Pagamento excluído com sucesso!');
                        loadPayments(idVenda); // Recarrega os pagamentos
                    } else {
                        alert('Erro ao excluir pagamento: ' + (response.message || 'Erro desconhecido.'));
                    }
                },
                error: function(xhr, status, error) {
                    console.group("Erro na requisição AJAX de exclusão");
                    console.error("Status:", status);
                    console.error("Mensagem:", error);
                    console.error("Resposta bruta:", xhr.responseText);
                    console.groupEnd();
                    alert('Erro na comunicação com o servidor. Consulte o console para detalhes.');
                }
            });
        }
    }

    // Função para carregar os pagamentos
    function loadPayments(idVenda) {
        $('#paymentsTableBody').html('<tr><td colspan="8" class="text-center"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Carregando...</span></div></td></tr>');
        $('#noPaymentsMessage').hide();

        $.ajax({
            url: 'get_pagamentos_por_comissao.asp',
            type: 'GET',
            dataType: 'json',
            data: { idVenda: idVenda },
            success: function(response) {
                console.log('Resposta recebida para ID_Venda=' + idVenda + ':', response);
                if (response && response.success && response.data && Array.isArray(response.data) && response.data.length > 0) {
                    let html = '';
                    response.data.forEach(function(payment, index) {
                        console.log('Pagamento ##' + (index + 1) + ':', payment);
                        // Converte ValorPago para número, lidando com string ou número
                        const valorPago = (typeof payment.ValorPago === 'string') ? parseFloat(payment.ValorPago.replace(',', '.')) : (payment.ValorPago || 0);
                        html += `
                            <tr>
                                <td>#${payment.ID_Pagamento || 'N/A'}</td>
                                <td>${payment.DataPagamento}</td>
                                <td class="text-end">${formatCurrency(valorPago)}</td>
                                <td>${payment.UsuariosNome || 'N/A'}</td>
                                <td>${(payment.TipoRecebedor || 'N/A').toUpperCase()}</td>
                                <td><span class="badge ${getStatusBadgeClass(payment.Status)}">${payment.Status || 'N/A'}</span></td>
                                <td>${payment.Obs || '-'}</td>
                                <td class="text-center">
                                    <button class="btn btn-danger btn-sm" onclick="deletePayment(${payment.ID_Pagamento}, ${idVenda})">
                                        <i class="fas fa-trash-alt"></i> Excluir
                                    </button>
                                </td>
                            </tr>`;
                    });
                    $('#paymentsTableBody').html(html);
                    $('#noPaymentsMessage').hide();
                } else {
                    console.warn('Nenhum pagamento encontrado ou resposta inválida:', response);
                    $('#paymentsTableBody').html('<tr><td colspan="8" class="text-center">Nenhum pagamento encontrado.</td></tr>');
                    $('#noPaymentsMessage').show();
                }
            },
            error: function(xhr, status, error) {
                console.group('Erro na requisição AJAX para ID_Venda=' + idVenda);
                console.error('Status:', status);
                console.error('Erro:', error);
                console.error('Resposta bruta:', xhr.responseText);
                console.error('Código HTTP:', xhr.status);
                console.groupEnd();
                let errorMessage = 'Erro ao carregar pagamentos. Por favor, tente novamente.';
                try {
                    const errorJson = JSON.parse(xhr.responseText);
                    if (errorJson && errorJson.error) {
                        errorMessage = 'Erro: ' + errorJson.error;
                    }
                } catch (e) {
                    console.warn('Resposta do servidor não é JSON válido:', xhr.responseText);
                }
                $('#paymentsTableBody').html(`
                    <tr>
                        <td colspan="8" class="text-center text-danger">
                            ${errorMessage}
                        </td>
                    </tr>
                `);
                $('#noPaymentsMessage').hide();
            }
        });
    }

    // Inicializa o DataTable e configura os eventos
    $(document).ready(function() {
        $('#comissoesTable').DataTable({
            responsive: true,
            order: [[1, "desc"]],
            pageLength: 100,
            lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "Todos"]],
            language: {
                url: 'https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json'
            },
            dom: '<"top"lf>rt<"bottom"ip>',
            initComplete: function() {
                this.api().columns.adjust().responsive.recalc();
            }
        });

        // Máscara para o campo de valor no modal de pagamento
        $('#modalValorAPagarInput').mask('#.##0,00', { reverse: true });

        // Preenche a data atual no campo de data do modal de pagamento
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        $('#modalDataPagamento').val(`${year}-${month}-${day}`);

        // Evento para abrir o modal de pagamento e preencher os dados
        $('#paymentModal').on('show.bs.modal', function(event) {
            const button = $(event.relatedTarget);
            const idComissao = button.data('id-comissao');
            const idVenda = button.data('id-venda');

            const diretoriaId = button.data('diretoria-id');
            const diretoriaNome = button.data('diretoria-nome');
            const diretoriaAPagar = parseCurrency(button.data('diretoria-apagar'));
            const diretoriaPago = parseCurrency(button.data('diretoria-pago'));

            const gerenciaId = button.data('gerencia-id');
            const gerenciaNome = button.data('gerencia-nome');
            const gerenciaAPagar = parseCurrency(button.data('gerencia-apagar'));
            const gerenciaPago = parseCurrency(button.data('gerencia-pago'));

            const corretorId = button.data('corretor-id');
            const corretorNome = button.data('corretor-nome');
            const corretorAPagar = parseCurrency(button.data('corretor-apagar'));
            const corretorPago = parseCurrency(button.data('corretor-pago'));

            const modal = $(this);
            modal.data('diretoria', { id: diretoriaId, nome: diretoriaNome, apagar: diretoriaAPagar, pago: diretoriaPago });
            modal.data('gerencia', { id: gerenciaId, nome: gerenciaNome, apagar: gerenciaAPagar, pago: gerenciaPago });
            modal.data('corretor', { id: corretorId, nome: corretorNome, apagar: corretorAPagar, pago: corretorPago });

            $('#modalComissaoId').val(idComissao);
            $('#modalVendaId').val(idVenda);

            const recipientSelect = $('#modalRecipient');
            recipientSelect.empty();
            recipientSelect.append('<option value="">Selecione...</option>');

            if (diretoriaId && diretoriaNome && diretoriaId !== 0) {
                recipientSelect.append(`<option value="diretoria" data-user-id="${diretoriaId}">${diretoriaNome} (Diretoria)</option>`);
            }
            if (gerenciaId && gerenciaNome && gerenciaId !== 0) {
                recipientSelect.append(`<option value="gerencia" data-user-id="${gerenciaId}">${gerenciaNome} (Gerência)</option>`);
            }
            if (corretorId && corretorNome && corretorId !== 0) {
                recipientSelect.append(`<option value="corretor" data-user-id="${corretorId}">${corretorNome} (Corretor)</option>`);
            }

            $('#modalValorAPagarTotal').text('R$ 0,00');
            $('#modalValorJaPago').text('R$ 0,00');
            $('#modalSaldoAPagar').text('R$ 0,00');
            $('#modalValorAPagarInput').val('');
            $('#modalUserId').val('');
            $('#modalObs').val('');
            $('#modalStatusPagamento').val('');
        });

        // Evento para atualizar os valores quando o destinatário é selecionado no modal de pagamento
        $('#modalRecipient').change(function() {
            const selectedType = $(this).val();
            const modal = $('#paymentModal');
            let data = null;
            let userId = '';

            if (selectedType === 'diretoria') {
                data = modal.data('diretoria');
                userId = data.id;
            } else if (selectedType === 'gerencia') {
                data = modal.data('gerencia');
                userId = data.id;
            } else if (selectedType === 'corretor') {
                data = modal.data('corretor');
                userId = data.id;
            }

            if (data) {
                const saldo = data.apagar - data.pago;
                $('#modalValorAPagarTotal').text('R$ ' + formatCurrency(data.apagar));
                $('#modalValorJaPago').text('R$ ' + formatCurrency(data.pago));
                $('#modalSaldoAPagar').text('R$ ' + formatCurrency(saldo));
                $('#modalValorAPagarInput').val(formatCurrency(saldo));
                $('#modalUserId').val(userId);
            } else {
                $('#modalValorAPagarTotal').text('R$ 0,00');
                $('#modalValorJaPago').text('R$ 0,00');
                $('#modalSaldoAPagar').text('R$ 0,00');
                $('#modalValorAPagarInput').val('');
                $('#modalUserId').val('');
            }
        });

        // Validação do formulário de pagamento antes de enviar
        $('#paymentForm').submit(function(e) {
            const valorPagoInput = $('#modalValorAPagarInput').val();
            const valorPago = parseCurrency(valorPagoInput);
            const saldoAPagar = parseCurrency($('#modalSaldoAPagar').text());

            if (valorPago <= 0) {
                alert('O valor a pagar deve ser maior que zero.');
                e.preventDefault();
                return;
            }

            if (valorPago > saldoAPagar) {
                alert('O valor a pagar não pode ser maior que o saldo a pagar.');
                e.preventDefault();
                return;
            }

            if ($('#modalRecipient').val() === '') {
                alert('Por favor, selecione para quem será o pagamento.');
                e.preventDefault();
                return;
            }

            if ($('#modalDataPagamento').val() === '') {
                alert('Por favor, selecione a data do pagamento.');
                e.preventDefault();
                return;
            }

            if ($('#modalStatusPagamento').val() === '') {
                alert('Por favor, selecione o status do pagamento.');
                e.preventDefault();
                return;
            }
        });

        // Evento para abrir o modal de visualização de pagamentos e carregar os dados
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
' Fechar recordsets e conexão
' ====================================================================
If IsObject(rsComissoes) Then
    rsComissoes.Close
    Set rsComissoes = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If
%>