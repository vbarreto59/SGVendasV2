<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% 
If Len(StrConn) = 0 Then 
%>
    <!--#include file="conexao.asp"-->
<%  
End If  

If Len(StrConnSales) = 0 Then  
%>
    <!--#include file="conSunSales.asp"-->
<%
End If
%>

<!--#include file="gestao_header.inc"-->

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"     

' Obter o ID da venda da query string
Dim idVenda
idVenda = Request.QueryString("id")

If idVenda = "" Or Not IsNumeric(idVenda) Then
    Response.Write "<div class='alert alert-danger'>ID da venda inválido.</div>"
    Response.End
End If

' Conexão com o banco
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Buscar informações da venda
Dim rsVenda, sqlVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
sqlVenda = "SELECT * FROM Vendas WHERE ID = " & idVenda
rsVenda.Open sqlVenda, connSales

If rsVenda.EOF Then
    Response.Write "<div class='alert alert-danger'>Venda não encontrada.</div>"
    rsVenda.Close
    Set rsVenda = Nothing
    connSales.Close
    Set connSales = Nothing
    Response.End
End If

' Buscar pagamentos da venda
Dim rsPagamentos, sqlPagamentos
Set rsPagamentos = Server.CreateObject("ADODB.Recordset")
sqlPagamentos = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & idVenda & " AND (Excluido <> -1 OR Excluido IS NULL) ORDER BY DataPagamento DESC"
rsPagamentos.Open sqlPagamentos, connSales
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pagamentos da Venda #<%= idVenda %> | Sistema</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            color: #2c3e50;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 20px;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
        }
        
        .card-header {
            background: linear-gradient(to right, #2c3e50, #3498db);
            color: white;
            border-bottom: none;
            padding: 1rem 1.5rem;
            font-weight: 600;
            border-radius: 12px 12px 0 0 !important;
        }
        
        .table th {
            background-color: #2c3e50;
            color: white;
            font-weight: 600;
            border: none;
            padding: 1rem 0.75rem;
        }
        
        .table td {
            padding: 0.75rem;
            vertical-align: middle;
            border-color: #e9ecef;
        }
        
        .btn-danger {
            background-color: #e74c3c;
            border-color: #e74c3c;
        }
        
        .btn-danger:hover {
            background-color: #c0392b;
            border-color: #c0392b;
        }
        
        .status-pago {
            background-color: #27ae60;
            color: black;
        }
        
        .status-pendente {
            background-color: #e74c3c;
            color: blue;
        }
        
        .info-venda {
            background-color: #e3f2fd;
            border-left: 4px solid #3498db;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        
        .total-pago {
            font-size: 1.2em;
            font-weight: bold;
            color: #27ae60;
        }
        
        .badge-sm {
            font-size: 0.7em;
            padding: 0.3em 0.6em;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row mb-4">
            <div class="col-12">
                <div class="d-flex justify-content-between align-items-center">
                    <h1><i class="fas fa-money-bill-wave me-2"></i>Pagamentos da Venda</h1>
                    <button type="button" onclick="window.close();" class="btn btn-secondary">
                        <i class="fas fa-times me-1"></i>Fechar
                    </button>
                </div>
            </div>
        </div>

        <!-- Informações da Venda -->
        <div class="card mb-4">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-info-circle me-2"></i>Informações da Venda #<%= idVenda %></h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-md-3">
                        <strong>Empreendimento:</strong><br>
                        <%= rsVenda("Empreend_ID") %> - <%= rsVenda("NomeEmpreendimento") %>
                    </div>
                    <div class="col-md-2">
                        <strong>Unidade:</strong><br>
                        <%= rsVenda("Unidade") %>
                    </div>
                    <div class="col-md-2">
                        <strong>Valor:</strong><br>
                        R$ <%= FormatNumber(rsVenda("ValorUnidade"), 2) %>
                    </div>
                    <div class="col-md-3">
                        <strong>Data:</strong><br>
                        <%= rsVenda("DiaVenda") %>/<%= rsVenda("MesVenda") %>/<%= rsVenda("AnoVenda") %>
                    </div>

                </div>
            </div>
        </div>

        <!-- Tabela de Pagamentos -->
        <div class="card">
            <div class="card-header d-flex justify-content-between align-items-center">
                <h5 class="mb-0"><i class="fas fa-list me-2"></i>Pagamentos Realizados</h5>
                <%
                ' Calcular total pago
                Dim totalPago
                totalPago = 0
                If Not rsPagamentos.EOF Then
                    rsPagamentos.MoveFirst
                    Do While Not rsPagamentos.EOF
                        If Not IsNull(rsPagamentos("ValorPago")) And IsNumeric(rsPagamentos("ValorPago")) Then
                            totalPago = totalPago + CDbl(rsPagamentos("ValorPago"))
                        End If
                        rsPagamentos.MoveNext
                    Loop
                    rsPagamentos.MoveFirst
                End If
                %>
                <span class="total-pago">Total Pago: R$ <%= FormatNumber(totalPago, 2) %></span>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table class="table table-hover table-striped">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Data Pagamento</th>
                                <th>Tipo Recebedor</th>
                                
                                <th>Valor Pago (R$)</th>
                                <th>Tipo Pagamento</th>
                                <th>Status</th>
                                <th>Observações</th>
                                <th>Registro</th>
                                <th width="100">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsPagamentos.EOF Then
                                Do While Not rsPagamentos.EOF
                            %>
                            <tr>
                                <td>
                                    <small class="text-muted">#<%= rsPagamentos("ID_Pagamento") %></small>
                                </td>
                                <td>
                                    <% If Not IsNull(rsPagamentos("DataPagamento")) Then %>
                                        <%= FormatDateTime(rsPagamentos("DataPagamento"), 2) %>
                                    <% Else %>
                                        <span class="text-muted">-</span>
                                    <% End If %>
                                </td>
                                <td>
                                    <span class="badge bg-primary badge-sm">
                                        <%= UCase(rsPagamentos("TipoRecebedor")) %>
                                    </span>
                                    <% If Not IsNull(rsPagamentos("UsuariosNome")) Then %>
                                        <br><small><%= rsPagamentos("UsuariosNome") %></small>
                                    <% End If %>
                                </td>
    
                                <td class="fw-bold text-success">
                                    R$ <%= FormatNumber(rsPagamentos("ValorPago"), 2) %>
                                </td>
                                <td>
                                    <span class="badge bg-info badge-sm">
                                        <%= rsPagamentos("TipoPagamento") %>
                                    </span>
                                </td>
                                <td>
                                    <span class="badge btn-success %>">
                                        <%= rsPagamentos("Status") %>
                                    </span>
                                </td>
                                <td>
                                    <% If Not IsNull(rsPagamentos("Obs")) And rsPagamentos("Obs") <> "" Then %>
                                        <span title="<%= Server.HTMLEncode(rsPagamentos("Obs")) %>">
                                            <i class="fas fa-info-circle text-info"></i>
                                            <%= Left(rsPagamentos("Obs"), 30) %>
                                            <% If Len(rsPagamentos("Obs")) > 30 Then %>...<% End If %>
                                        </span>
                                    <% Else %>
                                        <span class="text-muted">-</span>
                                    <% End If %>
                                </td>
                                <td>
                                    <small>
                                        <% If Not IsNull(rsPagamentos("DataHora")) Then %>
                                            <%= FormatDateTime(rsPagamentos("DataHora"), 2) %>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </small>
                                </td>
                                <td>
                                    <button type="button" class="btn btn-danger btn-sm" 
                                                onclick="excluirPagamento(<%= rsPagamentos("ID_Pagamento") %>, <%= idVenda %>)"
                                                title="Excluir Pagamento">
                                            <i class="fas fa-trash"></i>
                                    </button>
                                </td>
                            </tr>
                            <%
                                    rsPagamentos.MoveNext
                                Loop ' CORREÇÃO: Usado 'Loop' para fechar o Do While
                            Else
                            %>
                            <tr>
                                <td colspan="10" class="text-center text-muted py-4">
                                    <i class="fas fa-info-circle fa-2x mb-3"></i><br>
                                    Nenhum pagamento encontrado para esta venda.
                                </td>
                            </tr>
                            <%
                            End If ' Fechamento do bloco If
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Confirmação -->
    <div class="modal fade" id="modalExcluir" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Confirmar Exclusão</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <p>Tem certeza que deseja excluir este pagamento?</p>
                    <p class="fw-bold">Esta ação não pode ser desfeita.</p>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-danger" id="btnConfirmarExclusao">Excluir</button>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    


<script>
let pagamentoIdParaExcluir = null;
let vendaIdParaExcluir = null;

// Função chamada pelo botão de exclusão
function excluirPagamento(pagamentoId, vendaId) {
    pagamentoIdParaExcluir = pagamentoId;
    vendaIdParaExcluir = vendaId;
    $('#modalExcluir').modal('show');
}

// Confirmação da exclusão (dentro do modal)
$('#btnConfirmarExclusao').click(function() {
    if (pagamentoIdParaExcluir && vendaIdParaExcluir) {
        // Esconde o modal e mostra loading
        $('#modalExcluir').modal('hide'); 
        
        // Desabilita botão para evitar múltiplos cliques
        $(this).prop('disabled', true).html('<i class="fas fa-spinner fa-spin"></i> Excluindo...');

        // Fazer requisição para excluir o pagamento
        $.ajax({
            url: 'excluir_pagamento.asp',
            type: 'POST',
            data: {
                id_pagamento: pagamentoIdParaExcluir,
                id_venda: vendaIdParaExcluir
            },
            success: function(response) {
                // Reabilita o botão
                $('#btnConfirmarExclusao').prop('disabled', false).html('Excluir');
                
                // Verifica se a resposta do VBScript é SUCESSO
                if (response.trim() === 'SUCESSO') {
                    // Sucesso: alerta e recarrega
                    alert('Pagamento excluído com sucesso!');
                    location.reload();
                } else {
                    // Erro: exibe mensagem
                    alert('Erro ao excluir pagamento: ' + response);
                }
            },
            error: function(xhr, status, error) {
                // Reabilita o botão em caso de erro
                $('#btnConfirmarExclusao').prop('disabled', false).html('Excluir');
                
                console.error("Erro AJAX:", status, error, xhr.responseText);
                alert('Erro de comunicação. Tente novamente. Detalhes: ' + error);
            }
        });
    }
});
</script>
</body>
</html>

<%
' Fechar conexões
If Not rsVenda Is Nothing Then
    rsVenda.Close
    Set rsVenda = Nothing
End If

If Not rsPagamentos Is Nothing Then
    rsPagamentos.Close
    Set rsPagamentos = Nothing
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>