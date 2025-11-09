<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<% ' botao detalhes'
' Primeiro, vamos verificar se a tabela existe e tem dados
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

Response.Write "<!-- Conectado ao banco: " & StrConnSales & " --><br>"

' Verificar se a tabela VENDA_TEMP existe
On Error Resume Next
Set rsTest = conn.Execute("SELECT COUNT(*) as Total FROM VENDA_TEMP")
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>ERRO: Tabela VENDA_TEMP não existe ou não pode ser acessada. Erro: " & Err.Description & "</div>"
    hasRecords = False
Else
    If rsTest("Total") > 0 Then
       '' Response.Write "<div class='alert alert-success'>Tabela VENDA_TEMP encontrada com " & rsTest("Total") & " registros.</div>"
        hasRecords = True
    Else
        Response.Write "<div class='alert alert-warning'>Tabela VENDA_TEMP existe mas está vazia.</div>"
        hasRecords = False
    End If
    rsTest.Close
End If
On Error GoTo 0

' Se não tem registros, vamos verificar a tabela VENDAS original
If Not hasRecords Then
    Response.Write "<div class='alert alert-info'>Verificando tabela VENDAS original...</div>"
    
    On Error Resume Next
    Set rsVendas = conn.Execute("SELECT COUNT(*) as Total FROM VENDAS WHERE (Excluido <> -1 OR Excluido IS NULL)")
    If Err.Number = 0 Then
        Response.Write "<div class='alert alert-info'>Tabela VENDAS tem " & rsVendas("Total") & " registros ativos.</div>"
        
        ' Vamos mostrar algumas vendas como exemplo
        Set rsExemplo = conn.Execute("SELECT TOP 5 ID, Empreendimento, Corretor FROM VENDAS WHERE (Excluido <> -1 OR Excluido IS NULL) ORDER BY ID DESC")
        If Not rsExemplo.EOF Then
            Response.Write "<div class='alert alert-info'><strong>Últimas 5 vendas:</strong><br>"
            Do While Not rsExemplo.EOF
                Response.Write "ID: " & rsExemplo("ID") & " - " & rsExemplo("Empreendimento") & " - " & rsExemplo("Corretor") & "<br>"
                rsExemplo.MoveNext
            Loop
            Response.Write "</div>"
        End If
        rsExemplo.Close
    Else
        Response.Write "<div class='alert alert-danger'>Erro ao acessar tabela VENDAS: " & Err.Description & "</div>"
    End If
    On Error GoTo 0
End If

' Agora vamos tentar a consulta principal
If hasRecords Then
    sql = "SELECT " & _
          "VT.UserID, " & _
          "VT.Diretoria, " & _
          "VT.Gerencia, " & _
          "VT.Nome, " & _
          "VT.Cargo, " & _
          "VT.VUnid, " & _
          "VT.ID_Venda, " & _
          "VT.VBruto, " & _
          "VT.Desc, " & _
          "VT.VLiq, " & _
          "VT.Premio, " & _
          "VT.VTotal, " & _
          "(SELECT SUM(ValorPago) FROM PAGAMENTOS_COMISSOES PC " & _
          " WHERE PC.UsuariosUserId = VT.UserID AND PC.ID_Venda = VT.ID_Venda) AS SomaDeValorPago " & _
          "FROM VENDA_TEMP AS VT " & _
          "ORDER BY VT.ID_Venda, VT.Nome"

    Response.Write "<!-- SQL: " & sql & " -->"

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn
    hasRecords = Not rs.EOF
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Completo - Vendas e Pagamentos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        .table th { background-color: #800020; color: white; }
        .bg-pago { background-color: #d4edda; }
        .bg-pendente { background-color: #fff3cd; }
        .bg-parcial { background-color: #e2e3e5; }
        .valor-positivo { color: #198754; font-weight: bold; }
        .valor-negativo { color: #dc3545; font-weight: bold; }
        .valor-zero { color: #6c757d; }
        .saldo-pendente { background-color: #f8d7da; }
        .btn-detalhes { font-size: 0.8rem; padding: 0.25rem 0.5rem; }
        .valor-numero { font-family: 'Courier New', monospace; text-align: right; }
    </style>
</head>
<body>

<div class="container mt-4">
    <h1 class="mb-4 text-center"><i class="fas fa-file-invoice-dollar me-2"></i>Relatório Completo - Vendas e Pagamentos</h1>

    <% If hasRecords Then %>
    <div class="row mb-3">
        <div class="col-md-4">
            <div class="card text-white bg-success">
                <div class="card-body">
                    <h6 class="card-title">Total Comissões</h6>
                    <%
                    Dim totalComissoes
                    totalComissoes = 0
                    If hasRecords Then
                        Do While Not rs.EOF
                            If Not IsNull(rs("VTotal")) Then
                                totalComissoes = totalComissoes + CDbl(rs("VTotal"))
                            End If
                            rs.MoveNext
                        Loop
                        rs.MoveFirst
                    End If
                    %>
                    <h4><%= FormatNumber(totalComissoes, 2) %></h4>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card text-white bg-primary">
                <div class="card-body">
                    <h6 class="card-title">Total Pago</h6>
                    <%
                    Dim totalPago
                    totalPago = 0
                    If hasRecords Then
                        Do While Not rs.EOF
                            Dim valorPagoTemp
                            If Not IsNull(rs("SomaDeValorPago")) Then
                                valorPagoTemp = CDbl(rs("SomaDeValorPago"))
                            Else
                                valorPagoTemp = 0
                            End If
                            totalPago = totalPago + valorPagoTemp
                            rs.MoveNext
                        Loop
                        rs.MoveFirst
                    End If
                    %>
                    <h4><%= FormatNumber(totalPago, 2) %></h4>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card text-white bg-warning">
                <div class="card-body">
                    <h6 class="card-title">Saldo Pendente</h6>
                    <h4><%= FormatNumber(totalComissoes - totalPago, 2) %></h4>
                </div>
            </div>
        </div>
    </div>

    <table id="myTable" class="table table-striped table-bordered table-sm" style="width:100%">
        <thead>
            <tr>
                <th>ID Venda</th>
                <th>Nome</th>
                <th>Cargo</th>
                <th>Diretoria</th>
                <th>Gerencia</th>
                <th class="text-end">V. Unid</th>
                <th class="text-end">V. Bruto</th>
                <th class="text-end">Desconto</th>
                <th class="text-end">V. Líquido</th>
                <th class="text-end">Prêmio</th>
                <th class="text-end">Total Comissão</th>
                <th class="text-end">Valor Pago</th>
                <th class="text-end">Saldo</th>
                <th>Status</th>
                <th>Ações</th>
            </tr>
        </thead>
        <tbody>
            <% 
            Do While Not rs.EOF
                Dim vVBruto, vDesc, vVLiq, vPremio, vVTotal, vValorPago, vSaldo
                Dim statusPagamento, statusClass, saldoClass, badgeClass
                
                vVBruto = 0
                vDesc = 0
                vVLiq = 0
                vPremio = 0
                vVTotal = 0
                vValorPago = 0

                If Not IsNull(rs("VBruto")) Then vVBruto = CDbl(rs("VBruto"))
                If Not IsNull(rs("Desc")) Then vDesc = CDbl(rs("Desc"))
                If Not IsNull(rs("VLiq")) Then vVLiq = CDbl(rs("VLiq"))
                If Not IsNull(rs("Premio")) Then vPremio = CDbl(rs("Premio"))
                If Not IsNull(rs("VTotal")) Then vVTotal = CDbl(rs("VTotal"))

                If Not IsNull(rs("SomaDeValorPago")) Then 
                    vValorPago = CDbl(rs("SomaDeValorPago"))
                Else
                    vValorPago = 0
                End If

                vSaldo = vVTotal - vValorPago

                If vValorPago = 0 Then
                    statusPagamento = "Pendente"
                    statusClass = "bg-pendente"
                    badgeClass = "bg-warning text-dark"
                ElseIf vSaldo = 0 Then
                    statusPagamento = "Pago"
                    statusClass = "bg-pago"
                    badgeClass = "bg-success"
                ElseIf vValorPago > 0 And vSaldo > 0 Then
                    statusPagamento = "Parcial"
                    statusClass = "bg-parcial"
                    badgeClass = "bg-info"
                Else
                    statusPagamento = "Estornado"
                    statusClass = "bg-danger text-white"
                    badgeClass = "bg-danger"
                End If
                
                If vSaldo > 0 Then
                    saldoClass = "saldo-pendente"
                Else
                    saldoClass = ""
                End If
            %>
            <tr>
                <td><strong><%= "V"&rs("ID_Venda") %></strong></td>
                <td><strong><%= rs("Nome") %></strong></td>
                <td><%= rs("Cargo") %></td>
                <td><%= rs("Diretoria") %></td>
                <td><%= rs("Gerencia") %></td>
                <td class="valor-numero"><%= FormatNumber(rs("VUnid"),2) %></td>
                <td class="valor-numero"><%= FormatNumber(vVBruto, 2) %></td>
                <td class="valor-numero"><%= FormatNumber(vDesc, 2) %></td>
                <td class="valor-numero"><%= FormatNumber(vVLiq, 2) %></td>
                <td class="valor-numero"><%= FormatNumber(vPremio, 2) %></td>
                <td class="valor-numero"><strong><%= FormatNumber(vVTotal, 2) %></strong></td>
                <td class="valor-numero"><%= FormatNumber(vValorPago, 2) %></td>
                <td class="valor-numero <%= saldoClass %>">
                    <span class="<% 
                    If vSaldo > 0 Then 
                        Response.Write "valor-negativo"
                    ElseIf vSaldo < 0 Then
                        Response.Write "valor-positivo" 
                    Else
                        Response.Write "valor-zero"
                    End If %>">
                        <strong><%= FormatNumber(vSaldo, 2) %></strong>
                    </span>
                </td>
                <td class="<%= statusClass %>">
                    <span class="badge <%= badgeClass %>">
                        <%= statusPagamento %>
                    </span>
                </td>
                <td>
                    <button class="btn btn-sm btn-outline-primary btn-detalhes" 
                            data-bs-toggle="modal" 
                            data-bs-target="#modalDetalhes"
                            data-id-venda="<%= rs("ID_Venda") %>"
                            data-user-id="<%= rs("UserID") %>"
                            data-nome="<%= rs("Nome") %>">
                        <i class="fas fa-search me-1"></i>Detalhes
                    </button>
                </td>
            </tr>
            <% 
                rs.MoveNext
            Loop
            %>
        </tbody>
    </table>
    <% Else %>
    <div class="alert alert-danger text-center">
        <i class="fas fa-exclamation-triangle me-2"></i>
        <strong>Não foi possível gerar o relatório.</strong><br>
        A tabela VENDA_TEMP está vazia ou não existe.<br>
        <a href="AtualizarVendasTemp.asp" class="btn btn-warning mt-2">
            <i class="fas fa-sync-alt me-1"></i>Executar Atualização dos Dados
        </a>
    </div>
    <% End If %>
</div>

<!-- Modal para Detalhes dos Pagamentos -->
<div class="modal fade" id="modalDetalhes" tabindex="-1" aria-labelledby="modalDetalhesLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header bg-primary text-white">
                <h5 class="modal-title" id="modalDetalhesLabel">
                    <i class="fas fa-file-invoice-dollar me-2"></i>Detalhes dos Pagamentos
                </h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <div id="loadingDetalhes" class="text-center py-4">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                    <p class="mt-2">Carregando detalhes dos pagamentos...</p>
                </div>
                <div id="conteudoDetalhes" style="display: none;">
                    <h6 id="infoVenda" class="mb-3"></h6>
                    <div id="tabelaDetalhes"></div>
                </div>
                <div id="semDetalhes" class="text-center py-4" style="display: none;">
                    <i class="fas fa-info-circle fa-2x text-muted mb-3"></i>
                    <p class="text-muted">Nenhum pagamento encontrado para esta venda.</p>
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Fechar</button>
            </div>
        </div>
    </div>
</div>

<% If hasRecords Then
    rs.Close
    Set rs = Nothing
End If
conn.Close
Set conn = Nothing
%>

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>

<% If hasRecords Then %>
<script>
    $(document).ready(function() {
        $('#myTable').DataTable({
            "order": [[0, "desc"]],
            "pageLength": 50,
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            "columnDefs": [
                { "type": "num-fmt", "targets": [5,6,7,8,9,10,11,12] },
                { "orderable": false, "targets": [14] }
            ]
        });

        // Evento para abrir o modal de detalhes
        $('.btn-detalhes').on('click', function() {
            var idVenda = $(this).data('id-venda');
            var userId = $(this).data('user-id');
            var nome = $(this).data('nome');
            
            $('#modalDetalhesLabel').html('<i class="fas fa-file-invoice-dollar me-2"></i>Detalhes dos Pagamentos - Venda ' + idVenda);
            $('#infoVenda').html('<strong>Venda:</strong> ' + idVenda);
            
            $('#loadingDetalhes').show();
            $('#conteudoDetalhes').hide();
            $('#semDetalhes').hide();
            
            $.ajax({
                url: 'obter_detalhes_pagamentos.asp',
                type: 'POST',
                data: {
                    id_venda: idVenda,
                    nome: nome
                },
                success: function(response) {
                    $('#loadingDetalhes').hide();
                    $('#tabelaDetalhes').html(response);
                    $('#conteudoDetalhes').show();
                },
                error: function(xhr, status, error) {
                    $('#loadingDetalhes').hide();
                    $('#tabelaDetalhes').html(
                        '<div class="alert alert-danger">' +
                        '<i class="fas fa-exclamation-triangle me-2"></i>' +
                        'Erro ao carregar os detalhes dos pagamentos. Status: ' + status + ', Erro: ' + error +
                        '</div>'
                    );
                    $('#conteudoDetalhes').show();
                }
            });
        });
        
        $('#modalDetalhes').on('hidden.bs.modal', function() {
            $('#tabelaDetalhes').html('');
            $('#conteudoDetalhes').hide();
            $('#semDetalhes').hide();
            $('#loadingDetalhes').show();
        });
    });
</script>
<% End If %>
<!--#include file="footer.inc"-->
</body>
</html>