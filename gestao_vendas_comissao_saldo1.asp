<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' Consulta para o relatório
Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

' Primeiro, verificar se a view VENDAS_RESUMO existe e tem dados
Dim rsCheck, sqlCheck
sqlCheck = "SELECT COUNT(*) as total FROM VENDAS_RESUMO"
Set rsCheck = conn.Execute(sqlCheck)
Dim hasDataInView
hasDataInView = (rsCheck("total") > 0)
rsCheck.Close
Set rsCheck = Nothing

If hasDataInView Then
    ' Usar a view existente
    sql = "SELECT " & _
          "NomeRecebedor, " & _
          "TipoRecebedor, " & _
          "SUM(TotalVenda) AS TotalVenda, " & _
          "SUM(ComissaoDevida) AS ComissaoDevida, " & _
          "SUM(PremioDevido) AS PremioDevido, " & _
          "SUM(ComissaoPaga) AS ComissaoPaga, " & _
          "SUM(PremioPago) AS PremioPago, " & _
          "SUM(ComissaoDevida) - SUM(ComissaoPaga) AS SaldoComissao, " & _
          "SUM(PremioDevido) - SUM(PremioPago) AS SaldoPremio, " & _
          "(SUM(ComissaoDevida) - SUM(ComissaoPaga)) + (SUM(PremioDevido) - SUM(PremioPago)) AS SaldoTotal " & _
          "FROM VENDAS_RESUMO " & _
          "GROUP BY NomeRecebedor, TipoRecebedor " & _
          "ORDER BY NomeRecebedor"
Else
    ' Consulta alternativa usando dados diretos das tabelas
    sql = "SELECT " & _
          "p.UsuariosNome as NomeRecebedor, " & _
          "p.TipoRecebedor, " & _
          "0 AS TotalVenda, " & _
          "0 AS ComissaoDevida, " & _
          "0 AS PremioDevido, " & _
          "SUM(CASE WHEN p.TipoPagamento = 'Comissão' THEN p.ValorPago ELSE 0 END) AS ComissaoPaga, " & _
          "SUM(CASE WHEN p.TipoPagamento = 'Premiação' THEN p.ValorPago ELSE 0 END) AS PremioPago, " & _
          "0 - SUM(CASE WHEN p.TipoPagamento = 'Comissão' THEN p.ValorPago ELSE 0 END) AS SaldoComissao, " & _
          "0 - SUM(CASE WHEN p.TipoPagamento = 'Premiação' THEN p.ValorPago ELSE 0 END) AS SaldoPremio, " & _
          "(0 - SUM(CASE WHEN p.TipoPagamento = 'Comissão' THEN p.ValorPago ELSE 0 END)) + " & _
          "(0 - SUM(CASE WHEN p.TipoPagamento = 'Premiação' THEN p.ValorPago ELSE 0 END)) AS SaldoTotal " & _
          "FROM pagamentos_comissoes p " & _
          "WHERE p.Status = 'Realizado' " & _
          "GROUP BY p.UsuariosNome, p.TipoRecebedor " & _
          "ORDER BY p.UsuariosNome"
End If

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

Dim hasRecords
hasRecords = Not rs.EOF
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Consolidado - Vendas x Pagamentos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        .table th { background-color: #800020; color: white; }
        .bg-comissao { background-color: #e8f5e8; }
        .bg-premio { background-color: #fff3cd; }
        .bg-total { background-color: #d1ecf1; }
        .valor-positivo { color: #dc3545; font-weight: bold; }
        .valor-negativo { color: #198754; font-weight: bold; }
        .valor-zero { color: #6c757d; }
        .alert-debug { background-color: #fff3cd; border-color: #ffeaa7; }
    </style>
</head>
<body>

<div class="container mt-4">
    <h1 class="mb-4 text-center"><i class="fas fa-file-invoice-dollar me-2"></i>Relatório Consolidado</h1>

    <!-- Debug Information -->
    <div class="alert alert-debug mb-3">
        <i class="fas fa-bug me-2"></i>
        <strong>Informações do Sistema:</strong>
        <br>View VENDAS_RESUMO: <% If hasDataInView Then Response.Write "<span class='text-success'>Com dados (" & hasDataInView & ")</span>" Else Response.Write "<span class='text-danger'>Sem dados ou não existe</span>" End If %>
        <br>Consulta usada: <% If hasDataInView Then Response.Write "VENDAS_RESUMO" Else Response.Write "pagamentos_comissoes (alternativa)" End If %>
        <br>Registros encontrados: <%= hasRecords %>
    </div>

    <div class="alert alert-info">
        <i class="fas fa-info-circle me-2"></i>
        <strong>Fontes:</strong> Vendas (Valores Devidos) + Pagamentos (Valores Pagos)
        <a href="popular_resumo.asp" class="btn btn-warning btn-sm float-end">
            <i class="fas fa-sync me-1"></i>Atualizar Dados
        </a>
    </div>

    <table id="myTable" class="table table-striped table-bordered table-sm" style="width:100%">
        <thead>
            <tr>
                <th>Nome</th>
                <th>Tipo</th>
                <th class="text-end">Total Venda</th>
                <th class="text-end">Comissão Devida</th>
                <th class="text-end">Prêmio Devido</th>
                <th class="text-end">Comissão Paga</th>
                <th class="text-end">Prêmio Pago</th>
                <th class="text-end bg-comissao">Saldo Comissão</th>
                <th class="text-end bg-premio">Saldo Prêmio</th>
                <th class="text-end bg-total">Saldo Total</th>
            </tr>
        </thead>
        <tbody>
            <% 
            If hasRecords Then
                Do While Not rs.EOF
                    Dim vTotalVenda, vComissaoDevida, vPremioDevido, vComissaoPaga, vPremioPaga
                    Dim vSaldoComissao, vSaldoPremio, vSaldoTotal
                    
                    ' Usar CDbl com verificação para evitar erros
                    vTotalVenda = 0
                    vComissaoDevida = 0
                    vPremioDevido = 0
                    vComissaoPaga = 0
                    vPremioPaga = 0
                    vSaldoComissao = 0
                    vSaldoPremio = 0
                    vSaldoTotal = 0
                    
                    If Not IsNull(rs("TotalVenda")) Then vTotalVenda = CDbl(rs("TotalVenda"))
                    If Not IsNull(rs("ComissaoDevida")) Then vComissaoDevida = CDbl(rs("ComissaoDevida"))
                    If Not IsNull(rs("PremioDevido")) Then vPremioDevido = CDbl(rs("PremioDevido"))
                    If Not IsNull(rs("ComissaoPaga")) Then vComissaoPaga = CDbl(rs("ComissaoPaga"))
                    If Not IsNull(rs("PremioPago")) Then vPremioPaga = CDbl(rs("PremioPago"))
                    If Not IsNull(rs("SaldoComissao")) Then vSaldoComissao = CDbl(rs("SaldoComissao"))
                    If Not IsNull(rs("SaldoPremio")) Then vSaldoPremio = CDbl(rs("SaldoPremio"))
                    If Not IsNull(rs("SaldoTotal")) Then vSaldoTotal = CDbl(rs("SaldoTotal"))
            %>
            <tr>
                <td><strong><%= rs("NomeRecebedor") %></strong></td>
                <td>
                    <span class="badge 
                    <% Select Case rs("TipoRecebedor")
                        Case "corretor"
                            Response.Write "bg-primary"
                        Case "gerencia" 
                            Response.Write "bg-warning text-dark"
                        Case "diretoria"
                            Response.Write "bg-success"
                        Case Else
                            Response.Write "bg-secondary"
                    End Select %>">
                        <%= rs("TipoRecebedor") %>
                    </span>
                </td>
                <td class="text-end" data-numeric-value="<%= vTotalVenda %>">
                    R$ <%= FormatNumber(vTotalVenda, 2) %>
                </td>
                <td class="text-end" data-numeric-value="<%= vComissaoDevida %>">
                    R$ <%= FormatNumber(vComissaoDevida, 2) %>
                </td>
                <td class="text-end" data-numeric-value="<%= vPremioDevido %>">
                    R$ <%= FormatNumber(vPremioDevido, 2) %>
                </td>
                <td class="text-end" data-numeric-value="<%= vComissaoPaga %>">
                    R$ <%= FormatNumber(vComissaoPaga, 2) %>
                </td>
                <td class="text-end" data-numeric-value="<%= vPremioPaga %>">
                    R$ <%= FormatNumber(vPremioPaga, 2) %>
                </td>
                <td class="text-end bg-comissao" data-numeric-value="<%= vSaldoComissao %>">
                    <span class="<% 
                    If vSaldoComissao > 0 Then 
                        Response.Write "valor-positivo"
                    ElseIf vSaldoComissao < 0 Then
                        Response.Write "valor-negativo" 
                    Else
                        Response.Write "valor-zero"
                    End If %>">
                        R$ <%= FormatNumber(vSaldoComissao, 2) %>
                    </span>
                </td>
                <td class="text-end bg-premio" data-numeric-value="<%= vSaldoPremio %>">
                    <span class="<% 
                    If vSaldoPremio > 0 Then 
                        Response.Write "valor-positivo"
                    ElseIf vSaldoPremio < 0 Then
                        Response.Write "valor-negativo" 
                    Else
                        Response.Write "valor-zero"
                    End If %>">
                        R$ <%= FormatNumber(vSaldoPremio, 2) %>
                    </span>
                </td>
                <td class="text-end bg-total" data-numeric-value="<%= vSaldoTotal %>">
                    <span class="<% 
                    If vSaldoTotal > 0 Then 
                        Response.Write "valor-positivo"
                    ElseIf vSaldoTotal < 0 Then
                        Response.Write "valor-negativo" 
                    Else
                        Response.Write "valor-zero"
                    End If %>">
                        <strong>R$ <%= FormatNumber(vSaldoTotal, 2) %></strong>
                    </span>
                </td>
            </tr>
            <% 
                    rs.MoveNext
                Loop
            Else
            %>
            <tr>
                <td colspan="10" class="text-center text-danger">
                    <i class="fas fa-exclamation-triangle me-2"></i>
                    Nenhum registro encontrado. 
                    <br><small class="text-muted">Verifique se existem pagamentos com status 'Realizado' na tabela pagamentos_comissoes</small>
                    <div class="mt-2">
                        <a href="popular_resumo.asp" class="btn btn-sm btn-warning me-2">
                            <i class="fas fa-sync me-1"></i>Popular Tabela
                        </a>
                        <button onclick="location.reload()" class="btn btn-sm btn-primary">
                            <i class="fas fa-redo me-1"></i>Recarregar
                        </button>
                    </div>
                </td>
            </tr>
            <% 
            End If
            
            If hasRecords Then
                rs.Close
            End If
            Set rs = Nothing
            conn.Close
            Set conn = Nothing
            %>
        </tbody>
    </table>
</div>

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>

<script>
    $(document).ready(function() {
        $('#myTable').DataTable({
            "order": [[0, "asc"]],
            "pageLength": 100,
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            "columnDefs": [
                { "type": "num-fmt", "targets": [2,3,4,5,6,7,8,9] }
            ]
        });
    });
</script>

</body>
</html>