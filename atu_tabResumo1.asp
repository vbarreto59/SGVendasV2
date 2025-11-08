<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' Consulta para o relatório
Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

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
    </style>
</head>
<body>

<div class="container mt-4">
    <h1 class="mb-4 text-center"><i class="fas fa-file-invoice-dollar me-2"></i>Relatório Consolidado</h1>

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
                    
                    vTotalVenda = CDbl(rs("TotalVenda"))
                    vComissaoDevida = CDbl(rs("ComissaoDevida"))
                    vPremioDevido = CDbl(rs("PremioDevido"))
                    vComissaoPaga = CDbl(rs("ComissaoPaga"))
                    vPremioPaga = CDbl(rs("PremioPago"))
                    vSaldoComissao = CDbl(rs("SaldoComissao"))
                    vSaldoPremio = CDbl(rs("SaldoPremio"))
                    vSaldoTotal = CDbl(rs("SaldoTotal"))
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
                    <a href="popular_resumo.asp" class="btn btn-sm btn-warning ms-2">
                        <i class="fas fa-sync me-1"></i>Popular Tabela
                    </a>
                </td>
            </tr>
            <% 
            End If
            
            rs.Close
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
            }
        });
    });
</script>

</body>
</html>