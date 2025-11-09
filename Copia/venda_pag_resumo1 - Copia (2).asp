<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<!--#include file="AtualizarVendasTemp.asp"-->

<%
' Consulta para o relatório completo
'Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

' SQL corrigido - separar consultas para evitar duplicação
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
      " WHERE PC.UsuariosUserId = VT.UserID AND PC.ID_Venda = VT.ID_Venda) AS SomaDeValorPago, " & _
      "VT.VTotal - (SELECT SUM(ValorPago) FROM PAGAMENTOS_COMISSOES PC " & _
      " WHERE PC.UsuariosUserId = VT.UserID AND PC.ID_Venda = VT.ID_Venda) AS SaldoPendente " & _
      "FROM VENDA_TEMP AS VT " & _
      "ORDER BY VT.ID_Venda, VT.Nome"

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
    </style>
</head>
<body>

<div class="container mt-4">
    <h1 class="mb-4 text-center"><i class="fas fa-file-invoice-dollar me-2"></i>Relatório Completo - Vendas e Pagamentos</h1>

    <div class="alert alert-info">
        <i class="fas fa-info-circle me-2"></i>
        <strong>Relatório completo:</strong> Mostra todos os vendedores, incluindo quem já recebeu e quem ainda não recebeu pagamentos.
    </div>

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
                            If Not IsNull(rs("SomaDeValorPago")) Then
                                totalPago = totalPago + CDbl(rs("SomaDeValorPago"))
                            End If
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
                    <h4> <%= FormatNumber(totalComissoes - totalPago, 2) %></h4>
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
            </tr>
        </thead>
        <tbody>
            <% 
            If hasRecords Then
                Do While Not rs.EOF
                    Dim vVBruto, vDesc, vVLiq, vPremio, vVTotal, vValorPago, vSaldo
                    Dim statusPagamento, statusClass, saldoClass, badgeClass
                    
                    ' Tratar valores nulos
                    vVBruto = 0
                    vDesc = 0
                    vVLiq = 0
                    vPremio = 0
                    vVTotal = 0
                    vValorPago = 0
                    vSaldo = 0
                    
                    If Not IsNull(rs("VBruto")) Then vVBruto = CDbl(rs("VBruto"))
                    If Not IsNull(rs("Desc")) Then vDesc = CDbl(rs("Desc"))
                    If Not IsNull(rs("VLiq")) Then vVLiq = CDbl(rs("VLiq"))
                    If Not IsNull(rs("Premio")) Then vPremio = CDbl(rs("Premio"))
                    If Not IsNull(rs("VTotal")) Then vVTotal = CDbl(rs("VTotal"))
                    If Not IsNull(rs("SomaDeValorPago")) Then vValorPago = CDbl(rs("SomaDeValorPago"))
                    If Not IsNull(rs("SaldoPendente")) Then vSaldo = CDbl(rs("SaldoPendente"))
                    
                    ' Lógica CORRIGIDA para status
                    If vValorPago = 0 Then
                        statusPagamento = "Pendente"
                        statusClass = "bg-pendente"
                        badgeClass = "bg-warning text-dark"
                    ElseIf vSaldo = 0 Then
                        statusPagamento = "Pago"
                        statusClass = "bg-pago"
                        badgeClass = "bg-success"
                    Else
                        statusPagamento = "Parcial"
                        statusClass = "bg-parcial"
                        badgeClass = "bg-info"
                    End If
                    
                    ' Definir classe para saldo
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
                <td class="text-end"><%= FormatNumber(rs("VUnid"),2) %></td>
                <td class="text-end"> <%= FormatNumber(vVBruto, 2) %></td>
                <td class="text-end"> <%= FormatNumber(vDesc, 2) %></td>
                <td class="text-end"> <%= FormatNumber(vVLiq, 2) %></td>
                <td class="text-end"> <%= FormatNumber(vPremio, 2) %></td>
                <td class="text-end"><strong><%= FormatNumber(vVTotal, 2) %></strong></td>
                <td class="text-end"> <%= FormatNumber(vValorPago, 2) %></td>
                <td class="text-end <%= saldoClass %>">
                    <span class="<% 
                    If vSaldo > 0 Then 
                        Response.Write "valor-negativo"
                    ElseIf vSaldo < 0 Then
                        Response.Write "valor-positivo" 
                    Else
                        Response.Write "valor-zero"
                    End If %>">
                        <strong>R$ <%= FormatNumber(vSaldo, 2) %></strong>
                    </span>
                </td>
                <td class="<%= statusClass %>">
                    <span class="badge <%= badgeClass %>">
                        <%= statusPagamento %>
                    </span>
                </td>
            </tr>
            <% 
                    rs.MoveNext
                Loop
            Else
            %>
            <tr>
                <td colspan="14" class="text-center text-danger">
                    <i class="fas fa-exclamation-triangle me-2"></i>
                    Nenhum registro encontrado.
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
            "order": [[0, "desc"]],
            "pageLength": 50,
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            "columnDefs": [
                { "type": "num-fmt", "targets": [5,6,7,8,9,10,11,12] }
            ]
        });
    });
</script>

</body>
</html>