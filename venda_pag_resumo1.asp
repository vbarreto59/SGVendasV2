<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ... (Funções IIF e SafeCDbl mantidas para evitar erro VBScript) ...

Function IIF(condition, trueResult, falseResult)
    If condition Then
        IIF = trueResult
    Else
        IIF = falseResult
    End If
End Function

Function SafeCDbl(value)
    If IsNull(value) Or IsEmpty(value) Then
        SafeCDbl = 0
    ElseIf IsNumeric(value) Then
        SafeCDbl = CDbl(value)
    Else
        SafeCDbl = 0 
    End If
End Function

' 1. Conexão com o Banco de Dados (mantida)
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open StrConnSales
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>ERRO DE CONEXÃO: " & Err.Description & "</div>"
    Response.End
End If
On Error GoTo 0

' 2. Consulta SQL CORRIGIDA PARA MDB/ACCESS
sql = "SELECT " & _
      "VT.ID_Venda, " & _
      "VT.UserID, " & _
      "VT.Nome, " & _
      "MAX(VT.Diretoria) AS Diretoria, " & _
      "MAX(VT.Gerencia) AS Gerencia, " & _
      "SUM(VT.VTotal) AS TotalComissaoDevida, " & _
      "SUM(VT.VBruto) AS VBrutoConsolidado, " & _
      " ( SELECT SUM(PC.ValorPago) FROM PAGAMENTOS_COMISSOES AS PC " & _
      "   WHERE PC.UsuariosUserId = VT.UserID AND PC.ID_Venda = VT.ID_Venda " & _
      " ) AS ValorPagoPorVenda " & _
      "FROM VENDA_TEMP AS VT " & _
      "GROUP BY VT.ID_Venda, VT.UserID, VT.Nome, VT.Diretoria, VT.Gerencia " & _
      "ORDER BY VT.ID_Venda DESC, VT.Nome"

Set rs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rs.Open sql, conn
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>ERRO NA CONSULTA SQL: " & Err.Description & "</div>"
    hasRecords = False
Else
    hasRecords = Not rs.EOF
End If
On Error GoTo 0

Dim totalComissoesGeral, totalPagoGeral, totalSaldoGeral
totalComissoesGeral = 0
totalPagoGeral = 0

If hasRecords Then
    rs.MoveFirst
    Do While Not rs.EOF
        Dim vComissao, vPago
        vComissao = SafeCDbl(rs("TotalComissaoDevida"))
        vPago = SafeCDbl(rs("ValorPagoPorVenda"))
        
        totalComissoesGeral = totalComissoesGeral + vComissao
        totalPagoGeral = totalPagoGeral + vPago
        
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
    <title>Relatório Consolidado de Comissões e Saldos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        .table th { background-color: #008080; color: white; }
        .bg-pago { background-color: #d4edda; }
        .bg-pendente { background-color: #f8d7da; }
        .bg-parcial { background-color: #fff3cd; }
        .valor-numero { font-family: 'Courier New', monospace; text-align: right; }
        .table-responsive { overflow-x: auto; }
    </style>
</head>
<body>

<div class="container-fluid mt-4">
    <h1 class="mb-4 text-center"><i class="fas fa-search-dollar me-2"></i>Relatório Consolidado de Comissões e Saldos</h1>
    <p class="text-muted text-center">Uma linha por Pessoa/Venda, consolidando múltiplos cargos para cálculo correto do saldo.</p>
    
    <div class="row mb-4 justify-content-center">
        <div class="col-md-3">
            <div class="card text-white bg-success">
                <div class="card-body">
                    <h6 class="card-title">Total Comissões Devidas</h6>
                    <h4>R$ <%= FormatNumber(totalComissoesGeral, 2) %></h4>
                </div>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card text-white bg-primary">
                <div class="card-body">
                    <h6 class="card-title">Total Pago</h6>
                    <h4>R$ <%= FormatNumber(totalPagoGeral, 2) %></h4>
                </div>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card text-white <%= IIF(totalSaldoGeral > 0, "bg-danger", "bg-secondary") %>">
                <div class="card-body">
                    <h6 class="card-title">Saldo Pendente</h6>
                    <h4>R$ <%= FormatNumber(totalSaldoGeral, 2) %></h4>
                </div>
            </div>
        </div>
    </div>
    
    <% If hasRecords Then %>
    <div class="table-responsive">
        <table id="relatorioTable" class="table table-striped table-bordered table-sm" style="width:100%">
            <thead>
                <tr>
                    <th>ID Venda</th>
                    <th>Colaborador</th>
                    <th>Diretoria</th>
                    <th>Gerência</th>
                    <th class="text-end">V. Bruto Total</th>
                    <th class="text-end">Comissão Devida (Total)</th>
                    <th class="text-end">Valor Pago</th>
                    <th class="text-end">Saldo</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                Do While Not rs.EOF
                    'Dim vComissao, vPago, vSaldo
                    'Dim statusPagamento, statusClass
                    
                    vComissao = SafeCDbl(rs("TotalComissaoDevida"))
                    vPago = SafeCDbl(rs("ValorPagoPorVenda"))
                    vSaldo = vComissao - vPago
                    
                    If vComissao > 0 And vPago = 0 Then
                        statusPagamento = "Pendente"
                        statusClass = "bg-pendente"
                    ElseIf vSaldo <= 0 Then
                        statusPagamento = "Pago" & IIF(vSaldo < 0, " (Exc.)", "")
                        statusClass = "bg-pago"
                    Else
                        statusPagamento = "Parcial"
                        statusClass = "bg-parcial"
                    End If
                %>
                <tr class="<%= statusClass %>">
                    <td><strong>V<%= rs("ID_Venda") %></strong></td>
                    <td><%= rs("Nome") %> (ID: <%= rs("UserID") %>)</td>
                    <td><%= rs("Diretoria") %></td>
                    <td><%= rs("Gerencia") %></td>
                    <td class="valor-numero"><%= FormatNumber(SafeCDbl(rs("VBrutoConsolidado")), 2) %></td>
                    <td class="valor-numero"><strong><%= FormatNumber(vComissao, 2) %></strong></td>
                    <td class="valor-numero"><%= FormatNumber(vPago, 2) %></td>
                    <td class="valor-numero">
                        <span class="<%= IIF(vSaldo > 0, "text-danger fw-bold", IIF(vSaldo < 0, "text-success fw-bold", "text-muted")) %>">
                            <%= FormatNumber(vSaldo, 2) %>
                        </span>
                    </td>
                    <td><span class="badge <%= IIF(vSaldo > 0, "bg-danger", IIF(vSaldo < 0, "bg-success", "bg-secondary")) %>"><%= statusPagamento %></span></td>
                </tr>
                <%
                    rs.MoveNext
                Loop
                %>
            </tbody>
        </table>
    </div>
    <% Else %>
    <div class="alert alert-warning text-center">
        <i class="fas fa-info-circle me-2"></i>Nenhum registro de comissão encontrado para consolidar.
    </div>
    <% End If %>
</div>

<% 
If hasRecords Then
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

<script>
    $(document).ready(function() {
        $('#relatorioTable').DataTable({
            "order": [[0, "desc"], [1, "asc"]], 
            "pageLength": 50,
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            "columnDefs": [
                { "type": "num-fmt", "targets": [4, 5, 6, 7] } // As colunas mudaram, ajuste os índices
            ]
        });
    });
</script>
</body>
</html>