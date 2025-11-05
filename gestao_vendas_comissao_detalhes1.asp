<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>



<% ' funcional tentando incluir premiacao'
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Obter parâmetros
Dim anoDetalhe, mesDetalhe
anoDetalhe = Request.QueryString("ano")
mesDetalhe = Request.QueryString("mes")
diretoriaDetalhe =Request.QueryString("diretoria")
if diretoriaDetalhe = "" then
    vWhere = " WHERE 1=1 AND "
else
    vWhere = " WHERE V.diretoria = '" & diretoriaDetalhe & "' AND "   
end if    

If anoDetalhe = "" Or mesDetalhe = "" Then
    Response.Redirect "gestao_vendas_comissao_saldo2.asp"
End If

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Buscar nome do mês
Dim nomeMesDetalhe
Select Case CInt(mesDetalhe)
    Case 1: nomeMesDetalhe = "Janeiro"
    Case 2: nomeMesDetalhe = "Fevereiro"
    Case 3: nomeMesDetalhe = "Março"
    Case 4: nomeMesDetalhe = "Abril"
    Case 5: nomeMesDetalhe = "Maio"
    Case 6: nomeMesDetalhe = "Junho"
    Case 7: nomeMesDetalhe = "Julho"
    Case 8: nomeMesDetalhe = "Agosto"
    Case 9: nomeMesDetalhe = "Setembro"
    Case 10: nomeMesDetalhe = "Outubro"
    Case 11: nomeMesDetalhe = "Novembro"
    Case 12: nomeMesDetalhe = "Dezembro"
End Select

' Buscar vendas do mês
Set rsVendasMes = Server.CreateObject("ADODB.Recordset")

sqlVendasMes = "SELECT " & _
               "V.ID, " & _
               "V.Empreend_ID, " & _
               "V.NomeEmpreendimento, " & _
               "V.Unidade, " & _
               "V.ValorUnidade, " & _
               "V.DataVenda, " & _
               "V.Corretor, " & _
               "V.Diretoria, " & _
               "V.Gerencia, " & _
               "V.ComissaoPercentual, " & _
               "V.ValorComissaoGeral, " & _
               "V.ValorDiretoria, " & _
               "V.ValorGerencia, " & _
               "V.ValorCorretor, " & _
               "V.NomeDiretor, " & _
               "V.NomeGerente " & _
               "FROM Vendas V " & _
                vWhere & " V.AnoVenda = " & anoDetalhe & " " & _
               "AND V.MesVenda = " & mesDetalhe & " " & _
               "AND (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
               "ORDER BY V.DataVenda DESC, V.ID DESC"
               'Response.Write sqlVendasMes
               'Response.end 

rsVendasMes.Open sqlVendasMes, connSales

' Buscar pagamentos de comissões do mês
Set rsPagamentosMes = Server.CreateObject("ADODB.Recordset")
sqlPagamentosMes = "SELECT " & _
                   "PC.ID_Venda, " & _
                   "PC.TipoRecebedor, " & _
                   "PC.TipoPagamento, " & _
                   "PC.ValorPago, " & _
                   "PC.DataPagamento, " & _
                   "PC.Status, " & _
                   "V.NomeDiretor, " & _
                   "V.NomeGerente, " & _
                   "V.Corretor " & _
                   "FROM PAGAMENTOS_COMISSOES PC " & _
                   "INNER JOIN Vendas V ON PC.ID_Venda = V.ID " & _
                   "WHERE V.AnoVenda = " & anoDetalhe & " " & _
                   "AND V.MesVenda = " & mesDetalhe & " " & _
                   "AND (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                   "ORDER BY PC.DataPagamento DESC"

rsPagamentosMes.Open sqlPagamentosMes, connSales

' Calcular totais
Dim totalVGV, totalComissao, totalPago, totalComissaoPaga, totalPremiacaoPaga
totalVGV = 0
totalComissao = 0
totalPago = 0
totalComissaoPaga = 0
totalPremiacaoPaga = 0
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detalhes de Comissões | <%= nomeMesDetalhe %>/<%= anoDetalhe %></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding-top: 60px;
        }
        
        .header-bordo {
            background-color: #800000 !important;
            color: #ffffff !important;
            padding: 5px 20px; 
            margin-bottom: 8px !important; 
            font-size: 20px;
            font-weight: bold;
            text-align: left;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
            width: 100%;
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
        
        .valor-positivo {
            color: #28a745;
            font-weight: 600;
        }
        
        .badge-pago {
            background-color: #28a745;
            color: white;
        }
        
        .badge-pendente {
            background-color: #fd7e14;
            color: white;
        }
        
        .badge-comissao {
            background-color: #3498db;
            color: white;
        }
        
        .badge-premiacao {
            background-color: #9b59b6;
            color: white;
        }
        
        .info-card {
            background: white;
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 1rem;
            border-left: 4px solid #3498db;
        }
        
        .info-card-comissao {
            border-left: 4px solid #3498db;
        }
        
        .info-card-premiacao {
            border-left: 4px solid #9b59b6;
        }
        
        .nome-recebedor {
            font-weight: 600;
            color: #2c3e50;
        }
    </style>
</head>
<body>
    <header class="header-bordo">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 style="color: #ffffff !important; margin: 0; font-size: 20px;">
                        <i class="fas fa-search me-2"></i> Detalhes - <%= nomeMesDetalhe %>/<%= anoDetalhe %>
                    </h1>
                </div>
                <div class="col-md-6 text-end">
                    <a href="gestao_comissoes_resumo.asp?ano=<%= anoDetalhe %>" class="btn btn-light btn-sm" style="color: #333 !important;">
                        <i class="fas fa-arrow-left me-1"></i>Voltar ao Resumo
                    </a>
                </div>
            </div>
        </div>
    </header>

    <div class="container-fluid main-content">
        <!-- Cards de Resumo -->
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="info-card">
                    <h6><i class="fas fa-chart-line me-2"></i>VGV Total</h6>
                    <h4 class="valor-positivo">
                        <%
                        If Not rsVendasMes.EOF Then
                            rsVendasMes.MoveFirst
                            vCont = 0
                            Do While Not rsVendasMes.EOF
                                vCont = vCont + 1
                                totalVGV = totalVGV + CDbl(rsVendasMes("ValorUnidade"))
                                totalComissao = totalComissao + CDbl(rsVendasMes("ValorComissaoGeral"))
                                rsVendasMes.MoveNext
                            Loop
                            rsVendasMes.MoveFirst
                        End If
                        %>
                        R$ <%= FormatNumber(totalVGV, 2) %>
                    </h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card">
                    <h6><i class="fas fa-money-bill-wave me-2"></i>Comissão Total</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalComissao, 2) %></h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card info-card-comissao">
                    <h6><i class="fas fa-hand-holding-usd me-2"></i>Comissões Pagas</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalComissaoPaga, 2) %></h4>
                </div>
            </div>
            <div class="col-md-3">
                <div class="info-card info-card-premiacao">
                    <h6><i class="fas fa-trophy me-2"></i>Premiações Pagas</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalPremiacaoPaga, 2) %></h4>
                </div>
            </div>
        </div>

        <!-- Tabela de Vendas -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-shopping-cart me-2"></i>Vendas do Mês</h5>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table id="tabelaVendas" class="table table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Data</th>
                                <th>Empreendimento</th>
                                <th>Unidade</th>
                                <th>Corretor</th>
                                <th>Diretoria</th>
                                <th>Gerência</th>
                                <th>Valor (R$)</th>
                                <th>Comissão (R$)</th>
                                <th>%</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsVendasMes.EOF Then
                                Do While Not rsVendasMes.EOF
                            %>
                            <tr>
                                <td><strong><%= rsVendasMes("ID") %></strong></td>
                                <td><%= FormatDateTime(rsVendasMes("DataVenda"), 2) %></td>
                                <td>
                                    <strong><%= rsVendasMes("Empreend_ID") %></strong>
                                    <br><small class="text-muted"><%= rsVendasMes("NomeEmpreendimento") %></small>
                                </td>
                                <td><%= rsVendasMes("Unidade") %></td>
                                <td><%= rsVendasMes("Corretor") %></td>
                                <td><%= rsVendasMes("Diretoria") %></td>
                                <td><%= rsVendasMes("Gerencia") %></td>
                                <td class="valor-positivo">R$ <%= FormatNumber(rsVendasMes("ValorUnidade"), 2) %></td>
                                <td class="valor-positivo">R$ <%= FormatNumber(rsVendasMes("ValorComissaoGeral"), 2) %></td>
                                <td><span class="badge bg-info"><%= rsVendasMes("ComissaoPercentual") %>%</span></td>
                            </tr>
                            <%
                                    rsVendasMes.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="10" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhuma venda encontrada para <%= nomeMesDetalhe %>/<%= anoDetalhe %>.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Tabela de Pagamentos -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-receipt me-2"></i>Pagamentos de Comissões e Premiações</h5>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table id="tabelaPagamentos" class="table table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID Venda</th>
                                <th>Tipo Recebedor</th>
                                <th>Nome do Recebedor</th>
                                <th>Tipo Pagamento</th>
                                <th>Valor Pago (R$)</th>
                                <th>Data Pagamento</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsPagamentosMes.EOF Then
                                Do While Not rsPagamentosMes.EOF
                                    totalPago = totalPago + CDbl(rsPagamentosMes("ValorPago"))
                                    
                                    ' Acumular totais por tipo de pagamento
                                    If UCase(rsPagamentosMes("TipoPagamento")) = "COMISSÃO" Or UCase(rsPagamentosMes("TipoPagamento")) = "COMISSAO" Then
                                        totalComissaoPaga = totalComissaoPaga + CDbl(rsPagamentosMes("ValorPago"))
                                    ElseIf UCase(rsPagamentosMes("TipoPagamento")) = "PREMIACAO" Or UCase(rsPagamentosMes("TipoPagamento")) = "PREMIAÇÃO" Then
                                        totalPremiacaoPaga = totalPremiacaoPaga + CDbl(rsPagamentosMes("ValorPago"))
                                    End If
                                    
                                    Dim badgeClass, statusClass, tipoPagamentoClass, tipoPagamentoTexto, iconClass, nomeRecebedor
                                    
                                    ' Definir classe do badge para TipoRecebedor
                                    Select Case UCase(rsPagamentosMes("TipoRecebedor"))
                                        Case "DIRETORIA"
                                            badgeClass = "bg-primary"
                                            nomeRecebedor = rsPagamentosMes("NomeDiretor")
                                        Case "GERENCIA"
                                            badgeClass = "bg-warning"
                                            nomeRecebedor = rsPagamentosMes("NomeGerente")
                                        Case "CORRETOR"
                                            badgeClass = "bg-success"
                                            nomeRecebedor = rsPagamentosMes("Corretor")
                                        Case Else
                                            badgeClass = "bg-secondary"
                                            nomeRecebedor = "N/A"
                                    End Select
                                    
                                    ' Definir classe e texto para TipoPagamento
                                    If UCase(rsPagamentosMes("TipoPagamento")) = "COMISSÃO" Or UCase(rsPagamentosMes("TipoPagamento")) = "COMISSAO" Then
                                        tipoPagamentoClass = "badge-comissao"
                                        tipoPagamentoTexto = "Comissão"
                                        iconClass = "fa-money-bill-wave"
                                    ElseIf UCase(rsPagamentosMes("TipoPagamento")) = "PREMIACAO" Or UCase(rsPagamentosMes("TipoPagamento")) = "PREMIAÇÃO" Then
                                        tipoPagamentoClass = "badge-premiacao"
                                        tipoPagamentoTexto = "Premiação"
                                        iconClass = "fa-trophy"
                                    Else
                                        tipoPagamentoClass = "bg-secondary"
                                        tipoPagamentoTexto = rsPagamentosMes("TipoPagamento")
                                        iconClass = "fa-question"
                                    End If
                                    
                                    ' Definir classe do badge para Status
                                    If UCase(rsPagamentosMes("Status")) = "PAGO" Then
                                        statusClass = "badge-pago"
                                    Else
                                        statusClass = "badge-pendente"
                                    End If
                            %>
                            <tr>
                                <td><strong><%= rsPagamentosMes("ID_Venda") %></strong></td>
                                <td>
                                    <span class="badge <%= badgeClass %>">
                                        <%= UCase(rsPagamentosMes("TipoRecebedor")) %>
                                    </span>
                                </td>
                                <td>
                                    <span class="nome-recebedor">
                                        <%= nomeRecebedor %>
                                    </span>
                                </td>
                                <td>
                                    <span class="badge <%= tipoPagamentoClass %>">
                                        <i class="fas <%= iconClass %> me-1"></i>
                                        <%= tipoPagamentoTexto %>
                                    </span>
                                </td>
                                <td class="valor-positivo">R$ <%= FormatNumber(rsPagamentosMes("ValorPago"), 2) %></td>
                                <td><%= FormatDateTime(rsPagamentosMes("DataPagamento"), 2) %></td>
                                <td>
                                    <span class="badge <%= statusClass %>">
                                        <%= rsPagamentosMes("Status") %>
                                    </span>
                                </td>
                            </tr>
                            <%
                                    rsPagamentosMes.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="7" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhum pagamento encontrado para <%= nomeMesDetalhe %>/<%= anoDetalhe %>.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-light">
                                <th colspan="4" class="text-end">Total Pago:</th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalPago, 2) %></th>
                                <th colspan="2"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
    $(document).ready(function () {
        $('#tabelaVendas').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            pageLength: 25,
            order: [[0, 'desc']],
            responsive: true
        });

        $('#tabelaPagamentos').DataTable({
            language: {
                url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            pageLength: 25,
            order: [[5, 'desc']], // Ordena por DataPagamento
            responsive: true
        });
    });
    </script>
</body>
</html>

<%
' Fechar conexões
If Not rsVendasMes Is Nothing Then
    rsVendasMes.Close
    Set rsVendasMes = Nothing
End If

If Not rsPagamentosMes Is Nothing Then
    rsPagamentosMes.Close
    Set rsPagamentosMes = Nothing
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>