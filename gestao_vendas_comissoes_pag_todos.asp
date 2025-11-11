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

' Função para verificar se valor foi pago
Function IsValuePaid(valorPago, valorDevido)
    If valorDevido <= 0 Then
        IsValuePaid = True
    Else
        ' Usar comparação com tolerância para valores monetários
        IsValuePaid = (Abs(valorPago - valorDevido) < 0.01)
    End If
End Function

' ====================================================================
' Consulta principal para todas as vendas - ATUALIZADA COM VALORES LÍQUIDOS
' ====================================================================
Dim sqlVendas, rsVendas
sqlVendas = "SELECT " & _
           "v.ID, v.NomeEmpreendimento, v.Unidade, v.ValorUnidade, v.DataVenda, v.ValorComissaoGeral, " & _
           "v.Diretoria, v.Gerencia, v.Corretor, " & _
           "v.ValorDiretoria, v.PremioDiretoria, v.ValorLiqDiretoria, v.DescontoDiretoria, " & _
           "v.NomeDiretor, v.NomeGerente, v.Corretor " & _
           "v.ValorGerencia, v.PremioGerencia, v.ValorLiqGerencia, v.DescontoGerencia, " & _
           "v.ValorCorretor, v.PremioCorretor, v.ValorLiqCorretor, v.DescontoCorretor, " & _
           "v.DescontoPerc, v.DescontoBruto, v.DescontoDescricao, " & _
           "c.ID_Comissoes, c.StatusPagamento " & _
           "FROM Vendas AS v " & _
           "LEFT JOIN COMISSOES_A_PAGAR AS c ON v.ID = c.ID_Venda " & _
           "WHERE v.excluido = 0 " & _
           "ORDER BY v.DataVenda DESC, v.ID DESC"

sqlVendas = "SELECT " & _
           "v.* " & _
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
        .comissao-info {
            font-size: 0.8rem;
        }
        .valor-liquido {
            font-weight: bold;
            color: #28a745;
        }
        .valor-desconto {
            color: #dc3545;
            font-size: 0.75rem;
        }
        .valor-bruto {
            color: #6c757d;
            font-size: 0.75rem;
        }
        .badge-pago {
            background-color: #28a745;
            color: white;
            font-size: 0.7rem;
        }
        .badge-pendente {
            background-color: #dc3545;
            color: white;
            font-size: 0.7rem;
        }
        .badge-parcial {
            background-color: #ffc107;
            color: #212529;
            font-size: 0.7rem;
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
                        <th class="text-center">Desconto</th>                        
                        <th class="text-center">Diretoria</th>
                        <th class="text-center">Gerência</th>
                        <th class="text-center">Corretor</th>

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
                            If Not IsNull(rsVendas("ValorDiretoria")) Then
                                'totalComissao = totalComissao + CDbl(rsVendas("ValorComissaoGeral"))

                                totalComissao = totalComissao + CDbl(rsVendas("ValorLiqDiretoria")) + CDbl(rsVendas("ValorLiqGerencia")) + CDbl(rsVendas("ValorLiqCorretor"))
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
                            
                            ' ====================================================================
                            ' 🟢 TRATAMENTO DE VALORES LÍQUIDOS E DESCONTOS
                            ' ====================================================================
                            Dim dblValorLiqDiretoria, dblValorLiqGerencia, dblValorLiqCorretor
                            Dim dblDescontoDiretoria, dblDescontoGerencia, dblDescontoCorretor
                            Dim dblPremioDiretoria, dblPremioGerencia, dblPremioCorretor
                            
                            ' Valores Líquidos
                            If IsNull(rsVendas("ValorLiqDiretoria")) Then
                                dblValorLiqDiretoria = 0
                            Else
                                dblValorLiqDiretoria = CDbl(rsVendas("ValorLiqDiretoria"))
                            End If
                            
                            If IsNull(rsVendas("ValorLiqGerencia")) Then
                                dblValorLiqGerencia = 0
                            Else
                                dblValorLiqGerencia = CDbl(rsVendas("ValorLiqGerencia"))
                            End If
                            
                            If IsNull(rsVendas("ValorLiqCorretor")) Then
                                dblValorLiqCorretor = 0
                            Else
                                dblValorLiqCorretor = CDbl(rsVendas("ValorLiqCorretor"))
                            End If
                            
                            ' Descontos
                            If IsNull(rsVendas("DescontoDiretoria")) Then
                                dblDescontoDiretoria = 0
                            Else
                                dblDescontoDiretoria = CDbl(rsVendas("DescontoDiretoria"))
                            End If
                            
                            If IsNull(rsVendas("DescontoGerencia")) Then
                                dblDescontoGerencia = 0
                            Else
                                dblDescontoGerencia = CDbl(rsVendas("DescontoGerencia"))
                            End If
                            
                            If IsNull(rsVendas("DescontoCorretor")) Then
                                dblDescontoCorretor = 0
                            Else
                                dblDescontoCorretor = CDbl(rsVendas("DescontoCorretor"))
                            End If
                            
                            ' Prêmios
                            If IsNull(rsVendas("PremioDiretoria")) Then
                                dblPremioDiretoria = 0
                            Else
                                dblPremioDiretoria = CDbl(rsVendas("PremioDiretoria"))
                            End If
                            
                            If IsNull(rsVendas("PremioGerencia")) Then
                                dblPremioGerencia = 0
                            Else
                                dblPremioGerencia = CDbl(rsVendas("PremioGerencia"))
                            End If
                            
                            If IsNull(rsVendas("PremioCorretor")) Then
                                dblPremioCorretor = 0
                            Else
                                dblPremioCorretor = CDbl(rsVendas("PremioCorretor"))
                            End If
                            
                            ' ====================================================================
                            ' 🟢 VERIFICAÇÃO DE PAGAMENTOS COM VALORES LÍQUIDOS
                            ' ====================================================================
                            Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
                            Dim totalPremioPagoDiretoria, totalPremioPagoGerencia, totalPremioPagoCorretor
                            Dim comissaoDiretoriaPaga, comissaoGerenciaPaga, comissaoCorretorPaga
                            Dim premioDiretoriaPago, premioGerenciaPago, premioCorretorPago
                            
                            ' Buscar pagamentos realizados
                            totalPagoDiretoria = GetTotalPagoVenda(rsVendas("ID"), "diretoria", "Comissão")
                            totalPagoGerencia = GetTotalPagoVenda(rsVendas("ID"), "gerencia", "Comissão")
                            totalPagoCorretor = GetTotalPagoVenda(rsVendas("ID"), "corretor", "Comissão")
                            totalPremioPagoDiretoria = GetTotalPagoVenda(rsVendas("ID"), "diretoria", "Premiação")
                            totalPremioPagoGerencia = GetTotalPagoVenda(rsVendas("ID"), "gerencia", "Premiação")
                            totalPremioPagoCorretor = GetTotalPagoVenda(rsVendas("ID"), "corretor", "Premiação")
                            
                            ' Verificar se valores foram pagos (usando valores líquidos para comissões)
                            comissaoDiretoriaPaga = IsValuePaid(totalPagoDiretoria, dblValorLiqDiretoria)
                            comissaoGerenciaPaga = IsValuePaid(totalPagoGerencia, dblValorLiqGerencia)
                            comissaoCorretorPaga = IsValuePaid(totalPagoCorretor, dblValorLiqCorretor)
                            premioDiretoriaPago = IsValuePaid(totalPremioPagoDiretoria, dblPremioDiretoria)
                            premioGerenciaPago = IsValuePaid(totalPremioPagoGerencia, dblPremioGerencia)
                            premioCorretorPago = IsValuePaid(totalPremioPagoCorretor, dblPremioCorretor)
                            
                            ' Determinar status geral
                            Dim status, statusClass, rowClass
                            Dim todasComissoesPagas, todosPremiosPagos, statusCompleto
                            
                            todasComissoesPagas = True
                            If dblValorLiqDiretoria > 0 And Not comissaoDiretoriaPaga Then todasComissoesPagas = False
                            If dblValorLiqGerencia > 0 And Not comissaoGerenciaPaga Then todasComissoesPagas = False
                            If dblValorLiqCorretor > 0 And Not comissaoCorretorPaga Then todasComissoesPagas = False
                            
                            todosPremiosPagos = True
                            If dblPremioDiretoria > 0 And Not premioDiretoriaPago Then todosPremiosPagos = False
                            If dblPremioGerencia > 0 And Not premioGerenciaPago Then todosPremiosPagos = False
                            If dblPremioCorretor > 0 And Not premioCorretorPago Then todosPremiosPagos = False
                            
                            If todasComissoesPagas And todosPremiosPagos Then
                                statusCompleto = "PAGA"
                                rowClass = "row-paga"
                            ElseIf (totalPagoDiretoria + totalPagoGerencia + totalPagoCorretor + totalPremioPagoDiretoria + totalPremioPagoGerencia + totalPremioPagoCorretor) > 0 Then
                                statusCompleto = "PAGA PARCIALMENTE"
                                rowClass = "row-parcial"
                            Else
                                statusCompleto = "PENDENTE"
                                rowClass = "row-pendente"
                            End If
                            
                            status = statusCompleto
                            Select Case UCase(status)
                                Case "PAGA": statusClass = "status-pago"
                                Case "PAGA PARCIALMENTE": statusClass = "status-parcial"
                                Case "PENDENTE": statusClass = "status-pendente"
                                Case Else: statusClass = "bg-secondary"
                            End Select
                            
                            ' Verificar se há saldos pendentes
                            Dim temSaldoPendente
                            temSaldoPendente = (dblValorLiqDiretoria - totalPagoDiretoria > 0 Or _
                                              dblValorLiqGerencia - totalPagoGerencia > 0 Or _
                                              dblValorLiqCorretor - totalPagoCorretor > 0 Or _
                                              dblPremioDiretoria - totalPremioPagoDiretoria > 0 Or _
                                              dblPremioGerencia - totalPremioPagoGerencia > 0 Or _
                                              dblPremioCorretor - totalPremioPagoCorretor > 0)
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
                            <strong><%= rsVendas("Empreend_ID")%>-<%= rsVendas("NomeEmpreendimento") %></strong><br>

                            <small class="text-muted"><%= "Unid.: " & rsVendas("Unidade") %></small><br>
                            <small class="text-muted"><%= "R$ " & FormatNumber(rsVendas("ValorUnidade"),2) %></small><br>

                            <%vBrutoComissao = rsVendas("ValorDiretoria")+rsVendas("ValorGerencia")+rsVendas("ValorCorretor")%>
                            <small class="text-muted"><%= "Comissão: R$ " & FormatNumber(vBrutoComissao,2) %></small>

                            <%vComissaoLiq = rsVendas("ValorLiqDiretoria")+rsVendas("ValorLiqGerencia")+rsVendas("ValorLiqCorretor")%><br>
                            <small class="text-muted"><%= "Comissão Liq.: R$ " & FormatNumber(vComissaoLiq,2) %></small>

                        </td>

                        <!-- COLUNA DESCONTO TRIBUTÁRIO -->
                        <td class="text-center">
                            <% If Not IsNull(rsVendas("DescontoPerc")) And CDbl(rsVendas("DescontoPerc")) > 0 Then %>
                                <div class="small">
                                    <div class="fw-bold text-danger"><%= FormatNumber(rsVendas("DescontoPerc"), 2) %>%</div>
                                    <div class="text-muted">R$ <%= FormatNumber(rsVendas("DescontoBruto"), 2) %></div>
                                    <% If Not IsNull(rsVendas("DescontoDescricao")) And rsVendas("DescontoDescricao") <> "" Then %>
                                        <div class="text-info mt-1" title="<%= rsVendas("DescontoDescricao") %>">
                                            <i class="fas fa-info-circle"></i>
                                        </div>
                                    <% End If %>
                                </div>
                            <% Else %>
                                <span class="text-muted">-</span>
                            <% End If %>
                        </td>
                        
                        <!-- COLUNA DIRETORIA COM VALORES LÍQUIDOS -->
                        <td class="text-center">
                            <div class="comissao-info">
                                <strong><%= rsVendas("Diretoria") %><BR><%= rsVendas("NomeDiretor") %></strong><br>
                                
                                <!-- COMISSÃO DIRETORIA -->
                                <div class="valor-bruto">
                                    Bruto: R$ <%= FormatNumber(rsVendas("ValorDiretoria"), 2) %>
                                </div>
                                
                                <% If dblDescontoDiretoria > 0 Then %>
                                <div class="valor-desconto">
                                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoDiretoria, 2) %>
                                </div>
                                <% End If %>
                                
                                <div class="valor-liquido">
                                    <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqDiretoria, 2) %>
                                </div>
                                
                                <div class="mt-1">
                                    <% If comissaoDiretoriaPaga Then %>
                                        <span class="badge badge-pago">PAGA</span>
                                    <% ElseIf totalPagoDiretoria > 0 Then %>
                                        <span class="badge badge-parcial">PARCIAL</span>
                                        <small class="text-success">
                                            <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPagoDiretoria, 2) %>
                                        </small>
                                    <% Else %>
                                        <span class="badge badge-pendente">PENDENTE</span>
                                    <% End If %>
                                </div>
                                
                                <!-- PRÊMIO DIRETORIA -->
                                <% If dblPremioDiretoria > 0 Then %>
                                <div class="mt-2 pt-2 border-top">
                                    <div class="text-info fw-bold">
                                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioDiretoria, 2) %>
                                    </div>
                                    
                                    <div class="mt-1">
                                        <% If premioDiretoriaPago Then %>
                                            <span class="badge badge-pago">PAGA</span>
                                        <% ElseIf totalPremioPagoDiretoria > 0 Then %>
                                            <span class="badge badge-parcial">PARCIAL</span>
                                            <small class="text-success">
                                                <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPremioPagoDiretoria, 2) %>
                                            </small>
                                        <% Else %>
                                            <span class="badge badge-pendente">PENDENTE</span>
                                        <% End If %>
                                    </div>
                                </div>
                                <% End If %>
                            </div>
                        </td>
                        
                        <!-- COLUNA GERÊNCIA COM VALORES LÍQUIDOS -->
                        <td class="text-center">
                            <div class="comissao-info">
                                <strong><%= rsVendas("Gerencia") %><BR><%= rsVendas("NomeGerente") %></strong><br>
                                
                                <!-- COMISSÃO GERÊNCIA -->
                                <div class="valor-bruto">
                                    Bruto: R$ <%= FormatNumber(rsVendas("ValorGerencia"), 2) %>
                                </div>
                                
                                <% If dblDescontoGerencia > 0 Then %>
                                <div class="valor-desconto">
                                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoGerencia, 2) %>
                                </div>
                                <% End If %>
                                
                                <div class="valor-liquido">
                                    <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqGerencia, 2) %>
                                </div>
                                
                                <div class="mt-1">
                                    <% If comissaoGerenciaPaga Then %>
                                        <span class="badge badge-pago">PAGA</span>
                                    <% ElseIf totalPagoGerencia > 0 Then %>
                                        <span class="badge badge-parcial">PARCIAL</span>
                                        <small class="text-success">
                                            <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPagoGerencia, 2) %>
                                        </small>
                                    <% Else %>
                                        <span class="badge badge-pendente">PENDENTE</span>
                                    <% End If %>
                                </div>
                                
                                <!-- PRÊMIO GERÊNCIA -->
                                <% If dblPremioGerencia > 0 Then %>
                                <div class="mt-2 pt-2 border-top">
                                    <div class="text-info fw-bold">
                                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioGerencia, 2) %>
                                    </div>
                                    
                                    <div class="mt-1">
                                        <% If premioGerenciaPago Then %>
                                            <span class="badge badge-pago">PAGA</span>
                                        <% ElseIf totalPremioPagoGerencia > 0 Then %>
                                            <span class="badge badge-parcial">PARCIAL</span>
                                            <small class="text-success">
                                                <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPremioPagoGerencia, 2) %>
                                            </small>
                                        <% Else %>
                                            <span class="badge badge-pendente">PENDENTE</span>
                                        <% End If %>
                                    </div>
                                </div>
                                <% End If %>
                            </div>
                        </td>
                        
                        <!-- COLUNA CORRETOR COM VALORES LÍQUIDOS -->
                        <td class="text-center">
                            <div class="comissao-info">
                                <strong><%= rsVendas("Corretor") %></strong><br>
                                
                                <!-- COMISSÃO CORRETOR -->
                                <div class="valor-bruto">
                                    Bruto: R$ <%= FormatNumber(rsVendas("ValorCorretor"), 2) %>
                                </div>
                                
                                <% If dblDescontoCorretor > 0 Then %>
                                <div class="valor-desconto">
                                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoCorretor, 2) %>
                                </div>
                                <% End If %>
                                
                                <div class="valor-liquido">
                                    <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqCorretor, 2) %>
                                </div>
                                
                                <div class="mt-1">
                                    <% If comissaoCorretorPaga Then %>
                                        <span class="badge badge-pago">PAGA</span>
                                    <% ElseIf totalPagoCorretor > 0 Then %>
                                        <span class="badge badge-parcial">PARCIAL</span>
                                        <small class="text-success">
                                            <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPagoCorretor, 2) %>
                                        </small>
                                    <% Else %>
                                        <span class="badge badge-pendente">PENDENTE</span>
                                    <% End If %>
                                </div>
                                
                                <!-- PRÊMIO CORRETOR -->
                                <% If dblPremioCorretor > 0 Then %>
                                <div class="mt-2 pt-2 border-top">
                                    <div class="text-info fw-bold">
                                        <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioCorretor, 2) %>
                                    </div>
                                    
                                    <div class="mt-1">
                                        <% If premioCorretorPago Then %>
                                            <span class="badge badge-pago">PAGA</span>
                                        <% ElseIf totalPremioPagoCorretor > 0 Then %>
                                            <span class="badge badge-parcial">PARCIAL</span>
                                            <small class="text-success">
                                                <i class="fas fa-check-circle"></i> R$ <%= FormatNumber(totalPremioPagoCorretor, 2) %>
                                            </small>
                                        <% Else %>
                                            <span class="badge badge-pendente">PENDENTE</span>
                                        <% End If %>
                                    </div>
                                </div>
                                <% End If %>
                            </div>
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