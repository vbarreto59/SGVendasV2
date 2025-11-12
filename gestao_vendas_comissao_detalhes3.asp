detalhes com premio sem a segunda tabela
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% 'Nova estratégia - Array de premiações - 09/11/2025'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Obter parâmetros
Dim anoDetalhe, mesDetalhe
anoDetalhe = Request.QueryString("ano")
mesDetalhe = Request.QueryString("mes")
diretoriaDetalhe = Request.QueryString("diretoria")

If anoDetalhe = "" Or mesDetalhe = "" Then
    Response.Redirect "gestao_vendas_comissao_saldo2.asp"
End If

If diretoriaDetalhe = "" Then
    vWhere = " WHERE 1=1 AND "
Else
    vWhere = " WHERE V.diretoria = '" & diretoriaDetalhe & "' AND "   
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

' Calcular totais ANTES de exibir os cards
Dim totalVGV, totalComissaoBruta, totalDesconto, totalComissaoLiquida
Dim totalComissaoPaga, totalPremiacaoPaga, totalPremiacao, totalPremiacaoAPagar
Dim totalComissaoAPagar, totalPago, totalAPagar

totalVGV = 0
totalComissaoBruta = 0
totalDesconto = 0
totalComissaoLiquida = 0
totalComissaoPaga = 0
totalPremiacaoPaga = 0
totalPremiacao = 0
totalPremiacaoAPagar = 0
totalComissaoAPagar = 0
totalPago = 0
totalAPagar = 0

' =============================================
' NOVA ESTRATÉGIA: ARRAY DE PREMIAÇÕES
' =============================================

' Criar Dictionary para armazenar premiações por ID da venda
Set premiacoesDict = Server.CreateObject("Scripting.Dictionary")

' Buscar premiações de TODAS as fontes possíveis
On Error Resume Next

' 1. Buscar de tabela separada Premiacoes (se existir)
sqlPremiacoes = "SELECT ID_Venda, ValorPremiacao FROM Premiacoes WHERE ID_Venda IN (SELECT ID FROM Vendas WHERE AnoVenda = " & anoDetalhe & " AND MesVenda = " & mesDetalhe & ")"
Set rsPremiacoes = connSales.Execute(sqlPremiacoes)
If Err.Number = 0 Then
    Do While Not rsPremiacoes.EOF
        Dim idVendaPrem, valorPrem
        idVendaPrem = CStr(rsPremiacoes("ID_Venda"))
        valorPrem = CDbl(rsPremiacoes("ValorPremiacao"))
        
        If valorPrem > 0 Then
            premiacoesDict.Add idVendaPrem, valorPrem
        End If
        rsPremiacoes.MoveNext
    Loop
    rsPremiacoes.Close
End If
Err.Clear

' 2. Buscar da própria tabela Vendas em diferentes colunas
Dim colunasPremiacao
colunasPremiacao = Array("Premio", "Premiacao", "ValorPremiacao", "Premiação", "Premiacão", "Bonus", "ValorBonus")

For Each coluna In colunasPremiacao
    sqlTest = "SELECT ID, " & coluna & " FROM Vendas WHERE AnoVenda = " & anoDetalhe & " AND MesVenda = " & mesDetalhe & " AND " & coluna & " > 0"
    Set rsTest = connSales.Execute(sqlTest)
    If Err.Number = 0 Then
        Do While Not rsTest.EOF
            idVendaPrem = CStr(rsTest("ID"))
            valorPrem = CDbl(rsTest(coluna))
            
            If Not premiacoesDict.Exists(idVendaPrem) Then
                premiacoesDict.Add idVendaPrem, valorPrem
            End If
            rsTest.MoveNext
        Loop
        rsTest.Close
    End If
    Err.Clear
Next

' 3. Buscar de pagamentos de premiação
sqlPagamentosPremiacao = "SELECT PC.ID_Venda, SUM(PC.ValorPago) as TotalPremiacao " & _
                         "FROM PAGAMENTOS_COMISSOES PC " & _
                         "INNER JOIN Vendas V ON PC.ID_Venda = V.ID " & _
                         "WHERE V.AnoVenda = " & anoDetalhe & " AND V.MesVenda = " & mesDetalhe & " " & _
                         "AND (PC.TipoPagamento = 'PREMIACAO' OR PC.TipoPagamento = 'PREMIAÇÃO') " & _
                         "GROUP BY PC.ID_Venda"
Set rsPagPrem = connSales.Execute(sqlPagamentosPremiacao)
If Err.Number = 0 Then
    Do While Not rsPagPrem.EOF
        idVendaPrem = CStr(rsPagPrem("ID_Venda"))
        valorPrem = CDbl(rsPagPrem("TotalPremiacao"))
        
        If Not premiacoesDict.Exists(idVendaPrem) Then
            premiacoesDict.Add idVendaPrem, valorPrem
        End If
        rsPagPrem.MoveNext
    Loop
    rsPagPrem.Close
End If
Err.Clear

On Error GoTo 0

' =============================================
' FIM DA NOVA ESTRATÉGIA
' =============================================

' Buscar vendas do mês
Set rsVendasMes = Server.CreateObject("ADODB.Recordset")

' SQL SIMPLES apenas com colunas básicas
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
               "V.NomeGerente, " & _
               "V.DescontoBruto, " & _
               "V.ValorLiqGeral, " & _
               "(V.ValorDiretoria + V.ValorGerencia + V.ValorCorretor) as ComissaoBruta " & _
               "FROM Vendas V " & _
               vWhere & " V.AnoVenda = " & anoDetalhe & " " & _
               "AND V.MesVenda = " & mesDetalhe & " " & _
               "AND (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
               "ORDER BY V.DataVenda DESC, V.ID DESC"

rsVendasMes.Open sqlVendasMes, connSales

' Calcular totais das vendas
If Not rsVendasMes.EOF Then
    Do While Not rsVendasMes.EOF
        totalVGV = totalVGV + CDbl(rsVendasMes("ValorUnidade"))
        totalComissaoBruta = totalComissaoBruta + CDbl(rsVendasMes("ComissaoBruta"))
        totalDesconto = totalDesconto + CDbl(rsVendasMes("DescontoBruto"))
        totalComissaoLiquida = totalComissaoLiquida + CDbl(rsVendasMes("ValorLiqGeral"))
        
        ' Buscar premiação do array/dictionary
        Dim idVendaAtual, premiacaoVenda
        idVendaAtual = CStr(rsVendasMes("ID"))
        
        If premiacoesDict.Exists(idVendaAtual) Then
            premiacaoVenda = premiacoesDict(idVendaAtual)
        Else
            ' Se não encontrou premiação específica, verificar se há valores iguais de comissão
            Dim valDir, valGer, valCor
            valDir = CDbl(rsVendasMes("ValorDiretoria"))
            valGer = CDbl(rsVendasMes("ValorGerencia"))
            valCor = CDbl(rsVendasMes("ValorCorretor"))
            
            ' Se todos têm o mesmo valor (ex: R$ 500,00 cada), considerar como premiação
            If valDir > 0 And valGer > 0 And valCor > 0 Then
                If valDir = valGer And valGer = valCor Then
                    premiacaoVenda = valDir + valGer + valCor
                    ' Armazenar no dictionary para uso futuro
                    premiacoesDict.Add idVendaAtual, premiacaoVenda
                Else
                    premiacaoVenda = 0
                End If
            Else
                premiacaoVenda = 0
            End If
        End If
        
        totalPremiacao = totalPremiacao + premiacaoVenda
        
        rsVendasMes.MoveNext
    Loop
    rsVendasMes.MoveFirst
End If

' Buscar pagamentos e calcular totais pagos
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

' Calcular totais dos pagamentos
If Not rsPagamentosMes.EOF Then
    Do While Not rsPagamentosMes.EOF
        totalPago = totalPago + CDbl(rsPagamentosMes("ValorPago"))
        
        ' Acumular totais por tipo de pagamento
        If UCase(rsPagamentosMes("TipoPagamento")) = "COMISSÃO" Or UCase(rsPagamentosMes("TipoPagamento")) = "COMISSAO" Then
            totalComissaoPaga = totalComissaoPaga + CDbl(rsPagamentosMes("ValorPago"))
        ElseIf UCase(rsPagamentosMes("TipoPagamento")) = "PREMIACAO" Or UCase(rsPagamentosMes("TipoPagamento")) = "PREMIAÇÃO" Then
            totalPremiacaoPaga = totalPremiacaoPaga + CDbl(rsPagamentosMes("ValorPago"))
        End If
        rsPagamentosMes.MoveNext
    Loop
    rsPagamentosMes.MoveFirst
End If

' Se premiação total for 0 mas tem premiações pagas, usar como base
If totalPremiacao = 0 And totalPremiacaoPaga > 0 Then
    totalPremiacao = totalPremiacaoPaga
End If

' Calcular totais pendentes
totalComissaoAPagar = totalComissaoLiquida - totalComissaoPaga
If totalComissaoAPagar < 0 Then totalComissaoAPagar = 0

totalPremiacaoAPagar = totalPremiacao - totalPremiacaoPaga
If totalPremiacaoAPagar < 0 Then totalPremiacaoAPagar = 0

totalAPagar = totalComissaoAPagar + totalPremiacaoAPagar
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
        
        .valor-pendente {
            color: #dc3545;
            font-weight: 600;
        }
        
        .valor-desconto {
            color: #fd7e14;
            font-weight: 600;
        }
        
        .valor-liquido {
            color: #17a2b8;
            font-weight: 600;
        }
        
        .valor-premiacao {
            color: #9b59b6;
            font-weight: 600;
        }
        
        .valor-premiacao-positivo {
            color: #9b59b6;
            font-weight: 700;
            background-color: #f8f9fa;
            border: 1px solid #9b59b6;
            border-radius: 4px;
            padding: 2px 6px;
        }
        
        .premiacao-zero {
            color: #6c757d;
            font-style: italic;
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
            height: 100%;
        }
        
        .info-card-comissao {
            border-left: 4px solid #3498db;
        }
        
        .info-card-premiacao {
            border-left: 4px solid #9b59b6;
        }
        
        .info-card-desconto {
            border-left: 4px solid #fd7e14;
        }
        
        .info-card-liquido {
            border-left: 4px solid #17a2b8;
        }
        
        .info-card-total {
            border-left: 4px solid #2c3e50;
        }
        
        .nome-recebedor {
            font-weight: 600;
            color: #2c3e50;
        }
        
        .desconto-info {
            font-size: 0.8rem;
            color: #6c757d;
        }
        
        .section-comissao {
            border-left: 4px solid #3498db;
            background: linear-gradient(to right, #3498db, #2980b9);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .section-premiacao {
            border-left: 4px solid #9b59b6;
            background: linear-gradient(to right, #9b59b6, #8e44ad);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .section-total {
            border-left: 4px solid #2c3e50;
            background: linear-gradient(to right, #2c3e50, #34495e);
            color: white;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .info-card h6 {
            color: #6c757d;
            font-size: 0.9rem;
            margin-bottom: 0.5rem;
        }
        
        .info-card h4 {
            color: #2c3e50;
            margin-bottom: 0;
            font-weight: 700;
        }
        
        .debug-info {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 10px;
            font-size: 12px;
            color: #856404;
        }
        
        .sql-debug {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 10px;
            font-size: 11px;
            color: #155724;
            font-family: monospace;
            word-break: break-all;
        }
        
        .detalhes-premiacao {
            font-size: 0.8rem;
            color: #6c757d;
        }
    </style>
<style>
    body {
        /* Define a escala de 0.8 (80%) */
        transform: scale(0.8); 
        
        /* Define o ponto de origem para o canto superior esquerdo */
        transform-origin: 0 0; 
        
        /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
        /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
        width: calc(100% / 0.8); 
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
                        <% If diretoriaDetalhe <> "" Then %>
                        - <%= diretoriaDetalhe %>
                        <% End If %>
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
        <!-- Debug Info -->
        <div class="debug-info">
            <strong>Debug Info:</strong> 
            Premiações encontradas: <strong><%= premiacoesDict.Count %></strong> | 
            Premiação Total: <strong>R$ <%= FormatNumber(totalPremiacao, 2) %></strong> | 
            Premiações Pagas: <strong>R$ <%= FormatNumber(totalPremiacaoPaga, 2) %></strong> |
            Comissões Pagas: <strong>R$ <%= FormatNumber(totalComissaoPaga, 2) %></strong>
        </div>

        <!-- Cards de Resumo -->
        <div class="row mb-4">
            <div class="col-md-2">
                <div class="info-card">
                    <h6><i class="fas fa-chart-line me-2"></i>VGV Total</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalVGV, 2) %></h4>
                </div>
            </div>
            <div class="col-md-2">
                <div class="info-card info-card-comissao">
                    <h6><i class="fas fa-money-bill-wave me-2"></i>Comissão Líquida</h6>
                    <h4 class="valor-liquido">R$ <%= FormatNumber(totalComissaoLiquida, 2) %></h4>
                </div>
            </div>
            <div class="col-md-2">
                <div class="info-card info-card-premiacao">
                    <h6><i class="fas fa-trophy me-2"></i>Premiação Total</h6>
                    <h4 class="valor-premiacao">R$ <%= FormatNumber(totalPremiacao, 2) %></h4>
                </div>
            </div>
            <div class="col-md-2">
                <div class="info-card info-card-comissao">
                    <h6><i class="fas fa-hand-holding-usd me-2"></i>Comissões Pagas</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalComissaoPaga, 2) %></h4>
                </div>
            </div>
            <div class="col-md-2">
                <div class="info-card info-card-premiacao">
                    <h6><i class="fas fa-trophy me-2"></i>Premiações Pagas</h6>
                    <h4 class="valor-positivo">R$ <%= FormatNumber(totalPremiacaoPaga, 2) %></h4>
                </div>
            </div>
            <div class="col-md-2">
                <div class="info-card info-card-total">
                    <h6><i class="fas fa-calculator me-2"></i>Total a Pagar</h6>
                    <h4 class="valor-pendente">R$ <%= FormatNumber(totalAPagar, 2) %></h4>
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
                                <th>Comissão Bruta (R$)</th>
                                <th>Desconto (R$)</th>
                                <th>Comissão Líquida (R$)</th>
                                <th>Premiação (R$)</th>
                                <th>%</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If Not rsVendasMes.EOF Then
                                Do While Not rsVendasMes.EOF
                                    Dim classePremiacao, valorPremio, detalhesPremiacao
                                    idVendaAtual = CStr(rsVendasMes("ID"))
                                    
                                    ' Buscar premiação do array/dictionary
                                    If premiacoesDict.Exists(idVendaAtual) Then
                                        valorPremio = premiacoesDict(idVendaAtual)
                                        detalhesPremiacao = "Prêmio"
                                    Else
                                        ' Verificar se há valores iguais de comissão
                                       '' Dim valDir, valGer, valCor
                                        valDir = CDbl(rsVendasMes("ValorDiretoria"))
                                        valGer = CDbl(rsVendasMes("ValorGerencia"))
                                        valCor = CDbl(rsVendasMes("ValorCorretor"))
                                        
                                        If valDir > 0 And valGer > 0 And valCor > 0 Then
                                            If valDir = valGer And valGer = valCor Then
                                                valorPremio = valDir + valGer + valCor
                                                detalhesPremiacao = "Premiação Distribuída"
                                            Else
                                                valorPremio = 0
                                                detalhesPremiacao = ""
                                            End If
                                        Else
                                            valorPremio = 0
                                            detalhesPremiacao = ""
                                        End If
                                    End If
                                    
                                    If valorPremio > 0 Then
                                        classePremiacao = "valor-premiacao-positivo"
                                    Else
                                        classePremiacao = "premiacao-zero"
                                    End If
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
                                <td class="valor-positivo">R$ <%= FormatNumber(rsVendasMes("ComissaoBruta"), 2) %></td>
                                <td class="valor-desconto">
                                    R$ <%= FormatNumber(rsVendasMes("DescontoBruto"), 2) %>
                                    <% If CDbl(rsVendasMes("DescontoBruto")) > 0 Then %>
                                    <br><small class="desconto-info">
                                        <%= FormatNumber((rsVendasMes("DescontoBruto")/rsVendasMes("ComissaoBruta"))*100, 1) %>%
                                    </small>
                                    <% End If %>
                                </td>
                                <td class="valor-liquido">R$ <%= FormatNumber(rsVendasMes("ValorLiqGeral"), 2) %></td>
                                <td class="<%= classePremiacao %>">
                                    <% If valorPremio > 0 Then %>
                                        <strong>R$ <%= FormatNumber(valorPremio, 2) %></strong>
                                        <% If detalhesPremiacao <> "" Then %>
                                        <br><small class="detalhes-premiacao"><%= detalhesPremiacao %></small>
                                        <% End If %>
                                    <% Else %>
                                        R$ <%= FormatNumber(valorPremio, 2) %>
                                    <% End If %>
                                </td>
                                <td><span class="badge bg-info"><%= rsVendasMes("ComissaoPercentual") %>%</span></td>
                            </tr>
                            <%
                                    rsVendasMes.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="13" class="text-center py-4">
                                    <div class="alert alert-info mb-0">
                                        <i class="fas fa-info-circle me-2"></i>Nenhuma venda encontrada para <%= nomeMesDetalhe %>/<%= anoDetalhe %>.
                                    </div>
                                </td>
                            </tr>
                            <%
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-light">
                                <th colspan="7" class="text-end">Totais:</th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalVGV, 2) %></th>
                                <th class="valor-positivo">R$ <%= FormatNumber(totalComissaoBruta, 2) %></th>
                                <th class="valor-desconto">R$ <%= FormatNumber(totalDesconto, 2) %></th>
                                <th class="valor-liquido">R$ <%= FormatNumber(totalComissaoLiquida, 2) %></th>
                                <th class="valor-premiacao">R$ <%= FormatNumber(totalPremiacao, 2) %></th>
                                <th></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>

        <!-- Resto do código permanece igual -->
        <!-- ... -->

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

        <!-- Resumo de Comissões e Premiações -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-calculator me-2"></i>Resumo de Comissões e Premiações</h5>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table class="table table-bordered" style="width:100%">
                        <thead>
                            <tr>
                                <th></th>
                                
                                <!-- Seção Comissão -->
                                <th colspan="3" class="text-center section-comissao">
                                    <i class="fas fa-money-bill-wave me-2"></i>COMISSÃO
                                </th>
                                
                                <!-- Seção Premiação -->
                                <th colspan="3" class="text-center section-premiacao">
                                    <i class="fas fa-trophy me-2"></i>PREMAIAÇÃO
                                </th>
                                
                                <!-- Seção Total -->
                                <th colspan="2" class="text-center section-total">
                                    <i class="fas fa-calculator me-2"></i>TOTAL
                                </th>
                            </tr>
                            <tr>
                                <th></th>
                                
                                <!-- Subheaders Comissão -->
                                <th>Líquida</th>
                                <th>Paga</th>
                                <th>a Pagar</th>
                                
                                <!-- Subheaders Premiação -->
                                <th>Total</th>
                                <th>Paga</th>
                                <th>a Pagar</th>
                                
                                <!-- Subheaders Total -->
                                <th>Pago</th>
                                <th>a Pagar</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>Totais</strong></td>
                                
                                <!-- Dados Comissão -->
                                <td class="valor-liquido">R$ <%= FormatNumber(totalComissaoLiquida, 2) %></td>
                                <td class="valor-positivo">R$ <%= FormatNumber(totalComissaoPaga, 2) %></td>
                                <td class="valor-pendente">R$ <%= FormatNumber(totalComissaoAPagar, 2) %></td>
                                
                                <!-- Dados Premiação -->
                                <td class="valor-premiacao">R$ <%= FormatNumber(totalPremiacao, 2) %></td>
                                <td class="valor-premiacao">R$ <%= FormatNumber(totalPremiacaoPaga, 2) %></td>
                                <td class="valor-pendente">R$ <%= FormatNumber(totalPremiacaoAPagar, 2) %></td>
                                
                                <!-- Dados Total -->
                                <td class="valor-positivo">R$ <%= FormatNumber(totalPago, 2) %></td>
                                <td class="valor-pendente">R$ <%= FormatNumber(totalAPagar, 2) %></td>
                            </tr>
                        </tbody>
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
            order: [[5, 'desc']],
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
'final orinal'