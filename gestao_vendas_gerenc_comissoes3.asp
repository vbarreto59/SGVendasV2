<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="atualizarVendas.asp"-->
<!--#include file="atualizarVendas2.asp"-->

<% 'funcional - ajuste a esquerda colunas diretoria, gerencia e corretor'

Response.CodePage = 65001
Response.CharSet = "UTF-8"

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

' Função para verificar se valor foi pago
Function IsValuePaid(valorPago, valorDevido)
    If valorDevido <= 0 Then
        IsValuePaid = True
    Else
        ' Usar comparação com tolerância para valores monetários
        IsValuePaid = (Abs(valorPago - valorDevido) < 0.01)
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
' Sua consulta principal para as comissões a pagar (ATUALIZADA COM DESCONTOS)
' ====================================================================

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

sqlComissoes = "SELECT c.ID_Comissoes, c.ID_Venda, v.Empreend_ID, v.NomeEmpreendimento, v.Unidade, v.DataVenda, v.ValorComissaoGeral, " & _
               "c.UserIdDiretoria, c.NomeDiretor, v.ComissaoDIretoria, v.ValorDiretoria, v.PremioDiretoria, " & _
               "c.UserIdGerencia, c.NomeGerente, v.ComissaoGerencia, v.ValorGerencia, v.PremioGerencia, " & _
               "c.UserIdCorretor, c.NomeCorretor, v.ComissaoCorretor, v.ValorCorretor, v.PremioCorretor, v.ID, v.Diretoria, v.Gerencia," & _
               "c.StatusPagamento, " & _
               "v.DescontoPerc, v.DescontoBruto, v.DescontoDescricao, " & _
               "v.DescontoDiretoria, v.DescontoGerencia, v.DescontoCorretor, " & _
               "v.ValorLiqDiretoria, v.ValorLiqGerencia, v.ValorLiqCorretor " & _
               "FROM COMISSOES_A_PAGAR AS c INNER JOIN Vendas AS v ON c.ID_Venda = v.ID " & _
               "WHERE v.excluido = 0 ORDER BY c.ID_Comissoes DESC;"

Set rsComissoes = connSales.Execute(sqlComissoes)

' ====================================================================
' Script para Verificar e Atualizar Status de Comissões (PAGA/PENDENTE) - Otimizado
' ====================================================================
Response.Buffer = True
Response.Expires = -1
On Error Resume Next ' Usar On Error Resume Next para melhor tratamento

Dim rsCheckStatus
sqlCheckStatus = "SELECT c.ID_Comissoes, c.ID_Venda, c.StatusPagamento, " & _
                 "v.ValorLiqDiretoria, v.ValorLiqGerencia, v.ValorLiqCorretor " & _
                 "FROM COMISSOES_A_PAGAR c INNER JOIN Vendas v ON c.ID_Venda = v.ID ORDER by c.ID_Comissoes"

Set rsCheckStatus = connSales.Execute(sqlCheckStatus)

If Err.Number <> 0 Then
    Response.Write "Erro na consulta principal: " & Err.Description
    Err.Clear
End If

Do While Not rsCheckStatus.EOF
    Dim comissaoIdCheck, vendaIdCheck, currentStatusComissao
    Dim valorLiqDirCheck, valorLiqGerCheck, valorLiqCorCheck
    Dim totalDirPaid, totalGerPaid, totalCorPaid
    Dim newStatusComissao

    comissaoIdCheck = rsCheckStatus("ID_Comissoes")
    vendaIdCheck = rsCheckStatus("ID_Venda")
    currentStatusComissao = rsCheckStatus("StatusPagamento")
    valorLiqDirCheck = rsCheckStatus("ValorLiqDiretoria")
    valorLiqGerCheck = rsCheckStatus("ValorLiqGerencia")
    valorLiqCorCheck = rsCheckStatus("ValorLiqCorretor")

    totalDirPaid = 0
    totalGerPaid = 0
    totalCorPaid = 0

    Dim sqlGetPaid, rsGetPaid
    
    ' --- Verificar pagamentos para Diretoria (agora na conexão 'connSales') ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'diretoria' AND Left(TipoPagamento,4) = 'Comi'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de diretoria: " & Err.Description & "<br>"
        Err.Clear
    Else
        If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalDirPaid = rsGetPaid("TotalPago")
        If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing
    End If

    ' --- Verificar pagamentos para Gerência ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'gerencia' AND Left(TipoPagamento,4) = 'Comi'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de gerencia: " & Err.Description & "<br>"
        Err.Clear
    Else
        If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalGerPaid = rsGetPaid("TotalPago")
        If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing
    End If

    ' --- Verificar pagamentos para Corretor ---
    sqlGetPaid = "SELECT SUM(ValorPago) as TotalPago FROM PAGAMENTOS_COMISSOES " & _
                 "WHERE ID_Venda = " & vendaIdCheck & " AND TipoRecebedor = 'corretor' AND Left(TipoPagamento,4) = 'Comi'"
    Set rsGetPaid = connSales.Execute(sqlGetPaid)
    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de corretor: " & Err.Description & "<br>"
        Err.Clear
    Else
        If Not rsGetPaid.EOF And Not IsNull(rsGetPaid("TotalPago")) Then totalCorPaid = rsGetPaid("TotalPago")
        If Not rsGetPaid Is Nothing Then rsGetPaid.Close : Set rsGetPaid = Nothing
    End If

    newStatusComissao = "PAGA"
    If CDbl(valorLiqDirCheck) > 0 And CDbl(totalDirPaid) < CDbl(valorLiqDirCheck) Then newStatusComissao = "PENDENTE"
    If CDbl(valorLiqGerCheck) > 0 And CDbl(totalGerPaid) < CDbl(valorLiqGerCheck) Then newStatusComissao = "PENDENTE"
    If CDbl(valorLiqCorCheck) > 0 And CDbl(totalCorPaid) < CDbl(valorLiqCorCheck) Then newStatusComissao = "PENDENTE"

    If newStatusComissao <> currentStatusComissao Then
        sqlUpdateStatus = "UPDATE COMISSOES_A_PAGAR SET StatusPagamento = '" & newStatusComissao & "' WHERE ID_Comissoes = " & comissaoIdCheck
        connSales.Execute(sqlUpdateStatus)
        If Err.Number <> 0 Then
            Response.Write "Erro ao atualizar status: " & Err.Description & "<br>"
            Err.Clear
        End If
    End If

    rsCheckStatus.MoveNext
Loop

If Not rsCheckStatus Is Nothing Then rsCheckStatus.Close
Set rsCheckStatus = Nothing

On Error GoTo 0 ' Restaura o tratamento de erro padrão
%>    
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lista de Comissões a Pagar 1</title>
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
            background-color: #3123E3;
            color: white;
        }
        .status-pendente {
            background-color: #F692E9;
            color: #212529;
        }
        .status-parcial {
            background-color: #54BC4A;
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
        
        /* Estilo para botão de premiação */
        .btn-premio {
            background-color: #ff6b35;
            border-color: #ff6b35;
            color: white;
        }
        .btn-premio:hover {
            background-color: #e55a2b;
            border-color: #e55a2b;
            color: white;
        }
        
        /* NOVOS ESTILOS PARA STATUS DE PAGAMENTO */
        .row-paga {
            background-color: #e3f2fd !important; /* Azul claro */
        }
        .row-pendente {
            background-color: #ffebee !important; /* Vermelho claro */
        }
        .row-parcial {
            background-color: #fff3e0 !important; /* Laranja claro */
        }
        
        /* NOVOS ESTILOS PARA DESCONTOS */
        .desconto-info {
            font-size: 0.75rem;
            color: #6c757d;
        }
        .valor-liquido {
            font-weight: bold;
            color: #28a745;
        }
        .valor-desconto {
            color: #dc3545;
            font-size: 14px;
            text-align: left;
            padding-left: 0;
            margin-left: 0;
        }

        .valor-bruto {
            color: #6c757d;
            font-size: 1.0rem;
        }
        
        /* ESTILOS CORRIGIDOS PARA ÍCONES */
        .valor-bruto, .valor-liquido {
            display: flex;
            align-items: center;
            gap: 4px;
            min-height: 24px;
        }
        .fa-check-circle {
            font-size: 0.9em;
        }
        .payment-indicator {
            color: black;
            font-size: 14px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }
    </style>
</head>
<body>
    <!-- Teste do Font Awesome -->
    <div style="position: fixed; top: 10px; right: 10px; z-index: 10000; background: white; padding: 10px; border: 1px solid red; display: none;">
        <i class="fas fa-check-circle text-success"></i> Ícone teste
    </div>

    <div class="container-fluid mt-5">
        <h2 class="text-center mb-4 header-title"><i class="fas fa-coins me-2"></i>Comissões a Pagar - 1</h2>
        <a href="gestao_vendas_comissao_saldo1.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Saldos</a>
        
        <br><br>
        <div class="table-responsive">

<!-- ########################## -->
<table id="comissoesTable" class="table table-striped table-bordered align-middle nowrap" style="width:100%">
    <thead>
        <tr>
            <th class="text-center">Data/ID</th>
            <th class="text-center">Status</th>
            <th class="text-center">Venda</th>
            <th class="text-center">Diretoria</th>
            <th class="text-center">Gerência</th>
            <th class="text-center">Corretor</th>
            <th class="text-center">Desconto Trib.</th>
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
                ' 🟢 TRATAMENTO ROBUSTO DE VALORES DO RSCOMISSOES (COMISSÕES E PRÊMIOS)
                ' ====================================================================
                Dim dblValorDiretoriaAPagar, dblValorGerenciaAPagar, dblValorCorretorAPagar
                Dim dblPremioDiretoria, dblPremioGerencia, dblPremioCorretor
                Dim totalPremioPagoDiretoria, totalPremioPagoGerencia, totalPremioPagoCorretor
                
                ' 🆕 VARIÁVEIS PARA DESCONTOS E VALORES LÍQUIDOS
                Dim dblDescontoDiretoria, dblDescontoGerencia, dblDescontoCorretor
                Dim dblValorLiqDiretoria, dblValorLiqGerencia, dblValorLiqCorretor
                Dim dblDescontoPerc, dblDescontoBruto, strDescontoDescricao

                ' Valores a Pagar (Comissão - BRUTOS)
                If IsNull(rsComissoes("ValorDiretoria")) Then
                    dblValorDiretoriaAPagar = 0
                Else
                    dblValorDiretoriaAPagar = CDbl(rsComissoes("ValorDiretoria"))
                End If

                If IsNull(rsComissoes("ValorGerencia")) Then
                    dblValorGerenciaAPagar = 0
                Else
                    dblValorGerenciaAPagar = CDbl(rsComissoes("ValorGerencia"))
                End If

                If IsNull(rsComissoes("ValorCorretor")) Then
                    dblValorCorretorAPagar = 0
                Else
                    dblValorCorretorAPagar = CDbl(rsComissoes("ValorCorretor"))
                End If

                ' 🆕 VALORES DE DESCONTO
                If IsNull(rsComissoes("DescontoDiretoria")) Then
                    dblDescontoDiretoria = 0
                Else
                    dblDescontoDiretoria = CDbl(rsComissoes("DescontoDiretoria"))
                End If

                If IsNull(rsComissoes("DescontoGerencia")) Then
                    dblDescontoGerencia = 0
                Else
                    dblDescontoGerencia = CDbl(rsComissoes("DescontoGerencia"))
                End If

                If IsNull(rsComissoes("DescontoCorretor")) Then
                    dblDescontoCorretor = 0
                Else
                    dblDescontoCorretor = CDbl(rsComissoes("DescontoCorretor"))
                End If

                ' 🆕 VALORES LÍQUIDOS (PARA PAGAMENTO)

                If IsNull(rsComissoes("ValorLiqDiretoria")) Then
                    dblValorLiqDiretoria = dblValorDiretoriaAPagar - dblDescontoDiretoria
                Else
                    dblValorLiqDiretoria = CDbl(rsComissoes("ValorLiqDiretoria"))
                End If

                If IsNull(rsComissoes("ValorLiqGerencia")) Then
                    dblValorLiqGerencia = dblValorGerenciaAPagar - dblDescontoGerencia
                Else
                    dblValorLiqGerencia = CDbl(rsComissoes("ValorLiqGerencia"))
                End If

                If IsNull(rsComissoes("ValorLiqCorretor")) Then
                    dblValorLiqCorretor = dblValorCorretorAPagar - dblDescontoCorretor
                Else
                    dblValorLiqCorretor = CDbl(rsComissoes("ValorLiqCorretor"))
                End If

                ' 🆕 PERCENTUAL E DESCRIÇÃO DO DESCONTO
                If IsNull(rsComissoes("DescontoPerc")) Then
                    dblDescontoPerc = 0
                Else
                    dblDescontoPerc = CDbl(rsComissoes("DescontoPerc"))
                End If

                If IsNull(rsComissoes("DescontoBruto")) Then
                    dblDescontoBruto = dblDescontoDiretoria + dblDescontoGerencia + dblDescontoCorretor
                Else
                    dblDescontoBruto = CDbl(rsComissoes("DescontoBruto"))
                End If

                If IsNull(rsComissoes("DescontoDescricao")) Then
                    strDescontoDescricao = ""
                Else
                    strDescontoDescricao = rsComissoes("DescontoDescricao")
                End If

                ' Valores de Prêmio (NÃO SOFREM DESCONTO)
                If IsNull(rsComissoes("PremioDiretoria")) Then
                    dblPremioDiretoria = 0
                Else
                    dblPremioDiretoria = CDbl(rsComissoes("PremioDiretoria"))
                End If

                If IsNull(rsComissoes("PremioGerencia")) Then
                    dblPremioGerencia = 0
                Else
                    dblPremioGerencia = CDbl(rsComissoes("PremioGerencia"))
                End If

                If IsNull(rsComissoes("PremioCorretor")) Then
                    dblPremioCorretor = 0
                Else
                    dblPremioCorretor = CDbl(rsComissoes("PremioCorretor"))
                End If

                ' ====================================================================
                ' Consulta para obter os pagamentos já realizados (COMISSÕES)
                ' ====================================================================
                ' #### Pagamentos para Diretoria (Comissão)
                sqlPagamentos = "SELECT Sum(ValorPago) AS ValorTotalPago, MAX(DataPagamento) as DtPagamento  " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE PAGAMENTOS_COMISSOES.ID_Venda=" & vendaId & " " & _
                                "AND PAGAMENTOS_COMISSOES.UsuariosUserId=" & userIdDiretoria & " " & _
                                "AND PAGAMENTOS_COMISSOES.TipoRecebedor='diretoria' " & _
                                "AND PAGAMENTOS_COMISSOES.TipoPagamento='Comissão';"
                                        

                'Response.Write sqlPagamentos
                'Response.end                                         
                Set rsPagamentos = connSales.Execute(sqlPagamentos)

                Dim dataPagamentoDiretoria
                If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                    totalPagoDiretoria = rsPagamentos("ValorTotalPago")
                    If Not IsNull(rsPagamentos("DtPagamento")) Then
                        dataPagamentoDiretoria = FormatDateTime(rsPagamentos("DtPagamento"), 2)
                    End If
                End If
                If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing


                '============================================================='
                ' #### Pagamentos para Gerência (Comissão)
                '============================================================='
                sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago, MAX(DataPagamento) as DataPag " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE ID_Venda = " & vendaId & " AND UsuariosUserId = " & userIdGerencia & " AND TipoRecebedor = 'gerencia' AND Left(TipoPagamento,4) = 'Comi'"
                Set rsPagamentos = connSales.Execute(sqlPagamentos)
                Dim dataPagamentoGerencia
                If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                    totalPagoGerencia = rsPagamentos("ValorTotalPago")
                    If Not IsNull(rsPagamentos("DataPag")) Then
                        dataPagamentoGerencia = FormatDateTime(rsPagamentos("DataPag"), 2)
                    End If
                End If
                If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing



                '============================================================='
                ' #### Pagamentos para Corretor (Comissão)
                '============================================================='
                sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago, MAX(DataPagamento) as DataPag " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE ID_Venda = " & vendaId & " AND UsuariosUserId = " & userIdCorretor & " AND TipoRecebedor = 'corretor' AND Left(TipoPagamento,4) = 'Comi'"
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
                ' Consulta para obter os pagamentos já realizados (PRÊMIOS)
                ' ====================================================================
                ' #### Pagamentos para Diretoria (Prêmio)
                sqlPagamentos = "SELECT Sum(ValorPago) AS ValorTotalPago " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE PAGAMENTOS_COMISSOES.ID_Venda=" & vendaId & " " & _
                                "AND PAGAMENTOS_COMISSOES.UsuariosUserId=" & userIdDiretoria & " " & _
                                "AND PAGAMENTOS_COMISSOES.TipoRecebedor='diretoria' " & _
                                "AND PAGAMENTOS_COMISSOES.TipoPagamento='Premiação';"
                                        
                Set rsPagamentos = connSales.Execute(sqlPagamentos)
                If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                    totalPremioPagoDiretoria = rsPagamentos("ValorTotalPago")
                Else
                    totalPremioPagoDiretoria = 0
                End If
                If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing

                ' #### Pagamentos para Gerência (Prêmio)
                sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE ID_Venda = " & vendaId & " AND UsuariosUserId = " & userIdGerencia & " AND TipoRecebedor = 'gerencia' AND Left(TipoPagamento,3) = 'Pre'"
                Set rsPagamentos = connSales.Execute(sqlPagamentos)
                If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                    totalPremioPagoGerencia = rsPagamentos("ValorTotalPago")
                Else
                    totalPremioPagoGerencia = 0
                End If
                If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing

                ' #### Pagamentos para Corretor (Prêmio)
                sqlPagamentos = "SELECT SUM(ValorPago) as ValorTotalPago " & _
                                "FROM PAGAMENTOS_COMISSOES " & _
                                "WHERE ID_Venda = " & vendaId & " AND UsuariosUserId = " & userIdCorretor & " AND TipoRecebedor = 'corretor' AND Left(TipoPagamento,3) = 'Pre'"
                Set rsPagamentos = connSales.Execute(sqlPagamentos)
                If Not rsPagamentos.EOF And Not IsNull(rsPagamentos("ValorTotalPago")) Then
                    totalPremioPagoCorretor = rsPagamentos("ValorTotalPago")
                Else
                    totalPremioPagoCorretor = 0
                End If
                If IsObject(rsPagamentos) Then rsPagamentos.Close : Set rsPagamentos = Nothing
                
                ' ====================================================================
                ' 🟢 CORREÇÃO: Determina o status da comissão COM LÓGICA CORRIGIDA
                ' ====================================================================
                Dim comissaoDiretoriaPaga, comissaoGerenciaPaga, comissaoCorretorPaga
                Dim premioDiretoriaPago, premioGerenciaPago, premioCorretorPago

                ' ✅ CORREÇÃO APLICADA: Usar função personalizada para comparação monetária
                comissaoDiretoriaPaga = IsValuePaid(totalPagoDiretoria, dblValorLiqDiretoria)
                comissaoGerenciaPaga = IsValuePaid(totalPagoGerencia, dblValorLiqGerencia)
                comissaoCorretorPaga = IsValuePaid(totalPagoCorretor, dblValorLiqCorretor)

                premioDiretoriaPago = IsValuePaid(totalPremioPagoDiretoria, dblPremioDiretoria)
                premioGerenciaPago = IsValuePaid(totalPremioPagoGerencia, dblPremioGerencia)
                premioCorretorPago = IsValuePaid(totalPremioPagoCorretor, dblPremioCorretor)

                ' ====================================================================
                ' 🆕 VERIFICAÇÃO COMPLETA DO STATUS DE PAGAMENTO (COM VALORES LÍQUIDOS)
                ' ====================================================================
                Dim status, statusClass, rowClass
                status = rsComissoes("StatusPagamento")
                
                Dim todasComissoesPagas, todosPremiosPagos, statusCompleto
                
                ' Verifica se todas as comissões estão pagas (USANDO VALORES LÍQUIDOS)
                todasComissoesPagas = True
                If dblValorLiqDiretoria > 0 And Not comissaoDiretoriaPaga Then todasComissoesPagas = False
                If dblValorLiqGerencia > 0 And Not comissaoGerenciaPaga Then todasComissoesPagas = False
                If dblValorLiqCorretor > 0 And Not comissaoCorretorPaga Then todasComissoesPagas = False
                
                ' Verifica se todos os prêmios estão pagos (PRÊMIOS NÃO SOFREM DESCONTO)
                todosPremiosPagos = True
                If dblPremioDiretoria > 0 And Not premioDiretoriaPago Then todosPremiosPagos = False
                If dblPremioGerencia > 0 And Not premioGerenciaPago Then todosPremiosPagos = False
                If dblPremioCorretor > 0 And Not premioCorretorPago Then todosPremiosPagos = False
                
                ' Determina o status completo
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
                
                ' Atualiza o status no banco de dados se necessário
                If UCase(status) <> UCase(statusCompleto) Then
                    sqlUpdateStatus = "UPDATE COMISSOES_A_PAGAR SET StatusPagamento = '" & statusCompleto & "' WHERE ID_Comissoes = " & comissaoId
                    connSales.Execute(sqlUpdateStatus)
                    status = statusCompleto
                End If

                ' Define a classe do status
                Select Case UCase(status)
                    Case "PAGA": statusClass = "status-pago"
                    Case "PAGA PARCIALMENTE": statusClass = "status-parcial"
                    Case "PENDENTE": statusClass = "status-pendente"
                    Case Else: statusClass = "bg-secondary text-white"
                End Select
        %>
        <tr class="<%= rowClass %>">
            <td class="text-center"><%= Year(rsComissoes("DataVenda")) & "-" & Right("0" & Month(rsComissoes("DataVenda")),2) & "-" & Right("0" & Day(rsComissoes("DataVenda")),2) %><br><%="V"&vendaID%>-<%="C"& rsComissoes("ID_Comissoes")%>
            </td>
            <td 
            <%
              If statusClass = "status-pago" then %>
                class="text-center"><span class="status-badge bg-success <%= statusClass %>"><%= UCase(status) %></span>
              <%else%>  
                class="text-center"><span class="status-badge bg-info <%= statusClass %> text-white"><%= UCase(status) %></span>
             <%end if%>   
            </td>
            <td class="text-center">
                <small class="text-muted"><b><%= rsComissoes("Empreend_ID") %>-<%= rsComissoes("NomeEmpreendimento") %></b></small><br>
                <%= rsComissoes("Unidade") %> <br>

                <!-- somando as comissoes brutas 07 11 2025 -->
                <% vComisBruta = dblValorDiretoriaAPagar + dblValorGerenciaAPagar + dblValorCorretorAPagar%>

               <span class="badge bg-primary">
                    Comissão Bruta: R$ <%= FormatNumber(vComisBruta, 2) %>
                </span><br>

               <span class="badge bg-danger">
                    Desconto:  R$ <%= FormatNumber(dblDescontoBruto, 2) %>
               </span><br>

               <span class="badge bg-success">
                    Saldo:  R$ <%= FormatNumber(vComisBruta-dblDescontoBruto, 2) %>
               </span><br>

                
            </td>
            <!-- ####################################################################### -->
            <% ' ----------------------------------------------------------------- %>
            <% ' COLUNA DIRETORIA: COM DESCONTOS E VALORES LÍQUIDOS - CORRIGIDO %>
            <% ' ----------------------------------------------------------------- %>
            <td class="text-center">
                <div><b><%= rsComissoes("Diretoria") %><br><%= userIdDiretoria&"-"&rsComissoes("NomeDiretor") %></b></div>
                
                <% ' COMISSÃO DIRETORIA - COM DESCONTO %>
                <div class="valor-bruto">
                    <span class="badge bg-info">
                        <span class="payment-indicator">
                            <% If comissaoDiretoriaPaga Then %>
                                <i class="fas fa-plus text-success me-1" title="Comissão totalmente paga"></i>
                            <% End If %>
                            Bruto: R$ <%= FormatNumber(dblValorDiretoriaAPagar, 2) %>
                        </span>
                    </span>
                </div>
                
                <% If dblDescontoDiretoria > 0 Then %>
                <div class="valor-desconto">
                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoDiretoria, 2) %>
                </div>
                <% End If %>
                
                <div class="valor-liquido">
                    <span class="payment-indicator">
                        <% If comissaoDiretoriaPaga Then %>
                           <span class="badge bg-success me-1" title="Comissão totalmente paga">PAGA</span>
                        <% End If %>
                        <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqDiretoria, 2) %>
                    </span>
                </div>
                
                
                <% ' ############# PRÊMIO DIRETORIA (NÃO SOFRE DESCONTO) %>
                <% If dblPremioDiretoria > 0 Then %>
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <span class="text-info fw-bold payment-indicator">
                            <% If premioDiretoriaPago Then %>
                                <span class="badge bg-success me-1" title="Prêmio totalmente pago">PAGA</span>
                            <% ElseIf totalPremioPagoDiretoria > 0 Then %>
                                <i class="fas fa-check-circle text-warning me-1" title="Prêmio parcialmente pago"></i>
                            <% End If %>
                            <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioDiretoria, 2) %>
                        </span>
                    </div>
                <% End If %>

            <!-- informar o total pago: comissao + premio 07 11 2025-->
                <% 
                    ' 1. Calcular o total
                    Dim dblTotalPagoDiretoria
                    dblTotalPagoDiretoria = dblValorLiqDiretoria + dblPremioDiretoria
                    
                    ' 2. Verificar o status de pagamento total
                    Dim totalDiretoriaPago
                    totalDiretoriaPago = comissaoDiretoriaPaga And premioDiretoriaPago

                    ' 3. Verificar se houve pagamento parcial (opcional, mas útil)
                    Dim totalPagoParcial
                    totalPagoParcial = (comissaoDiretoriaPaga Or totalPremioPagoDiretoria > 0) And Not totalDiretoriaPago
                %>

                <div class="total-pago mt-2 pt-2 border-top border-2">
                    <div class="d-flex justify-content-between align-items-center">

                        <span class="payment-total-indicator style='font-size: 14px;'">
                            <% If totalDiretoriaPago Then %>
                                <span class="badge bg-success me-1" title="Comissão e Prêmio totalmente pagos">TOTAL PAGO</span>
                            <% ElseIf totalPagoParcial Then %>
                                <span class="badge bg-warning me-1" title="Alguns itens pagos parcialmente">PARCIAL</span>
                            <% End If %>
                            
                            <i class="fas fa-money-bill-wave text-primary"></i> 
                            <strong class="text-primary">R$ <%= FormatNumber(dblTotalPagoDiretoria, 2) %></strong>
                        </span>
                    </div>
                </div>


            </td>

            <% ' ----------------------------------------------------------------- %>
            <% ' COLUNA GERÊNCIA: COM DESCONTOS E VALORES LÍQUIDOS - CORRIGIDO %>
            <% ' ----------------------------------------------------------------- %>
            <td class="text-center">
                <div><b><%= rsComissoes("Gerencia") %><br><%= userIdGerencia&"-"& rsComissoes("NomeGerente") %></b></div>
                
                <% ' COMISSÃO GERÊNCIA - COM DESCONTO %>
                <div class="valor-bruto">
                    <span class="badge bg-info">
                        <span class="payment-indicator">
                            <% If comissaoGerenciaPaga Then %>
                                <i class="fas fa-plus text-success me-1" title="Comissão totalmente paga"></i>
                            <% End If %>
                            Bruto: R$ <%= FormatNumber(dblValorGerenciaAPagar, 2) %>
                        </span>
                    </span>    
                </div>
                
                <% If dblDescontoGerencia > 0 Then %>
                <div class="valor-desconto">
                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoGerencia, 2) %>
                </div>
                <% End If %>
                
                <div class="valor-liquido">
                    <span class="payment-indicator">
                        <% If comissaoGerenciaPaga Then %>
                            <span class="badge bg-success me-1" title="Comissão totalmente paga">PAGA</span>
                        <% End If %>
                        <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqGerencia, 2) %>
                    </span>
                </div>
                

                
                <% ' PRÊMIO GERÊNCIA (NÃO SOFRE DESCONTO) %>
                <% If dblPremioGerencia > 0 Then %>
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <span class="text-info fw-bold payment-indicator">
                            <% If premioGerenciaPago Then %>
                                <span class="badge bg-success me-1" title="Comissão totalmente paga">PAGA</span>
                            <% ElseIf totalPremioPagoGerencia > 0 Then %>
                                <i class="fas fa-plus text-warning me-1" title="Prêmio parcialmente pago"></i>
                            <% End If %>
                            <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioGerencia, 2) %>
                        </span>
                    </div>
                <% End If %>


                <% 
                    ' ############# CÁLCULO TOTAL PAGO GERÊNCIA #############

                    ' 1. Calcular o total (Comissão Líquida + Prêmio)
                    Dim dblTotalPagoGerencia
                    ' Assumindo que essas variáveis contêm os valores corretos da Gerência
                    dblTotalPagoGerencia = dblValorLiqGerencia + dblPremioGerencia
                    
                    ' 2. Verificar o status de pagamento total (Somente se ambos pagos)
                    Dim totalGerenciaPago
                    totalGerenciaPago = comissaoGerenciaPaga And premioGerenciaPago

                    ' 3. Verificar se houve pagamento parcial (Pelo menos um pago, mas não o total)
                    Dim totalPagoParcialGerencia
                    totalPagoParcialGerencia = (comissaoGerenciaPaga Or totalPremioPagoGerencia > 0) And Not totalGerenciaPago
                %>

                <div class="total-pago mt-2 pt-2 border-top border-2">
                    <div class="d-flex justify-content-between align-items-center">
                        <span class="payment-total-indicator" style="font-size: 14px;">
                            <% If totalGerenciaPago Then %>
                                <span class="badge bg-success me-1" title="Comissão e Prêmio totalmente pagos">TOTAL PAGO</span>
                            <% ElseIf totalPagoParcialGerencia Then %>
                                <span class="badge bg-warning me-1" title="Alguns itens pagos parcialmente">PARCIAL</span>
                            <% End If %>
                            
                            <i class="fas fa-money-bill-wave text-primary"></i> 
                            <strong class="text-primary">R$ <%= FormatNumber(dblTotalPagoGerencia, 2) %></strong>
                        </span>
                    </div>
                </div>

            </td>

            <% ' ----------------------------------------------------------------- %>
            <% ' COLUNA CORRETOR: COM DESCONTOS E VALORES LÍQUIDOS - CORRIGIDO %>
            <% ' ----------------------------------------------------------------- %>
            <td class="text-center">
                <div><b><%= userIdCorretor &"-"&rsComissoes("NomeCorretor") %></b></div>
                
                <% ' COMISSÃO CORRETOR - COM DESCONTO %>
                <div class="valor-bruto">
                    <span class="badge bg-info">
                        <span class="payment-indicator">
                            <% If comissaoCorretorPaga Then %>
                                <i class="fas fa-plus text-dark me-1" title="Comissão totalmente paga"></i>
                            <% End If %>
                            Bruto: R$ <%= FormatNumber(dblValorCorretorAPagar, 2) %>
                        </span>
                    </span>
                </div>
                
                <% If dblDescontoCorretor > 0 Then %>
                <div class="valor-desconto">
                    <i class="fas fa-minus-circle"></i> R$ <%= FormatNumber(dblDescontoCorretor, 2) %>
                </div>
                <% End If %>
                
                <div class="valor-liquido">
                    <span class="payment-indicator">
                        <% If comissaoCorretorPaga Then %>
                            <span class="badge bg-success me-1" title="Comissão totalmente paga">PAGA</span>
                        <% End If %>
                        <i class="fas fa-hand-holding-usd"></i> R$ <%= FormatNumber(dblValorLiqCorretor, 2) %>
                    </span>
                </div>
                

                
                <% ' PRÊMIO CORRETOR (NÃO SOFRE DESCONTO) %>
                <% If dblPremioCorretor > 0 Then %>
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <span class="text-info fw-bold payment-indicator">
                            <% If premioCorretorPago Then %>
                                <span class="badge bg-success me-1" title="Comissão totalmente paga">PAGA</span>
                            <% ElseIf totalPremioPagoCorretor > 0 Then %>
                                <i class="fas fa-plus text-warning me-1" title="Prêmio parcialmente pago"></i>
                            <% End If %>
                            <i class="fas fa-trophy"></i> R$ <%= FormatNumber(dblPremioCorretor, 2) %>
                        </span>
                    </div>
                    
                <% End If %>


            <% 
                ' ############# CÁLCULO TOTAL PAGO CORRETOR #############

                ' 1. Calcular o total (Comissão Líquida + Prêmio)
                Dim dblTotalPagoCorretor
                ' Certifique-se de que as variáveis do Corretor estão disponíveis
                dblTotalPagoCorretor = dblValorLiqCorretor + dblPremioCorretor
                
                ' 2. Verificar o status de pagamento total (Somente se ambos pagos)
                Dim totalCorretorPago
                totalCorretorPago = comissaoCorretorPaga And premioCorretorPago

                ' 3. Verificar se houve pagamento parcial (Pelo menos um pago, mas não o total)
                Dim totalPagoParcialCorretor
                totalPagoParcialCorretor = (comissaoCorretorPaga Or totalPremioPagoCorretor > 0) And Not totalCorretorPago
            %>

            <div class="total-pago mt-2 pt-2 border-top border-2">
                <div class="d-flex justify-content-between align-items-center">

                    <span class="payment-total-indicator" style="font-size: 14px;">
                        <% If totalCorretorPago Then %>
                            <span class="badge bg-success me-1" title="Comissão e Prêmio totalmente pagos">TOTAL PAGO</span>
                        <% ElseIf totalPagoParcialCorretor Then %>
                            <span class="badge bg-warning me-1" title="Alguns itens pagos parcialmente">PARCIAL</span>
                        <% End If %>
                        
                        <i class="fas fa-money-bill-wave text-primary"></i> 
                        <strong class="text-primary">R$ <%= FormatNumber(dblTotalPagoCorretor, 2) %></strong>
                    </span>
                </div>
            </div>



            </td>
           <!-- ################################################# -->
            <% ' ----------------------------------------------------------------- %>
            <% ' COLUNA DESCONTO TRIBUTÁRIO %>
            <% ' ----------------------------------------------------------------- %>
            <td class="text-center">
                <% If dblDescontoPerc > 0 Then %>
                    <div class="desconto-info">
                        <strong><%= FormatNumber(dblDescontoPerc, 2) %>%</strong>
                        <br>
                        <small>Total: R$ <%= FormatNumber(dblDescontoBruto, 2) %></small>
                        <% If strDescontoDescricao <> "" Then %>
                            <br>
                            <small title="<%= strDescontoDescricao %>">
                                <i class="fas fa-info-circle"></i> <%= Left(strDescontoDescricao, 20) & "..." %>
                            </small>
                        <% End If %>
                    </div>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>

            <td class="text-center">
                <button class="btn btn-primary btn-sm mb-1" 
                    data-bs-toggle="modal" data-bs-target="#paymentModal"
                    data-id-comissao="<%= rsComissoes("ID_Comissoes") %>"
                    data-id-venda="<%= rsComissoes("ID_Venda") %>"
                    data-diretoria-id="<%= userIdDiretoria %>"
                    data-diretoria-nome="<%= rsComissoes("NomeDiretor") %>"
                    data-diretoria-apagar="<%= FormatNumber(dblValorLiqDiretoria, 2) %>"
                    data-diretoria-pago="<%= FormatNumber(totalPagoDiretoria, 2) %>"
                    data-gerencia-id="<%= userIdGerencia %>"
                    data-gerencia-nome="<%= rsComissoes("NomeGerente") %>"
                    data-gerencia-apagar="<%= FormatNumber(dblValorLiqGerencia, 2) %>"
                    data-gerencia-pago="<%= FormatNumber(totalPagoGerencia, 2) %>"
                    data-corretor-id="<%= userIdCorretor %>"
                    data-corretor-nome="<%= rsComissoes("NomeCorretor") %>"
                    data-corretor-apagar="<%= FormatNumber(dblValorLiqCorretor, 2) %>"
                    data-corretor-pago="<%= FormatNumber(totalPagoCorretor, 2) %>"
                >
                    <i class="fas fa-hand-holding-usd"></i> Pagar Comiss.
                </button>

                <% If dblPremioDiretoria > 0 Or dblPremioGerencia > 0 Or dblPremioCorretor > 0 Then %>
                <button class="btn btn-premio btn-sm mb-1" 
                    data-bs-toggle="modal" data-bs-target="#premioModal"
                    data-id-comissao="<%= rsComissoes("ID_Comissoes") %>"
                    data-id-venda="<%= rsComissoes("ID_Venda") %>"
                    data-diretoria-id="<%= userIdDiretoria %>"
                    data-diretoria-nome="<%= rsComissoes("NomeDiretor") %>"
                    data-diretoria-premio="<%= FormatNumber(dblPremioDiretoria, 2) %>"
                    data-diretoria-premio-pago="<%= FormatNumber(totalPremioPagoDiretoria, 2) %>"
                    data-gerencia-id="<%= userIdGerencia %>"
                    data-gerencia-nome="<%= rsComissoes("NomeGerente") %>"
                    data-gerencia-premio="<%= FormatNumber(dblPremioGerencia, 2) %>"
                    data-gerencia-premio-pago="<%= FormatNumber(totalPremioPagoGerencia, 2) %>"
                    data-corretor-id="<%= userIdCorretor %>"
                    data-corretor-nome="<%= rsComissoes("NomeCorretor") %>"
                    data-corretor-premio="<%= FormatNumber(dblPremioCorretor, 2) %>"
                    data-corretor-premio-pago="<%= FormatNumber(totalPremioPagoCorretor, 2) %>"
                >
                    <i class="fas fa-trophy"></i> Pagar Prêmio
                </button><br>
                <% End If %>

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
<!-- ################################# -->
        </div>
    </div>

    <!-- Modal de Pagamento de Comissão -->
    <div class="modal fade" id="paymentModal" tabindex="-1" aria-labelledby="paymentModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="paymentModalLabel">Realizar Pagamento de Comissão</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <form id="paymentForm" action="gestao_vendas_salvar_pagamento.asp" method="post">
                    <input type="hidden" name="TipoPagamento" value="Comissao">
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
                            <label class="form-label">Valor Líquido a Pagar:</label>
                            <p class="form-control-plaintext" id="modalValorAPagarTotal">R$ 0,00</p>
                            <small class="text-muted">(Valor já considera desconto tributário)</small>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Valor Já Pago:</label>
                            <p class="form-control-plaintext" id="modalValorJaPago">R$ 0,00</p>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Saldo Líquido a Pagar:</label>
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

    <!-- Modal de Pagamento de Prêmio -->
    <div class="modal fade" id="premioModal" tabindex="-1" aria-labelledby="premioModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header bg-warning">
                    <h5 class="modal-title" id="premioModalLabel">Realizar Pagamento de Prêmio</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <form id="premioForm" action="gestao_vendas_salvar_premio.asp" method="post">
                    <input type="hidden" name="TipoPagamento" value="Premiacao">
                    <div class="modal-body">
                        <input type="hidden" id="premioModalComissaoId" name="ID_Comissao">
                        <input type="hidden" id="premioModalVendaId" name="ID_Venda">
                        <input type="hidden" id="premioModalUserId" name="UserId">

                        <div class="mb-3">
                            <label for="premioModalRecipient" class="form-label">Para quem será o pagamento do prêmio?</label>
                            <select class="form-select" id="premioModalRecipient" name="RecipientType" required>
                                <option value="">Selecione...</option>
                                <!-- Opções preenchidas via JS -->
                            </select>
                        </div>

                        <div class="mb-3">
                            <label class="form-label">Valor Total do Prêmio:</label>
                            <p class="form-control-plaintext" id="premioModalValorTotal">R$ 0,00</p>
                            <small class="text-muted">(Prêmio não sofre desconto tributário)</small>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Prêmio Já Pago:</label>
                            <p class="form-control-plaintext" id="premioModalValorJaPago">R$ 0,00</p>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Saldo do Prêmio a Pagar:</label>
                            <p class="form-control-plaintext" id="premioModalSaldoAPagar"
                               style="font-weight: bold; color: #ff0000; font-size: 1.2em;">R$ 0,00</p>
                        </div>

                        <div class="mb-3">
                            <label for="premioModalValorAPagarInput" class="form-label">Valor do Prêmio a Pagar (nesta transação) *</label>
                            <input type="text" class="form-control" id="premioModalValorAPagarInput" name="ValorPago" required>
                        </div>
                        <div class="mb-3">
                            <label for="premioModalDataPagamento" class="form-label">Data do Pagamento *</label>
                            <input type="date" class="form-control" id="premioModalDataPagamento" name="DataPagamento" required>
                        </div>
                        <div class="mb-3">
                            <label for="premioModalStatusPagamento" class="form-label">Status do Pagamento *</label>
                            <select class="form-select" id="premioModalStatusPagamento" name="Status" required>
                                <option value="">Selecione...</option>
                                <option value="Em processamento">Em processamento</option>
                                <option value="Agendado">Agendado</option>
                                <option value="Realizado">Realizado</option>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="premioModalObs" class="form-label">Observações</label>
                            <textarea class="form-control" id="premioModalObs" name="Obs" rows="3"></textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="submit" class="btn btn-warning">Salvar Pagamento do Prêmio</button>
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
                                    <th>Tipo</th>
                                    <th class="text-end">Valor</th>
                                    <th>Destinatário</th>
                                    <th>Cargo</th>
                                    <th>Status</th>
                                    <th>Observações</th>
                                    <th>Ações</th>
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
        $('#paymentsTableBody').html('<tr><td colspan="9" class="text-center"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Carregando...</span></div></td></tr>');
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
                                <td>${payment.TipoPagamento}</td>
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
                    $('#paymentsTableBody').html('<tr><td colspan="9" class="text-center">Nenhum pagamento encontrado.</td></tr>');
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
                        <td colspan="9" class="text-center text-danger">
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
            order: [[0, "desc"]],
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

        // Máscara para os campos de valor nos modais
        $('#modalValorAPagarInput, #premioModalValorAPagarInput').mask('#.##0,00', { reverse: true });

        // Preenche a data atual nos campos de data dos modais
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const todayStr = `${year}-${month}-${day}`;
        $('#modalDataPagamento, #premioModalDataPagamento').val(todayStr);

        // ====================================================================
        // EVENTOS PARA MODAL DE COMISSÃO (USANDO VALORES LÍQUIDOS)
        // ====================================================================
        $('#paymentModal').on('show.bs.modal', function(event) {
            const button = $(event.relatedTarget);
            const idComissao = button.data('id-comissao');
            const idVenda = button.data('id-venda');

            const diretoriaId = button.data('diretoria-id');
            const diretoriaNome = button.data('diretoria-nome');
            const diretoriaAPagar = parseCurrency(button.data('diretoria-apagar')); // VALOR LÍQUIDO
            const diretoriaPago = parseCurrency(button.data('diretoria-pago'));

            const gerenciaId = button.data('gerencia-id');
            const gerenciaNome = button.data('gerencia-nome');
            const gerenciaAPagar = parseCurrency(button.data('gerencia-apagar')); // VALOR LÍQUIDO
            const gerenciaPago = parseCurrency(button.data('gerencia-pago'));

            const corretorId = button.data('corretor-id');
            const corretorNome = button.data('corretor-nome');
            const corretorAPagar = parseCurrency(button.data('corretor-apagar')); // VALOR LÍQUIDO
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

            if (diretoriaId && diretoriaNome && diretoriaId !== 0 && diretoriaAPagar > 0) {
                recipientSelect.append(`<option value="diretoria" data-user-id="${diretoriaId}">${diretoriaNome} (Diretoria)</option>`);
            }
            if (gerenciaId && gerenciaNome && gerenciaId !== 0 && gerenciaAPagar > 0) {
                recipientSelect.append(`<option value="gerencia" data-user-id="${gerenciaId}">${gerenciaNome} (Gerência)</option>`);
            }
            if (corretorId && corretorNome && corretorId !== 0 && corretorAPagar > 0) {
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

        // ====================================================================
        // EVENTOS PARA MODAL DE PRÊMIO (VALORES BRUTOS - NÃO SOFREM DESCONTO)
        // ====================================================================
        $('#premioModal').on('show.bs.modal', function(event) {
            const button = $(event.relatedTarget);
            const idComissao = button.data('id-comissao');
            const idVenda = button.data('id-venda');

            const diretoriaId = button.data('diretoria-id');
            const diretoriaNome = button.data('diretoria-nome');
            const diretoriaPremio = parseCurrency(button.data('diretoria-premio'));
            const diretoriaPremioPago = parseCurrency(button.data('diretoria-premio-pago'));

            const gerenciaId = button.data('gerencia-id');
            const gerenciaNome = button.data('gerencia-nome');
            const gerenciaPremio = parseCurrency(button.data('gerencia-premio'));
            const gerenciaPremioPago = parseCurrency(button.data('gerencia-premio-pago'));

            const corretorId = button.data('corretor-id');
            const corretorNome = button.data('corretor-nome');
            const corretorPremio = parseCurrency(button.data('corretor-premio'));
            const corretorPremioPago = parseCurrency(button.data('corretor-premio-pago'));

            const modal = $(this);
            modal.data('diretoria', { id: diretoriaId, nome: diretoriaNome, premio: diretoriaPremio, premioPago: diretoriaPremioPago });
            modal.data('gerencia', { id: gerenciaId, nome: gerenciaNome, premio: gerenciaPremio, premioPago: gerenciaPremioPago });
            modal.data('corretor', { id: corretorId, nome: corretorNome, premio: corretorPremio, premioPago: corretorPremioPago });

            $('#premioModalComissaoId').val(idComissao);
            $('#premioModalVendaId').val(idVenda);

            const recipientSelect = $('#premioModalRecipient');
            recipientSelect.empty();
            recipientSelect.append('<option value="">Selecione...</option>');

            if (diretoriaId && diretoriaNome && diretoriaId !== 0 && diretoriaPremio > 0) {
                recipientSelect.append(`<option value="diretoria" data-user-id="${diretoriaId}">${diretoriaNome} (Diretoria)</option>`);
            }
            if (gerenciaId && gerenciaNome && gerenciaId !== 0 && gerenciaPremio > 0) {
                recipientSelect.append(`<option value="gerencia" data-user-id="${gerenciaId}">${gerenciaNome} (Gerência)</option>`);
            }
            if (corretorId && corretorNome && corretorId !== 0 && corretorPremio > 0) {
                recipientSelect.append(`<option value="corretor" data-user-id="${corretorId}">${corretorNome} (Corretor)</option>`);
            }

            $('#premioModalValorTotal').text('R$ 0,00');
            $('#premioModalValorJaPago').text('R$ 0,00');
            $('#premioModalSaldoAPagar').text('R$ 0,00');
            $('#premioModalValorAPagarInput').val('');
            $('#premioModalUserId').val('');
            $('#premioModalObs').val('');
            $('#premioModalStatusPagamento').val('');
        });

        $('#premioModalRecipient').change(function() {
            const selectedType = $(this).val();
            const modal = $('#premioModal');
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
                const saldo = data.premio - data.premioPago;
                $('#premioModalValorTotal').text('R$ ' + formatCurrency(data.premio));
                $('#premioModalValorJaPago').text('R$ ' + formatCurrency(data.premioPago));
                $('#premioModalSaldoAPagar').text('R$ ' + formatCurrency(saldo));
                $('#premioModalValorAPagarInput').val(formatCurrency(saldo));
                $('#premioModalUserId').val(userId);
            } else {
                $('#premioModalValorTotal').text('R$ 0,00');
                $('#premioModalValorJaPago').text('R$ 0,00');
                $('#premioModalSaldoAPagar').text('R$ 0,00');
                $('#premioModalValorAPagarInput').val('');
                $('#premioModalUserId').val('');
            }
        });

        // ====================================================================
        // VALIDAÇÕES DOS FORMULÁRIOS
        // ====================================================================
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

        $('#premioForm').submit(function(e) {
            const valorPagoInput = $('#premioModalValorAPagarInput').val();
            const valorPago = parseCurrency(valorPagoInput);
            const saldoAPagar = parseCurrency($('#premioModalSaldoAPagar').text());

            if (valorPago <= 0) {
                alert('O valor do prêmio a pagar deve ser maior que zero.');
                e.preventDefault();
                return;
            }

            if (valorPago > saldoAPagar) {
                alert('O valor do prêmio a pagar não pode ser maior que o saldo do prêmio a pagar.');
                e.preventDefault();
                return;
            }

            if ($('#premioModalRecipient').val() === '') {
                alert('Por favor, selecione para quem será o pagamento do prêmio.');
                e.preventDefault();
                return;
            }

            if ($('#premioModalDataPagamento').val() === '') {
                alert('Por favor, selecione a data do pagamento.');
                e.preventDefault();
                return;
            }

            if ($('#premioModalStatusPagamento').val() === '') {
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

If IsObject(connSales) Then
    If connSales.State = 1 Then ' adStateOpen
        connSales.Close
    End If
    Set connSales = Nothing
End If

If IsObject(conn) Then
    If conn.State = 1 Then ' adStateOpen
        conn.Close
    End If
    Set conn = Nothing
End If
%>