<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: BBKVHDODGV          -->
<!-- OBS: Painel Bolsa de Valores - VGVs Gerências - VISÃO TABELA -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Adicionar Option Explicit no início para forçar declaração de variáveis

%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if 
%>

<%
' ===============================================
' CONFIGURAÇÕES INICIAIS
' ===============================================
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' ===============================================
' FUNÇÃO SIMPLES PARA VALOR
' ===============================================
Function GetNum(val)
    On Error Resume Next
    If IsNull(val) Or Trim(val & "") = "" Then
        GetNum = 0
    ElseIf IsNumeric(val) Then
        GetNum = CDbl(val)
    Else
        GetNum = 0
    End If
    On Error GoTo 0
End Function

' ===============================================
' FILTROS
' ===============================================
Dim anoFiltro, gerenciaFiltro
anoFiltro = Request.QueryString("ano")
gerenciaFiltro = Request.QueryString("gerencia")

If anoFiltro = "" Or Not IsNumeric(anoFiltro) Then 
    anoFiltro = Year(Date())
Else
    anoFiltro = CInt(anoFiltro)
End If

' ===============================================
' BUSCA DADOS BÁSICOS
' ===============================================
Dim sql, rs
sql = "SELECT Gerencia, MesVenda, ValorUnidade FROM Vendas "
sql = sql & "WHERE Excluido = 0 AND AnoVenda = " & anoFiltro
sql = sql & " AND Gerencia IS NOT NULL AND Gerencia <> ''"
sql = sql & " AND ValorUnidade IS NOT NULL"

If gerenciaFiltro <> "" Then
    sql = sql & " AND Gerencia = '" & Replace(gerenciaFiltro, "'", "''") & "'"
End If

sql = sql & " ORDER BY Gerencia, MesVenda"

Set rs = connSales.Execute(sql)

' ===============================================
' ARRAYS SIMPLES
' ===============================================
Dim gerenciaNomes(), gerenciaTotais(), gerenciaVariacoes()
Dim gerenciaMeses1(), gerenciaMeses2(), gerenciaMeses3()
Dim gerenciaMeses4(), gerenciaMeses5(), gerenciaMeses6()
Dim gerenciaMeses7(), gerenciaMeses8(), gerenciaMeses9()
Dim gerenciaMeses10(), gerenciaMeses11(), gerenciaMeses12()

Dim count, totalGeral, altas, baixas
count = 0
totalGeral = 0
altas = 0
baixas = 0

Dim currentGerencia, lastGerencia
Dim meses(12), totalGerencia
Dim lastVal1, lastVal2, lastMes1, lastMes2

If Not rs.EOF Then
    lastGerencia = ""
    
    Do While Not rs.EOF
        Dim gNome, gMes, gValor
        gNome = Trim(rs("Gerencia"))
        gMes = GetNum(rs("MesVenda"))
        gValor = GetNum(rs("ValorUnidade"))
        
        ' Nova gerência?
        If lastGerencia <> gNome Then
            ' Processa gerência anterior se existir
            If lastGerencia <> "" Then
                ' Calcula total
                totalGerencia = 0
                For i = 1 To 12
                    totalGerencia = totalGerencia + meses(i)
                Next
                
                ' Encontra últimos valores
                lastVal1 = 0
                lastVal2 = 0
                lastMes1 = 0
                lastMes2 = 0
                
                For i = 12 To 1 Step -1
                    If meses(i) > 0 Then
                        If lastMes1 = 0 Then
                            lastMes1 = i
                            lastVal1 = meses(i)
                        ElseIf lastMes2 = 0 Then
                            lastMes2 = i
                            lastVal2 = meses(i)
                            Exit For
                        End If
                    End If
                Next
                
                ' Calcula variação
                Dim variacao
                variacao = 0
                If lastVal2 > 0 Then
                    variacao = ((lastVal1 - lastVal2) / lastVal2) * 100
                ElseIf lastVal1 > 0 Then
                    variacao = 100
                ElseIf lastVal1 = 0 And lastVal2 > 0 Then
                    variacao = -100
                End If
                
                ' Armazena nos arrays
                ReDim Preserve gerenciaNomes(count)
                ReDim Preserve gerenciaTotais(count)
                ReDim Preserve gerenciaVariacoes(count)
                
                ReDim Preserve gerenciaMeses1(count)
                ReDim Preserve gerenciaMeses2(count)
                ReDim Preserve gerenciaMeses3(count)
                ReDim Preserve gerenciaMeses4(count)
                ReDim Preserve gerenciaMeses5(count)
                ReDim Preserve gerenciaMeses6(count)
                ReDim Preserve gerenciaMeses7(count)
                ReDim Preserve gerenciaMeses8(count)
                ReDim Preserve gerenciaMeses9(count)
                ReDim Preserve gerenciaMeses10(count)
                ReDim Preserve gerenciaMeses11(count)
                ReDim Preserve gerenciaMeses12(count)
                
                gerenciaNomes(count) = lastGerencia
                gerenciaTotais(count) = totalGerencia
                gerenciaVariacoes(count) = variacao
                
                gerenciaMeses1(count) = meses(1)
                gerenciaMeses2(count) = meses(2)
                gerenciaMeses3(count) = meses(3)
                gerenciaMeses4(count) = meses(4)
                gerenciaMeses5(count) = meses(5)
                gerenciaMeses6(count) = meses(6)
                gerenciaMeses7(count) = meses(7)
                gerenciaMeses8(count) = meses(8)
                gerenciaMeses9(count) = meses(9)
                gerenciaMeses10(count) = meses(10)
                gerenciaMeses11(count) = meses(11)
                gerenciaMeses12(count) = meses(12)
                
                totalGeral = totalGeral + totalGerencia
                
                If variacao > 0 Then altas = altas + 1
                If variacao < 0 Then baixas = baixas + 1
                
                count = count + 1
            End If
            
            ' Prepara nova gerência
            lastGerencia = gNome
            For i = 1 To 12
                meses(i) = 0
            Next
        End If
        
        ' Acumula valor
        If gMes >= 1 And gMes <= 12 Then
            meses(gMes) = meses(gMes) + gValor
        End If
        
        rs.MoveNext
    Loop
    
    ' Processa última gerência
    If lastGerencia <> "" Then
        totalGerencia = 0
        For i = 1 To 12
            totalGerencia = totalGerencia + meses(i)
        Next
        
        lastVal1 = 0
        lastVal2 = 0
        lastMes1 = 0
        lastMes2 = 0
        
        For i = 12 To 1 Step -1
            If meses(i) > 0 Then
                If lastMes1 = 0 Then
                    lastMes1 = i
                    lastVal1 = meses(i)
                ElseIf lastMes2 = 0 Then
                    lastMes2 = i
                    lastVal2 = meses(i)
                    Exit For
                End If
            End If
        Next
        
        variacao = 0
        If lastVal2 > 0 Then
            variacao = ((lastVal1 - lastVal2) / lastVal2) * 100
        ElseIf lastVal1 > 0 Then
            variacao = 100
        ElseIf lastVal1 = 0 And lastVal2 > 0 Then
            variacao = -100
        End If
        
        ReDim Preserve gerenciaNomes(count)
        ReDim Preserve gerenciaTotais(count)
        ReDim Preserve gerenciaVariacoes(count)
        
        ReDim Preserve gerenciaMeses1(count)
        ReDim Preserve gerenciaMeses2(count)
        ReDim Preserve gerenciaMeses3(count)
        ReDim Preserve gerenciaMeses4(count)
        ReDim Preserve gerenciaMeses5(count)
        ReDim Preserve gerenciaMeses6(count)
        ReDim Preserve gerenciaMeses7(count)
        ReDim Preserve gerenciaMeses8(count)
        ReDim Preserve gerenciaMeses9(count)
        ReDim Preserve gerenciaMeses10(count)
        ReDim Preserve gerenciaMeses11(count)
        ReDim Preserve gerenciaMeses12(count)
        
        gerenciaNomes(count) = lastGerencia
        gerenciaTotais(count) = totalGerencia
        gerenciaVariacoes(count) = variacao
                
        gerenciaMeses1(count) = meses(1)
        gerenciaMeses2(count) = meses(2)
        gerenciaMeses3(count) = meses(3)
        gerenciaMeses4(count) = meses(4)
        gerenciaMeses5(count) = meses(5)
        gerenciaMeses6(count) = meses(6)
        gerenciaMeses7(count) = meses(7)
        gerenciaMeses8(count) = meses(8)
        gerenciaMeses9(count) = meses(9)
        gerenciaMeses10(count) = meses(10)
        gerenciaMeses11(count) = meses(11)
        gerenciaMeses12(count) = meses(12)
        
        totalGeral = totalGeral + totalGerencia
        
        If variacao > 0 Then altas = altas + 1
        If variacao < 0 Then baixas = baixas + 1
        
        count = count + 1
    End If
End If

rs.Close
Set rs = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Painel - VGVs Gerências - Visão Tabela</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
    body { 
        background: #0a0e17; 
        color: #e0e0e0; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        padding-bottom: 20px;
    }
    
    .stock-header {
        background: linear-gradient(135deg, #1a1f2e 0%, #0a0e17 100%);
        border-bottom: 1px solid #2a3142;
        padding: 15px 0;
        margin-bottom: 20px;
    }
    
    .filter-panel {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        padding: 20px;
        margin-bottom: 25px;
    }
    
    .btn-stock {
        background: linear-gradient(135deg, #3a7bd5 0%, #00d2ff 100%);
        border: none;
        color: white;
        font-weight: 600;
        padding: 10px 25px;
        border-radius: 8px;
        transition: all 0.3s;
    }
    
    .btn-stock:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(58, 123, 213, 0.4);
        color: white;
    }
    
    .form-control-stock {
        background: #121826;
        border: 1px solid #2a3142;
        color: #e0e0e0;
        border-radius: 8px;
    }
    
    .market-status {
        background: #121826;
        border-radius: 20px;
        padding: 8px 15px;
        font-size: 0.9rem;
        border: 1px solid #2a3142;
    }
    
    .total-market {
        font-size: 1.8rem;
        font-weight: 700;
        color: #00d2ff;
        margin: 10px 0;
    }
    
    /* Estilos para a tabela */
    .stock-table {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        overflow: hidden;
        margin-bottom: 30px;
    }
    
    .table-header {
        background: #121826;
        padding: 15px 20px;
        border-bottom: 1px solid #2a3142;
    }
    
    .table-container {
        overflow-x: auto;
    }
    
    .table-dark-custom {
        background: #1a1f2e;
        color: #e0e0e0;
        margin-bottom: 0;
    }
    
    .table-dark-custom th {
        background: #121826;
        border-color: #2a3142;
        font-weight: 600;
        padding: 12px 15px;
        white-space: nowrap;
    }
    
    .table-dark-custom td {
        border-color: #2a3142;
        padding: 12px 15px;
        vertical-align: middle;
    }
    
    .table-dark-custom tbody tr:hover {
        background: rgba(58, 123, 213, 0.1);
    }
    
    .gerencia-cell {
        font-weight: 600;
        color: #ffffff;
        min-width: 200px;
        position: sticky;
        left: 0;
        background: #1a1f2e;
        z-index: 1;
    }
    
    .month-cell {
        text-align: center;
        min-width: 100px;
    }
    
    .valor-cell {
        font-weight: 600;
        font-size: 0.95rem;
    }
    
    .valor-up {
        color: #00c853;
        background: rgba(0, 200, 83, 0.1);
    }
    
    .valor-down {
        color: #ff3d00;
        background: rgba(255, 61, 0, 0.1);
    }
    
    .valor-neutral {
        color: #9e9e9e;
    }
    
    .trend-icon {
        margin-left: 5px;
        font-size: 0.8rem;
    }
    
    .total-cell {
        font-weight: 700;
        color: #00d2ff;
        background: rgba(0, 210, 255, 0.1);
        text-align: right;
        min-width: 150px;
    }
    
    .variacao-cell {
        text-align: center;
        min-width: 120px;
    }
    
    .change-up {
        color: #00c853;
        background: rgba(0, 200, 83, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
    
    .change-down {
        color: #ff3d00;
        background: rgba(255, 61, 0, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
    
    .change-neutral {
        color: #9e9e9e;
        background: rgba(158, 158, 158, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
    
    .empty-cell {
        color: #8a94a6;
        font-style: italic;
    }
    
    .table-total-row {
        background: #121826 !important;
        font-weight: 700;
        border-top: 2px solid #2a3142;
    }
    
    .table-total-row td {
        border-top: 2px solid #2a3142;
    }
    
    .simbolo-col {
        color: #3a7bd5;
        font-weight: 700;
        font-size: 1.1rem;
        width: 80px;
    }
    
    .participacao-bar {
        width: 100px;
        height: 6px;
        background: #121826;
        border-radius: 3px;
        overflow: hidden;
        display: inline-block;
        margin-right: 10px;
    }
    
    .participacao-fill {
        height: 100%;
        background: linear-gradient(90deg, #3a7bd5 0%, #00d2ff 100%);
        border-radius: 3px;
    }
    
    .participacao-cell {
        min-width: 120px;
    }
    
    /* Cards de resumo */
    .resumo-card {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        padding: 15px;
        margin-bottom: 20px;
        height: 100%;
    }
    
    .resumo-title {
        font-size: 0.9rem;
        color: #8a94a6;
        margin-bottom: 10px;
    }
    
    .resumo-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #00d2ff;
        margin-bottom: 5px;
    }
</style>
</head>
<body>

<!-- HEADER -->
<div class="stock-header">
    <div class="container">
        <div class="row align-items-center">
            <div class="col-md-6">
                <h1 class="mb-0">
                    <i class="fas fa-table text-primary me-2"></i>
                    PAINEL <span class="text-info">VGVs</span> - VISÃO TABELA
                </h1>
                <p class="text-muted mb-0">Monitoramento comparativo das gerências</p>
            </div>
            <div class="col-md-6 text-end">
                <div class="market-status d-inline-block">
                    <span class="me-2">Mercado:</span>
                    <span class="text-success">
                        <i class="fas fa-arrow-up me-1"></i>ABERTO
                    </span>
                    <span class="ms-3">
                        <i class="far fa-calendar me-1"></i><%= anoFiltro %>
                    </span>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- FILTROS -->
<div class="container">
    <div class="filter-panel">
        <form method="get" class="row g-3 align-items-end">
            <div class="col-md-3">
                <label class="form-label">Ano</label>
                <select name="ano" class="form-select" style="background:#121826;border:1px solid #2a3142;color:#e0e0e0;" onchange="this.form.submit()">
                    <% 
                    Dim rsAnos
                    Set rsAnos = connSales.Execute("SELECT DISTINCT AnoVenda FROM Vendas WHERE Excluido = 0 ORDER BY AnoVenda DESC")
                    
                    Do While Not rsAnos.EOF
                        If Not IsNull(rsAnos("AnoVenda")) Then
                            Dim anoVal
                            anoVal = CStr(rsAnos("AnoVenda"))
                            Response.Write "<option value=""" & anoVal & """"
                            If CStr(anoFiltro) = anoVal Then Response.Write " selected"
                            Response.Write ">" & anoVal & "</option>"
                        End If
                        rsAnos.MoveNext
                    Loop
                    
                    rsAnos.Close
                    Set rsAnos = Nothing
                    %>
                </select>
            </div>
            
            <div class="col-md-5">
                <label class="form-label">Gerência</label>
                <select name="gerencia" class="form-select" style="background:#121826;border:1px solid #2a3142;color:#e0e0e0;" onchange="this.form.submit()">
                    <option value="">Todas as Gerências</option>
                    <%
                    If count > 0 Then
                        For i = 0 To count - 1
                            Response.Write "<option value=""" & gerenciaNomes(i) & """"
                            If CStr(gerenciaFiltro) = CStr(gerenciaNomes(i)) Then Response.Write " selected"
                            Response.Write ">" & gerenciaNomes(i) & "</option>"
                        Next
                    End If
                    %>
                </select>
            </div>
            
            <div class="col-md-4 text-end">
                <button type="submit" class="btn btn-stock me-2">
                    <i class="fas fa-sync-alt me-2"></i>ATUALIZAR
                </button>
                <a href="?ano=<%= anoFiltro %>" class="btn btn-outline-light">LIMPAR FILTRO</a>
            </div>
        </form>
    </div>
</div>

<!-- RESUMO DO MERCADO -->
<div class="container mb-4">
    <div class="row">
        <div class="col-md-3">
            <div class="resumo-card text-center">
                <div class="resumo-title">Total Mercado</div>
                <div class="resumo-value">
                    R$ 
                    <%
                    If totalGeral > 0 Then
                        Response.Write FormatNumber(totalGeral, 2)
                    Else
                        Response.Write "0,00"
                    End If
                    %>
                </div>
                <div class="text-muted"><%= anoFiltro %></div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Gerências</div>
                <div class="resumo-value">
                    <%= count %>
                </div>
                <div class="text-muted">Ativas</div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Alta</div>
                <div class="resumo-value text-success">
                    <%= altas %>
                </div>
                <div class="text-success">
                    <i class="fas fa-arrow-up me-1"></i>Valorizando
                </div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Baixa</div>
                <div class="resumo-value text-danger">
                    <%= baixas %>
                </div>
                <div class="text-danger">
                    <i class="fas fa-arrow-down me-1"></i>Desvalorizando
                </div>
            </div>
        </div>
        
        <div class="col-md-3">
            <div class="resumo-card text-center">
                <div class="resumo-title">Média por Gerência</div>
                <div class="resumo-value">
                    R$ 
                    <%
                    If count > 0 And totalGeral > 0 Then
                        Response.Write FormatNumber(totalGeral / count, 2)
                    Else
                        Response.Write "0,00"
                    End If
                    %>
                </div>
                <div class="text-muted">Média anual</div>
            </div>
        </div>
    </div>
</div>

<!-- TABELA PRINCIPAL -->
<div class="container">
    <div class="stock-table">
        <div class="table-header">
            <h5 class="mb-0">
                <i class="fas fa-chart-line me-2"></i>
                Desempenho por Gerência e Mês - <%= anoFiltro %>
            </h5>
        </div>
        
        <div class="table-container">
            <table class="table table-dark-custom">
                <thead>
                    <tr>
                        <th class="gerencia-cell">Gerência</th>
                        <th class="month-cell">JAN</th>
                        <th class="month-cell">FEV</th>
                        <th class="month-cell">MAR</th>
                        <th class="month-cell">ABR</th>
                        <th class="month-cell">MAI</th>
                        <th class="month-cell">JUN</th>
                        <th class="month-cell">JUL</th>
                        <th class="month-cell">AGO</th>
                        <th class="month-cell">SET</th>
                        <th class="month-cell">OUT</th>
                        <th class="month-cell">NOV</th>
                        <th class="month-cell">DEZ</th>
                        <th class="total-cell">TOTAL ANUAL</th>
                        <th class="variacao-cell">VARIAÇÃO</th>
                        <th class="participacao-cell">PARTICIPAÇÃO</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If count > 0 Then
                        ' Declarar todas as variáveis temporárias usadas na ordenação
                        Dim tempNome, tempTotal, tempVar
                        Dim tempM1, tempM2, tempM3, tempM4, tempM5, tempM6
                        Dim tempM7, tempM8, tempM9, tempM10, tempM11, tempM12
                        Dim i, j, m
                        
                        ' Ordena por total usando GetNum para garantir valores numéricos
                        For i = 0 To count - 2
                            For j = i + 1 To count - 1
                                If GetNum(gerenciaTotais(i)) < GetNum(gerenciaTotais(j)) Then
                                    ' Troca nomes
                                    tempNome = gerenciaNomes(i)
                                    gerenciaNomes(i) = gerenciaNomes(j)
                                    gerenciaNomes(j) = tempNome
                                    
                                    ' Troca totais
                                    tempTotal = gerenciaTotais(i)
                                    gerenciaTotais(i) = gerenciaTotais(j)
                                    gerenciaTotais(j) = tempTotal
                                    
                                    ' Troca variações
                                    tempVar = gerenciaVariacoes(i)
                                    gerenciaVariacoes(i) = gerenciaVariacoes(j)
                                    gerenciaVariacoes(j) = tempVar
                                    
                                    ' Troca meses
                                    tempM1 = gerenciaMeses1(i)
                                    gerenciaMeses1(i) = gerenciaMeses1(j)
                                    gerenciaMeses1(j) = tempM1
                                    
                                    tempM2 = gerenciaMeses2(i)
                                    gerenciaMeses2(i) = gerenciaMeses2(j)
                                    gerenciaMeses2(j) = tempM2
                                    
                                    tempM3 = gerenciaMeses3(i)
                                    gerenciaMeses3(i) = gerenciaMeses3(j)
                                    gerenciaMeses3(j) = tempM3
                                    
                                    tempM4 = gerenciaMeses4(i)
                                    gerenciaMeses4(i) = gerenciaMeses4(j)
                                    gerenciaMeses4(j) = tempM4
                                    
                                    tempM5 = gerenciaMeses5(i)
                                    gerenciaMeses5(i) = gerenciaMeses5(j)
                                    gerenciaMeses5(j) = tempM5
                                    
                                    tempM6 = gerenciaMeses6(i)
                                    gerenciaMeses6(i) = gerenciaMeses6(j)
                                    gerenciaMeses6(j) = tempM6
                                    
                                    tempM7 = gerenciaMeses7(i)
                                    gerenciaMeses7(i) = gerenciaMeses7(j)
                                    gerenciaMeses7(j) = tempM7
                                    
                                    tempM8 = gerenciaMeses8(i)
                                    gerenciaMeses8(i) = gerenciaMeses8(j)
                                    gerenciaMeses8(j) = tempM8
                                    
                                    tempM9 = gerenciaMeses9(i)
                                    gerenciaMeses9(i) = gerenciaMeses9(j)
                                    gerenciaMeses9(j) = tempM9
                                    
                                    tempM10 = gerenciaMeses10(i)
                                    gerenciaMeses10(i) = gerenciaMeses10(j)
                                    gerenciaMeses10(j) = tempM10
                                    
                                    tempM11 = gerenciaMeses11(i)
                                    gerenciaMeses11(i) = gerenciaMeses11(j)
                                    gerenciaMeses11(j) = tempM11
                                    
                                    tempM12 = gerenciaMeses12(i)
                                    gerenciaMeses12(i) = gerenciaMeses12(j)
                                    gerenciaMeses12(j) = tempM12
                                End If
                            Next
                        Next
                        
                        ' Exibe cada gerência
                        For i = 0 To count - 1
                            'Dim gNome, gTotal, gVariacao
                            Dim changeClass, changeIcon, simbolo, palavras, palavra
                            Dim participacao, totalFormatado, variacaoFormatada, participacaoFormatada
                            
                            gNome = gerenciaNomes(i)
                            gTotal = gerenciaTotais(i)
                            gVariacao = gerenciaVariacoes(i)
                            
                            ' Determina estilo da variação
                            If GetNum(gVariacao) > 0 Then
                                changeClass = "change-up"
                                changeIcon = "fas fa-arrow-up"
                            ElseIf GetNum(gVariacao) < 0 Then
                                changeClass = "change-down"
                                changeIcon = "fas fa-arrow-down"
                            Else
                                changeClass = "change-neutral"
                                changeIcon = "fas fa-minus"
                            End If
                            
                            ' Símbolo
                            simbolo = ""
                            palavras = Split(gNome, " ")
                            For Each palavra In palavras
                                If Len(palavra) > 0 Then
                                    simbolo = simbolo & UCase(Left(palavra, 1))
                                End If
                            Next
                            If Len(simbolo) > 4 Then simbolo = Left(simbolo, 4)
                            
                            ' Participação
                            If GetNum(totalGeral) > 0 Then
                                participacao = (GetNum(gTotal) / GetNum(totalGeral)) * 100
                            Else
                                participacao = 0
                            End If
                            
                            ' Formata números
                            totalFormatado = FormatNumber(GetNum(gTotal), 2)
                            variacaoFormatada = FormatNumber(Abs(GetNum(gVariacao)), 2)
                            participacaoFormatada = FormatNumber(GetNum(participacao), 1)
                            
' Array de valores mensais
Dim mesesValores

' *** AÇÃO NECESSÁRIA: Redimensionar o array! ***
' Usamos ReDim para definir o tamanho do array de 1 a 12
ReDim mesesValores(12) 

mesesValores(1) = GetNum(gerenciaMeses1(i))
mesesValores(2) = GetNum(gerenciaMeses2(i))
mesesValores(3) = GetNum(gerenciaMeses3(i))
mesesValores(4) = GetNum(gerenciaMeses4(i))
mesesValores(5) = GetNum(gerenciaMeses5(i))
mesesValores(6) = GetNum(gerenciaMeses6(i))
mesesValores(7) = GetNum(gerenciaMeses7(i))
mesesValores(8) = GetNum(gerenciaMeses8(i))
mesesValores(9) = GetNum(gerenciaMeses9(i))
mesesValores(10) = GetNum(gerenciaMeses10(i))
mesesValores(11) = GetNum(gerenciaMeses11(i))
mesesValores(12) = GetNum(gerenciaMeses12(i))
                    %>
                    <tr>
                        <td class="gerencia-cell">
                            <div class="simbolo-col d-inline-block me-2"><%= simbolo %></div>
                            <%= Left(gNome, 25) %>
                        </td>
                        
                        <%
                        ' Exibe valores mensais com cores e setas
                        For m = 1 To 12
                            Dim valorMes, valorAnterior, valorClass, valorExibir
                            Dim trendIcon, variacaoMes
                            
                            valorMes = mesesValores(m)
                            
                            ' Calcula variação em relação ao mês anterior
                            valorAnterior = 0
                            If m > 1 Then
                                valorAnterior = mesesValores(m-1)
                            End If
                            
                            ' Determina classe de cor e ícone
                            If valorMes = 0 Then
                                valorClass = "valor-neutral"
                                trendIcon = ""
                            Else
                                If m = 1 Then
                                    ' Primeiro mês: apenas valor positivo
                                    valorClass = "valor-up"
                                    trendIcon = ""
                                Else
                                    ' Compara com mês anterior
                                    If valorAnterior = 0 Then
                                        ' Cresceu de zero para algum valor
                                        valorClass = "valor-up"
                                        trendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf valorMes > valorAnterior Then
                                        ' Cresceu em relação ao anterior
                                        valorClass = "valor-up"
                                        trendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf valorMes < valorAnterior Then
                                        ' Caiu em relação ao anterior
                                        valorClass = "valor-down"
                                        trendIcon = "<i class='fas fa-arrow-down trend-icon'></i>"
                                    Else
                                        ' Igual ao anterior
                                        valorClass = "valor-neutral"
                                        trendIcon = "<i class='fas fa-minus trend-icon'></i>"
                                    End If
                                End If
                            End If
                            
                            ' Formata valor para exibição
                            If valorMes > 0 Then
                                valorExibir = FormatNumber(valorMes / 1000, 0) & "k"
                            Else
                                valorExibir = "-"
                                valorClass = "empty-cell"
                            End If
                        %>
                        <td class="month-cell valor-cell <%= valorClass %>">
                            <%= valorExibir %>
                            <% If valorMes > 0 And trendIcon <> "" Then %>
                            <%= trendIcon %>
                            <% End If %>
                        </td>
                        <% Next %>
                        
                        <td class="total-cell">
                            R$ <%= totalFormatado %>
                        </td>
                        
                        <td class="variacao-cell">
                            <span class="<%= changeClass %>">
                                <i class="<%= changeIcon %> me-1"></i>
                                <%= variacaoFormatada %>%
                            </span>
                        </td>
                        
                        <td class="participacao-cell">
                            <div class="d-flex align-items-center">
                                <div class="participacao-bar">
                                    <div class="participacao-fill" style="width: <%= participacaoFormatada %>%"></div>
                                </div>
                                <div>
                                    <%= participacaoFormatada %>%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <% 
                        Next 
                    %>
                    
                    <!-- LINHA DE TOTAIS -->
                    <tr class="table-total-row">
                        <td class="gerencia-cell">
                            <strong>TOTAL MERCADO</strong>
                        </td>
                        <%
                        ' Calcula totais mensais
                        Dim totalMensal(12)
                        For m = 1 To 12
                            totalMensal(m) = 0
                        Next
                        
                        If count > 0 Then
                            For i = 0 To count - 1
                                totalMensal(1) = totalMensal(1) + GetNum(gerenciaMeses1(i))
                                totalMensal(2) = totalMensal(2) + GetNum(gerenciaMeses2(i))
                                totalMensal(3) = totalMensal(3) + GetNum(gerenciaMeses3(i))
                                totalMensal(4) = totalMensal(4) + GetNum(gerenciaMeses4(i))
                                totalMensal(5) = totalMensal(5) + GetNum(gerenciaMeses5(i))
                                totalMensal(6) = totalMensal(6) + GetNum(gerenciaMeses6(i))
                                totalMensal(7) = totalMensal(7) + GetNum(gerenciaMeses7(i))
                                totalMensal(8) = totalMensal(8) + GetNum(gerenciaMeses8(i))
                                totalMensal(9) = totalMensal(9) + GetNum(gerenciaMeses9(i))
                                totalMensal(10) = totalMensal(10) + GetNum(gerenciaMeses10(i))
                                totalMensal(11) = totalMensal(11) + GetNum(gerenciaMeses11(i))
                                totalMensal(12) = totalMensal(12) + GetNum(gerenciaMeses12(i))
                            Next
                        End If
                        
                        For m = 1 To 12
                            Dim totalMesFormatado, totalMesClass, totalMesTrendIcon
                            Dim totalMesAnterior
                            
                            ' Determina tendência do total mensal
                            totalMesAnterior = 0
                            If m > 1 Then
                                totalMesAnterior = totalMensal(m-1)
                            End If
                            
                            If totalMensal(m) > 0 Then
                                totalMesFormatado = FormatNumber(totalMensal(m) / 1000, 0) & "k"
                                
                                If m = 1 Then
                                    totalMesClass = "valor-up"
                                    totalMesTrendIcon = ""
                                Else
                                    If totalMesAnterior = 0 Then
                                        totalMesClass = "valor-up"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf totalMensal(m) > totalMesAnterior Then
                                        totalMesClass = "valor-up"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf totalMensal(m) < totalMesAnterior Then
                                        totalMesClass = "valor-down"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-down trend-icon'></i>"
                                    Else
                                        totalMesClass = "valor-neutral"
                                        totalMesTrendIcon = "<i class='fas fa-minus trend-icon'></i>"
                                    End If
                                End If
                            Else
                                totalMesFormatado = "-"
                                totalMesClass = "empty-cell"
                                totalMesTrendIcon = ""
                            End If
                        %>
                        <td class="month-cell valor-cell <%= totalMesClass %>">
                            <strong><%= totalMesFormatado %></strong>
                            <% If totalMensal(m) > 0 And totalMesTrendIcon <> "" Then %>
                            <%= totalMesTrendIcon %>
                            <% End If %>
                        </td>
                        <% Next %>
                        
                        <td class="total-cell">
                            <strong>R$ <%= FormatNumber(totalGeral, 2) %></strong>
                        </td>
                        
                        <td class="variacao-cell">
                            <span class="text-muted">-</span>
                        </td>
                        
                        <td class="participacao-cell">
                            <div class="d-flex align-items-center">
                                <div class="participacao-bar">
                                    <div class="participacao-fill" style="width: 100%"></div>
                                </div>
                                <div>
                                    100%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <% Else %>
                    <tr>
                        <td colspan="15" class="text-center py-5">
                            <i class="fas fa-chart-bar fa-3x text-muted mb-3"></i>
                            <h3 class="text-muted">Nenhuma gerência encontrada</h3>
                            <p class="text-muted">Tente ajustar os filtros ou verificar os dados do ano <%= anoFiltro %></p>
                        </td>
                    </tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- LEGENDA -->
<div class="container mt-3 mb-4">
    <div class="row">
        <div class="col-md-6">
            <div class="resumo-card">
                <h6 class="mb-3"><i class="fas fa-info-circle me-2"></i>Legenda</h6>
                <div class="row">
                    <div class="col-md-6">
                        <div class="mb-2">
                            <span class="valor-up me-2"><i class="fas fa-arrow-up"></i></span>
                            <span class="text-muted">Crescimento em relação ao mês anterior</span>
                        </div>
                        <div class="mb-2">
                            <span class="valor-down me-2"><i class="fas fa-arrow-down"></i></span>
                            <span class="text-muted">Queda em relação ao mês anterior</span>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="mb-2">
                            <span class="change-up me-2"><i class="fas fa-arrow-up"></i> 25%</span>
                            <span class="text-muted">Variação positiva anual</span>
                        </div>
                        <div class="mb-2">
                            <span class="change-down me-2"><i class="fas fa-arrow-down"></i> 25%</span>
                            <span class="text-muted">Variação negativa anual</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="resumo-card">
                <h6 class="mb-3"><i class="fas fa-chart-pie me-2"></i>Interpretação</h6>
                <ul class="text-muted small mb-0">
                    <li><strong>Verde com seta ↑</strong>: Mês melhor que o anterior</li>
                    <li><strong>Vermelho com seta ↓</strong>: Mês pior que o anterior</li>
                    <li><strong>Total Anual</strong>: Soma de todos os meses</li>
                    <li><strong>Variação</strong>: Comparação dos últimos 2 meses com dados</li>
                    <li><strong>Participação</strong>: Percentual do total do mercado</li>
                    <li><strong>"k"</strong>: Representa × 1.000 (ex: 1.5k = 1.500)</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!-- RODAPÉ -->
<div class="container mt-4">
    <div class="text-center text-muted">
        <p>
            <i class="fas fa-info-circle me-1"></i>
            Painel Bolsa de VGVs - Visão Tabela - Dados referentes ao ano <%= anoFiltro %> 
            | Atualizado em <%= Now() %> 
            | Total de gerências: <%= count %>
        </p>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>

<%
' Fecha conexões
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
If Not connSales Is Nothing Then
    If connSales.State = 1 Then connSales.Close
    Set connSales = Nothing
End If
%>