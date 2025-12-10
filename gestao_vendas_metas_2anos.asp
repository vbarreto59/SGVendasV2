<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 04/12/2025                       -->
<!-- CODIGO_ARQUIVO: YABLTUXLJI             -->
<!-- OBS: Relatório modificado para comparar 2 anos mais recentes -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% ' Funcionalidade de Cores Corrigida e Melhorada com Data Labels '
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<!--#include file="usr_acoes_v4GVendas.inc"-->

<!--#include file="gestao_header.inc"-->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if 
%>
<!--#include file="usr_acoes_v4GVendas.inc"-->


<%
' ===========================================================
' LOG de acesso (mantido)
' ===========================================================
    if Not BloqueioEmail() AND (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
        On Error Resume Next 
        set objMail = server.createobject("CDONTS.NewMail")
        if Err.Number <> 0 then 
            set objMail = Nothing
        else
            objMail.From = "sendmail@gabnetweb.com.br"
            objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
            objMail.Subject = "SV-REL2ANOS-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
            objMail.MailFormat = 0 ' 0 = Texto Simples
            objMail.Body = "Página Relatório com VGV de 2 anos. " & Ucase(Session("Usuario"))
            objMail.Send
            set objMail = Nothing
        end if 
        On Error GoTo 0 
    end if
%>

<%
' ===========================================================
' Função substituta para Nz() do Access (mantida)
' ===========================================================
Function Nz(valor, opcional)
    If IsNull(valor) Or IsEmpty(valor) Or valor = "" Then
        If IsMissing(opcional) Then
            Nz = 0
        Else
            Nz = opcional
        End If
    Else
        Nz = valor
    End If
End Function
%>


<%
' ==============================================================================
' INCLUSÕES E CONFIGURAÇÕES
' ==============================================================================

Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' ==============================================================================
' DETERMINAR OS 2 ANOS MAIS RECENTES COM VENDAS (do menor para o maior)
' ==============================================================================
Dim anoAntigo, anoRecente
Dim rsAnos, sqlAnos

' Buscar todos os anos distintos com vendas em ordem crescente
sqlAnos = "SELECT DISTINCT AnoVenda FROM Vendas WHERE (Excluido <> -1 OR Excluido IS NULL) ORDER BY AnoVenda"
Set rsAnos = Server.CreateObject("ADODB.Recordset")
rsAnos.CursorType = 0 ' adOpenForwardOnly - para evitar problemas com MoveLast
rsAnos.Open sqlAnos, connSales

Dim anosArray()
ReDim anosArray(0) ' Array dinâmico
Dim anosCount
anosCount = 0

' Coletar todos os anos em um array
If Not rsAnos.EOF Then
    Do While Not rsAnos.EOF
        ReDim Preserve anosArray(anosCount)
        anosArray(anosCount) = rsAnos("AnoVenda")
        anosCount = anosCount + 1
        rsAnos.MoveNext
    Loop
End If
rsAnos.Close
Set rsAnos = Nothing

' Determinar os anos para comparação
If anosCount >= 2 Then
    ' Temos pelo menos 2 anos
    ' Pegar os 2 últimos do array (os mais recentes)
    anoAntigo = anosArray(anosCount - 2) ' Penúltimo
    anoRecente = anosArray(anosCount - 1) ' Último
ElseIf anosCount = 1 Then
    ' Só tem um ano
    anoRecente = anosArray(0)
    anoAntigo = anoRecente - 1
Else
    ' Não tem nenhum ano
    anoRecente = Year(Date())
    anoAntigo = anoRecente - 1
End If

' Array com nomes dos meses
Dim meses(12)
meses(1) = "Jan"
meses(2) = "Fev"
meses(3) = "Mar"
meses(4) = "Abr"
meses(5) = "Mai"
meses(6) = "Jun"
meses(7) = "Jul"
meses(8) = "Ago"
meses(9) = "Set"
meses(10) = "Out"
meses(11) = "Nov"
meses(12) = "Dez"

' Arrays para armazenar dados dos 2 anos
Dim vendasAntigo(12), vendasRecente(12), metasRecente(12)
Dim diferencasAnuais(12), variacaoPercentual(12)

' Inicializar arrays (Jan a Dez = 1 a 12)
For i = 1 To 12
    vendasAntigo(i) = 0
    vendasRecente(i) = 0
    metasRecente(i) = 0
    diferencasAnuais(i) = 0
    variacaoPercentual(i) = 0
Next

' ==============================================================================
' BUSCAR VENDAS DO ANO ANTIGO (MAIS ANTIGO DOS 2)
' ==============================================================================
Set rsVendasAntigo = Server.CreateObject("ADODB.Recordset")
sqlVendasAntigo = "SELECT MesVenda, SUM(ValorUnidade) as TotalVendas " & _
                "FROM Vendas " & _
                "WHERE AnoVenda = " & anoAntigo & " AND (Excluido <> -1 OR Excluido IS NULL) " & _
                "GROUP BY MesVenda " & _
                "ORDER BY MesVenda"

rsVendasAntigo.Open sqlVendasAntigo, connSales

If Not rsVendasAntigo.EOF Then
    Do While Not rsVendasAntigo.EOF
        mes = CInt(rsVendasAntigo("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            If Not IsNull(rsVendasAntigo("TotalVendas")) Then
                vendasAntigo(mes) = CDbl(rsVendasAntigo("TotalVendas"))
            Else
                vendasAntigo(mes) = 0
            End If
        End If
        rsVendasAntigo.MoveNext
    Loop
End If
rsVendasAntigo.Close
Set rsVendasAntigo = Nothing

' ==============================================================================
' BUSCAR VENDAS DO ANO RECENTE (MAIS RECENTE DOS 2)
' ==============================================================================
Set rsVendasRecente = Server.CreateObject("ADODB.Recordset")
sqlVendasRecente = "SELECT MesVenda, SUM(ValorUnidade) as TotalVendas " & _
                   "FROM Vendas " & _
                   "WHERE AnoVenda = " & anoRecente & " AND (Excluido <> -1 OR Excluido IS NULL) " & _
                   "GROUP BY MesVenda " & _
                   "ORDER BY MesVenda"

rsVendasRecente.Open sqlVendasRecente, connSales

If Not rsVendasRecente.EOF Then
    Do While Not rsVendasRecente.EOF
        mes = CInt(rsVendasRecente("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            If Not IsNull(rsVendasRecente("TotalVendas")) Then
                vendasRecente(mes) = CDbl(rsVendasRecente("TotalVendas"))
            Else
                vendasRecente(mes) = 0
            End If
        End If
        rsVendasRecente.MoveNext
    Loop
End If
rsVendasRecente.Close
Set rsVendasRecente = Nothing

' ==============================================================================
' CALCULAR DIFERENÇAS E VARIAÇÕES ENTRE ANOS (RECENTE - ANTIGO)
' ==============================================================================
For i = 1 To 12
    diferencasAnuais(i) = vendasRecente(i) - vendasAntigo(i)
    If vendasAntigo(i) > 0 Then
        variacaoPercentual(i) = ((vendasRecente(i) - vendasAntigo(i)) / vendasAntigo(i)) * 100
    Else
        If vendasRecente(i) > 0 Then
            variacaoPercentual(i) = 100 ' Crescimento infinito (de 0 para algum valor)
        Else
            variacaoPercentual(i) = 0 ' Ambos são zero
        End If
    End If
Next

' ==============================================================================
' BUSCAR METAS DO ANO RECENTE
' ==============================================================================
Set rsMetasRecente = Server.CreateObject("ADODB.Recordset")
sqlMetasRecente = "SELECT Mes, Meta FROM MetaEmpresa WHERE Ano = " & anoRecente & " ORDER BY Mes"
rsMetasRecente.Open sqlMetasRecente, connSales

If Not rsMetasRecente.EOF Then
    Do While Not rsMetasRecente.EOF
        mes = CInt(rsMetasRecente("Mes"))
        If mes >= 1 And mes <= 12 Then
            metasRecente(mes) = CDbl(Nz(rsMetasRecente("Meta"), 0))
        End If
        rsMetasRecente.MoveNext
    Loop
End If
rsMetasRecente.Close
Set rsMetasRecente = Nothing

' ==============================================================================
' CÁLCULOS GERAIS PARA AMBOS OS ANOS
' ==============================================================================
Dim totalVendasAntigo, totalVendasRecente, totalMetaRecente
Dim totalUnidadesAntigo, totalUnidadesRecente
Dim ticketMedioAntigo, ticketMedioRecente
Dim diferencaTotalAnual, variacaoTotalPercentual

totalVendasAntigo = 0
totalVendasRecente = 0
totalMetaRecente = 0

' Somar totais
For i = 1 To 12
    totalVendasAntigo = totalVendasAntigo + vendasAntigo(i)
    totalVendasRecente = totalVendasRecente + vendasRecente(i)
    totalMetaRecente = totalMetaRecente + metasRecente(i)
Next

diferencaTotalAnual = totalVendasRecente - totalVendasAntigo
If totalVendasAntigo > 0 Then
    variacaoTotalPercentual = (diferencaTotalAnual / totalVendasAntigo) * 100
Else
    If totalVendasRecente > 0 Then
        variacaoTotalPercentual = 100
    Else
        variacaoTotalPercentual = 0
    End If
End If

' Buscar quantidade total de unidades vendidas
Set rsUnidadesAntigo = Server.CreateObject("ADODB.Recordset")
sqlUnidadesAntigo = "SELECT COUNT(*) as TotalUnidades FROM Vendas WHERE AnoVenda = " & anoAntigo & " AND (Excluido <> -1 OR Excluido IS NULL)"
rsUnidadesAntigo.Open sqlUnidadesAntigo, connSales
If Not rsUnidadesAntigo.EOF Then
    totalUnidadesAntigo = rsUnidadesAntigo("TotalUnidades")
Else
    totalUnidadesAntigo = 0
End If
rsUnidadesAntigo.Close
Set rsUnidadesAntigo = Nothing

Set rsUnidadesRecente = Server.CreateObject("ADODB.Recordset")
sqlUnidadesRecente = "SELECT COUNT(*) as TotalUnidades FROM Vendas WHERE AnoVenda = " & anoRecente & " AND (Excluido <> -1 OR Excluido IS NULL)"
rsUnidadesRecente.Open sqlUnidadesRecente, connSales
If Not rsUnidadesRecente.EOF Then
    totalUnidadesRecente = rsUnidadesRecente("TotalUnidades")
Else
    totalUnidadesRecente = 0
End If
rsUnidadesRecente.Close
Set rsUnidadesRecente = Nothing

' Calcular Ticket Médio
If totalUnidadesAntigo > 0 Then
    ticketMedioAntigo = totalVendasAntigo / totalUnidadesAntigo
Else
    ticketMedioAntigo = 0
End If

If totalUnidadesRecente > 0 Then
    ticketMedioRecente = totalVendasRecente / totalUnidadesRecente
Else
    ticketMedioRecente = 0
End If

' ==============================================================================
' BUSCAR ÚLTIMAS VENDAS (APENAS DO ANO RECENTE)
' ==============================================================================
Set rsUltimasVendas = Server.CreateObject("ADODB.Recordset")
sqlUltimasVendas = "SELECT TOP 3 V.ID, V.NomeEmpreendimento, V.Unidade, V.ValorUnidade, V.DataVenda, V.Corretor, V.Localidade, V.MesVenda, V.AnoVenda, V.ComissaoPercentual, V.ValorComissaoGeral, V.Diretoria, V.Gerencia " & _
                   "FROM Vendas V " & _
                   "WHERE V.AnoVenda = " & anoRecente & " AND (V.Excluido <> -1 OR V.Excluido IS NULL) " & _
                   "ORDER BY V.DataVenda DESC, V.ID DESC"
rsUltimasVendas.Open sqlUltimasVendas, connSales

' ==============================================================================
' PREPARAR DADOS PARA O JAVASCRIPT
' ==============================================================================
Dim strVendasAntigoJS, strVendasRecenteJS, strVariacaoPercentualJS

strVendasAntigoJS = ""
strVendasRecenteJS = ""
strVariacaoPercentualJS = ""

For i = 1 To 12
    ' Formatar números com ponto decimal para JavaScript
    If i > 1 Then
        strVendasAntigoJS = strVendasAntigoJS & ","
        strVendasRecenteJS = strVendasRecenteJS & ","
        strVariacaoPercentualJS = strVariacaoPercentualJS & ","
    End If
    
    strVendasAntigoJS = strVendasAntigoJS & Replace(FormatNumber(vendasAntigo(i), 2, , , False), ",", ".")
    strVendasRecenteJS = strVendasRecenteJS & Replace(FormatNumber(vendasRecente(i), 2, , , False), ",", ".")
    strVariacaoPercentualJS = strVariacaoPercentualJS & Replace(FormatNumber(variacaoPercentual(i), 1, , , False), ",", ".")
Next
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Comparativo Anual | Gestão de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --accent: #e74c3c;
            --success: #28a745;
            --warning: #fd7e14;
            --light-bg: #f8f9fa;
        }
        
        body {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding-top: 80px;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1rem 0;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1rem 1.5rem;
            font-weight: 600;
        }
        
        .mes-card {
            transition: transform 0.2s;
            height: 100%;
        }
        
        .mes-card:hover {
            transform: translateY(-2px);
        }
        
        .filter-section {
            background: white;
            border-radius: 12px;
            padding: 1rem;
            margin-bottom: 1.5rem;
        }
        
        .ultimas-vendas {
            background: white;
            padding: 1.5rem;
            margin-top: 2rem;
        }
        
        .venda-item {
            border-left: 4px solid var(--secondary);
            padding: 1rem;
            margin-bottom: 1rem;
            background: #f8f9fa;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
        }
        
        .btn-refresh {
            background-color: var(--warning);
            border-color: var(--warning);
            color: white;
        }
        
        .chart-container {
            position: relative;
            height: 400px;
            width: 100%;
        }
        
        /* CORES PARA CARDS DE COMPARAÇÃO */
        .mes-card.crescimento {
            background-color: #e8f5e8;
            border-left: 4px solid #28a745;
        }
        
        .mes-card.declinio {
            background-color: #ffebee;
            border-left: 4px solid #dc3545;
        }
        
        .mes-card.estavel {
            background-color: #f8f9fa;
            border-left: 4px solid #6c757d;
        }
        
        .mes-card.sem-dados {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
        }
        
        /* Cores para indicadores de variação */
        .variacao-positiva {
            color: #28a745;
            font-weight: bold;
        }
        
        .variacao-negativa {
            color: #dc3545;
            font-weight: bold;
        }
        
        .variacao-neutra {
            color: #6c757d;
            font-weight: bold;
        }
        
        /* Estilo para cards de comparação */
        .comparativo-card {
            text-align: center;
            padding: 1rem;
        }
        
        .comparativo-valor {
            font-size: 1.2rem;
            font-weight: bold;
            margin: 0.5rem 0;
        }
        
        .comparativo-variacao {
            font-size: 0.9rem;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            display: inline-block;
        }
        
        .ano-antigo {
            color: #95a5a6;
        }
        
        .ano-recente {
            color: #3498db;
        }

        .legenda-grafico {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 15px;
            flex-wrap: wrap;
        }
        
        .item-legenda {
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 0.9rem;
        }
        
        .cor-legenda {
            width: 15px;
            height: 15px;
            border-radius: 3px;
        }
        
        .ano-label {
            font-size: 0.8rem;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <header class="app-header">
        <div class="container-fluid">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <h1 class="app-title"><i class="fas fa-chart-line me-2"></i> Comparativo Anual</h1>
                    <small><%=Session("Usuario")%></small>
                </div>
                <div class="col-md-6 text-end">
                    <a href="javascript:window.close()" class="btn btn-light btn-sm">
                        <i class="fas fa-times me-1"></i>Fechar
                    </a>
                </div>
            </div>
        </div>
    </header>


    <div class="container-fluid main-content mt-5">
        <!-- Cabeçalho de comparação -->
        <div class="row mb-4">
            <div class="col-md-6">
                <div class="card bg-secondary text-white">
                    <div class="card-body text-center">
                        <h5 class="card-title">ANO ANTIGO</h5>
                        <h2><%= anoAntigo %></h2>
                        <div class="comparativo-valor">R$ <%= FormatNumber(totalVendasAntigo, 2) %></div>
                        <div class="ano-label">(Mais Antigo)</div>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card bg-primary text-white">
                    <div class="card-body text-center">
                        <h5 class="card-title">ANO RECENTE</h5>
                        <h2><%= anoRecente %></h2>
                        <div class="comparativo-valor">R$ <%= FormatNumber(totalVendasRecente, 2) %></div>
                        <div class="ano-label">(Mais Recente)</div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Gráfico comparativo -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Comparativo de Vendas <%= anoAntigo %> vs <%= anoRecente %></h5>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="graficoComparativo"></canvas>
                </div>
                <div class="legenda-grafico">
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #95a5a6;"></div>
                        <span><%= anoAntigo %> (Antigo)</span>
                    </div>
                    <div class="item-legenda">
                        <div class="cor-legenda" style="background-color: #3498db;"></div>
                        <span><%= anoRecente %> (Recente)</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Cards de totais comparativos -->
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="card comparativo-card">
                    <div class="card-body">
                        <h6 class="card-title">Variação Total</h6>
                        <div class="comparativo-valor">
                            R$ <%= FormatNumber(diferencaTotalAnual, 2) %>
                        </div>
                        <%
                        Dim classeVariacaoTotal
                        If diferencaTotalAnual > 0 Then
                            classeVariacaoTotal = "variacao-positiva"
                            Response.Write "<span class='comparativo-variacao " & classeVariacaoTotal & "'><i class='fas fa-arrow-up me-1'></i>" & FormatNumber(variacaoTotalPercentual, 1) & "%</span>"
                        Else 
                            If diferencaTotalAnual < 0 Then
                                classeVariacaoTotal = "variacao-negativa"
                                Response.Write "<span class='comparativo-variacao " & classeVariacaoTotal & "'><i class='fas fa-arrow-down me-1'></i>" & FormatNumber(variacaoTotalPercentual, 1) & "%</span>"
                            Else
                                classeVariacaoTotal = "variacao-neutra"
                                Response.Write "<span class='comparativo-variacao " & classeVariacaoTotal & "'>0.0%</span>"
                            End If
                        End If
                        %>
                        <div class="mt-2">
                            <small class="text-muted">De <%= anoAntigo %> para <%= anoRecente %></small>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card comparativo-card">
                    <div class="card-body">
                        <h6 class="card-title">Unidades Vendidas</h6>
                        <div class="comparativo-valor ano-recente">
                            <%= totalUnidadesRecente %>
                        </div>
                        <div class="text-muted small"><%= anoAntigo %>: <%= totalUnidadesAntigo %></div>
                        <%
                        Dim difUnidades, percUnidades
                        difUnidades = totalUnidadesRecente - totalUnidadesAntigo
                        If totalUnidadesAntigo > 0 Then
                            percUnidades = (difUnidades / totalUnidadesAntigo) * 100
                        Else
                            If totalUnidadesRecente > 0 Then
                                percUnidades = 100
                            Else
                                percUnidades = 0
                            End If
                        End If
                        
                        If difUnidades > 0 Then
                            Response.Write "<span class='comparativo-variacao variacao-positiva'><i class='fas fa-arrow-up me-1'></i>" & difUnidades & " (" & FormatNumber(percUnidades, 1) & "%)</span>"
                        Else 
                            If difUnidades < 0 Then
                                Response.Write "<span class='comparativo-variacao variacao-negativa'><i class='fas fa-arrow-down me-1'></i>" & difUnidades & " (" & FormatNumber(percUnidades, 1) & "%)</span>"
                            Else
                                Response.Write "<span class='comparativo-variacao variacao-neutra'>0 (0.0%)</span>"
                            End If
                        End If
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card comparativo-card">
                    <div class="card-body">
                        <h6 class="card-title">Ticket Médio</h6>
                        <div class="comparativo-valor ano-recente">
                            R$ <%= FormatNumber(ticketMedioRecente, 2) %>
                        </div>
                        <div class="text-muted small"><%= anoAntigo %>: R$ <%= FormatNumber(ticketMedioAntigo, 2) %></div>
                        <%
                        Dim difTicket, percTicket
                        difTicket = ticketMedioRecente - ticketMedioAntigo
                        If ticketMedioAntigo > 0 Then
                            percTicket = (difTicket / ticketMedioAntigo) * 100
                        Else
                            If ticketMedioRecente > 0 Then
                                percTicket = 100
                            Else
                                percTicket = 0
                            End If
                        End If
                        
                        If difTicket > 0 Then
                            Response.Write "<span class='comparativo-variacao variacao-positiva'><i class='fas fa-arrow-up me-1'></i>" & FormatNumber(difTicket, 2) & " (" & FormatNumber(percTicket, 1) & "%)</span>"
                        Else 
                            If difTicket < 0 Then
                                Response.Write "<span class='comparativo-variacao variacao-negativa'><i class='fas fa-arrow-down me-1'></i>" & FormatNumber(difTicket, 2) & " (" & FormatNumber(percTicket, 1) & "%)</span>"
                            Else
                                Response.Write "<span class='comparativo-variacao variacao-neutra'>0.00 (0.0%)</span>"
                            End If
                        End If
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card comparativo-card">
                    <div class="card-body">
                        <h6 class="card-title">Meta do Ano Recente</h6>
                        <div class="comparativo-valor ano-recente">
                            R$ <%= FormatNumber(totalMetaRecente, 2) %>
                        </div>
                        <div class="text-muted small">Meta: <%= FormatNumber(totalMetaRecente, 2) %></div>
                        <%
                        Dim atingimentoMeta
                        Dim classeAtingimentoMeta
                        If totalMetaRecente > 0 Then
                            atingimentoMeta = (totalVendasRecente / totalMetaRecente) * 100
                        Else
                            atingimentoMeta = 0
                        End If
                        
                        If atingimentoMeta >= 100 Then
                            classeAtingimentoMeta = "variacao-positiva"
                        Else
                            classeAtingimentoMeta = "variacao-negativa"
                        End If
                        %>
                        <span class="comparativo-variacao <%= classeAtingimentoMeta %>">
                            <%= FormatNumber(atingimentoMeta, 1) %>% da meta
                        </span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Detalhamento mensal comparativo -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-calendar-alt me-2"></i>Comparativo Mensal - <%= anoAntigo %> vs <%= anoRecente %></h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <%
                    For i = 1 To 12
                        Dim classeCard, iconeVariacao, textoVariacao, classeVariacao
                        Dim percVar
                        
                        ' Determinar classe do card baseada na variação (Recente - Antigo)
                        If variacaoPercentual(i) > 0 Then
                            classeCard = "crescimento"
                            iconeVariacao = "fa-arrow-up"
                            textoVariacao = "+" & FormatNumber(variacaoPercentual(i), 1) & "%"
                            classeVariacao = "variacao-positiva"
                        Else 
                            If variacaoPercentual(i) < 0 Then
                                classeCard = "declinio"
                                iconeVariacao = "fa-arrow-down"
                                textoVariacao = FormatNumber(variacaoPercentual(i), 1) & "%"
                                classeVariacao = "variacao-negativa"
                            Else
                                If vendasRecente(i) = 0 And vendasAntigo(i) = 0 Then
                                    classeCard = "sem-dados"
                                    iconeVariacao = "fa-minus"
                                    textoVariacao = "Sem dados"
                                    classeVariacao = "variacao-neutra"
                                Else
                                    classeCard = "estavel"
                                    iconeVariacao = "fa-equals"
                                    textoVariacao = "0.0%"
                                    classeVariacao = "variacao-neutra"
                                End If
                            End If
                        End If
                    %>
                    <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6 mb-3">
                        <div class="card mes-card <%= classeCard %>">
                            <div class="card-body text-center p-2">
                                <h6 class="card-title fw-bold"><%= meses(i) %></h6>
                                
                                <!-- Ano Antigo (primeiro) -->
                                <div class="mb-1">
                                    <small class="text-muted"><%= anoAntigo %>:</small>
                                    <div class="fw-bold ano-antigo">R$ <%= FormatNumber(vendasAntigo(i), 2) %></div>
                                </div>
                                
                                <!-- Ano Recente (segundo) -->
                                <div class="mb-2">
                                    <small class="text-muted"><%= anoRecente %>:</small>
                                    <div class="ano-recente">R$ <%= FormatNumber(vendasRecente(i), 2) %></div>
                                </div>
                                
                                <!-- Meta do ano recente (se houver) -->
                                <% If metasRecente(i) > 0 Then %>
                                <div class="mb-1">
                                    <small class="text-muted">Meta <%= anoRecente %>:</small>
                                    <div>R$ <%= FormatNumber(metasRecente(i), 2) %></div>
                                </div>
                                <% End If %>
                                
                                <!-- Variação -->
                                <div class="mt-2">
                                    <span class="badge <%= classeVariacao %>">
                                        <i class="fas <%= iconeVariacao %> me-1"></i>
                                        <%= textoVariacao %>
                                    </span>
                                </div>
                                
                                <!-- Diferença em R$ -->
                                <% If diferencasAnuais(i) <> 0 Then %>
                                <div class="mt-1">
                                    <small class="<%= classeVariacao %>">
                                        <% 
                                        If diferencasAnuais(i) > 0 Then
                                            Response.Write "+R$ " & FormatNumber(diferencasAnuais(i), 2)
                                        Else
                                            Response.Write "-R$ " & FormatNumber(Abs(diferencasAnuais(i)), 2)
                                        End If
                                        %>
                                    </small>
                                </div>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                    <% Next %>
                </div>
            </div>
        </div>

        <!-- Últimas Vendas -->
        <div class="card ultimas-vendas">
            <h5 class="card-title"><i class="fas fa-history me-2"></i> Últimas Vendas Registradas em <%= anoRecente %></h5>
            <div class="row mt-3">
                <% If Not rsUltimasVendas.EOF Then %>
                <% Do While Not rsUltimasVendas.EOF %>
                <div class="col-md-4">
                    <div class="venda-item">
                        <span class="fw-bold"><%= rsUltimasVendas("NomeEmpreendimento") %> - Unidade <%= rsUltimasVendas("Unidade") %></span>
                        <br>
                        Valor: <span class="text-success fw-bold">R$ <%= FormatNumber(rsUltimasVendas("ValorUnidade"), 2) %></span>
                        <br>
                        Data: <%= rsUltimasVendas("DataVenda") %> | Corretor: <%= rsUltimasVendas("Corretor") %>
                    </div>
                </div>
                <% rsUltimasVendas.MoveNext %>
                <% Loop %>
                <% Else %>
                <div class="col-12"><p class="text-muted text-center">Nenhuma venda registrada para o ano <%= anoRecente %>.</p></div>
                <% End If %>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        // Dados do VBScript
        const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        const vendasAntigo = [<%= strVendasAntigoJS %>];
        const vendasRecente = [<%= strVendasRecenteJS %>];
        const variacaoPercentual = [<%= strVariacaoPercentualJS %>];
        
        // Configuração do gráfico comparativo
        const ctx = document.getElementById('graficoComparativo').getContext('2d');
        const graficoComparativo = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: meses,
                datasets: [
                    {
                        label: '<%= anoAntigo %>',
                        data: vendasAntigo,
                        backgroundColor: '#95a5a6',
                        borderColor: '#7f8c8d',
                        borderWidth: 1,
                        barPercentage: 0.4,
                    },
                    {
                        label: '<%= anoRecente %>',
                        data: vendasRecente,
                        backgroundColor: '#3498db',
                        borderColor: '#2980b9',
                        borderWidth: 1,
                        barPercentage: 0.4,
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Valor (R$)'
                        },
                        ticks: {
                            callback: function(value) {
                                return 'R$ ' + value.toLocaleString('pt-BR', {minimumFractionDigits: 0});
                            }
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Meses'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                label += 'R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                                return label;
                            },
                            afterLabel: function(context) {
                                const index = context.dataIndex;
                                const variacao = variacaoPercentual[index];
                                
                                if (context.datasetIndex === 1) { // Apenas para o ano recente
                                    if (variacao !== 0) {
                                        const sinal = variacao > 0 ? '+' : '';
                                        return `Variação: ${sinal}${variacao.toFixed(1)}%`;
                                    }
                                }
                                return null;
                            }
                        }
                    }
                }
            }
        });

        // Função para atualizar a página
        function atualizarPagina() {
            location.reload();
        }

        // Atualizar a página a cada 60 segundos
        setInterval(atualizarPagina, 60000);
    </script>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>

<%
' Fechar conexões
If Not rsUltimasVendas Is Nothing Then
    If Not rsUltimasVendas.State = 0 Then rsUltimasVendas.Close
    Set rsUltimasVendas = Nothing
End If

If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If
%>