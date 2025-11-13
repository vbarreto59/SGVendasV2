<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar ao banco de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

Function GetUniqueValues(tableName, columnName, whereClause)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    sqlQuery = "SELECT DISTINCT " & columnName & " FROM " & tableName & " "
    sqlQuery = sqlQuery & whereClause & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = connSales.Execute(sqlQuery)
    If Err.Number <> 0 Then
        GetUniqueValues = Array()
        Exit Function
    End If
    On Error GoTo 0
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                dict.Add CStr(rs(0)), 1
            End If
            rs.MoveNext
        Loop
    End If
    
    If Not rs Is Nothing Then 
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

Function RemoverNumeros(texto)
    If IsNull(texto) Or texto = "" Then
        RemoverNumeros = ""
        Exit Function
    End If
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    RemoverNumerosEAsteriscos = regex.Replace(texto, "")
    RemoverNumeros = Trim(Replace(RemoverNumerosEAsteriscos, "  ", " "))
End Function

' ===============================================
' OBTER DADOS DAS VENDAS
' ===============================================

Dim filtroAno, filtroMes, filtroCorretor, filtroEmpreendimento
filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")

' Popular selects
Dim uniqueAnos, uniqueMeses, uniqueCorretores, uniqueEmpreendimentos
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueMeses = GetUniqueValues("Vendas", "MesVenda", "WHERE MesVenda IS NOT NULL")
uniqueCorretores = GetUniqueValues("Vendas", "Corretor", "WHERE Corretor IS NOT NULL AND Corretor <> ''")
uniqueEmpreendimentos = GetUniqueValues("Vendas", "NomeEmpreendimento", "WHERE NomeEmpreendimento IS NOT NULL AND NomeEmpreendimento <> ''")

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Jan"
arrMesesNome(2) = "Fev"
arrMesesNome(3) = "Mar"
arrMesesNome(4) = "Abr"
arrMesesNome(5) = "Mai"
arrMesesNome(6) = "Jun"
arrMesesNome(7) = "Jul"
arrMesesNome(8) = "Ago"
arrMesesNome(9) = "Set"
arrMesesNome(10) = "Out"
arrMesesNome(11) = "Nov"
arrMesesNome(12) = "Dez"

' Consulta principal
Dim rsVendas, sqlVendas
Set rsVendas = Server.CreateObject("ADODB.Recordset")

sqlVendas = "SELECT " & _
            "Vendas.ID, " & _
            "Vendas.ValorUnidade, " & _
            "Vendas.DataVenda, " & _
            "Vendas.AnoVenda, " & _            
            "Vendas.MesVenda, " & _            
            "Vendas.DiaVenda, " & _            
            "Vendas.NomeCliente, " & _
            "Vendas.Corretor, " & _
            "Vendas.CorretorId, " & _
            "Vendas.NomeEmpreendimento, " & _
            "Vendas.Empreend_ID, " & _
            "Vendas.Unidade, " & _
            "Vendas.UnidadeM2, " & _
            "Vendas.Diretoria, " & _
            "Vendas.Gerencia, " & _
            "Vendas.MesVenda, " & _
            "Vendas.AnoVenda, " & _
            "Vendas.Trimestre, " & _
            "Vendas.ComissaoPercentual, " & _
            "Vendas.ValorComissaoGeral, " & _
            "Vendas.ValorDiretoria, " & _
            "Vendas.ValorGerencia, " & _
            "Vendas.ValorCorretor, " & _
            "Vendas.ValorLiqDiretoria, " & _
            "Vendas.ValorLiqGerencia, " & _
            "Vendas.ValorLiqCorretor, " & _            
            "Vendas.ComissaoDiretoria, " & _
            "Vendas.ComissaoGerencia, " & _
            "Vendas.ComissaoCorretor, " & _
            "Vendas.DataRegistro, " & _
            "Vendas.Usuario, " & _
            "Vendas.NomeDiretor, " & _
            "Vendas.NomeGerente, " & _
            "Vendas.UserIdDiretoria, " & _
            "Vendas.UserIdGerencia " & _
            "FROM Vendas " & _
            "WHERE (Vendas.Excluido <> -1 OR Vendas.Excluido IS NULL) " 

            'response.write sqlVendas
            'Response.end 
            

If filtroAno <> "" Then sqlVendas = sqlVendas & " AND Vendas.AnoVenda = " & filtroAno
If filtroMes <> "" Then sqlVendas = sqlVendas & " AND Vendas.MesVenda = " & filtroMes
If filtroCorretor <> "" Then sqlVendas = sqlVendas & " AND Vendas.Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
If filtroEmpreendimento <> "" Then sqlVendas = sqlVendas & " AND Vendas.NomeEmpreendimento = '" & Replace(filtroEmpreendimento, "'", "''") & "'"

sqlVendas = sqlVendas & " ORDER BY Vendas.AnoVenda DESC, Vendas.MesVenda DESC, Vendas.DiaVenda DESC"

On Error Resume Next
rsVendas.Open sqlVendas, connSales

If Err.Number <> 0 Then
    Response.Write "Erro na consulta: " & Err.Description
    Response.End
End If
On Error GoTo 0

' Calcular totais
Dim totalValor, totalComissao
totalValor = 0
totalComissao = 0

If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        totalValor = totalValor + CDbl(rsVendas("ValorUnidade"))
        totalComissao = totalComissao + CDbl(rsVendas("ValorComissaoGeral"))
        rsVendas.MoveNext
    Loop
    rsVendas.MoveFirst
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Vendas - Mobile</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 10px;
            font-size: 14px;
        }
        .container {
            max-width: 100%;
            margin: 0 auto;
        }
        .header {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 15px;
            margin-bottom: 15px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }
        .page-title {
            color: #2c3e50;
            font-size: 20px;
            font-weight: 700;
            text-align: center;
            margin-bottom: 5px;
        }
        .page-subtitle {
            color: #7f8c8d;
            font-size: 12px;
            text-align: center;
            margin-bottom: 15px;
        }
        .filter-card {
            background: rgba(255, 255, 255, 0.9);
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        .filter-label {
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 6px;
            font-size: 12px;
        }
        .form-select {
            border-radius: 8px;
            border: 1px solid #e9ecef;
            padding: 8px;
            font-size: 12px;
            margin-bottom: 8px;
        }
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            border-radius: 8px;
            padding: 10px;
            font-weight: 600;
            font-size: 12px;
        }
        .btn-secondary {
            background: #6c757d;
            border: none;
            border-radius: 8px;
            padding: 10px;
            font-weight: 600;
            font-size: 12px;
            color: white;
        }
        .venda-card {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 15px;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.3);
            position: relative;
        }
        .valor-venda {
            font-size: 22px;
            font-weight: 800;
            color: #27ae60;
            text-align: center;
            margin-bottom: 10px;
        }
        .empreendimento {
            font-size: 16px;
            font-weight: 700;
            color: #2c3e50;
            text-align: center;
            margin-bottom: 8px;
        }
        .info-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin: 10px 0;
        }
        .info-item {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 8px;
            font-size: 11px;
        }
        .info-label {
            font-weight: 600;
            color: #6c757d;
            display: block;
            font-size: 10px;
        }
        .info-value {
            color: #2c3e50;
            font-weight: 500;
        }
        .comissao-grid {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 6px;
            margin: 10px 0;
        }
        .comissao-item {
            background: #34495e;
            color: white;
            border-radius: 8px;
            padding: 8px;
            text-align: center;
            font-size: 10px;
        }
        .comissao-valor {
            font-size: 12px;
            font-weight: 700;
            margin-top: 2px;
        }
        .badge-pago {
            background: #27ae60;
            color: white;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 9px;
            font-weight: 600;
        }
        .badge-mes {
            background: #3498db;
            color: white;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 10px;
            font-weight: 600;
            position: absolute;
            top: 10px;
            right: 10px;
        }
        .actions {
            display: flex;
            gap: 8px;
            margin-top: 10px;
        }
        .btn-action {
            flex: 1;
            padding: 6px;
            font-size: 11px;
            border-radius: 6px;
        }
        .total-card {
            background: #2c3e50;
            color: white;
            border-radius: 12px;
            padding: 12px;
            margin-top: 10px;
            text-align: center;
        }
        .total-label {
            font-size: 12px;
            opacity: 0.9;
        }
        .total-value {
            font-size: 16px;
            font-weight: 700;
        }
        .no-results {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            padding: 30px 20px;
            text-align: center;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        .no-results i {
            font-size: 36px;
            color: #bdc3c7;
            margin-bottom: 10px;
        }
        .no-results h4 {
            color: #7f8c8d;
            margin-bottom: 8px;
            font-size: 16px;
        }
        .filter-buttons {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin-top: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Cabeçalho -->
        <div class="header">
            <h1 class="page-title">
                <i class="fas fa-handshake"></i> Gestão de Vendas
            </h1>
            <p class="page-subtitle">Versão Mobile</p>
            
            <!-- Filtros -->
            <div class="filter-card">
                <form id="filterForm" method="get">
                    <div class="row g-2">
                        <div class="col-6">
                            <label class="filter-label">Ano</label>
                            <select class="form-select" name="ano" id="anoFilter">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueAnos) Then
                                    For Each ano In uniqueAnos
                                        Response.Write "<option value=""" & ano & """"
                                        If CStr(filtroAno) = CStr(ano) Then Response.Write " selected"
                                        Response.Write ">" & ano & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-6">
                            <label class="filter-label">Mês</label>
                            <select class="form-select" name="mes" id="mesFilter">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueMeses) Then
                                    For Each mes In uniqueMeses
                                        If Not IsEmpty(mes) Then
                                            Dim mesNum
                                            mesNum = CInt(mes)
                                            Response.Write "<option value=""" & mes & """"
                                            If CStr(filtroMes) = CStr(mes) Then Response.Write " selected"
                                            Response.Write ">" & arrMesesNome(mesNum) & "</option>"
                                        End If
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-12">
                            <label class="filter-label">Corretor</label>
                            <select class="form-select" name="corretor" id="corretorFilter">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueCorretores) Then
                                    For Each corretor In uniqueCorretores
                                        Response.Write "<option value=""" & corretor & """"
                                        If filtroCorretor = corretor Then Response.Write " selected"
                                        Response.Write ">" & corretor & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-12">
                            <label class="filter-label">Empreendimento</label>
                            <select class="form-select" name="empreendimento" id="empreendimentoFilter">
                                <option value="">Todos</option>
                                <%
                                If IsArray(uniqueEmpreendimentos) Then
                                    For Each empreendimento In uniqueEmpreendimentos
                                        Response.Write "<option value=""" & empreendimento & """"
                                        If filtroEmpreendimento = empreendimento Then Response.Write " selected"
                                        Response.Write ">" & empreendimento & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="col-12">
                            <div class="filter-buttons">
                                <button type="submit" class="btn btn-primary">
                                    <i class="fas fa-filter"></i> Aplicar
                                </button>
                                <button type="button" class="btn btn-secondary" onclick="limparFiltros()">
                                    <i class="fas fa-times"></i> Limpar
                                </button>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
        </div>

        <!-- Totais -->
        <div class="total-card">
            <div class="row">
                <div class="col-6">
                    <div class="total-label">Total Vendas</div>
                    <div class="total-value">R$ <%= FormatNumber(totalValor, 2) %></div>
                </div>
                <div class="col-6">
                    <div class="total-label">Total Comissões</div>
                    <div class="total-value">R$ <%= FormatNumber(totalComissao, 2) %></div>
                </div>
            </div>
        </div>

        <!-- Lista de Vendas -->
        <%
        If Not rsVendas.EOF Then
            Do While Not rsVendas.EOF
                ' Verificar pagamentos (igual à versão desktop)
                Dim sqlPagamentos, rsPagamentos
                Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
                Dim pagoDiretoria, pagoGerencia, pagoCorretor
                
                totalPagoDiretoria = 0
                totalPagoGerencia = 0
                totalPagoCorretor = 0
                pagoDiretoria = False
                pagoGerencia = False
                pagoCorretor = False

                sqlPagamentos = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rsVendas("ID") & " ORDER BY DataPagamento ASC;"
                Set rsPagamentos = connSales.Execute(sqlPagamentos)

                If Not rsPagamentos.EOF Then
                    Do While Not rsPagamentos.EOF
                        Select Case LCase(rsPagamentos("TipoRecebedor"))
                            Case "diretoria"
                                totalPagoDiretoria = totalPagoDiretoria + CDbl(rsPagamentos("ValorPago"))
                            Case "gerencia"
                                totalPagoGerencia = totalPagoGerencia + CDbl(rsPagamentos("ValorPago"))
                            Case "corretor"
                                totalPagoCorretor = totalPagoCorretor + CDbl(rsPagamentos("ValorPago"))
                        End Select
                        rsPagamentos.MoveNext
                    Loop
                End If
                rsPagamentos.Close
                Set rsPagamentos = Nothing

                If rsVendas("ValorDiretoria") > 0 And totalPagoDiretoria >= CDbl(rsVendas("ValorDiretoria")) Then pagoDiretoria = True
                If rsVendas("ValorDiretoria") = 0 Then pagoDiretoria = True
                If rsVendas("ValorGerencia") > 0 And totalPagoGerencia >= CDbl(rsVendas("ValorGerencia")) Then pagoGerencia = True
                If rsVendas("ValorGerencia") = 0 Then pagoGerencia = True
                If rsVendas("ValorCorretor") > 0 And totalPagoCorretor >= CDbl(rsVendas("ValorCorretor")) Then pagoCorretor = True
                If rsVendas("ValorCorretor") = 0 Then pagoCorretor = True

                ' Verificar se comissão já existe
                Dim rsComissaoCheck, comissaoExiste
                Set rsComissaoCheck = Server.CreateObject("ADODB.Recordset")
                rsComissaoCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(rsVendas("ID")), connSales
                comissaoExiste = Not rsComissaoCheck.EOF
                rsComissaoCheck.Close
                Set rsComissaoCheck = Nothing
        %>
        <!-- ============================= -->
<div class="venda-card">
    <span class="badge-mes"><%= arrMesesNome(rsVendas("MesVenda")) %>/<%= Right(rsVendas("AnoVenda"), 2) %></span>
    
    <div class="valor-venda">
        R$ <%= FormatNumber(rsVendas("ValorUnidade"), 2) %>
    </div>
    
    <div class="empreendimento">
        ID Venda: <%= rsVendas("ID") %> <br>(<%= rsVendas("Empreend_ID") %>) <%= RemoverNumeros(rsVendas("NomeEmpreendimento")) %>
    </div>
    
    <div class="info-grid">
        <div class="info-item">
            <span class="info-label">Unidade</span>
            <span class="info-value"><%= rsVendas("Unidade") %></span>
        </div>
        <div class="info-item">
            <span class="info-label">M²</span>
            <span class="info-value"><%= rsVendas("UnidadeM2") %></span>
        </div>
        <div class="info-item">
            <span class="info-label">Data Venda:</span>
            <span class="info-value"><%= rsVendas("DiaVenda") & "/" & rsVendas("MesVenda") & "/" & rsVendas("AnoVenda") %></span>
        </div>
        <div class="info-item">
            <span class="info-label">Corretor</span>
            <span class="info-value"><%= rsVendas("CorretorId") & "-" & rsVendas("Corretor") %></span>
        </div>
    </div>
    
    <!-- Nova seção: Informações de Diretoria, Gerência e Cliente -->
    <div class="info-grid" style="margin-top: 8px;">
        <% If rsVendas("Diretoria") <> "" Then %>
        <div class="info-item">
            <span class="info-label">Diretoria</span>
            <span class="info-value"><%= rsVendas("Diretoria") %></span>
        </div>
        <% End If %>
        
        <% If rsVendas("NomeDiretor") <> "" Then %>
        <div class="info-item">
            <span class="info-label">Diretor</span>
            <span class="info-value"><%= rsVendas("NomeDiretor") %></span>
        </div>
        <% End If %>
        
        <% If rsVendas("Gerencia") <> "" Then %>
        <div class="info-item">
            <span class="info-label">Gerência</span>
            <span class="info-value"><%= rsVendas("Gerencia") %></span>
        </div>
        <% End If %>
        
        <% If rsVendas("NomeGerente") <> "" Then %>
        <div class="info-item">
            <span class="info-label">Gerente</span>
            <span class="info-value"><%= rsVendas("NomeGerente") %></span>
        </div>
        <% End If %>
        
        <% If rsVendas("NomeCliente") <> "" Then %>
        <div class="info-item" style="grid-column: 1 / -1;">
            <span class="info-label">Cliente</span>
            <span class="info-value"><%= rsVendas("NomeCliente") %></span>
        </div>
        <% End If %>
    </div>
    
    <div class="comissao-grid">
        <div class="comissao-item">
            <div>Diretoria</div>
            <div class="comissao-valor"><%= rsVendas("ComissaoDiretoria") %>%</div>
            <div>R$ <%= FormatNumber(rsVendas("ValorLiqDiretoria"), 2) %></div>
            <% If pagoDiretoria Then %><span class="badge-pago">PAGO</span><% End If %>
        </div>
        <div class="comissao-item">
            <div>Gerência</div>
            <div class="comissao-valor"><%= rsVendas("ComissaoGerencia") %>%</div>
            <div>R$ <%= FormatNumber(rsVendas("ValorLiqGerencia"), 2) %></div>
            <% If pagoGerencia Then %><span class="badge-pago">PAGO</span><% End If %>
        </div>
        <div class="comissao-item">
            <div>Corretor</div>
            <div class="comissao-valor"><%= rsVendas("ComissaoCorretor") %>%</div>
            <div>R$ <%= FormatNumber(rsVendas("ValorLiqCorretor"), 2) %></div>
            <% If pagoCorretor Then %><span class="badge-pago">PAGO</span><% End If %>
        </div>
    </div>
    
    <div class="actions">
        <% If UCase(Session("Usuario")) = "BARRETO" Then %>

            <% If Not comissaoExiste Then %>
            <a href="gestao_vendas_inserir_comissao1.asp?id=<%= rsVendas("ID") %>" class="btn btn-primary btn-action">
                <i class="fas fa-plus"></i> Comissão
            </a>
            <% End If %>

        <% End If %>
    </div>
    
    <div style="font-size: 10px; color: #6c757d; text-align: center; margin-top: 8px;">
        Registro: <%= FormatDateTime(rsVendas("DataRegistro"), 2) %> por <%= rsVendas("Usuario") %>
    </div>
</div>


        <!-- ============================================================ -->
        <%
                rsVendas.MoveNext
            Loop
        Else
        %>
        <div class="no-results">
            <i class="fas fa-search"></i>
            <h4>Nenhuma venda encontrada</h4>
            <p>Não foram encontradas vendas para os filtros selecionados.</p>
        </div>
        <%
        End If

        
        If rsVendas.State = 1 Then rsVendas.Close
        Set rsVendas = Nothing
        connSales.Close
        Set connSales = Nothing
        %>
    </div>
    <!-- div container -->

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function limparFiltros() {
            // Redirecionar para a mesma página sem parâmetros de filtro
            window.location.href = window.location.pathname;
        }

        // Verificar se há filtros ativos e mostrar indicador visual
        document.addEventListener('DOMContentLoaded', function() {
            const urlParams = new URLSearchParams(window.location.search);
            const hasFilters = urlParams.has('ano') || urlParams.has('mes') || urlParams.has('corretor') || urlParams.has('empreendimento');
            
            if (hasFilters) {
                // Adicionar classe aos selects que têm valores selecionados
                const selects = document.querySelectorAll('select');
                selects.forEach(select => {
                    if (select.value !== '') {
                        select.style.borderColor = '#28a745';
                        select.style.backgroundColor = '#f8fff9';
                    }
                });
            }
        });
    </script>
</body>
</html>