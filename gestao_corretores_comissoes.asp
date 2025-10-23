<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

' Abrir conexão apenas com o banco Sales
Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar ao banco de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno, filtroMes
filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")

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

' ===============================================
' POPULAR OS SELECTS DO FORMULÁRIO
' ===============================================

Dim uniqueAnos, uniqueMeses
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueMeses = GetUniqueValues("Vendas", "MesVenda", "WHERE MesVenda IS NOT NULL")

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Março"
arrMesesNome(4) = "Abril"
arrMesesNome(5) = "Maio"
arrMesesNome(6) = "Junho"
arrMesesNome(7) = "Julho"
arrMesesNome(8) = "Agosto"
arrMesesNome(9) = "Setembro"
arrMesesNome(10) = "Outubro"
arrMesesNome(11) = "Novembro"
arrMesesNome(12) = "Dezembro"

' ===============================================
' OBTER DADOS DE COMISSÕES (APENAS SE ANO ESTIVER PREENCHIDO)
' ===============================================

Dim comissoesData, totalGeralComissoes, totalCorretores, totalVendasGeral, totalVGVGeral
Set comissoesData = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Construir consulta SQL
    Dim sqlComissoes, rsComissoes
    sqlComissoes = "SELECT " & _
                   "Corretor, " & _
                   "Diretoria, " & _
                   "Gerencia, " & _
                   "MesVenda, " & _
                   "SUM(ValorCorretor) as TotalComissao, " & _
                   "COUNT(*) as TotalVendas, " & _
                   "SUM(ValorUnidade) as TotalVGV " & _
                   "FROM Vendas " & _
                   "WHERE Excluido = 0 " & _
                   "AND AnoVenda = " & filtroAno
    
    If filtroMes <> "" Then
        sqlComissoes = sqlComissoes & " AND MesVenda = " & filtroMes
    End If
    
    sqlComissoes = sqlComissoes & " GROUP BY Corretor, Diretoria, Gerencia, MesVenda " & _
                   "ORDER BY Corretor, MesVenda"

    Set rsComissoes = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsComissoes.Open sqlComissoes, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta de comissões: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlComissoes)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados de comissões
    totalGeralComissoes = 0
    totalCorretores = 0
    totalVendasGeral = 0
    totalVGVGeral = 0
    
    Dim corretoresProcessados
    Set corretoresProcessados = Server.CreateObject("Scripting.Dictionary")

    If Not rsComissoes.EOF Then
        Do While Not rsComissoes.EOF
            Dim corretor, diretoria, gerencia, mes, comissaoMes, vendasMes, vgvMes
            corretor = CStr(rsComissoes("Corretor"))
            diretoria = CStr(rsComissoes("Diretoria"))
            gerencia = CStr(rsComissoes("Gerencia"))
            mes = CStr(rsComissoes("MesVenda"))
            comissaoMes = CDbl(rsComissoes("TotalComissao"))
            vendasMes = CLng(rsComissoes("TotalVendas"))
            vgvMes = CDbl(rsComissoes("TotalVGV"))
            
            ' Adicionar corretor à lista de processados
            If Not corretoresProcessados.Exists(corretor) Then
                corretoresProcessados.Add corretor, 1
                totalCorretores = totalCorretores + 1
            End If
            
            ' Criar estrutura para o corretor se não existir
            If Not comissoesData.Exists(corretor) Then
                Dim infoCorretor
                Set infoCorretor = Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "Diretoria", diretoria
                infoCorretor.Add "Gerencia", gerencia
                infoCorretor.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "TotalComissao", 0
                infoCorretor.Add "TotalVendas", 0
                infoCorretor.Add "TotalVGV", 0
                comissoesData.Add corretor, infoCorretor
            End If
            
            ' Atualizar dados do mês
            Set infoCorretor = comissoesData(corretor)
            infoCorretor("Meses").Add mes, Array(comissaoMes, vendasMes, vgvMes)
            
            ' Atualizar totais do corretor
            infoCorretor("TotalComissao") = infoCorretor("TotalComissao") + comissaoMes
            infoCorretor("TotalVendas") = infoCorretor("TotalVendas") + vendasMes
            infoCorretor("TotalVGV") = infoCorretor("TotalVGV") + vgvMes
            
            ' Atualizar totais gerais
            totalGeralComissoes = totalGeralComissoes + comissaoMes
            totalVendasGeral = totalVendasGeral + vendasMes
            totalVGVGeral = totalVGVGeral + vgvMes
            
            rsComissoes.MoveNext
        Loop
    End If

    If rsComissoes.State = 1 Then rsComissoes.Close
    Set rsComissoes = Nothing
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Comissões dos Corretores</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #A5A2A2;
            padding: 20px;
            color: white;
        }
        .card-kpi {
            background-color: #F0ECEC;
            color: black;
            padding: 15px;
            margin-top: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
        }
        .container-fluid {
            max-width: 1800px;
            margin: 0 auto;
        }
        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 20px;
            border-radius: 8px;
            font-size: 1rem;
            margin-bottom: 10px;
            min-height: 120px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .kpi-card h5 {
            font-size: 1rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            line-height: 1.2;
            font-size: 0.9rem;
        }
        .kpi-card i {
            font-size: 1.5rem;
            margin-bottom: 8px;
        }
        .bg-primary-kpi { background-color: #007bff; }
        .bg-success-kpi { background-color: #28a745; }
        .bg-info-kpi { background-color: #17a2b8; }
        .bg-warning-kpi { background-color: #ffc107; color: #000; }
        .bg-danger-kpi { background-color: #dc3545; }
        .bg-secondary-kpi { background-color: #6c757d; }
        .bg-dark-kpi { background-color: #343a40; }
        .bg-maroon-kpi { background-color: #800000; }
        
        .filter-container {
            background-color: #Fff;
            color: black;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .table-responsive {
            background-color: white;
            border-radius: 8px;
            font-size: 0.85rem;
        }
        .table th {
            background-color: #800000;
            color: white;
            position: sticky;
            top: 0;
            font-size: 0.8rem;
        }
        .text-right-v { text-align: right; }
        .text-center-v { text-align: center; }
        .corretor-header {
            background-color: #e9ecef !important;
            font-weight: bold;
        }
        .mes-header {
            background-color: #17a2b8;
            color: white;
            font-weight: bold;
        }
        .total-row {
            background-color: #800000;
            color: white;
            font-weight: bold;
        }
        .alert-warning {
            background-color: #fff3cd;
            border-color: #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .table-hover tbody tr:hover {
            background-color: rgba(0,0,0,.075);
        }
        .comissao-cell {
            font-weight: bold;
            color: #28a745;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-money-bill-wave"></i> SGVendas - Comissões dos Corretores
        </h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-4">
                        <div class="filter-label">Ano</div>
                        <select class="form-select" name="ano" id="anoFilter" required>
                            <option value="">Selecione o ano</option>
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
                    
                    <div class="col-md-4">
                        <div class="filter-label">Mês</div>
                        <select class="form-select" name="mes" id="mesFilter">
                            <option value="">Todos os meses</option>
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
                    
                    <div class="col-md-4">
                        <div class="filter-label">&nbsp;</div>
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Relatório
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Then %>
            <div class="alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar o relatório de comissões.
            </div>
        <% Else %>
        
        <!-- KPIs Principais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-money-bill-wave"></i>
                    <h5>Total Comissões <%= filtroAno %></h5>
                    <p>R$ <%= FormatNumber(totalGeralComissoes, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <i class="fas fa-user-tie"></i>
                    <h5>Corretores com Vendas</h5>
                    <p><%= totalCorretores %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-calendar-alt"></i>
                    <h5>Período</h5>
                    <p>
                        <% 
                        If filtroMes <> "" Then 
                            Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                        Else 
                            Response.Write "Ano " & filtroAno & " (Todos os meses)"
                        End If 
                        %>
                    </p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-chart-line"></i>
                    <h5>Comissão Média</h5>
                    <p>
                        <% 
                        If totalCorretores > 0 Then 
                            Response.Write "R$ " & FormatNumber(totalGeralComissoes / totalCorretores, 2)
                        Else 
                            Response.Write "R$ 0,00"
                        End If 
                        %>
                    </p>
                </div>
            </div>
        </div>

        <!-- Tabela de Comissões -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">
                Comissões por Corretor - 
                <% 
                If filtroMes <> "" Then 
                    Response.Write arrMesesNome(CInt(filtroMes)) & " de " & filtroAno
                Else 
                    Response.Write "Ano " & filtroAno
                End If 
                %>
            </h3>
            
            <div class="table-responsive" style="max-height: 600px; overflow-y: auto;">
                <table class="table table-striped table-hover table-bordered">
                    <thead>
                        <tr>
                            <th class="text-center-v">Corretor</th>
                            <th class="text-center-v">Diretoria</th>
                            <th class="text-center-v">Gerência</th>
                            <%
                            ' Cabeçalhos dos meses (apenas se não tiver filtro de mês específico)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Response.Write "<th class='text-center-v mes-header'>" & Left(arrMesesNome(i), 3) & "</th>"
                                Next
                            Else
                                Response.Write "<th class='text-center-v mes-header'>Comissão " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>Vendas " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                                Response.Write "<th class='text-center-v'>VGV " & Left(arrMesesNome(CInt(filtroMes)), 3) & "</th>"
                            End If
                            %>
                            <th class="text-center-v bg-success text-white">Total Comissão</th>
                            <th class="text-center-v bg-primary text-white">Total Vendas</th>
                            <th class="text-center-v bg-info text-white">Total VGV</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        If comissoesData.Count > 0 Then
                            Dim arrCorretores, corretorKey
                            arrCorretores = comissoesData.Keys
                            
                            ' Ordenar corretores por total de comissão (decrescente)
                            For i = 0 To UBound(arrCorretores)
                                For j = i + 1 To UBound(arrCorretores)
                                    If comissoesData(arrCorretores(j))("TotalComissao") > comissoesData(arrCorretores(i))("TotalComissao") Then
                                        Dim temp
                                        temp = arrCorretores(i)
                                        arrCorretores(i) = arrCorretores(j)
                                        arrCorretores(j) = temp
                                    End If
                                Next
                            Next
                            
                            For Each corretorKey In arrCorretores
                                Set infoCorretor = comissoesData(corretorKey)
                                Dim mesesCorretor
                                Set mesesCorretor = infoCorretor("Meses")
                        %>
                        <tr>
                            <td class="corretor-header"><%= corretorKey %></td>
                            <td><%= infoCorretor("Diretoria") %></td>
                            <td><%= infoCorretor("Gerencia") %></td>
                            
                            <%
                            ' Dados dos meses
                            If filtroMes = "" Then
                                ' Mostrar todos os meses
                                For i = 1 To 12
                                    Dim mesKey
                                    mesKey = CStr(i)
                                    If mesesCorretor.Exists(mesKey) Then
                                        Dim dadosMes
                                        dadosMes = mesesCorretor(mesKey)
                                        Response.Write "<td class='text-right-v comissao-cell'>R$ " & FormatNumber(dadosMes(0), 2) & "</td>"
                                    Else
                                        Response.Write "<td class='text-center-v'>-</td>"
                                    End If
                                Next
                            Else
                                ' Mostrar apenas o mês filtrado com detalhes
                                If mesesCorretor.Exists(filtroMes) Then
                                    Dim dadosMesFiltrado
                                    dadosMesFiltrado = mesesCorretor(filtroMes)
                                    Response.Write "<td class='text-right-v comissao-cell'>R$ " & FormatNumber(dadosMesFiltrado(0), 2) & "</td>"
                                    Response.Write "<td class='text-center-v'>" & dadosMesFiltrado(1) & "</td>"
                                    Response.Write "<td class='text-right-v'>R$ " & FormatNumber(dadosMesFiltrado(2), 2) & "</td>"
                                Else
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                    Response.Write "<td class='text-center-v'>-</td>"
                                End If
                            End If
                            %>
                            
                            <!-- Totais do corretor -->
                            <td class="text-right-v bg-success text-white"><strong>R$ <%= FormatNumber(infoCorretor("TotalComissao"), 2) %></strong></td>
                            <td class="text-center-v bg-primary text-white"><strong><%= infoCorretor("TotalVendas") %></strong></td>
                            <td class="text-right-v bg-info text-white"><strong>R$ <%= FormatNumber(infoCorretor("TotalVGV"), 2) %></strong></td>
                        </tr>
                        <%
                            Next
                        Else
                        %>
                        <tr>
                            <td 
                            <% If filtroMes = "" Then %>
                                colspan="16"
                            <% Else %>
                                colspan="8"
                            <% End If %>
                            class="text-center-v">Nenhum dado encontrado para os filtros selecionados.</td>
                        </tr>
                        <%
                        End If
                        %>
                    </tbody>
                    <tfoot>
                        <tr class="total-row">
                            <td colspan="3"><strong>TOTAIS GERAIS</strong></td>
                            <%
                            ' Totais por mês (apenas se não tiver filtro de mês)
                            If filtroMes = "" Then
                                For i = 1 To 12
                                    Dim totalMes
                                    totalMes = 0
                                    For Each corretorKey In comissoesData.Keys
                                        Set mesesCorretor = comissoesData(corretorKey)("Meses")
                                        If mesesCorretor.Exists(CStr(i)) Then
                                            totalMes = totalMes + mesesCorretor(CStr(i))(0)
                                        End If
                                    Next
                                    Response.Write "<td class='text-right-v'><strong>R$ " & FormatNumber(totalMes, 2) & "</strong></td>"
                                Next
                            Else
                                Response.Write "<td colspan='3'></td>"
                            End If
                            %>
                            <td class="text-right-v"><strong>R$ <%= FormatNumber(totalGeralComissoes, 2) %></strong></td>
                            <td class="text-center-v"><strong><%= totalVendasGeral %></strong></td>
                            <td class="text-right-v"><strong>R$ <%= FormatNumber(totalVGVGeral, 2) %></strong></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <!-- Resumo Estatístico -->
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">Top 5 Corretores (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>Posição</th>
                                    <th>Corretor</th>
                                    <th class="text-right-v">Comissão (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                If comissoesData.Count > 0 Then
                                    Dim contador
                                    contador = 0
                                    For Each corretorKey In arrCorretores
                                        If contador < 5 Then
                                            Set infoCorretor = comissoesData(corretorKey)
                                %>
                                <tr>
                                    <td><%= contador + 1 %></td>
                                    <td><%= corretorKey %></td>
                                    <td class="text-right-v">R$ <%= FormatNumber(infoCorretor("TotalComissao"), 2) %></td>
                                </tr>
                                <%
                                            contador = contador + 1
                                        Else
                                            Exit For
                                        End If
                                    Next
                                Else
                                %>
                                <tr>
                                    <td colspan="3" class="text-center">Nenhum dado disponível</td>
                                </tr>
                                <%
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">Estatísticas</h4>
                    <%
                    If comissoesData.Count > 0 Then
                        Dim maiorComissao, menorComissao, corretorMaior, corretorMenor
                        maiorComissao = 0
                        menorComissao = 999999999
                        
                        For Each corretorKey In comissoesData.Keys
                            Set infoCorretor = comissoesData(corretorKey)
                            If infoCorretor("TotalComissao") > maiorComissao Then
                                maiorComissao = infoCorretor("TotalComissao")
                                corretorMaior = corretorKey
                            End If
                            If infoCorretor("TotalComissao") < menorComissao Then
                                menorComissao = infoCorretor("TotalComissao")
                                corretorMenor = corretorKey
                            End If
                        Next
                    %>
                    <p><strong>Maior Comissão:</strong><br>
                    <%= corretorMaior %> - R$ <%= FormatNumber(maiorComissao, 2) %></p>
                    
                    <p><strong>Menor Comissão:</strong><br>
                    <%= corretorMenor %> - R$ <%= FormatNumber(menorComissao, 2) %></p>
                    
                    <p><strong>Média de Comissões:</strong><br>
                    R$ <%= FormatNumber(totalGeralComissoes / comissoesData.Count, 2) %></p>
                    <%
                    Else
                    %>
                    <p class="text-center">Nenhum dado disponível</p>
                    <%
                    End If
                    %>
                </div>
            </div>
        </div>

        <% End If %>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>