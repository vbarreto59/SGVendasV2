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

Dim filtroAno, filtroCorretor
filtroAno = Request.QueryString("ano")
filtroCorretor = Request.QueryString("corretor")

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
' POPULAR OS SELECTS DO FORMULÁRIO PRIMEIRO
' ===============================================

Dim uniqueAnos, uniqueCorretores
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")
uniqueCorretores = GetUniqueValues("Vendas", "Corretor", "WHERE Corretor IS NOT NULL AND Corretor <> ''")

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
' OBTER DADOS DO CORRETOR POR MÊS (APENAS SE FILTROS ESTIVEREM PREENCHIDOS)
' ===============================================

Dim monthlyData, totalGeralVGV, totalGeralUnidades, totalGeralComissao
Dim infoCorretor, totalVendasCorretor, mediaVGV
Dim diretoriaCorretor, gerenciaCorretor, totalEmpreendimentosCorretor

If filtroAno <> "" And filtroCorretor <> "" Then
    Set monthlyData = Server.CreateObject("Scripting.Dictionary")

    ' Inicializar dados para todos os meses
    For i = 1 To 12
        monthlyData.Add CStr(i), Array(0, 0, 0) ' [VGV, Unidades, Comissao]
    Next

    ' Consulta para obter dados mensais do corretor
    Dim sqlMonthly, rsMonthly
    sqlMonthly = "SELECT " & _
                 "MesVenda, " & _
                 "SUM(ValorUnidade) as TotalVGV, " & _
                 "COUNT(*) as TotalUnidades, " & _
                 "SUM(ValorCorretor) as TotalComissao " & _
                 "FROM Vendas " & _
                 "WHERE Excluido = 0 " & _
                 "AND AnoVenda = " & filtroAno & " " & _
                 "AND Corretor = '" & Replace(filtroCorretor, "'", "''") & "' " & _
                 "GROUP BY MesVenda " & _
                 "ORDER BY MesVenda"

    Set rsMonthly = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsMonthly.Open sqlMonthly, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta mensal: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlMonthly)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados mensais
    totalGeralVGV = 0
    totalGeralUnidades = 0
    totalGeralComissao = 0

    If Not rsMonthly.EOF Then
        Do While Not rsMonthly.EOF
            Dim mes, vgvMes, unidadesMes, comissaoMes
            mes = CStr(rsMonthly("MesVenda"))
            vgvMes = CDbl(rsMonthly("TotalVGV"))
            unidadesMes = CLng(rsMonthly("TotalUnidades"))
            comissaoMes = CDbl(rsMonthly("TotalComissao"))
            
            ' Atualizar dados do mês
            monthlyData(mes) = Array(vgvMes, unidadesMes, comissaoMes)
            
            ' Atualizar totais gerais
            totalGeralVGV = totalGeralVGV + vgvMes
            totalGeralUnidades = totalGeralUnidades + unidadesMes
            totalGeralComissao = totalGeralComissao + comissaoMes
            
            rsMonthly.MoveNext
        Loop
    End If

    If rsMonthly.State = 1 Then rsMonthly.Close
    Set rsMonthly = Nothing

    ' Obter informações básicas do corretor
    Dim sqlInfo, rsInfo
    sqlInfo = "SELECT TOP 1 Diretoria, Gerencia FROM Vendas WHERE Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
    
    Set rsInfo = Server.CreateObject("ADODB.Recordset")
    rsInfo.Open sqlInfo, connSales
    
    If Not rsInfo.EOF Then
        diretoriaCorretor = rsInfo("Diretoria")
        gerenciaCorretor = rsInfo("Gerencia")
    Else
        diretoriaCorretor = "N/A"
        gerenciaCorretor = "N/A"
    End If
    
    If rsInfo.State = 1 Then rsInfo.Close
    Set rsInfo = Nothing

    ' Contar empreendimentos distintos usando abordagem compatível com JET
    Dim sqlEmp, rsEmp
    sqlEmp = "SELECT COUNT(*) as TotalEmp FROM (" & _
             "SELECT DISTINCT NomeEmpreendimento " & _
             "FROM Vendas " & _
             "WHERE Corretor = '" & Replace(filtroCorretor, "'", "''") & "'" & _
             ") as Empreendimentos"
    
    Set rsEmp = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsEmp.Open sqlEmp, connSales
    
    If Err.Number = 0 And Not rsEmp.EOF Then
        totalEmpreendimentosCorretor = rsEmp("TotalEmp")
    Else
        ' Fallback: contar de forma mais simples
        sqlEmp = "SELECT COUNT(*) as TotalEmp FROM Vendas WHERE Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
        Set rsEmp = connSales.Execute(sqlEmp)
        If Not rsEmp.EOF Then
            totalEmpreendimentosCorretor = rsEmp("TotalEmp")
        Else
            totalEmpreendimentosCorretor = 0
        End If
    End If
    On Error GoTo 0
    
    If Not rsEmp Is Nothing Then
        If rsEmp.State = 1 Then rsEmp.Close
        Set rsEmp = Nothing
    End If

    ' Contar total de vendas do corretor
    Dim sqlTotal, rsTotal
    sqlTotal = "SELECT COUNT(*) as TotalVendas FROM Vendas WHERE Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
    
    Set rsTotal = Server.CreateObject("ADODB.Recordset")
    rsTotal.Open sqlTotal, connSales
    
    If Not rsTotal.EOF Then
        totalVendasCorretor = rsTotal("TotalVendas")
    Else
        totalVendasCorretor = 0
    End If
    
    If rsTotal.State = 1 Then rsTotal.Close
    Set rsTotal = Nothing

    ' Calcular média de VGV por venda
    If totalVendasCorretor > 0 Then
        mediaVGV = totalGeralVGV / totalVendasCorretor
    Else
        mediaVGV = 0
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório do Corretor</title>
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
            max-width: 1400px;
            margin: 0 auto;
        }
        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 20px;
            border-radius: 8px;
            font-size: 1rem;
            margin-bottom: 10px;
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .kpi-card h5 {
            font-size: 1.1rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            line-height: 1.2;
        }
        .kpi-card i {
            font-size: 1.8rem;
            margin-bottom: 10px;
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
        }
        .table th {
            background-color: #800000;
            color: white;
        }
        .text-right-v { text-align: right; }
        .text-center-v { text-align: center; }
        .corretor-info {
            background-color: #F0ECEC;
            color: #000;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .alert-warning {
            background-color: #fff3cd;
            border-color: #ffeaa7;
            color: #856404;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
    </style>
    <style>
    body {
        /* Define a escala de 0.8 (80%) */
        transform: scale(0.7); 
        
        /* Define o ponto de origem para o canto superior esquerdo */
        transform-origin: 0 0; 
        
        /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
        /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
        width: calc(100% / 0.7); 
    }
</style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-user-tie"></i> SGVendas - Relatório do Corretor
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
                    
                    <div class="col-md-6">
                        <div class="filter-label">Corretor</div>
                        <select class="form-select" name="corretor" id="corretorFilter" required>
                            <option value="">Selecione o corretor</option>
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
                    
                    <div class="col-md-2">
                        <div class="filter-label">&nbsp;</div>
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Relatório
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Or filtroCorretor = "" Then %>
            <div class="alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano e um corretor para visualizar o relatório.
            </div>
        <% Else %>
        
        <!-- Informações do Corretor -->
        <div class="corretor-info">
            <div class="row">
                <div class="col-md-12">
                    <h3 style="color: #800000;"><%= filtroCorretor %></h3>
                    <div class="row">
                        <div class="col-md-4">
                            <strong>Diretoria:</strong> <%= diretoriaCorretor %>
                        </div>
                        <div class="col-md-4">
                            <strong>Gerência:</strong> <%= gerenciaCorretor %>
                        </div>
                        <div class="col-md-4">
                            <strong>Empreendimentos Atuados:</strong> <%= totalEmpreendimentosCorretor %>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- KPIs Principais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-handshake"></i>
                    <h5>Total VGV <%= filtroAno %></h5>
                    <p>R$ <%= FormatNumber(totalGeralVGV, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <i class="fas fa-home"></i>
                    <h5>Unidades Vendidas</h5>
                    <p><%= totalGeralUnidades %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-money-bill-wave"></i>
                    <h5>Total Comissões</h5>
                    <p>R$ <%= FormatNumber(totalGeralComissao, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-chart-line"></i>
                    <h5>Média por Venda</h5>
                    <p>R$ <%= FormatNumber(mediaVGV, 2) %></p>
                </div>
            </div>
        </div>

        <!-- Tabela de Desempenho Mensal -->
        <div class="card-kpi mt-4">
            <h3 class="text-dark mb-4">Desempenho Mensal - Ano <%= filtroAno %></h3>
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th class="text-center-v">Mês</th>
                            <th class="text-right-v">VGV (R$)</th>
                            <th class="text-center-v">Unidades Vendidas</th>
                            <th class="text-right-v">Comissão (R$)</th>
                            <th class="text-right-v">VGV Média (R$)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        For i = 1 To 12
                            'Dim mesData, vgvMes, unidadesMes, comissaoMes, vgvMediaMes
                            mesData = monthlyData(CStr(i))
                            vgvMes = mesData(0)
                            unidadesMes = mesData(1)
                            comissaoMes = mesData(2)
                            
                            If unidadesMes > 0 Then
                                vgvMediaMes = vgvMes / unidadesMes
                            Else
                                vgvMediaMes = 0
                            End If
                        %>
                        <tr>
                            <td class="text-center-v"><strong><%= arrMesesNome(i) %></strong></td>
                            <td class="text-right-v"><%= FormatNumber(vgvMes, 2) %></td>
                            <td class="text-center-v"><%= unidadesMes %></td>
                            <td class="text-right-v"><%= FormatNumber(comissaoMes, 2) %></td>
                            <td class="text-right-v"><%= FormatNumber(vgvMediaMes, 2) %></td>
                        </tr>
                        <%
                        Next
                        %>
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #800000; color: white;">
                            <td class="text-center-v"><strong>TOTAL GERAL</strong></td>
                            <td class="text-right-v"><strong>R$ <%= FormatNumber(totalGeralVGV, 2) %></strong></td>
                            <td class="text-center-v"><strong><%= totalGeralUnidades %></strong></td>
                            <td class="text-right-v"><strong>R$ <%= FormatNumber(totalGeralComissao, 2) %></strong></td>
                            <td class="text-right-v"><strong>R$ <%= FormatNumber(mediaVGV, 2) %></strong></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <!-- Resumo Estatístico -->
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">Resumo Estatístico</h4>
                    <div class="row">
                        <div class="col-md-6">
                            <p><strong>Melhor Mês (VGV):</strong><br>
                            <%
                            Dim melhorMes, melhorMesVGV
                            melhorMesVGV = 0
                            For i = 1 To 12
                                If monthlyData(CStr(i))(0) > melhorMesVGV Then
                                    melhorMesVGV = monthlyData(CStr(i))(0)
                                    melhorMes = i
                                End If
                            Next
                            If melhorMesVGV > 0 Then
                                Response.Write arrMesesNome(melhorMes) & " - R$ " & FormatNumber(melhorMesVGV, 2)
                            Else
                                Response.Write "N/A"
                            End If
                            %>
                            </p>
                        </div>
                        <div class="col-md-6">
                            <p><strong>Mês com Mais Unidades:</strong><br>
                            <%
                            Dim melhorMesUnidades, maxUnidades
                            maxUnidades = 0
                            For i = 1 To 12
                                If monthlyData(CStr(i))(1) > maxUnidades Then
                                    maxUnidades = monthlyData(CStr(i))(1)
                                    melhorMesUnidades = i
                                End If
                            Next
                            If maxUnidades > 0 Then
                                Response.Write arrMesesNome(melhorMesUnidades) & " - " & maxUnidades & " un."
                            Else
                                Response.Write "N/A"
                            End If
                            %>
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card-kpi">
                    <h4 class="text-dark">Performance</h4>
                    <p><strong>VGV Mensal Média:</strong> R$ <%= FormatNumber(totalGeralVGV / 12, 2) %></p>
                    <p><strong>Unidades Mensais Médias:</strong> <%= FormatNumber(totalGeralUnidades / 12, 1) %></p>
                    <p><strong>Comissão Mensal Média:</strong> R$ <%= FormatNumber(totalGeralComissao / 12, 2) %></p>
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