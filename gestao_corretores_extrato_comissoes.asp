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
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno
filtroAno = Request.QueryString("ano")

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

Dim uniqueAnos
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "jan"
arrMesesNome(2) = "fev"
arrMesesNome(3) = "mar"
arrMesesNome(4) = "abr"
arrMesesNome(5) = "mai"
arrMesesNome(6) = "jun"
arrMesesNome(7) = "jul"
arrMesesNome(8) = "ago"
arrMesesNome(9) = "set"
arrMesesNome(10) = "out"
arrMesesNome(11) = "nov"
arrMesesNome(12) = "dez"

' ===============================================
' OBTER DADOS DAS COMISSÕES (APENAS SE ANO ESTIVER PREENCHIDO)
' ===============================================

Dim comissoesData, totalGeralComissoes, totalGeralVendas
Set comissoesData = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Consulta para obter dados de comissões por corretor e mês
    Dim sqlComissoes, rsComissoes
    sqlComissoes = "SELECT " & _
                "Corretor, " & _
                "Diretoria, " & _
                "Gerencia, " & _
                "MesVenda, " & _
                "SUM(ValorCorretor) as TotalComissao, " & _
                "COUNT(*) as TotalVendas " & _
                "FROM Vendas " & _
                "WHERE Excluido = 0 " & _
                "AND AnoVenda = " & filtroAno & " " & _
                "GROUP BY Corretor, Diretoria, Gerencia, MesVenda " & _
                "ORDER BY Diretoria, Gerencia, Corretor, MesVenda"

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
    totalGeralVendas = 0

    If Not rsComissoes.EOF Then
        Do While Not rsComissoes.EOF
            Dim corretor, diretoria, gerencia, mes, comissaoMes, vendasMes
            corretor = CStr(rsComissoes("Corretor"))
            diretoria = CStr(rsComissoes("Diretoria"))
            gerencia = CStr(rsComissoes("Gerencia"))
            mes = CStr(rsComissoes("MesVenda"))
            comissaoMes = CDbl(rsComissoes("TotalComissao"))
            vendasMes = CLng(rsComissoes("TotalVendas"))
            
            ' Criar chave única para o corretor
            Dim chaveCorretor
            chaveCorretor = diretoria & "|" & gerencia & "|" & corretor
            
            ' Criar estrutura para o corretor se não existir
            If Not comissoesData.Exists(chaveCorretor) Then
                Dim infoCorretor
                Set infoCorretor = Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "Corretor", corretor
                infoCorretor.Add "Diretoria", diretoria
                infoCorretor.Add "Gerencia", gerencia
                infoCorretor.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                infoCorretor.Add "TotalComissao", 0
                infoCorretor.Add "TotalVendas", 0
                infoCorretor.Add "MediaComissao", 0
                comissoesData.Add chaveCorretor, infoCorretor
            End If
            
            ' Atualizar dados do mês
            Set infoCorretor = comissoesData(chaveCorretor)
            infoCorretor("Meses").Add mes, Array(comissaoMes, vendasMes)
            
            ' Atualizar totais do corretor
            infoCorretor("TotalComissao") = infoCorretor("TotalComissao") + comissaoMes
            infoCorretor("TotalVendas") = infoCorretor("TotalVendas") + vendasMes
            
            ' Atualizar totais gerais
            totalGeralComissoes = totalGeralComissoes + comissaoMes
            totalGeralVendas = totalGeralVendas + vendasMes
            
            rsComissoes.MoveNext
        Loop
    End If

    If rsComissoes.State = 1 Then rsComissoes.Close
    Set rsComissoes = Nothing

    ' Calcular médias para cada corretor
    Dim chave
    For Each chave In comissoesData.Keys
        Set infoCorretor = comissoesData(chave)
        If infoCorretor("TotalVendas") > 0 Then
            infoCorretor("MediaComissao") = infoCorretor("TotalComissao") / infoCorretor("TotalVendas")
        Else
            infoCorretor("MediaComissao") = 0
        End If
    Next
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório de Comissões</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background-color: #f8f9fa;
            padding: 10px;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }
        .container-fluid {
            max-width: 95%;
            margin: 0 auto;
        }
        .filter-container {
            background-color: white;
            padding: 10px;
            border-radius: 5px;
            margin-bottom: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 3px;
            font-size: 12px;
        }
        .table-responsive {
            background-color: white;
            border-radius: 5px;
            font-size: 11px;
        }
        .table {
            margin-bottom: 0;
        }
        .table th {
            background-color: #2c3e50;
            color: white;
            font-weight: bold;
            padding: 6px 4px;
            text-align: center;
            font-size: 10px;
            border: 1px solid #dee2e6;
        }
        .table td {
            padding: 4px 3px;
            border: 1px solid #dee2e6;
            vertical-align: middle;
        }
        .text-right {
            text-align: right;
        }
        .text-center {
            text-align: center;
        }
        .diretoria-header {
            background-color: #34495e !important;
            color: white;
            font-weight: bold;
        }
        .gerencia-header {
            background-color: #7f8c8d !important;
            color: white;
            font-weight: bold;
        }
        .corretor-row {
            background-color: #f8f9fa;
        }
        .corretor-row:hover {
            background-color: #e9ecef;
        }
        .mes-header {
            background-color: #3498db;
            color: white;
        }
        .total-header {
            background-color: #27ae60;
            color: white;
        }
        .media-header {
            background-color: #e67e22;
            color: white;
        }
        .vendas-header {
            background-color: #9b59b6;
            color: white;
        }
        .btn-sm {
            padding: 3px 8px;
            font-size: 11px;
        }
        .form-select-sm {
            padding: 3px 6px;
            font-size: 11px;
        }
        h2 {
            color: #2c3e50;
            font-size: 18px;
            margin: 10px 0;
            text-align: center;
        }
        .total-geral {
            background-color: #2c3e50;
            color: white;
            font-weight: bold;
        }
        .comissao-cell {
            color: #27ae60;
            font-weight: bold;
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
    <div class="container-fluid">
        <h2>SGVendas - Relatório de Comissões</h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row align-items-center">
                    <div class="col-md-3">
                        <div class="filter-label">Ano</div>
                        <select class="form-select form-select-sm" name="ano" id="anoFilter" required>
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
                    <div class="col-md-2">
                        <div class="filter-label">&nbsp;</div>
                        <button type="submit" class="btn btn-primary btn-sm w-100">
                            Gerar Relatório
                        </button>
                    </div>
                    <div class="col-md-7 text-end">
                        <%
                        If filtroAno <> "" Then
                            Response.Write "<small class='text-muted'>Ano " & filtroAno & " | Total Comissões: R$ " & FormatNumber(totalGeralComissoes, 2) & " | Total Vendas: " & totalGeralVendas & "</small>"
                        End If
                        %>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Then %>
            <div class="alert alert-warning text-center">
                Por favor, selecione um ano para visualizar o relatório de comissões.
            </div>
        <% ElseIf comissoesData.Count = 0 Then %>
            <div class="alert alert-info text-center">
                Nenhuma comissão encontrada para o ano <%= filtroAno %>.
            </div>
        <% Else %>
        
        <div class="table-responsive">
            <table class="table table-bordered table-sm">
                <thead>
                    <tr>
                        <th rowspan="2">Diretoria</th>
                        <th rowspan="2">Gerencia</th>
                        <th rowspan="2">Corretor</th>
                        <%
                        ' Cabeçalhos dos meses
                        For i = 1 To 12
                            Response.Write "<th class='mes-header'>" & arrMesesNome(i) & "</th>"
                        Next
                        %>
                        <th class="total-header">Total</th>
                        <th class="media-header">Média</th>
                        <th class="vendas-header">QTD vendas</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim currentDiretoria, currentGerencia, firstRow
                    currentDiretoria = ""
                    currentGerencia = ""
                    firstRow = True
                    
                    Dim arrChaves
                    arrChaves = comissoesData.Keys
                    
                    ' Ordenar por diretoria, gerencia e corretor
                    For i = 0 To UBound(arrChaves)
                        For j = i + 1 To UBound(arrChaves)
                            If arrChaves(j) < arrChaves(i) Then
                                Dim tempChave
                                tempChave = arrChaves(i)
                                arrChaves(i) = arrChaves(j)
                                arrChaves(j) = tempChave
                            End If
                        Next
                    Next
                    
                    For Each chave In arrChaves
                        Set infoCorretor = comissoesData(chave)
                        Dim showDiretoria, showGerencia
                        showDiretoria = (infoCorretor("Diretoria") <> currentDiretoria)
                        showGerencia = (infoCorretor("Gerencia") <> currentGerencia Or showDiretoria)
                        
                        If showDiretoria Then
                            currentDiretoria = infoCorretor("Diretoria")
                            currentGerencia = ""
                        End If
                        
                        If showGerencia Then
                            currentGerencia = infoCorretor("Gerencia")
                        End If
                    %>
                    
                    <% If showDiretoria Then %>
                    <tr class="diretoria-header">
                        <td colspan="3"><strong><%= infoCorretor("Diretoria") %></strong></td>
                        <%
                        For i = 1 To 12
                            Response.Write "<td></td>"
                        Next
                        %>
                        <td colspan="3"></td>
                    </tr>
                    <% End If %>
                    
                    <% If showGerencia Then %>
                    <tr class="gerencia-header">
                        <td colspan="2"><strong><%= infoCorretor("Gerencia") %></strong></td>
                        <td></td>
                        <%
                        For i = 1 To 12
                            Response.Write "<td></td>"
                        Next
                        %>
                        <td colspan="3"></td>
                    </tr>
                    <% End If %>
                    
                    <tr class="corretor-row">
                        <td></td>
                        <td></td>
                        <td><strong><%= infoCorretor("Corretor") %></strong></td>
                        <%
                        Dim mesesCorretor
                        Set mesesCorretor = infoCorretor("Meses")
                        
                        ' Dados dos meses - COMISSÕES
                        For i = 1 To 12
                            Dim mesKey
                            mesKey = CStr(i)
                            If mesesCorretor.Exists(mesKey) Then
                                Dim dadosMes
                                dadosMes = mesesCorretor(mesKey)
                                Response.Write "<td class='text-right comissao-cell'>R$ " & FormatNumber(dadosMes(0), 2) & "</td>"
                            Else
                                Response.Write "<td class='text-center'>-</td>"
                            End If
                        Next
                        %>
                        
                        <!-- Totais do corretor - COMISSÕES -->
                        <td class="text-right comissao-cell"><strong>R$ <%= FormatNumber(infoCorretor("TotalComissao"), 2) %></strong></td>
                        <td class="text-right"><strong>R$ <%= FormatNumber(infoCorretor("MediaComissao"), 2) %></strong></td>
                        <td class="text-center"><strong><%= infoCorretor("TotalVendas") %></strong></td>
                    </tr>
                    <%
                    Next
                    %>
                    
                    <!-- Total Geral -->
                    <tr class="total-geral">
                        <td colspan="3" class="text-center"><strong>TOTAL GERAL</strong></td>
                        <%
                        ' Totais por mês - COMISSÕES
                        For i = 1 To 12
                            Dim totalMes
                            totalMes = 0
                            For Each chave In comissoesData.Keys
                                Set mesesCorretor = comissoesData(chave)("Meses")
                                If mesesCorretor.Exists(CStr(i)) Then
                                    totalMes = totalMes + mesesCorretor(CStr(i))(0)
                                End If
                            Next
                            Response.Write "<td class='text-right comissao-cell'><strong>R$ " & FormatNumber(totalMes, 2) & "</strong></td>"
                        Next
                        %>
                        <td class="text-right comissao-cell"><strong>R$ <%= FormatNumber(totalGeralComissoes, 2) %></strong></td>
                        <td class="text-right">
                            <strong>
                                <% 
                                If totalGeralVendas > 0 Then 
                                    Response.Write "R$ " & FormatNumber(totalGeralComissoes / totalGeralVendas, 2)
                                Else 
                                    Response.Write "R$ 0,00"
                                End If 
                                %>
                            </strong>
                        </td>
                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- Resumo -->
        <div class="row mt-3">
            <div class="col-md-12">
                <div class="card">
                    <div class="card-body py-2">
                        <div class="row">
                            <div class="col-md-3">
                                <small><strong>Total de Corretores:</strong> <%= comissoesData.Count %></small>
                            </div>
                            <div class="col-md-3">
                                <small><strong>Comissão Média por Corretor:</strong> 
                                <% 
                                If comissoesData.Count > 0 Then 
                                    Response.Write "R$ " & FormatNumber(totalGeralComissoes / comissoesData.Count, 2)
                                Else 
                                    Response.Write "R$ 0,00"
                                End If 
                                %>
                                </small>
                            </div>
                            <div class="col-md-3">
                                <small><strong>Vendas Médias por Corretor:</strong> 
                                <% 
                                If comissoesData.Count > 0 Then 
                                    Response.Write FormatNumber(totalGeralVendas / comissoesData.Count, 1)
                                Else 
                                    Response.Write "0"
                                End If 
                                %>
                                </small>
                            </div>
                            <div class="col-md-3">
                                <small><strong>Comissão Média por Venda:</strong> 
                                <% 
                                If totalGeralVendas > 0 Then 
                                    Response.Write "R$ " & FormatNumber(totalGeralComissoes / totalGeralVendas, 2)
                                Else 
                                    Response.Write "R$ 0,00"
                                End If 
                                %>
                                </small>
                            </div>
                        </div>
                    </div>
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