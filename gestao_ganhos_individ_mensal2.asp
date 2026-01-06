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
' FUNÇÃO PARA CALCULAR MÉDIA MENSAL (TOTAL/12)
' ===============================================
Function CalcularMediaMensal(totalGeral)
    Dim mediaMensal
    
    mediaMensal = 0
    
    ' Calcular média dividindo por 12 meses
    If totalGeral > 0 Then
        mediaMensal = totalGeral / 12
    End If
    
    CalcularMediaMensal = mediaMensal
End Function

' ===============================================
' POPULAR OS SELECTS DO FORMULÁRIO
' ===============================================

Dim uniqueAnos
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")

' ===============================================
' OBTER DADOS DA CONSULTA (APENAS SE ANO ESTIVER PREENCHIDO)
' ===============================================

Dim dadosPessoas, totalGeralVTotal
Set dadosPessoas = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    ' Construir consulta SQL para dados principais
    Dim sqlConsulta, rsConsulta
    sqlConsulta = "SELECT Vendas.AnoVenda, VENDA_TEMP.Nome, Sum(VENDA_TEMP.VBruto) AS SomaDeVTotal " & _
                  "FROM VENDA_TEMP INNER JOIN Vendas ON VENDA_TEMP.ID_Venda = Vendas.Id " & _
                  "WHERE Vendas.AnoVenda = " & filtroAno & " " & _
                  "GROUP BY Vendas.AnoVenda, VENDA_TEMP.Nome " & _
                  "ORDER BY VENDA_TEMP.Nome, Sum(VENDA_TEMP.VBruto) DESC"

    Set rsConsulta = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsConsulta.Open sqlConsulta, connSales

    If Err.Number <> 0 Then
        Response.Write "Erro na consulta: " & Err.Description & "<br>"
        Response.Write "SQL: " & Server.HTMLEncode(sqlConsulta)
        Response.End
    End If
    On Error GoTo 0

    ' Processar dados principais
    totalGeralVTotal = 0
    
    If Not rsConsulta.EOF Then
        Do While Not rsConsulta.EOF
            Dim nomePessoa, vTotalPessoa
            nomePessoa = Trim(CStr(rsConsulta("Nome")))
            vTotalPessoa = CDbl(rsConsulta("SomaDeVTotal"))
            
            ' Adicionar pessoa ao dicionário
            If Not dadosPessoas.Exists(nomePessoa) Then
                dadosPessoas.Add nomePessoa, vTotalPessoa
            Else
                ' Se pessoa já existe, somar ao valor existente
                dadosPessoas(nomePessoa) = dadosPessoas(nomePessoa) + vTotalPessoa
            End If
            
            ' Atualizar total geral
            totalGeralVTotal = totalGeralVTotal + vTotalPessoa
            
            rsConsulta.MoveNext
        Loop
    End If

    If rsConsulta.State = 1 Then rsConsulta.Close
    Set rsConsulta = Nothing
End If
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Listagem Simples por Média Mensal</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <!-- DataTables CSS -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css">
    <style>
        body {
            background-color: #f8f9fa;
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        .header-title {
            color: #800000;
            border-bottom: 2px solid #800000;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        .filter-container {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            border: 1px solid #dee2e6;
        }
        .card-summary {
            background: linear-gradient(135deg, #800000, #a00000);
            color: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .table th {
            background-color: #800000;
            color: white;
            cursor: pointer;
        }
        .table th:hover {
            background-color: #a00000;
        }
        .table-striped tbody tr:nth-of-type(odd) {
            background-color: rgba(128, 0, 0, 0.05);
        }
        .badge-position {
            font-size: 0.8rem;
            min-width: 50px;
            text-align: center;
        }
        .text-success-dark {
            color: #28a745;
            font-weight: bold;
        }
        .text-primary-dark {
            color: #007bff;
            font-weight: bold;
        }
        .posicao-1 {
            background-color: #fff3cd !important;
            border-left: 4px solid #ffc107 !important;
        }
        .posicao-2 {
            background-color: #e9ecef !important;
            border-left: 4px solid #6c757d !important;
        }
        .posicao-3 {
            background-color: #f8d7da !important;
            border-left: 4px solid #dc3545 !important;
        }
        .posicao-top {
            background-color: #d1ecf1 !important;
            border-left: 4px solid #17a2b8 !important;
        }
        .btn-export {
            margin-right: 5px;
            margin-bottom: 5px;
        }
        .total-row {
            background-color: #800000 !important;
            color: white !important;
            font-weight: bold;
        }
        .sorting::after {
            content: "↕";
            float: right;
            opacity: 0.5;
        }
        .sorting_asc::after {
            content: "↑";
            float: right;
            opacity: 1;
        }
        .sorting_desc::after {
            content: "↓";
            float: right;
            opacity: 1;
        }
        .dataTables_wrapper .dataTables_length,
        .dataTables_wrapper .dataTables_filter,
        .dataTables_wrapper .dataTables_info,
        .dataTables_wrapper .dataTables_paginate {
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2 class="header-title text-center">
            <i class="fas fa-chart-line"></i> SGVendas - Listagem por Média Mensal
        </h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-8">
                        <label class="form-label fw-bold">Ano</label>
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
                        <label class="form-label">&nbsp;</label>
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-search"></i> Buscar
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <% If filtroAno = "" Then %>
            <div class="alert alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar a listagem.
            </div>
        <% Else %>
        
        <!-- Resumo -->
        <div class="card-summary">
            <div class="row">
                <div class="col-md-4 text-center">
                    <h5><i class="fas fa-users"></i> Total de Pessoas</h5>
                    <h3><%= dadosPessoas.Count %></h3>
                </div>
                <div class="col-md-4 text-center">
                    <h5><i class="fas fa-money-bill-wave"></i> Total Geral <%= filtroAno %></h5>
                    <h3><%= FormatNumber(totalGeralVTotal, 2) %></h3>
                </div>
                <div class="col-md-4 text-center">
                    <h5><i class="fas fa-calculator"></i> Média Geral Mensal</h5>
                    <h3><%= FormatNumber(CalcularMediaMensal(totalGeralVTotal), 2) %></h3>
                </div>
            </div>
        </div>

        <!-- Botões de Exportação -->
        <div class="mb-3">
            <button class="btn btn-success btn-export" onclick="exportToExcel()">
                <i class="fas fa-file-excel"></i> Exportar Excel
            </button>
            <button class="btn btn-danger btn-export" onclick="window.print()">
                <i class="fas fa-print"></i> Imprimir
            </button>
        </div>

        <!-- Tabela Simples -->
        <div class="table-responsive">
            <table id="tabelaSimples" class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th class="text-center">Posição</th>
                        <th class="text-left">Nome</th>
                        <th class="text-right">Total Recebido (R$)</th>
                        <th class="text-right">Média Mensal (R$)</th>
                        <th class="text-center">% do Total</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If dadosPessoas.Count > 0 Then
                        ' Criar array para ordenação por média mensal
                        Dim arrOrdenado(), pessoaKey, i, j
                        ReDim arrOrdenado(dadosPessoas.Count - 1, 2)
                        
                        i = 0
                        For Each pessoaKey In dadosPessoas.Keys
                            arrOrdenado(i, 0) = pessoaKey
                            arrOrdenado(i, 1) = dadosPessoas(pessoaKey)
                            arrOrdenado(i, 2) = CalcularMediaMensal(dadosPessoas(pessoaKey))
                            i = i + 1
                        Next
                        
                        ' Ordenar por média mensal (DECRESCENTE - maior média primeiro)
                        For i = 0 To UBound(arrOrdenado, 1)
                            For j = i + 1 To UBound(arrOrdenado, 1)
                                If arrOrdenado(j, 2) > arrOrdenado(i, 2) Then
                                    ' Trocar posições
                                    Dim tempNome, tempTotal, tempMedia
                                    tempNome = arrOrdenado(i, 0)
                                    tempTotal = arrOrdenado(i, 1)
                                    tempMedia = arrOrdenado(i, 2)
                                    
                                    arrOrdenado(i, 0) = arrOrdenado(j, 0)
                                    arrOrdenado(i, 1) = arrOrdenado(j, 1)
                                    arrOrdenado(i, 2) = arrOrdenado(j, 2)
                                    
                                    arrOrdenado(j, 0) = tempNome
                                    arrOrdenado(j, 1) = tempTotal
                                    arrOrdenado(j, 2) = tempMedia
                                End If
                            Next
                        Next
                        
                        ' Exibir dados ordenados
                        For i = 0 To UBound(arrOrdenado, 1)
                            Dim nome, total, media, percentual, classePosicao, mediaNum, totalNum
                            nome = arrOrdenado(i, 0)
                            total = arrOrdenado(i, 1)
                            media = arrOrdenado(i, 2)
                            percentual = (total / totalGeralVTotal) * 100
                            
                            ' Format numbers without thousands separator for DataTables sorting
                            mediaNum = Replace(FormatNumber(media, 2), ".", "")
                            mediaNum = Replace(mediaNum, ",", ".")
                            totalNum = Replace(FormatNumber(total, 2), ".", "")
                            totalNum = Replace(totalNum, ",", ".")
                            
                            ' Determinar classe da posição
                            Select Case i + 1
                                Case 1: classePosicao = "posicao-1"
                                Case 2: classePosicao = "posicao-2"
                                Case 3: classePosicao = "posicao-3"
                                Case Else
                                    If i + 1 <= 10 Then
                                        classePosicao = "posicao-top"
                                    Else
                                        classePosicao = ""
                                    End If
                            End Select
                    %>
                    <tr class="<%= classePosicao %>">
                        <td class="text-center align-middle" data-order="<%= i + 1 %>">
                            <span class="badge bg-secondary badge-position">
                                <%= i + 1 %>º
                            </span>
                        </td>
                        <td class="text-left align-middle"><strong><%= nome %></strong></td>
                        <td class="text-right align-middle text-success-dark" data-order="<%= totalNum %>">
                            <strong><%= FormatNumber(total, 2) %></strong>
                        </td>
                        <td class="text-right align-middle text-primary-dark" data-order="<%= mediaNum %>">
                            <strong><%= FormatNumber(media, 2) %></strong>
                        </td>
                        <td class="text-center align-middle" data-order="<%= percentual %>">
                            <span class="badge bg-info">
                                <%= FormatNumber(percentual, 2) %>%
                            </span>
                        </td>
                    </tr>
                    <%
                        Next
                    Else
                    %>
                    <tr>
                        <td colspan="5" class="text-center py-4">
                            <div class="alert alert-info">
                                <i class="fas fa-info-circle"></i> Nenhum dado encontrado para o ano <%= filtroAno %>.
                            </div>
                        </td>
                    </tr>
                    <%
                    End If
                    %>
                </tbody>
                <tfoot>
                    <tr class="total-row">
                        <td class="text-center" colspan="2">
                            <strong>TOTAL GERAL - <%= filtroAno %></strong>
                        </td>
                        <td class="text-right">
                            <strong><%= FormatNumber(totalGeralVTotal, 2) %></strong>
                        </td>
                        <td class="text-right">
                            <strong><%= FormatNumber(CalcularMediaMensal(totalGeralVTotal), 2) %></strong>
                        </td>
                        <td class="text-center">
                            <strong>100%</strong>
                        </td>
                    </tr>
                </tfoot>
            </table>
        </div>

        <!-- Estatísticas -->
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-info text-white">
                        <h5 class="mb-0"><i class="fas fa-trophy"></i> Top 5 - Maiores Médias Mensais</h5>
                    </div>
                    <div class="card-body">
                        <%
                        If dadosPessoas.Count > 0 Then
                            For i = 0 To 4
                                If i <= UBound(arrOrdenado, 1) Then
                                    Dim nomeTop, mediaTop, percentualTop
                                    nomeTop = arrOrdenado(i, 0)
                                    mediaTop = arrOrdenado(i, 2)
                                    percentualTop = (arrOrdenado(i, 1) / totalGeralVTotal) * 100
                        %>
                        <div class="d-flex justify-content-between align-items-center mb-2 pb-2 border-bottom">
                            <div>
                                <span class="badge 
                                <% 
                                Select Case i + 1
                                    Case 1: Response.Write "bg-warning text-dark"
                                    Case 2: Response.Write "bg-secondary"
                                    Case 3: Response.Write "bg-danger"
                                    Case Else: Response.Write "bg-info"
                                End Select 
                                %>">
                                    <%= i + 1 %>º
                                </span>
                                <strong><%= nomeTop %></strong>
                            </div>
                            <div>
                                <span class="text-primary fw-bold"><%= FormatNumber(mediaTop, 2) %></span>
                                <small class="text-muted ms-2">(<%= FormatNumber(percentualTop, 2) %>%)</small>
                            </div>
                        </div>
                        <%
                                End If
                            Next
                        Else
                        %>
                        <p class="text-center text-muted">Nenhum dado disponível</p>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-success text-white">
                        <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Estatísticas</h5>
                    </div>
                    <div class="card-body">
                        <%
                        If dadosPessoas.Count > 0 Then
                            ' Calcular estatísticas
                            Dim somaMedias, maiorMedia, menorMedia, countAcimaMediaGeral
                            somaMedias = 0
                            maiorMedia = 0
                            menorMedia = 999999999
                            countAcimaMediaGeral = 0
                            
                            For i = 0 To UBound(arrOrdenado, 1)
                                Dim mediaAtual
                                mediaAtual = arrOrdenado(i, 2)
                                somaMedias = somaMedias + mediaAtual
                                
                                If mediaAtual > maiorMedia Then maiorMedia = mediaAtual
                                If mediaAtual < menorMedia Then menorMedia = mediaAtual
                                
                                If mediaAtual > CalcularMediaMensal(totalGeralVTotal) Then
                                    countAcimaMediaGeral = countAcimaMediaGeral + 1
                                End If
                            Next
                            
                            Dim mediaDasMedias
                            mediaDasMedias = somaMedias / dadosPessoas.Count
                        %>
                        <div class="row">
                            <div class="col-6">
                                <p class="mb-1"><strong>Maior Média:</strong></p>
                                <p class="mb-1"><strong>Menor Média:</strong></p>
                                <p class="mb-1"><strong>Média das Médias:</strong></p>
                                <p class="mb-1"><strong>Acima da Média Geral:</strong></p>
                            </div>
                            <div class="col-6 text-end">
                                <p class="mb-1 text-success"><strong><%= FormatNumber(maiorMedia, 2) %></strong></p>
                                <p class="mb-1 text-danger"><strong><%= FormatNumber(menorMedia, 2) %></strong></p>
                                <p class="mb-1 text-primary"><strong><%= FormatNumber(mediaDasMedias, 2) %></strong></p>
                                <p class="mb-1"><strong><%= countAcimaMediaGeral %> pessoas</strong></p>
                            </div>
                        </div>
                        <%
                        Else
                        %>
                        <p class="text-center text-muted">Nenhuma estatística disponível</p>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>

        <% End If %>
    </div>

    <!-- Scripts -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <!-- DataTables -->
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
    $(document).ready(function() {
        // Inicializar DataTable com configurações específicas para ordenação numérica
        var table = $('#tabelaSimples').DataTable({
            dom: '<"row"<"col-md-6"l><"col-md-6"f>>rt<"row"<"col-md-6"i><"col-md-6"p>>',
            language: {
                url: '//cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json',
                decimal: ',',
                thousands: '.'
            },
            pageLength: 25,
            order: [[3, 'desc']], // Ordenar inicialmente por média mensal (coluna 3) decrescente
            columnDefs: [
                {
                    targets: 0, // Coluna Posição
                    orderable: true,
                    type: 'num',
                    className: 'text-center'
                },
                {
                    targets: 1, // Coluna Nome
                    orderable: true,
                    type: 'string',
                    className: 'text-left'
                },
                {
                    targets: 2, // Coluna Total Recebido
                    orderable: true,
                    type: 'num',
                    className: 'text-right'
                },
                {
                    targets: 3, // Coluna Média Mensal - CRÍTICO: garantir ordenação numérica
                    orderable: true,
                    type: 'num',
                    className: 'text-right',
                    render: function(data, type, row) {
                        // Para ordenação, retornar o valor numérico
                        if (type === 'sort' || type === 'type') {
                            // Extrair número do formato brasileiro
                            return parseFloat(data.replace(/\./g, '').replace(',', '.'));
                        }
                        // Para exibição, retornar o valor formatado
                        return data;
                    }
                },
                {
                    targets: 4, // Coluna % do Total
                    orderable: true,
                    type: 'num',
                    className: 'text-center'
                }
            ],
            // Remover formatação para ordenação
            columnDefs: [
                { 
                    targets: [2, 3, 4], 
                    render: function(data, type, row) {
                        if (type === 'sort') {
                            // Para ordenação, retornar número puro
                            return parseFloat(data.replace(/[^\d,.-]/g, '').replace('.', '').replace(',', '.'));
                        }
                        // Para exibição, manter o formato original
                        return data;
                    }
                }
            ]
        });
        
        // Garantir que a ordenação inicial funcione corretamente
        table.order([3, 'desc']).draw();
    });

    function exportToExcel() {
        // Criar tabela HTML para exportação
        var table = document.getElementById('tabelaSimples');
        var html = '<table border="1">';
        
        // Cabeçalho
        html += '<thead><tr>';
        html += '<th>Posição</th>';
        html += '<th>Nome</th>';
        html += '<th>Total Recebido (R$)</th>';
        html += '<th>Média Mensal (R$)</th>';
        html += '<th>% do Total</th>';
        html += '</tr></thead>';
        
        // Corpo
        html += '<tbody>';
        var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
        for (var i = 0; i < rows.length; i++) {
            var cells = rows[i].getElementsByTagName('td');
            html += '<tr>';
            for (var j = 0; j < cells.length; j++) {
                // Remover HTML interno (badges, etc.)
                var cellText = cells[j].innerText || cells[j].textContent;
                html += '<td>' + cellText + '</td>';
            }
            html += '</tr>';
        }
        html += '</tbody></table>';
        
        // Criar blob e download
        var blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        var url = window.URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'media_mensal_' + new Date().toISOString().split('T')[0] + '.xls';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }
    </script>
</body>
</html>

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>