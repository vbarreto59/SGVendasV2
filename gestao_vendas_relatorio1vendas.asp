<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->


<%
' ===============================================
' FUNÇÕES GLOBAIS - v6
' ===============================================

Function RemoverNumeros(texto)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    RemoverNumeros = Trim(Replace(regex.Replace(texto, ""), " ", " "))
End Function

Function SortDictionaryByValue(dict, valueKey)
    Dim arrKeys, i, j, temp
    If dict.Count > 0 Then
        arrKeys = dict.Keys
        For i = 0 To UBound(arrKeys)
            For j = i + 1 To UBound(arrKeys)
                If dict(arrKeys(i))(valueKey) < dict(arrKeys(j))(valueKey) Then
                    temp = arrKeys(i)
                    arrKeys(i) = arrKeys(j)
                    arrKeys(j) = temp
                End If
            Next
        Next
    Else
        SortDictionaryByValue = Array()
        Exit Function
    End If
    SortDictionaryByValue = arrKeys
End Function

Function SortDictionaryByKey(dict)
    Dim arrKeys, i, j, temp
    If dict.Count > 0 Then
        arrKeys = dict.Keys
        For i = 0 To UBound(arrKeys)
            For j = i + 1 To UBound(arrKeys)
                If arrKeys(i) < arrKeys(j) Then
                    temp = arrKeys(i)
                    arrKeys(i) = arrKeys(j)
                    arrKeys(j) = temp
                End If
            Next
        Next
    Else
        SortDictionaryByKey = Array()
        Exit Function
    End If
    SortDictionaryByKey = arrKeys
End Function

Function mesNome(numMes)
    Dim meses
    meses = Array("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", _
                                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
    If numMes >= 1 And numMes <= 12 Then
        mesNome = meses(CInt(numMes))
    Else
        mesNome = numMes
    End If
End Function

' Função para obter a paleta de cores por ano
Function GetYearColorPalette(year)
        Dim palette
        Set palette = Server.CreateObject("Scripting.Dictionary")
          
        Select Case CInt(year)
            Case 2024
                    ' Azul profissional
                    palette.Add "primary", "#1565C0"
                    palette.Add "secondary", "#5E92F3"
                    palette.Add "accent", "#003C8F"
                    palette.Add "light", "#E3F2FD"
            Case 2025
                    ' Verde corporativo
                    palette.Add "primary", "#2E7D32"
                    palette.Add "secondary", "#66BB6A"
                    palette.Add "accent", "#005005"
                    palette.Add "light", "#E8F5E9"
            Case 2026
                    ' Roxo elegante
                    palette.Add "primary", "#6A1B9A"
                    palette.Add "secondary", "#9C4DCC"
                    palette.Add "accent", "#38006B"
                    palette.Add "light", "#F3E5F5"
            Case 2027
                    ' Vermelho vibrante
                    palette.Add "primary", "#C62828"
                    palette.Add "secondary", "#EF5350"
                    palette.Add "accent", "#8E0000"
                    palette.Add "light", "#FFEBEE"
            Case 2028
                    ' Laranja energético
                    palette.Add "primary", "#EF6C00"
                    palette.Add "secondary", "#FFA726"
                    palette.Add "accent", "#B53D00"
                    palette.Add "light", "#FFF3E0"
            Case 2029
                    ' Azul-turquesa
                    palette.Add "primary", "#00897B"
                    palette.Add "secondary", "#4DB6AC"
                    palette.Add "accent", "#00695C"
                    palette.Add "light", "#E0F2F1"
            Case 2030
                    ' Vermelho vinho
                    palette.Add "primary", "#8E24AA"
                    palette.Add "secondary", "#BA68C8"
                    palette.Add "accent", "#5C007A"
                    palette.Add "light", "#F3E5F5"
            Case Else
                    ' Padrão (verde escuro)
                    palette.Add "primary", "#004d40"
                    palette.Add "secondary", "#39796b"
                    palette.Add "accent", "#00251a"
                    palette.Add "light", "#e8f5e9"
        End Select
          
        Set GetYearColorPalette = palette
End Function

' ===============================================
' PROCESSAMENTO PRINCIPAL
' ===============================================

Dim conn, rs, sql, sqlWhere
Dim filtroAno, filtroMes, filtroTrimestre, filtroDiretoria, filtroGerencia, filtroCorretor
Dim reportData, rankingPorSemestre, rankingPorTrimestre, diretoriaRanking, gerenciaRanking

' Inicializar objetos de dados
Set reportData = Server.CreateObject("Scripting.Dictionary")
Set rankingPorSemestre = Server.CreateObject("Scripting.Dictionary")
Set rankingPorTrimestre = Server.CreateObject("Scripting.Dictionary")
Set diretoriaRanking = Server.CreateObject("Scripting.Dictionary")
Set gerenciaRanking = Server.CreateObject("Scripting.Dictionary")

' Obter filtros da query string
filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")

' Construir SQL WHERE
sqlWhere = " WHERE Vendas.Excluido = 0 "
If filtroAno <> "" Then sqlWhere = sqlWhere & " AND Vendas.AnoVenda = " & filtroAno
If filtroMes <> "" Then sqlWhere = sqlWhere & " AND Vendas.MesVenda = " & filtroMes
If filtroTrimestre <> "" Then sqlWhere = sqlWhere & " AND Vendas.Trimestre = " & filtroTrimestre
If filtroDiretoria <> "" Then sqlWhere = sqlWhere & " AND Vendas.Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
If filtroGerencia <> "" Then sqlWhere = sqlWhere & " AND Vendas.Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
If filtroCorretor <> "" Then sqlWhere = sqlWhere & " AND Vendas.Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"

' Consulta principal
sql = "SELECT AnoVenda, MesVenda, Trimestre, Corretor, Diretoria, Gerencia, ValorUnidade, ValorComissaoGeral FROM Vendas" & sqlWhere & " ORDER BY AnoVenda DESC, MesVenda DESC, Trimestre DESC;"

' Executar consulta
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

' Processar dados
If Not rs.EOF Then
    Dim anoKey, semestreKey, trimestreKey, mesKey, corretorKey
    Dim valorUnidade, valorComissao
      
    Do While Not rs.EOF
        anoKey = CStr(rs("AnoVenda"))
        mesKey = CStr(rs("MesVenda"))
        trimestreKey = CStr(rs("Trimestre"))
        corretorKey = CStr(rs("Corretor"))
        valorUnidade = CDbl(rs("ValorUnidade"))
        valorComissao = CDbl(rs("ValorComissaoGeral"))

        ' Lógica para definir o semestre
        If CInt(mesKey) >= 1 And CInt(mesKey) <= 6 Then
            semestreKey = "1"
        Else
            semestreKey = "2"
        End If
          
        ' Estrutura ANO > SEMESTRE > TRIMESTRE > MÊS > CORRETOR
        If Not reportData.Exists(anoKey) Then reportData.Add anoKey, Server.CreateObject("Scripting.Dictionary")
        If Not reportData(anoKey).Exists(semestreKey) Then reportData(anoKey).Add semestreKey, Server.CreateObject("Scripting.Dictionary")
        If Not reportData(anoKey)(semestreKey).Exists(trimestreKey) Then reportData(anoKey)(semestreKey).Add trimestreKey, Server.CreateObject("Scripting.Dictionary")
        If Not reportData(anoKey)(semestreKey)(trimestreKey).Exists(mesKey) Then reportData(anoKey)(semestreKey)(trimestreKey).Add mesKey, Server.CreateObject("Scripting.Dictionary")
          
        If Not reportData(anoKey)(semestreKey)(trimestreKey)(mesKey).Exists(corretorKey) Then
            Dim corretorData
            Set corretorData = Server.CreateObject("Scripting.Dictionary")
            corretorData.Add "vendas", 0
            corretorData.Add "valor", 0
            corretorData.Add "comissao", 0
            reportData(anoKey)(semestreKey)(trimestreKey)(mesKey).Add corretorKey, corretorData
        End If
          
        With reportData(anoKey)(semestreKey)(trimestreKey)(mesKey)(corretorKey)
            .Item("vendas") = .Item("vendas") + 1
            .Item("valor") = .Item("valor") + valorUnidade
            .Item("comissao") = .Item("comissao") + valorComissao
        End With
          
        ' Ranking por semestre
        Dim semestreFullKey
        semestreFullKey = anoKey & "S" & semestreKey
        If Not rankingPorSemestre.Exists(semestreFullKey) Then rankingPorSemestre.Add semestreFullKey, Server.CreateObject("Scripting.Dictionary")
        If Not rankingPorSemestre(semestreFullKey).Exists(corretorKey) Then
            Dim rankData
            Set rankData = Server.CreateObject("Scripting.Dictionary")
            rankData.Add "vendas", 0
            rankData.Add "valor", 0
            rankingPorSemestre(semestreFullKey).Add corretorKey, rankData
        End If
        With rankingPorSemestre(semestreFullKey)(corretorKey)
            .Item("vendas") = .Item("vendas") + 1
            .Item("valor") = .Item("valor") + valorUnidade
        End With

        ' Ranking por trimestre
        Dim trimestreFullKey
        trimestreFullKey = anoKey & "T" & trimestreKey
        If Not rankingPorTrimestre.Exists(trimestreFullKey) Then rankingPorTrimestre.Add trimestreFullKey, Server.CreateObject("Scripting.Dictionary")
        If Not rankingPorTrimestre(trimestreFullKey).Exists(corretorKey) Then
            Dim rankData2
            Set rankData2 = Server.CreateObject("Scripting.Dictionary")
            rankData2.Add "vendas", 0
            rankData2.Add "valor", 0
            rankingPorTrimestre(trimestreFullKey).Add corretorKey, rankData2
        End If
        With rankingPorTrimestre(trimestreFullKey)(corretorKey)
            .Item("vendas") = .Item("vendas") + 1
            .Item("valor") = .Item("valor") + valorUnidade
        End With
          
        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Relatório de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            font-family: Arial, sans-serif;
            color: #333;
            background-color: #f8f9fa;
            padding: 20px;
        }
        .report-container {
            max-width: 1000px;
            margin: 0 auto;
            background: #fff;
            padding: 30px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 20px;
        }
        .header .logo {
            width: 120px;
            height: auto;
        }
        .header h1 {
            font-size: 24px;
            font-weight: bold;
            margin: 0;
        }
        .summary-card {
            color: #fff;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        .summary-card h3 {
            font-size: 18px;
            margin-bottom: 15px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.3);
            padding-bottom: 10px;
        }
        .summary-card p {
            margin: 5px 0;
            font-size: 16px;
        }
        .summary-card p span {
            font-weight: bold;
            font-size: 20px;
            display: block;
        }
        .section-title {
            font-size: 18px;
            font-weight: bold;
            padding-bottom: 5px;
            margin-bottom: 20px;
        }
        .data-table th, .data-table td {
            border: 1px solid #dee2e6;
            padding: 8px;
            text-align: left;
        }
        .data-table thead th {
            color: #fff;
            border-color: #004d40;
            font-weight: normal;
        }
        .data-table tbody tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .data-table tbody tr:hover {
            background-color: #e9ecef;
        }
        .ranking-table th, .ranking-table td {
                border-color: #004d40 !important;
        }
        .ranking-table thead th {
            color: #fff;
        }
        .ranking-table tbody tr:nth-child(odd) {
            background-color: #e9f5f5;
        }
        .ranking-table tbody tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .text-right {
            text-align: right !important;
        }
        .text-center {
                text-align: center;
        }
        .badge-year, .badge-semester, .badge-quarter, .badge-month {
            display: block;
            text-align: center;
            font-weight: bold;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            color: #fff;
        }
        .badge-year {
            font-size: 30px;
        }
        .badge-semester {
            font-size: 24px;
        }
        .badge-quarter {
            font-size: 20px;
        }
        .badge-month {
            font-size: 18px;
        }
        .filter-bar .form-select-sm {
                height: calc(1.5em + 0.5rem + 2px); /* Ajuste de altura para o select */
        }
          
        @media print {
                body {
                    font-size: 10px;
                }
                .report-container {
                    box-shadow: none;
                    padding: 0;
                }
                .no-print {
                    display: none;
                }
                table {
                    page-break-inside: avoid;
                }
                .page-break {
                    page-break-before: always;
                }
                .col-md-4 {
                    width: 33.333333%;
                    float: left;
                }
        }
    </style>
</head>
<body>
    <div class="report-container">
        <div class="header">
            <img src="img/logoTocaa2.png?text=LOGO" alt="Logo" class="logo">
            <h1>SGVendas - Relatório de Vendas</h1>
        </div>

        <div class="filter-bar no-print mb-4 p-3 bg-light rounded">
            <form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="get" id="filterForm">
                <div class="row g-2">
                    <div class="col-md-2">
                        <label for="ano" class="form-label mb-0">Ano</label>
                        <select class="form-select form-select-sm" id="ano" name="ano">
                            <option value="">Todos</option>
                            <%
                                ' Exemplo de loop dinâmico (você pode buscar do banco)
                                For anoValue = 2028 To 2024 Step -1
                                    Response.Write "<option value='" & anoValue & "'"
                                    If CStr(filtroAno) = CStr(anoValue) Then Response.Write " selected"
                                    Response.Write ">" & anoValue & "</option>"
                                Next
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <label for="mes" class="form-label mb-0">Mês</label>
                        <select class="form-select form-select-sm" id="mes" name="mes">
                            <option value="">Todos</option>
                            <%
                                For i = 1 To 12
                                    Response.Write "<option value='" & i & "'"
                                    If CStr(filtroMes) = CStr(i) Then Response.Write " selected"
                                    Response.Write ">" & mesNome(i) & "</option>"
                                Next
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <label for="trimestre" class="form-label mb-0">Trimestre</label>
                        <select class="form-select form-select-sm" id="trimestre" name="trimestre">
                            <option value="">Todos</option>
                            <option value="1" <% If filtroTrimestre = "1" Then Response.Write "selected" %>>1º</option>
                            <option value="2" <% If filtroTrimestre = "2" Then Response.Write "selected" %>>2º</option>
                            <option value="3" <% If filtroTrimestre = "3" Then Response.Write "selected" %>>3º</option>
                            <option value="4" <% If filtroTrimestre = "4" Then Response.Write "selected" %>>4º</option>
                        </select>
                    </div>
                      
                    <div class="col-md-3">
                        <label for="submit" class="form-label mb-0 d-none d-md-block">&nbsp;</label>
                        <button type="submit" class="btn btn-primary btn-sm w-100" id="submit">
                            <i class="fas fa-filter"></i> Aplicar Filtros
                        </button>
                    </div>
                      
                    <div class="col-md-3">
                        <label for="clear" class="form-label mb-0 d-none d-md-block">&nbsp;</label>
                        <a href="<%= Request.ServerVariables("SCRIPT_NAME") %>" class="btn btn-secondary btn-sm w-100" id="clear">
                            <i class="fas fa-eraser"></i> Limpar Filtros
                        </a>
                    </div>
                </div>
            </form>
        </div>
        <hr>
        <div class="mb-4">
            <span class="text-muted">Filtros Aplicados:   
            <% If filtroAno <> "" Then Response.Write "Ano: " & filtroAno & "; " %>
            <% If filtroMes <> "" Then Response.Write "Mês: " & mesNome(filtroMes) & "; " %>
            <% If filtroTrimestre <> "" Then Response.Write "Trimestre: " & filtroTrimestre & "; " %>
            <% If filtroDiretoria <> "" Then Response.Write "Diretoria: " & filtroDiretoria & "; " %>
            <% If filtroGerencia <> "" Then Response.Write "Gerência: " & filtroGerencia & "; " %>
            <% If filtroCorretor <> "" Then Response.Write "Corretor: " & filtroCorretor & "; " %>
            </span>
        </div>

        <%
        Dim arrAnos, ano, arrSemestres, semestre, arrTrimestres, trimestre, arrMeses, mes
        Dim totalAnoVendas, totalAnoValor, totalAnoComissao
        Dim totalSemestreVendas, totalSemestreValor, totalSemestreComissao
        Dim totalTrimestreVendas, totalTrimestreValor, totalTrimestreComissao
        Dim totalMesVendas, totalMesValor, totalMesComissao
        Dim yearPalette
          
        arrAnos = SortDictionaryByKey(reportData)
          
        For Each ano In arrAnos
                Set yearPalette = GetYearColorPalette(ano)
                  
                Response.Write "<style>"
                Response.Write ".year-" & ano & " .section-title { border-bottom: 2px solid " & yearPalette("primary") & "; color: " & yearPalette("primary") & "; }"
                Response.Write ".year-" & ano & " .data-table thead th { background-color: " & yearPalette("primary") & "; }"
                Response.Write ".year-" & ano & " .ranking-table thead th { background-color: " & yearPalette("primary") & "; }"
                Response.Write ".year-" & ano & " .badge-year { background-color: " & yearPalette("primary") & "; }"
                Response.Write ".year-" & ano & " .badge-semester { background-color: " & yearPalette("secondary") & "; }"
                Response.Write ".year-" & ano & " .badge-quarter { background-color: " & yearPalette("accent") & "; }"
                Response.Write ".year-" & ano & " .badge-month { background-color: " & yearPalette("light") & "; color: #333; }"
                Response.Write ".year-" & ano & " .summary-card { background-color: " & yearPalette("primary") & "; }"
                Response.Write "</style>"
                  
                Response.Write "<div class='section-block year-" & ano & "'>"
                  
                Response.Write "<h2 class='section-title'><span class='badge-year'>ANO: " & ano & "</span></h2>"
                  
                totalAnoVendas = 0
                totalAnoValor = 0
                totalAnoComissao = 0
                  
                arrSemestres = SortDictionaryByKey(reportData(ano))
                  
                For Each semestre In arrSemestres
                        Response.Write "<h3 class='mt-4'><span class='badge-semester'>SEMESTRE: " & semestre & "º</span></h3>"
                          
                        ' Exibir TOP 10 corretores do semestre
                        semestreFullKey = ano & "S" & semestre
                        If rankingPorSemestre.Exists(semestreFullKey) Then
                                arrRankingCorretores = SortDictionaryByValue(rankingPorSemestre(semestreFullKey), "valor")
                                Response.Write "<p class='mb-2 fw-bold text-muted'>Top 10 Corretores do Semestre " & semestre & "º</p>"
                                Response.Write "<table class='table table-sm data-table ranking-table'>"
                                ' CORREÇÃO AQUI: Vendas agora está centralizada
                                Response.Write "<thead><tr><th style='width:50px;'>Pos</th><th>Corretor</th><th class='text-center'>Vendas</th><th class='text-right'>Valor (R$)</th></tr></thead><tbody>"
                                maxPos = 10
                                If UBound(arrRankingCorretores) + 1 < maxPos Then maxPos = UBound(arrRankingCorretores) + 1
                                For pos = 0 To maxPos - 1
                                        rankCorretor = arrRankingCorretores(pos)
                                        Set rankData = rankingPorSemestre(semestreFullKey)(rankCorretor)
                                        ' CORREÇÃO AQUI: Dados de Vendas centralizados
                                        Response.Write "<tr><td>" & (pos + 1) & "º</td><td>" & rankCorretor & "</td><td class='text-center'>" & rankData("vendas") & "</td><td class='text-right'>" & FormatNumber(rankData("valor"), 2) & "</td></tr>"
                                Next
                                Response.Write "</tbody></table>"
                        End If

                        totalSemestreVendas = 0
                        totalSemestreValor = 0
                        totalSemestreComissao = 0
                          
                        arrTrimestres = SortDictionaryByKey(reportData(ano)(semestre))
                          
                        For Each trimestre In arrTrimestres
                                Response.Write "<h4 class='mt-3'><span class='badge-quarter'>TRIMESTRE: " & trimestre & "</span></h4>"
                                  
                                ' Exibir TOP 10 corretores do trimestre
                                trimestreFullKey = ano & "T" & trimestre
                                If rankingPorTrimestre.Exists(trimestreFullKey) Then
                                        arrRankingCorretores = SortDictionaryByValue(rankingPorTrimestre(trimestreFullKey), "valor")
                                        Response.Write "<p class='mb-2 fw-bold text-muted'>Top 10 Corretores do Trimestre " & trimestre & "</p>"
                                        Response.Write "<table class='table table-sm data-table ranking-table'>"
                                        ' CORREÇÃO AQUI: Vendas agora está centralizada
                                        Response.Write "<thead><tr><th style='width:50px;'>Pos</th><th>Corretor</th><th class='text-center'>Vendas</th><th class='text-right'>Valor (R$)</th></tr></thead><tbody>"
                                        maxPos = 10
                                        If UBound(arrRankingCorretores) + 1 < maxPos Then maxPos = UBound(arrRankingCorretores) + 1
                                        For pos = 0 To maxPos - 1
                                                rankCorretor = arrRankingCorretores(pos)
                                                Set rankData = rankingPorTrimestre(trimestreFullKey)(rankCorretor)
                                                ' CORREÇÃO AQUI: Dados de Vendas centralizados
                                                Response.Write "<tr><td>" & (pos + 1) & "º</td><td>" & rankCorretor & "</td><td class='text-center'>" & rankData("vendas") & "</td><td class='text-right'>" & FormatNumber(rankData("valor"), 2) & "</td></tr>"
                                        Next
                                        Response.Write "</tbody></table>"
                                End If

                                totalTrimestreVendas = 0
                                totalTrimestreValor = 0
                                totalTrimestreComissao = 0

                                arrMeses = SortDictionaryByKey(reportData(ano)(semestre)(trimestre))
                                  
                                For Each mes In arrMeses
                                        Response.Write "<h5 class='mt-2'><span class='badge-month'>MÊS: " & mesNome(mes) & "</span></h5>"
                                        totalMesVendas = 0
                                        totalMesValor = 0
                                        totalMesComissao = 0
                                          
                                        Response.Write "<table class='table table-sm data-table'>"
                                        ' CORREÇÃO AQUI: Vendas agora está centralizada
                                        Response.Write "<thead><tr><th style='width:50px;'>Pos</th><th>Corretor</th><th class='text-center'>Vendas</th><th class='text-right'>Valor (R$)</th><th class='text-right'>Comissão (R$)</th></tr></thead><tbody>"
                                          
                                        arrCorretores = SortDictionaryByValue(reportData(ano)(semestre)(trimestre)(mes), "valor")
                                          
                                        posicao = 0
                                        For Each corretorKey In arrCorretores
                                                posicao = posicao + 1
                                                Set corretorData = reportData(ano)(semestre)(trimestre)(mes)(corretorKey)
                                                ' CORREÇÃO AQUI: Dados de Vendas centralizados
                                                Response.Write "<tr><td>" & posicao & "º</td><td>" & corretorKey & "</td><td class='text-center'>" & corretorData("vendas") & "</td><td class='text-right'>" & FormatNumber(corretorData("valor"), 2) & "</td><td class='text-right'>" & FormatNumber(corretorData("comissao"), 2) & "</td></tr>"
                                                  
                                                totalMesVendas = totalMesVendas + corretorData("vendas")
                                                totalMesValor = totalMesValor + corretorData("valor")
                                                totalMesComissao = totalMesComissao + corretorData("comissao")
                                        Next
                                          
                                        ' CORREÇÃO AQUI: Total de Vendas centralizado
                                        Response.Write "<tr class='table-info'><td colspan='2' class='fw-bold'>TOTAL DO MÊS</td><td class='text-center fw-bold'>" & totalMesVendas & "</td><td class='text-right fw-bold'>" & FormatNumber(totalMesValor, 2) & "</td><td class='text-right fw-bold'>" & FormatNumber(totalMesComissao, 2) & "</td></tr>"
                                        Response.Write "</tbody></table>"
                                          
                                        totalTrimestreVendas = totalTrimestreVendas + totalMesVendas
                                        totalTrimestreValor = totalTrimestreValor + totalMesValor
                                        totalTrimestreComissao = totalTrimestreComissao + totalMesComissao
                                Next
                                  
                                totalSemestreVendas = totalSemestreVendas + totalTrimestreVendas
                                totalSemestreValor = totalSemestreValor + totalTrimestreValor
                                totalSemestreComissao = totalSemestreComissao + totalTrimestreComissao
                        Next
                          
                        totalAnoVendas = totalAnoVendas + totalSemestreVendas
                        totalAnoValor = totalAnoValor + totalSemestreValor
                        totalAnoComissao = totalSemestreComissao + totalSemestreComissao
                Next
                  
                ' Início da nova estrutura de layout
                Response.Write "<div class='row my-3 g-3'>"
                  
                ' Card de Trimestre
                Response.Write "<div class='col-md-4'>"
                Response.Write "<div class='summary-card'><h3 class='text-white'>Total do Trimestre</h3><p>Total de Vendas: <span>" & totalTrimestreVendas & "</span></p><p>Valor Total: <span>" & FormatNumber(totalTrimestreValor, 2) & " R$</span></p><p>Comissão Total: <span>" & FormatNumber(totalTrimestreComissao, 2) & " R$</span></p></div>"
                Response.Write "</div>"
                  
                ' Card de Semestre
                Response.Write "<div class='col-md-4'>"
                Response.Write "<div class='summary-card'><h3 class='text-white'>Total do Semestre</h3><p>Total de Vendas: <span>" & totalSemestreVendas & "</span></p><p>Valor Total: <span>" & FormatNumber(totalSemestreValor, 2) & " R$</span></p><p>Comissão Total: <span>" & FormatNumber(totalSemestreComissao, 2) & " R$</span></p></div>"
                Response.Write "</div>"

                ' Card de Ano
                Response.Write "<div class='col-md-4'>"
                Response.Write "<div class='summary-card'><h3 class='text-white'>Total do Ano " & ano & "</h3><p>Total de Vendas: <span>" & totalAnoVendas & "</span></p><p>Valor Total: <span>" & FormatNumber(totalAnoValor, 2) & " R$</span></p><p>Comissão Total: <span>" & FormatNumber(totalAnoComissao, 2) & " R$</span></p></div>"
                Response.Write "</div>"
                  
                Response.Write "</div>" ' Fim da row

                Response.Write "</div>" ' Fecha o bloco do ano
                If Not ano = arrAnos(UBound(arrAnos)) Then
                        Response.Write "<div class='page-break'></div>"
                End If
        Next
        %>

        <div class="text-center mt-5 no-print">
            <button onclick="window.print()" class="btn btn-primary">
                <i class="fas fa-print"></i> Imprimir Relatório
            </button>
        </div>
    </div>
</body>
</html>