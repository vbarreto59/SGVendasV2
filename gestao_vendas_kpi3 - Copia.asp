<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="gestao_header.inc"-->

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

' Verifica conexões
If IsEmpty(StrConn) Then
    Response.Write "<h3 style='color:red'>ERRO: StrConn não está definido em conexao.asp</h3>"
    Response.End
End If

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================
Function GetUniqueValues(connObj, tableName, columnName, whereClause)
    Dim dict, rsU, sqlU
    Set dict = Server.CreateObject("Scripting.Dictionary")
    sqlU = "SELECT DISTINCT " & columnName & " FROM " & tableName & whereClause & " ORDER BY " & columnName
    On Error Resume Next
    Set rsU = connObj.Execute(sqlU)
    If Err.Number <> 0 Then
        GetUniqueValues = Array()
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If Not rsU.EOF Then
        Do While Not rsU.EOF
            If Not IsNull(rsU(0)) Then
                dict.Add CStr(rsU(0)), 1
            End If
            rsU.MoveNext
        Loop
    End If

    If Not rsU Is Nothing Then rsU.Close
    Set rsU = Nothing

    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

Sub ProcessData(mainDict, key, qtd, valor)
    If key = "" Or IsNull(key) Then Exit Sub
    If Not mainDict.Exists(CStr(key)) Then
        Dim newD
        Set newD = Server.CreateObject("Scripting.Dictionary")
        newD.Add "vendas", 0
        newD.Add "valor", 0
        mainDict.Add CStr(key), newD
    End If
    mainDict(CStr(key))("vendas") = mainDict(CStr(key))("vendas") + qtd
    mainDict(CStr(key))("valor") = mainDict(CStr(key))("valor") + valor
End Sub

Function SortKeysNumeric(dictObj)
    Dim arrKeys, i, j, tmp
    If dictObj.Count = 0 Then
        SortKeysNumeric = Array()
        Exit Function
    End If
    arrKeys = dictObj.Keys
    For i = 0 To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            If CInt(arrKeys(i)) > CInt(arrKeys(j)) Then
                tmp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = tmp
            End If
        Next
    Next
    SortKeysNumeric = arrKeys
End Function

Function SortKeysAlpha(dictObj)
    Dim arrKeys, i, j, tmp
    If dictObj.Count = 0 Then
        SortKeysAlpha = Array()
        Exit Function
    End If
    arrKeys = dictObj.Keys
    For i = 0 To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            If CStr(arrKeys(i)) > CStr(arrKeys(j)) Then
                tmp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = tmp
            End If
        Next
    Next
    SortKeysAlpha = arrKeys
End Function

' ===============================================
' FILTROS (QueryString)
' ===============================================
Dim filtroAno, filtroSemestre, filtroMes, filtroTrimestre
Dim filtroDiretoria, filtroGerencia, filtroCorretor
Dim filtroEmpreendimento, filtroEmpresa

filtroAno = Request.QueryString("ano")
filtroSemestre = Request.QueryString("semestre")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")
filtroEmpresa = Request.QueryString("empresa")

' ===============================================
' BUSCA VALORES ÚNICOS PARA OS FILTROS
' ===============================================
Dim uniqueAnos, uniqueMeses, uniqueTrimestres, uniqueDiretorias, uniqueGerencias
Dim uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas

uniqueAnos = GetUniqueValues(connSales, "Vendas", "AnoVenda", " WHERE Excluido = 0")
uniqueMeses = GetUniqueValues(connSales, "Vendas", "MesVenda", " WHERE Excluido = 0")
uniqueTrimestres = GetUniqueValues(connSales, "Vendas", "Trimestre", " WHERE Excluido = 0")
uniqueDiretorias = GetUniqueValues(connSales, "Vendas", "Diretoria", " WHERE Excluido = 0 AND Diretoria IS NOT NULL")
uniqueGerencias = GetUniqueValues(connSales, "Vendas", "Gerencia", " WHERE Excluido = 0 AND Gerencia IS NOT NULL")
uniqueCorretores = GetUniqueValues(connSales, "Vendas", "Corretor", " WHERE Excluido = 0 AND Corretor IS NOT NULL")
uniqueEmpreendimentos = GetUniqueValues(conn, "Empreendimento", "NomeEmpreendimento", " WHERE Excluido = 0")
uniqueEmpresas = GetUniqueValues(conn, "Empresa", "NomeEmpresa", " WHERE Excluido = 0")

' ===============================================
' NOMES DOS MESES
' ===============================================
Dim arrMesesNome
ReDim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Marco"
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
' INICIALIZA KPI DATA
' ===============================================
Dim kpiData, categories, cat
Set kpiData = Server.CreateObject("Scripting.Dictionary")
categories = Array("Ano","Semestre","Trimestre","Mes","TopCorretores","TopDiretorias","TopGerencias","TopEmpreendimentos","TopEmpresas")
For Each cat In categories
    Set kpiData(cat) = Server.CreateObject("Scripting.Dictionary")
Next

' ===============================================
' CONSULTA PRINCIPAL (somente ValorUnidade)
' ===============================================
Dim sqlVendas, rsVendas
sqlVendas = "SELECT ID, DataVenda, AnoVenda, MesVenda, Trimestre, ValorUnidade, Diretoria, Gerencia, Corretor, Empreend_ID FROM Vendas WHERE Excluido = 0"

If filtroAno <> "" Then sqlVendas = sqlVendas & " AND AnoVenda = " & Replace(filtroAno,"'","''")
If filtroSemestre <> "" Then
    If filtroSemestre = "1" Then
        sqlVendas = sqlVendas & " AND MesVenda BETWEEN 1 AND 6"
    ElseIf filtroSemestre = "2" Then
        sqlVendas = sqlVendas & " AND MesVenda BETWEEN 7 AND 12"
    End If
End If
If filtroMes <> "" Then sqlVendas = sqlVendas & " AND MesVenda = " & Replace(filtroMes,"'","''")
If filtroTrimestre <> "" Then sqlVendas = sqlVendas & " AND Trimestre = " & Replace(filtroTrimestre,"'","''")
If filtroDiretoria <> "" Then sqlVendas = sqlVendas & " AND Diretoria = '" & Replace(filtroDiretoria,"'","''") & "'"
If filtroGerencia <> "" Then sqlVendas = sqlVendas & " AND Gerencia = '" & Replace(filtroGerencia,"'","''") & "'"
If filtroCorretor <> "" Then sqlVendas = sqlVendas & " AND Corretor = '" & Replace(filtroCorretor,"'","''") & "'"
If filtroEmpreendimento <> "" Then sqlVendas = sqlVendas & " AND Empreend_ID IN (SELECT Empreend_ID FROM Empreendimento WHERE NomeEmpreendimento = '" & Replace(filtroEmpreendimento,"'","''") & "')"
If filtroEmpresa <> "" Then sqlVendas = sqlVendas & " AND Empreend_ID IN (SELECT Empreend_ID FROM Empreendimento WHERE Empresa_ID IN (SELECT Empresa_ID FROM Empresa WHERE NomeEmpresa = '" & Replace(filtroEmpresa,"'","''") & "'))"

sqlVendas = sqlVendas & " ORDER BY AnoVenda, MesVenda"

Set rsVendas = connSales.Execute(sqlVendas)

' ===============================================
' Dicionário auxiliar para gráfico
' ===============================================
Dim chartDict
Set chartDict = Server.CreateObject("Scripting.Dictionary")
Dim m
For m = 1 To 12
    chartDict.Add CStr(m), 0
Next

' ===============================================
' PROCESSA REGISTROS — SOMENTE ValorUnidade
' ===============================================
If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        Dim valorUnidade, anoVenda, mesVenda, trimestreVenda, semestreVenda
        Dim diretoria, gerencia, corretor, empreend_ID, empreendimento, empresa

        If Not IsNull(rsVendas("ValorUnidade")) And rsVendas("ValorUnidade") <> "" Then
            valorUnidade = CDbl(rsVendas("ValorUnidade"))
        Else
            valorUnidade = 0
        End If

        anoVenda = ""
        If Not IsNull(rsVendas("AnoVenda")) Then anoVenda = CStr(rsVendas("AnoVenda"))

        If Not IsNull(rsVendas("MesVenda")) Then
            mesVenda = CInt(rsVendas("MesVenda"))
        Else
            mesVenda = 0
        End If

        If Not IsNull(rsVendas("Trimestre")) Then
            trimestreVenda = CStr(rsVendas("Trimestre"))
        Else
            trimestreVenda = ""
        End If

        If mesVenda > 0 And mesVenda <= 6 Then
            semestreVenda = "1"
        ElseIf mesVenda >= 7 And mesVenda <= 12 Then
            semestreVenda = "2"
        Else
            semestreVenda = ""
        End If

        If Not IsNull(rsVendas("Diretoria")) Then diretoria = CStr(rsVendas("Diretoria")) Else diretoria = ""
        If Not IsNull(rsVendas("Gerencia")) Then gerencia = CStr(rsVendas("Gerencia")) Else gerencia = ""
        If Not IsNull(rsVendas("Corretor")) Then corretor = CStr(rsVendas("Corretor")) Else corretor = ""
        If Not IsNull(rsVendas("Empreend_ID")) Then empreend_ID = rsVendas("Empreend_ID") Else empreend_ID = ""

        empreendimento = ""
        empresa = ""

        If empreend_ID <> "" Then
            On Error Resume Next
            Dim rsEmp, sqlEmp
            sqlEmp = "SELECT NomeEmpreendimento, Empresa_ID FROM Empreendimento WHERE Empreend_ID = " & empreend_ID
            Set rsEmp = conn.Execute(sqlEmp)
            On Error GoTo 0

            If Not rsEmp Is Nothing Then
                If Not rsEmp.EOF Then
                    If Not IsNull(rsEmp("NomeEmpreendimento")) Then empreendimento = CStr(rsEmp("NomeEmpreendimento"))
                    If Not IsNull(rsEmp("Empresa_ID")) Then
                        Dim rsEmp2, sqlEmp2
                        sqlEmp2 = "SELECT NomeEmpresa FROM Empresa WHERE Empresa_ID = " & rsEmp("Empresa_ID")
                        Set rsEmp2 = conn.Execute(sqlEmp2)
                        If Not rsEmp2 Is Nothing Then
                            If Not rsEmp2.EOF Then
                                empresa = CStr(rsEmp2("NomeEmpresa"))
                            End If
                            rsEmp2.Close
                            Set rsEmp2 = Nothing
                        End If
                    End If
                End If
                rsEmp.Close
                Set rsEmp = Nothing
            End If
        End If

        ' Atualiza KPIs
        If anoVenda <> "" Then Call ProcessData(kpiData("Ano"), anoVenda, 1, valorUnidade)
        If semestreVenda <> "" Then Call ProcessData(kpiData("Semestre"), semestreVenda, 1, valorUnidade)
        If trimestreVenda <> "" Then Call ProcessData(kpiData("Trimestre"), trimestreVenda, 1, valorUnidade)
        If mesVenda > 0 Then Call ProcessData(kpiData("Mes"), CStr(mesVenda), 1, valorUnidade)

        If corretor <> "" Then Call ProcessData(kpiData("TopCorretores"), corretor, 1, valorUnidade)
        If diretoria <> "" Then Call ProcessData(kpiData("TopDiretorias"), diretoria, 1, valorUnidade)
        If gerencia <> "" Then Call ProcessData(kpiData("TopGerencias"), gerencia, 1, valorUnidade)
        If empreendimento <> "" Then Call ProcessData(kpiData("TopEmpreendimentos"), empreendimento, 1, valorUnidade)
        If empresa <> "" Then Call ProcessData(kpiData("TopEmpresas"), empresa, 1, valorUnidade)

        ' Atualiza chartDict
        If mesVenda >= 1 And mesVenda <= 12 Then
            chartDict(CStr(mesVenda)) = chartDict(CStr(mesVenda)) + valorUnidade
        End If

        rsVendas.MoveNext
    Loop
End If

If Not rsVendas Is Nothing Then rsVendas.Close
Set rsVendas = Nothing

' =========================================================
' Prepara dados para o gráfico
' =========================================================
Dim chartLabels, chartData, mesesFiltrados, totalMeses, j, i, mesNum
If CStr(filtroSemestre) = "1" Then
    mesesFiltrados = Array(1,2,3,4,5,6)
ElseIf CStr(filtroSemestre) = "2" Then
    mesesFiltrados = Array(7,8,9,10,11,12)
Else
    mesesFiltrados = Array(1,2,3,4,5,6,7,8,9,10,11,12)
End If

If filtroMes <> "" And IsNumeric(filtroMes) Then mesesFiltrados = Array(CInt(filtroMes))

If IsArray(mesesFiltrados) Then
    totalMeses = UBound(mesesFiltrados) - LBound(mesesFiltrados) + 1
Else
    totalMeses = 0
End If

If totalMeses > 0 Then
    ReDim chartLabels(totalMeses - 1), chartData(totalMeses - 1)
    For j = 0 To totalMeses - 1
        i = mesesFiltrados(j)
        chartLabels(j) = arrMesesNome(i)
        Dim mesKey
        mesKey = CStr(i)
        If kpiData("Mes").Exists(mesKey) Then
            chartData(j) = CDbl(kpiData("Mes")(mesKey)("valor"))
        Else
            chartData(j) = 0
        End If
    Next
Else
    ReDim chartLabels(-1), chartData(-1)
End If

' ===============================================
' PREPARA KPIs SUMÁRIOS
' ===============================================
Dim anoRef
If filtroAno <> "" Then anoRef = filtroAno Else anoRef = Year(Date())

Dim totalAno, mediaMensal, maiorValor, menorValor
totalAno = 0 : maiorValor = 0 : menorValor = 0

If kpiData("Ano").Exists(CStr(anoRef)) Then
    totalAno = kpiData("Ano")(CStr(anoRef))("valor")
End If

Dim mCount
mCount = 12
If totalAno > 0 Then
    mediaMensal = totalAno / mCount
Else
    mediaMensal = 0
End If

' maior/menor entre meses do anoRef
maiorValor = -1
menorValor = -1
For idx = 1 To 12
    Dim vtmp
    vtmp = chartDict(CStr(idx))
    If vtmp > 0 Then
        If maiorValor = -1 Or vtmp > maiorValor Then maiorValor = vtmp
        If menorValor = -1 Or vtmp < menorValor Then menorValor = vtmp
    End If
Next

If maiorValor = -1 Then maiorValor = 0
If menorValor = -1 Then menorValor = 0

' Prepara top corretores
Function GetTopList(dictObj, topN)
    Dim keysArr, i, j, tmpKey
    If dictObj.Count = 0 Then
        GetTopList = Array()
        Exit Function
    End If
    keysArr = dictObj.Keys
    For i = 0 To UBound(keysArr) - 1
        For j = i + 1 To UBound(keysArr)
            If dictObj(keysArr(i))("valor") < dictObj(keysArr(j))("valor") Then
                tmpKey = keysArr(i)
                keysArr(i) = keysArr(j)
                keysArr(j) = tmpKey
            End If
        Next
    Next
    Dim result(), take
    take = topN
    If take > UBound(keysArr) + 1 Then take = UBound(keysArr) + 1
    ReDim result(take - 1)
    For i = 0 To take - 1
        result(i) = Array(keysArr(i), dictObj(keysArr(i))("vendas"), dictObj(keysArr(i))("valor"))
    Next
    GetTopList = result
End Function

Dim topCorretores, topDiretorias, topGerencias
topCorretores = GetTopList(kpiData("TopCorretores"), 10)
topDiretorias = GetTopList(kpiData("TopDiretorias"), 10)
topGerencias  = GetTopList(kpiData("TopGerencias"), 10)
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Gestão Vendas - KPIs</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
    body { background:#f5f7fb; padding-bottom:50px; }
    .kpi-card { 
        border-radius:8px; 
        color: black;

        padding:15px; 
        background:#ABF5AB; 
        box-shadow:0 1px 4px rgba(0,0,0,0.08); 
    }
    .small-muted { font-size:0.9rem; color:#666; }
    .month-card { 
        border-radius:8px; 
        padding:15px; 
        background:#DEDEDB; 
        box-shadow:0 1px 4px rgba(0,0,0,0.08);
        text-align: center;
        height: 100%;
        transition: transform 0.2s ease;
        border-left: 4px solid #dee2e6;
    }
    .month-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .month-value {
        font-weight: bold;
        font-size: 1.2rem;
        color: #2c3e50;
        margin: 8px 0;
    }
    .month-name {
        font-size: 0.9rem;
        color: #7f8c8d;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .month-badge {
        font-size: 0.75rem;
        padding: 4px 8px;
        border-radius: 12px;
    }
    .cards-container {
        margin-bottom: 25px;
    }
    .card-highlight {
        border-left-color: #28a745 !important;
        background: linear-gradient(135deg, #fff, #f8fff8);
    }
    .card-warning {
        border-left-color: #ffc107 !important;
        background: linear-gradient(135deg, #fff, #fffbf0);
    }
    .card-info {
        border-left-color: #17a2b8 !important;
        background: linear-gradient(135deg, #fff, #f0f9ff);
    }
</style>
</head>
<body>
<nav class="navbar navbar-expand-lg navbar-dark bg-dark">
  <div class="container-fluid">
    <a class="navbar-brand" href="#">Gestão Vendas</a>
    <div class="collapse navbar-collapse">
      <ul class="navbar-nav ms-auto">
        <li class="nav-item"><a class="nav-link" href="gestao_logout.asp">Sair</a></li>
      </ul>
    </div>
  </div>
</nav>

<div class="container mt-4">
    <!-- FILTROS -->
    <div class="card mb-3">
        <div class="card-body">
            <form id="filterForm" method="get" class="row g-2 align-items-end">
                <div class="col-md-2">
                    <label class="form-label">Ano</label>
                    <select name="ano" class="form-select" onchange="this.form.submit()">
                        <option value="">Todos</option>
                        <% If IsArray(uniqueAnos) Then
                            For i = 0 To UBound(uniqueAnos)
                                If Not IsNull(uniqueAnos(i)) And uniqueAnos(i) <> "" Then
                                    Response.Write "<option value=""" & Server.HTMLEncode(uniqueAnos(i)) & """"
                                    If CStr(filtroAno) = CStr(uniqueAnos(i)) Then Response.Write " selected"
                                    Response.Write ">" & Server.HTMLEncode(uniqueAnos(i)) & "</option>"
                                End If
                            Next
                        End If %>
                    </select>
                </div>

                <div class="col-md-2">
                    <label class="form-label">Semestre</label>
                    <select name="semestre" class="form-select" onchange="this.form.submit()">
                        <option value="">Todos</option>
                        <option value="1" <% If filtroSemestre = "1" Then Response.Write "selected" %>>1º Semestre</option>
                        <option value="2" <% If filtroSemestre = "2" Then Response.Write "selected" %>>2º Semestre</option>
                    </select>
                </div>

                <div class="col-md-2">
                    <label class="form-label">Mês</label>
                    <select name="mes" class="form-select" onchange="this.form.submit()">
                        <option value="">Todos</option>
                        <% For i = 1 To 12
                            Response.Write "<option value=""" & i & """"
                            If CStr(filtroMes) = CStr(i) Then Response.Write " selected"
                            Response.Write ">" & arrMesesNome(i) & "</option>"
                        Next %>
                    </select>
                </div>

                <div class="col-md-2">
                    <label class="form-label">Trimestre</label>
                    <select name="trimestre" class="form-select" onchange="this.form.submit()">
                        <option value="">Todos</option>
                        <% If IsArray(uniqueTrimestres) Then
                            For i = 0 To UBound(uniqueTrimestres)
                                If Not IsNull(uniqueTrimestres(i)) And uniqueTrimestres(i) <> "" Then
                                    Response.Write "<option value=""" & Server.HTMLEncode(uniqueTrimestres(i)) & """"
                                    If CStr(filtroTrimestre) = CStr(uniqueTrimestres(i)) Then Response.Write " selected"
                                    Response.Write ">" & Server.HTMLEncode(uniqueTrimestres(i)) & "</option>"
                                End If
                            Next
                        End If %>
                    </select>
                </div>

                <div class="col-md-2">
                    <label class="form-label">Diretoria</label>
                    <select name="diretoria" class="form-select" onchange="this.form.submit()">
                        <option value="">Todos</option>
                        <% If IsArray(uniqueDiretorias) Then
                            For i = 0 To UBound(uniqueDiretorias)
                                If Not IsNull(uniqueDiretorias(i)) And uniqueDiretorias(i) <> "" Then
                                    Response.Write "<option value=""" & Server.HTMLEncode(uniqueDiretorias(i)) & """"
                                    If CStr(filtroDiretoria) = CStr(uniqueDiretorias(i)) Then Response.Write " selected"
                                    Response.Write ">" & Server.HTMLEncode(uniqueDiretorias(i)) & "</option>"
                                End If
                            Next
                        End If %>
                    </select>
                </div>

                <div class="col-md-2 text-end">
                    <button type="button" class="btn btn-secondary" onclick="window.location.href=window.location.pathname">Limpar</button>
                </div>
            </form>
        </div>
    </div>

    <!-- CARDS DOS MESES -->
    <div class="cards-container">
        <div class="card">
            <div class="card-header d-flex justify-content-between align-items-center">
                <h6 class="mb-0">Vendas por Mês - Ano: <%= Server.HTMLEncode(anoRef) %></h6>
                <span class="badge bg-primary">Total: R$ <%= FormatNumber(totalAno, 2) %></span>
            </div>
            <div class="card-body">
                <div class="row g-3">
                    <%
                    Dim mesesParaExibir
                    If CStr(filtroSemestre) = "1" Then
                        mesesParaExibir = Array(1,2,3,4,5,6)
                    ElseIf CStr(filtroSemestre) = "2" Then
                        mesesParaExibir = Array(7,8,9,10,11,12)
                    Else
                        mesesParaExibir = Array(1,2,3,4,5,6,7,8,9,10,11,12)
                    End If
                    
                    If filtroMes <> "" And IsNumeric(filtroMes) Then mesesParaExibir = Array(CInt(filtroMes))
                    
                    For Each mesNum In mesesParaExibir
                        Dim valorMes, classeCard, badgeClass, badgeText
                        valorMes = chartDict(CStr(mesNum))
                        
                        ' Define a classe e badge baseado no valor
                        If valorMes = 0 Then
                            classeCard = ""
                            badgeClass = "bg-secondary"
                            badgeText = "Sem vendas"
                        ElseIf valorMes = maiorValor And maiorValor > 0 Then
                            classeCard = "card-highlight"
                            badgeClass = "bg-success"
                            badgeText = "Melhor mês"
                        ElseIf valorMes >= maiorValor * 0.7 Then
                            classeCard = "card-warning"
                            badgeClass = "bg-warning"
                            badgeText = "Alto desempenho"
                        Else
                            classeCard = "card-info"
                            badgeClass = "bg-info"
                            badgeText = "Em andamento"
                        End If
                    %>
                    <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                        <div class="month-card <%= classeCard %>">
                            <div class="month-name"><%= arrMesesNome(mesNum) %></div>
                            <div class="month-value">R$ <%= FormatNumber(valorMes, 2) %></div>
                            <div>
                                <span class="badge <%= badgeClass %> month-badge"><%= badgeText %></span>
                            </div>
                            <div class="small text-muted mt-1">
                                <% If valorMes > 0 And totalAno > 0 Then %>
                                    <%= FormatNumber((valorMes / totalAno) * 100, 1) %>%
                                <% Else %>
                                    0%
                                <% End If %>
                                do total
                            </div>
                        </div>
                    </div>
                    <% Next %>
                </div>
                
                <!-- Resumo -->
                <div class="row mt-4">
                    <div class="col-12">
                        <div class="alert alert-light border">
                            <div class="row text-center">
                                <div class="col-md-3">
                                    <small class="text-muted">Média Mensal</small>
                                    <div class="fw-bold text-primary">R$ <%= FormatNumber(mediaMensal, 2) %></div>
                                </div>
                                <div class="col-md-3">
                                    <small class="text-muted">Maior Mês</small>
                                    <div class="fw-bold text-success">R$ <%= FormatNumber(maiorValor, 2) %></div>
                                </div>
                                <div class="col-md-3">
                                    <small class="text-muted">Menor Mês</small>
                                    <div class="fw-bold text-warning">R$ <%= FormatNumber(menorValor, 2) %></div>
                                </div>
                                <div class="col-md-3">
                                    <small class="text-muted">Meses com Vendas</small>
                                    <div class="fw-bold text-info">
                                        <%
                                        Dim mesesComVendas
                                        mesesComVendas = 0
                                        For i = 1 To 12
                                            If chartDict(CStr(i)) > 0 Then mesesComVendas = mesesComVendas + 1
                                        Next
                                        Response.Write mesesComVendas & "/" & UBound(mesesParaExibir) + 1
                                        %>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- KPIs PRINCIPAIS -->
    <div class="row mb-3">
        <div class="col-md-3">
            <div class="kpi-card text-center">
                <div class="small-muted">Total (ano: <%= Server.HTMLEncode(anoRef) %>)</div>
                <h4>R$ <%= FormatNumber(totalAno,2) %></h4>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card text-center">
                <div class="small-muted">Média mensal</div>
                <h4>R$ <%= FormatNumber(mediaMensal,2) %></h4>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card text-center">
                <div class="small-muted">Maior mês</div>
                <h4>R$ <%= FormatNumber(maiorValor,2) %></h4>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card text-center">
                <div class="small-muted">Menor mês</div>
                <h4>R$ <%= FormatNumber(menorValor,2) %></h4>
            </div>
        </div>
    </div>

    <!-- Top listas -->
    <div class="row mb-4">
        <div class="col-md-4">
            <div class="card p-2">
                <h6>Top Corretores</h6>
                <ul class="list-group list-group-flush">
                    <% If IsArray(topCorretores) And UBound(topCorretores) >= 0 Then
                        For i = 0 To UBound(topCorretores)
                            Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'>" & _
                                Server.HTMLEncode(topCorretores(i)(0)) & "<span>R$ " & FormatNumber(topCorretores(i)(2),2) & "</span></li>"
                        Next
                    Else
                        Response.Write "<li class='list-group-item'>Nenhum registro</li>"
                    End If %>
                </ul>
            </div>
        </div>

        <div class="col-md-4">
            <div class="card p-2">
                <h6>Top Diretorias</h6>
                <ul class="list-group list-group-flush">
                    <% If IsArray(topDiretorias) And UBound(topDiretorias) >= 0 Then
                        For i = 0 To UBound(topDiretorias)
                            Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'>" & _
                                Server.HTMLEncode(topDiretorias(i)(0)) & "<span>R$ " & FormatNumber(topDiretorias(i)(2),2) & "</span></li>"
                        Next
                    Else
                        Response.Write "<li class='list-group-item'>Nenhum registro</li>"
                    End If %>
                </ul>
            </div>
        </div>

        <div class="col-md-4">
            <div class="card p-2">
                <h6>Top Gerências</h6>
                <ul class="list-group list-group-flush">
                    <% If IsArray(topGerencias) And UBound(topGerencias) >= 0 Then
                        For i = 0 To UBound(topGerencias)
                            Response.Write "<li class='list-group-item d-flex justify-content-between align-items-center'>" & _
                                Server.HTMLEncode(topGerencias(i)(0)) & "<span>R$ " & FormatNumber(topGerencias(i)(2),2) & "</span></li>"
                        Next
                    Else
                        Response.Write "<li class='list-group-item'>Nenhum registro</li>"
                    End If %>
                </ul>
            </div>
        </div>
    </div>

    <!-- GRÁFICO -->
    <div class="card mb-5">
        <div class="card-body">
            <h5>Vendas por mês </h5>
            <canvas id="monthlySalesChart" height="120"></canvas>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<script>
    // Dados para o gráfico - CORREÇÃO APLICADA
    const chartLabels = [<%
        If IsArray(chartLabels) And UBound(chartLabels) >= 0 Then
            For i = 0 To UBound(chartLabels)
                If i > 0 Then Response.Write ","
                Response.Write """" & Replace(Server.HTMLEncode(chartLabels(i)), """", "\""") & """"
            Next
        Else
            Response.Write ""
        End If
    %>];
    
    const chartData = [<%
        If IsArray(chartData) And UBound(chartData) >= 0 Then
            For i = 0 To UBound(chartData)
                If i > 0 Then Response.Write ","
                Response.Write Replace(CStr(chartData(i)), ",", ".")
            Next
        Else
            Response.Write ""
        End If
    %>];

    console.log("DEBUG GRÁFICO - Labels:", chartLabels);
    console.log("DEBUG GRÁFICO - Data:", chartData);
    console.log("DEBUG GRÁFICO - Total de pontos:", chartLabels.length);

    // Verifica se há dados para o gráfico
    if (chartLabels.length === 0 || chartData.length === 0) {
        console.warn("Dados do gráfico vazios ou inválidos");
        document.getElementById('monthlySalesChart').closest('.card-body').innerHTML = 
            '<div class="alert alert-info text-center">Nenhum dado disponível para o gráfico com os filtros atuais.</div>';
    } else {
        // Debug detalhado no console
        console.log("DEBUG: populando gráfico — ano, mês, valor (somente ValorUnidade):");
        for (let i = 0; i < chartLabels.length; i++) {
            console.log(`${chartLabels[i]} : R$ ${Number(chartData[i]).toLocaleString('pt-BR', {minimumFractionDigits:2})}`);
        }

        const ctx = document.getElementById('monthlySalesChart').getContext('2d');
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartLabels,
                datasets: [{
                    label: 'Valor de Vendas (R$)',
                    data: chartData,
                    borderWidth: 1,
                    backgroundColor: '#F68811'
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) { 
                                return 'R$ ' + value.toLocaleString('pt-BR'); 
                            }
                        }
                    }
                },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return 'R$ ' + Number(context.parsed.y).toLocaleString('pt-BR', {minimumFractionDigits:2});
                            }
                        }
                    }
                }
            }
        });
    }
</script>

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