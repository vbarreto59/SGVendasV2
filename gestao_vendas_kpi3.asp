<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="gestao_header.inc"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"  
%>
<%
' ===============================================
' CONFIGURAÇÕES INICIAIS
' ===============================================

' Obter caminho do banco de dados
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)

' Mensagem do sistema (se houver)
Dim mensagem
mensagem = Request.QueryString("mensagem")

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

' Função para obter valores únicos de uma coluna
Function GetUniqueValues(conn, tableName, columnName, whereClause)
    Dim dict, rs, sql
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    sql = "SELECT DISTINCT " & columnName & " FROM " & tableName & whereClause & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
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
    
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

' Processa e agrega dados em um dicionário
Sub ProcessData(mainDict, key, vendas, valor, comissao)
    If Not mainDict.Exists(key) Then
        Dim newSubDict
        Set newSubDict = Server.CreateObject("Scripting.Dictionary")
        newSubDict.Add "vendas", 0
        newSubDict.Add "valor", 0
        newSubDict.Add "comissao", 0
        mainDict.Add key, newSubDict
    End If
    mainDict(key)("vendas") = mainDict(key)("vendas") + vendas
    mainDict(key)("valor") = mainDict(key)("valor") + valor
    mainDict(key)("comissao") = mainDict(key)("comissao") + comissao
End Sub

' Ordena dicionário por valor específico (decrescente)
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

' Ordena dicionário por chave (crescente)
Function SortDictionaryByKey(dict)
    Dim arrKeys, i, j, temp
    If dict.Count > 0 Then
        arrKeys = dict.Keys
        For i = 0 To UBound(arrKeys)
            For j = i + 1 To UBound(arrKeys)
                If CInt(arrKeys(i)) > CInt(arrKeys(j)) Then
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

' ===============================================
' PROCESSAMENTO PRINCIPAL
' ===============================================

' Abre conexões com os bancos de dados
Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Obtém filtros da query string
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

' Buscar valores únicos para os filtros
Dim uniqueAnos, uniqueMeses, uniqueTrimestres, uniqueDiretorias, uniqueGerencias
Dim uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas
Dim arrMesesNome(12)
Dim kpiData
Dim categories, cat

uniqueAnos = GetUniqueValues(connSales, "Vendas", "AnoVenda", " WHERE Excluido = 0")
uniqueMeses = GetUniqueValues(connSales, "Vendas", "MesVenda", " WHERE Excluido = 0")
uniqueTrimestres = GetUniqueValues(connSales, "Vendas", "Trimestre", " WHERE Excluido = 0")
uniqueDiretorias = GetUniqueValues(connSales, "Vendas", "Diretoria", " WHERE Excluido = 0 AND Diretoria IS NOT NULL")
uniqueGerencias = GetUniqueValues(connSales, "Vendas", "Gerencia", " WHERE Excluido = 0 AND Gerencia IS NOT NULL")
uniqueCorretores = GetUniqueValues(connSales, "Vendas", "Corretor", " WHERE Excluido = 0 AND Corretor IS NOT NULL")
uniqueEmpreendimentos = GetUniqueValues(conn, "Empreendimento", "NomeEmpreendimento", " WHERE Excluido = 0")
uniqueEmpresas = GetUniqueValues(conn, "Empresa", "NomeEmpresa", " WHERE Excluido = 0")

' Nomes dos meses para exibição
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

' Inicializa a estrutura kpiData
Set kpiData = Server.CreateObject("Scripting.Dictionary")

' Inicializa todas as categorias de KPIs
categories = Array("Ano", "Semestre", "Trimestre", "Mes", "TopCorretores", "TopDiretorias", "TopGerencias", "TopEmpreendimentos", "TopEmpresas")

For Each cat In categories
    Set kpiData(cat) = Server.CreateObject("Scripting.Dictionary")
Next

' Consulta principal de vendas
Dim sqlVendas, rsVendas
sqlVendas = "SELECT * FROM Vendas WHERE Excluido = 0"

' Aplica filtros
If filtroAno <> "" Then sqlVendas = sqlVendas & " AND AnoVenda = " & filtroAno
If filtroSemestre <> "" Then
    If filtroSemestre = "1" Then
        sqlVendas = sqlVendas & " AND MesVenda >= 1 AND MesVenda <= 6"
    ElseIf filtroSemestre = "2" Then
        sqlVendas = sqlVendas & " AND MesVenda >= 7 AND MesVenda <= 12"
    End If
End If
If filtroMes <> "" Then sqlVendas = sqlVendas & " AND MesVenda = " & filtroMes
If filtroTrimestre <> "" Then sqlVendas = sqlVendas & " AND Trimestre = " & filtroTrimestre
If filtroDiretoria <> "" Then sqlVendas = sqlVendas & " AND Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
If filtroGerencia <> "" Then sqlVendas = sqlVendas & " AND Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
If filtroCorretor <> "" Then sqlVendas = sqlVendas & " AND Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"

sqlVendas = sqlVendas & " ORDER BY ID DESC"

Set rsVendas = connSales.Execute(sqlVendas)

' Processa os dados das vendas
If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        Dim valorUnidade, valorComissao, ano, mes, trimestre, semestre
        Dim diretoria, gerencia, corretor, empreendimento, empresa
        Dim empreend_ID
        
        On Error Resume Next
        valorUnidade = CDbl(rsVendas("ValorUnidade"))
        valorComissao = CDbl(rsVendas("ValorComissaoGeral"))
        empreend_ID = rsVendas("Empreend_ID")
        
        ano = CStr(rsVendas("AnoVenda"))
        mes = CStr(rsVendas("MesVenda"))
        trimestre = CStr(rsVendas("Trimestre"))
        
        If CInt(mes) <= 6 Then
            semestre = "1"
        Else
            semestre = "2"
        End If
        
        diretoria = CStr(rsVendas("Diretoria"))
        gerencia = CStr(rsVendas("Gerencia"))
        corretor = CStr(rsVendas("Corretor"))
        
        ' Busca informações adicionais do empreendimento e empresa
        If Not IsNull(empreend_ID) Then
            Dim sqlEmp, rsEmp
            sqlEmp = "SELECT NomeEmpreendimento, Empresa_ID FROM Empreendimento WHERE Empreend_ID = " & empreend_ID
            Set rsEmp = conn.Execute(sqlEmp)
            
            If Not rsEmp.EOF Then
                empreendimento = CStr(rsEmp("NomeEmpreendimento"))
                
                ' Busca nome da empresa
                Dim sqlEmpresa, rsEmpresa
                sqlEmpresa = "SELECT NomeEmpresa FROM Empresa WHERE Empresa_ID = " & rsEmp("Empresa_ID")
                Set rsEmpresa = conn.Execute(sqlEmpresa)
                
                If Not rsEmpresa.EOF Then
                    empresa = CStr(rsEmpresa("NomeEmpresa"))
                End If
                
                If Not rsEmpresa Is Nothing Then rsEmpresa.Close
                Set rsEmpresa = Nothing
            End If
            
            If Not rsEmp Is Nothing Then rsEmp.Close
            Set rsEmp = Nothing
        End If
        On Error GoTo 0
        
        ' Atualiza KPIs
        Call ProcessData(kpiData("Ano"), ano, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("Semestre"), semestre, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("Trimestre"), trimestre, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("Mes"), mes, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("TopCorretores"), corretor, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("TopDiretorias"), diretoria, 1, valorUnidade, valorComissao)
        Call ProcessData(kpiData("TopGerencias"), gerencia, 1, valorUnidade, valorComissao)
        
        If empreendimento <> "" Then 
            Call ProcessData(kpiData("TopEmpreendimentos"), empreendimento, 1, valorUnidade, valorComissao)
        End If
        If empresa <> "" Then 
            Call ProcessData(kpiData("TopEmpresas"), empresa, 1, valorUnidade, valorComissao)
        End If
        
        rsVendas.MoveNext
    Loop
End If

' Fecha recordsets
If Not rsVendas Is Nothing Then rsVendas.Close
Set rsVendas = Nothing

' =========================================================
' Prepara dados para o gráfico de vendas por mês e o filtro
' =========================================================
Dim chartLabels, chartData, mesesFiltrados, totalMeses, j, i, mesNum

' Inicializa mesesFiltrados como um array vazio para evitar o erro de tipo
mesesFiltrados = Array()

' Define quais meses devem ser exibidos com base no filtro de semestre.
' O array é sempre preenchido aqui.
If CStr(filtroSemestre) = "1" Then
    mesesFiltrados = Array(1, 2, 3, 4, 5, 6)
ElseIf CStr(filtroSemestre) = "2" Then
    mesesFiltrados = Array(7, 8, 9, 10, 11, 12)
Else
    mesesFiltrados = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
End If

' Verifica se o array é válido e dimensiona os arrays do gráfico com segurança.
If IsArray(mesesFiltrados) Then
    totalMeses = UBound(mesesFiltrados) - LBound(mesesFiltrados) + 1
Else
    totalMeses = 0
End If

If totalMeses > 0 Then
    ReDim chartLabels(totalMeses - 1), chartData(totalMeses - 1)
Else
    ' Se não houver meses para exibir, redimensione para -1 para criar um array vazio.
    ReDim chartLabels(-1), chartData(-1)
End If

For j = 0 To totalMeses - 1
    i = mesesFiltrados(j)
    chartLabels(j) = arrMesesNome(i)
    chartData(j) = 0
    If kpiData("Mes").Exists(CStr(i)) Then
        chartData(j) = kpiData("Mes")(CStr(i))("valor")
    End If
Next
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #f8f9fa;
            margin-top: 20px;
   
        }
        .card-kpi {
            background-color: white;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-container {
            background-color: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .kpi-card {
            text-align: center;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 8px;
            background-color: #800000;
            color: white;
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
        <h3 class="card-title mt-4 mb-1" style="margin-top: 1.5rem !important; color: black;">
            <i class="fas fa-handshake"></i> Filtro de Vendas
        </h3>
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row filter-row">
                    <div class="col-md-2">
                        <div class="filter-label">Ano</div>
                        <select class="form-select" name="ano" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueAnos) Then
                                For i = 0 To UBound(uniqueAnos)
                                    If Not IsEmpty(uniqueAnos(i)) And Not IsNull(uniqueAnos(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueAnos(i)) & """"
                                        If CStr(filtroAno) = CStr(uniqueAnos(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueAnos(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Semestre</div>
                        <select class="form-select" name="semestre" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <option value="1" <% If CStr(filtroSemestre) = "1" Then Response.Write "selected" %>>1º Semestre</option>
                            <option value="2" <% If CStr(filtroSemestre) = "2" Then Response.Write "selected" %>>2º Semestre</option>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Mês</div>
                        <select class="form-select" name="mes" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(mesesFiltrados) Then
                                For i = LBound(mesesFiltrados) To UBound(mesesFiltrados)
                                    mesNum = mesesFiltrados(i)
                                    If Not IsEmpty(mesNum) And Not IsNull(mesNum) Then
                                        Response.Write "<option value=""" & mesNum & """"
                                        If CStr(filtroMes) = CStr(mesNum) Then Response.Write " selected"
                                        Response.Write ">" & arrMesesNome(mesNum) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Trimestre</div>
                        <select class="form-select" name="trimestre" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueTrimestres) Then
                                For i = 0 To UBound(uniqueTrimestres)
                                    If Not IsEmpty(uniqueTrimestres(i)) And Not IsNull(uniqueTrimestres(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueTrimestres(i)) & """"
                                        If CStr(filtroTrimestre) = CStr(uniqueTrimestres(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueTrimestres(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Diretoria</div>
                        <select class="form-select" name="diretoria" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueDiretorias) Then
                                For i = 0 To UBound(uniqueDiretorias)
                                    If Not IsEmpty(uniqueDiretorias(i)) And Not IsNull(uniqueDiretorias(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueDiretorias(i)) & """"
                                        If CStr(filtroDiretoria) = CStr(uniqueDiretorias(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueDiretorias(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Gerência</div>
                        <select class="form-select" name="gerencia" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueGerencias) Then
                                For i = 0 To UBound(uniqueGerencias)
                                    If Not IsEmpty(uniqueGerencias(i)) And Not IsNull(uniqueGerencias(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueGerencias(i)) & """"
                                        If CStr(filtroGerencia) = CStr(uniqueGerencias(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueGerencias(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>
                </div>

                <div class="row filter-row mt-3">
                    <div class="col-md-2">
                        <div class="filter-label">Corretor</div>
                        <select class="form-select" name="corretor" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueCorretores) Then
                                For i = 0 To UBound(uniqueCorretores)
                                    If Not IsEmpty(uniqueCorretores(i)) And Not IsNull(uniqueCorretores(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueCorretores(i)) & """"
                                        If CStr(filtroCorretor) = CStr(uniqueCorretores(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueCorretores(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Empreendimento</div>
                        <select class="form-select" name="empreendimento" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueEmpreendimentos) Then
                                For i = 0 To UBound(uniqueEmpreendimentos)
                                    If Not IsEmpty(uniqueEmpreendimentos(i)) And Not IsNull(uniqueEmpreendimentos(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueEmpreendimentos(i)) & """"
                                        If CStr(filtroEmpreendimento) = CStr(uniqueEmpreendimentos(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueEmpreendimentos(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Empresa</div>
                        <select class="form-select" name="empresa" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueEmpresas) Then
                                For i = 0 To UBound(uniqueEmpresas)
                                    If Not IsEmpty(uniqueEmpresas(i)) And Not IsNull(uniqueEmpresas(i)) Then
                                        Response.Write "<option value=""" & Server.HTMLEncode(uniqueEmpresas(i)) & """"
                                        If CStr(filtroEmpresa) = CStr(uniqueEmpresas(i)) Then Response.Write " selected"
                                        Response.Write ">" & Server.HTMLEncode(uniqueEmpresas(i)) & "</option>"
                                    End If
                                Next
                            End If
                            %>
                        </select>
                    </div>

                    <div class="col-md-6 text-end">
                        <button type="button" class="btn btn-secondary" onclick="limparFiltros()">
                            <i class="fas fa-times"></i> Limpar Filtros
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <div class="card-kpi">
            <h3>KPIs de Vendas</h3>
            
            <h4 class="mt-4">Vendas por Ano</h4>
            <div class="row">
                <%
                If kpiData("Ano").Count > 0 Then
                    Dim arrAnos
                    arrAnos = SortDictionaryByKey(kpiData("Ano"))
                    
                    For Each ano In arrAnos
                        Dim anoData
                        Set anoData = kpiData("Ano")(ano)
                %>
                <div class="col-md-3">
                    <div class="kpi-card">
                        <h5>Ano <%= ano %></h5>
                        <p>QTD: <%= anoData("vendas") %></p>
                        <p>VALOR: R$ <%= FormatNumber(anoData("valor"), 2) %></p>
                    </div>
                </div>
                <%
                    Next
                End If
                %>
            </div>
            
            <h4 class="mt-4">Vendas por Mês</h4>
            <div class="row">
                <%
                If IsArray(mesesFiltrados) Then
                    For i = 0 To UBound(mesesFiltrados)
                        Dim mesKey
                        mesKey = CStr(mesesFiltrados(i))
                        
                        Dim vendasMes, valorMes
                        vendasMes = 0
                        valorMes = 0
                        
                        If kpiData("Mes").Exists(mesKey) Then
                            Dim mesData
                            Set mesData = kpiData("Mes")(mesKey)
                            vendasMes = mesData("vendas")
                            valorMes = mesData("valor")
                        End If

                        If vendasMes > 0 Then
                %>
                <div class="col-md-2">
                    <div class="kpi-card">
                        <h5><%= arrMesesNome(CInt(mesKey)) %></h5>
                        <p>QTD: <%= vendasMes %></p>
                        <p>VALOR: R$ <%= FormatNumber(valorMes, 2) %></p>
                    </div>
                </div>
                <%
                        End If
                    Next
                End If
                %>
            </div>
        </div>

        <div class="card-kpi">
            <h3>Gráfico de Vendas Mensais</h3>
            <canvas id="monthlySalesChart" height="100"></canvas>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        function limparFiltros() {
            window.location.href = window.location.pathname;
        }

        // Gráfico de vendas mensais
        const ctx = document.getElementById('monthlySalesChart');
        if (ctx) {
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: [<% For i=0 To UBound(chartLabels) : Response.Write """" & chartLabels(i) & """" : If i < UBound(chartLabels) Then Response.Write "," : End If : Next %>],
                    datasets: [{
                        label: 'Valor de Vendas',
                        data: [<% For i=0 To UBound(chartData) : Response.Write chartData(i) : If i < UBound(chartData) Then Response.Write "," : End If : Next %>],
                        backgroundColor: '#F68811',
                        borderColor: 'black',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    scales: {
                        y: {
                            beginAtZero: true
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
If conn.State = 1 Then conn.Close
If connSales.State = 1 Then connSales.Close
Set conn = Nothing
Set connSales = Nothing
%>