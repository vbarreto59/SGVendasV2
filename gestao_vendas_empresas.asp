<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO DE BANCOS DE DADOS
' ===============================================

' Extrair caminhos dos bancos de dados
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunSalesPath = Split(StrConnSales, "Data Source=")(1)

' Verificar se os caminhos foram extraídos corretamente
If dbSunnyPath = "" Or dbSunSalesPath = "" Then
    Response.Write "Erro: Não foi possível extrair caminhos dos bancos de dados"
    Response.End
End If

' ===============================================
' INICIALIZAR CONEXÕES
' As conexões agora são abertas uma vez e usadas por todo o script.
' ===============================================

Dim connSales, connSunny
Set connSales = Server.CreateObject("ADODB.Connection")
Set connSunny = Server.CreateObject("ADODB.Connection")

On Error Resume Next
connSales.Open StrConnSales
connSunny.Open StrConn

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar aos bancos de dados: " & Err.Description
    Response.End
End If
On Error GoTo 0

'=============================================

' Esta é a instrução SQL para atualizar o campo Semestre.
sql = "UPDATE Vendas " & _
      "SET Semestre = SWITCH(" & _
      "    Trimestre IN (1, 2), 1, " & _
      "    Trimestre IN (3, 4), 2" & _
      ") " & _
      "WHERE Trimestre IS NOT NULL;"

' Assumimos que a conexão de banco de dados 'connSales' já está aberta.
' connSales é o objeto de conexão do banco de dados sunsales.mdb.
On Error Resume Next
connSales.Execute sql

' Verificação de erros.
If Err.Number <> 0 Then
    Response.Write "Ocorreu um erro ao atualizar o campo Semestre: " & Err.Description
Else
    ' Response.Write "O campo Semestre foi atualizado com sucesso para todos os registros."
End If
On Error GoTo 0

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

' Função otimizada para obter valores únicos de qualquer tabela com ou sem JOIN
Function GetUniqueValuesAdvanced(tableName, columnName, whereClause, joinClause)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    ' Construir a consulta SQL dinamicamente
    sqlQuery = "SELECT DISTINCT " & columnName & " FROM " & tableName & " "
    If Not IsEmpty(joinClause) Then
        sqlQuery = sqlQuery & joinClause & " "
    End If
    sqlQuery = sqlQuery & whereClause & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = connSales.Execute(sqlQuery)
    If Err.Number <> 0 Then
        ' Em caso de erro, retorna um array vazio e trata o erro na execução
        GetUniqueValuesAdvanced = Array()
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
        GetUniqueValuesAdvanced = dict.Keys
    Else
        GetUniqueValuesAdvanced = Array()
    End If
End Function

Function RemoverNumeros(texto)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    RemoverNumeros = Trim(Replace(regex.Replace(texto, ""), " ", " "))
End Function

Sub ProcessaComissoes(mainDict, key, comissaoValor)
    If Not mainDict.Exists(key) Then mainDict.Add key, 0
    mainDict(key) = mainDict(key) + comissaoValor
End Sub

Function SortDictionaryByValue(dict)
    Dim arrKeys, i, j, temp
    If dict.Count > 0 Then
        arrKeys = dict.Keys
        For i = 0 To UBound(arrKeys)
            For j = i + 1 To UBound(arrKeys)
                If dict(arrKeys(i)) < dict(arrKeys(j)) Then
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

' ===============================================
' OBTER PARÂMETROS DE FILTRO E POPULAR VARIÁVEIS
' ===============================================

' Obter parâmetros de filtro
Dim filtroAno, filtroMes, filtroTrimestre, filtroSemestre
Dim filtroDiretoria, filtroGerencia, filtroCorretor
Dim filtroEmpreendimento, filtroEmpresa

filtroAno = Request.QueryString("ano")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroSemestre = Request.QueryString("semestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")
filtroEmpresa = Request.QueryString("empresa")

' Construir cláusula WHERE
Dim sqlWhere
sqlWhere = " WHERE V.Excluido = 0 "

If filtroAno <> "" Then sqlWhere = sqlWhere & " AND V.AnoVenda = " & filtroAno
If filtroMes <> "" Then sqlWhere = sqlWhere & " AND V.MesVenda = " & filtroMes
If filtroTrimestre <> "" Then sqlWhere = sqlWhere & " AND V.Trimestre = " & filtroTrimestre
If filtroSemestre <> "" Then
    sqlWhere = sqlWhere & " AND V.MesVenda IN (" & IIf(filtroSemestre = "1", "1,2,3,4,5,6", "7,8,9,10,11,12") & ")"
End If
If filtroDiretoria <> "" Then sqlWhere = sqlWhere & " AND V.Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
If filtroGerencia <> "" Then sqlWhere = sqlWhere & " AND V.Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
If filtroCorretor <> "" Then sqlWhere = sqlWhere & " AND V.Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
If filtroEmpreendimento <> "" Then sqlWhere = sqlWhere & " AND E.NomeEmpreendimento = '" & Replace(filtroEmpreendimento, "'", "''") & "'"
If filtroEmpresa <> "" Then sqlWhere = sqlWhere & " AND E.NomeEmpresa = '" & Replace(filtroEmpresa, "'", "''") & "'"

'-------------------------------------------------------------------------------'

' Consulta SQL principal
Dim sql
sql = "SELECT V.*, E.NomeEmpreendimento, E.ComissaoVenda, E.NomeEmpresa " & _
      "FROM Vendas AS V " & _
      "INNER JOIN [;DATABASE=" & dbSunnyPath & "].Empreendimento AS E " & _
      "ON V.Empreend_ID = E.Empreend_ID " & _
      sqlWhere & " AND E.ComissaoVenda IS NOT NULL AND E.ComissaoVenda > 0 " & _
      "ORDER BY V.ID DESC"

' Execução da consulta principal
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rs.Open sql, connSales

If Err.Number <> 0 Then
    Response.Write "Erro na consulta: " & Err.Description & "<br>"
    Response.Write "SQL: " & Server.HTMLEncode(sql)
    Response.End
End If
On Error GoTo 0

'-----------------------------------------'

' Processar resultados
Dim kpiData, totalVGV, totalComissoes, totalComissoesDiretoria, totalComissoesGerencia, totalComissoesCorretor, totalUnidadesVendidas
Set kpiData = Server.CreateObject("Scripting.Dictionary")
Set kpiData("Mes") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopDiretorias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopGerencias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopCorretores") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpreendimentos") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpresas") = Server.CreateObject("Scripting.Dictionary")

' Inicializar as novas variáveis para os totais
totalVGV = 0
totalComissoes = 0
totalComissoesDiretoria = 0
totalComissoesGerencia = 0
totalComissoesCorretor = 0
totalUnidadesVendidas = 0 ' Nova variável para contar unidades vendidas

If Not rs.EOF Then
    Do While Not rs.EOF
        Dim valorUnidade, comissaoVenda, comissaoImobiliaria
        Dim comissaoDiretor, comissaoGerente, comissaoCorretor
        Dim mes, diretoria, gerencia, corretor, empreendimento, empresa
        
        valorUnidade = CDbl(rs("ValorUnidade"))
        
        ' Verificar se ComissaoVenda é válido
        If IsNull(rs("ComissaoVenda")) Or CDbl(rs("ComissaoVenda")) <= 0 Then
            ' Ignorar registro e registrar para depuração (opcional)
            Response.Write "<!-- Aviso: Venda ID " & rs("ID") & " ignorada devido a ComissaoVenda nula ou zero -->"
            rs.MoveNext
            Exit Do ' Sai do loop atual e vai para a próxima iteração
        End If
        
        comissaoVenda = CDbl(rs("ComissaoVenda"))
        
        comissaoImobiliaria = (valorUnidade * comissaoVenda) / 100
        
        ' Garantir que comissaoImobiliaria seja maior que zero
        If comissaoImobiliaria <= 0 Then
            Response.Write "<!-- Aviso: Venda ID " & rs("ID") & " ignorada devido a comissaoImobiliaria zero -->"
            rs.MoveNext
            Exit Do ' Sai do loop atual e vai para a próxima iteração
        End If
        
        comissaoDiretor = comissaoImobiliaria * 0.05
        comissaoGerente = comissaoImobiliaria * 0.10
        comissaoCorretor = comissaoImobiliaria * 0.35
        
        mes = CStr(rs("MesVenda"))
        diretoria = CStr(rs("Diretoria"))
        gerencia = CStr(rs("Gerencia"))
        corretor = CStr(rs("Corretor"))
        empreendimento = CStr(rs("E.NomeEmpreendimento"))
        empresa = CStr(rs("NomeEmpresa"))

        ' Atualizar KPIs apenas se os campos não forem nulos
        If Not IsEmpty(mes) Then Call ProcessaComissoes(kpiData("Mes"), mes, comissaoImobiliaria)
        If Not IsEmpty(diretoria) Then Call ProcessaComissoes(kpiData("TopDiretorias"), diretoria, comissaoDiretor)
        If Not IsEmpty(gerencia) Then Call ProcessaComissoes(kpiData("TopGerencias"), gerencia, comissaoGerente)
        If Not IsEmpty(corretor) Then Call ProcessaComissoes(kpiData("TopCorretores"), corretor, comissaoCorretor)
        If Not IsEmpty(empreendimento) Then Call ProcessaComissoes(kpiData("TopEmpreendimentos"), empreendimento, comissaoImobiliaria)
        If Not IsEmpty(empresa) Then Call ProcessaComissoes(kpiData("TopEmpresas"), empresa, comissaoImobiliaria)

        ' Atualizar totais
        totalVGV = totalVGV + valorUnidade
        totalComissoes = totalComissoes + comissaoImobiliaria
        totalComissoesDiretoria = totalComissoesDiretoria + comissaoDiretor
        totalComissoesGerencia = totalComissoesGerencia + comissaoGerente
        totalComissoesCorretor = totalComissoesCorretor + comissaoCorretor
        totalUnidadesVendidas = totalUnidadesVendidas + 1 ' Incrementar total de unidades vendidas

        rs.MoveNext
    Loop
End If

' Fechar recordset principal
rs.Close
Set rs = Nothing

' ===============================================
' POPULAR OS SELECTS DO FORMULÁRIO
' ===============================================

Dim uniqueAnos, uniqueMeses, uniqueTrimestres, uniqueDiretorias, uniqueGerencias, uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas
Dim empreendimentoJoinClause, empreendimentoWhereClause

' Definir cláusulas para consultas com JOIN
empreendimentoJoinClause = " INNER JOIN [;DATABASE=" & dbSunnyPath & "].Empreendimento AS E ON V.Empreend_ID = E.Empreend_ID "
empreendimentoWhereClause = "WHERE E.NomeEmpreendimento IS NOT NULL"
empresaWhereClause = "WHERE E.NomeEmpresa IS NOT NULL"

' Popule os arrays de valores únicos usando a nova função
uniqueAnos = GetUniqueValuesAdvanced("Vendas AS V", "V.AnoVenda", " WHERE V.AnoVenda IS NOT NULL", "")
uniqueMeses = GetUniqueValuesAdvanced("Vendas AS V", "V.MesVenda", " WHERE V.MesVenda IS NOT NULL", "")
uniqueTrimestres = GetUniqueValuesAdvanced("Vendas AS V", "V.Trimestre", " WHERE V.Trimestre IS NOT NULL", "")
uniqueDiretorias = GetUniqueValuesAdvanced("Vendas AS V", "V.Diretoria", " WHERE V.Diretoria IS NOT NULL", "")
uniqueGerencias = GetUniqueValuesAdvanced("Vendas AS V", "V.Gerencia", " WHERE V.Gerencia IS NOT NULL", "")
uniqueCorretores = GetUniqueValuesAdvanced("Vendas AS V", "V.Corretor", " WHERE V.Corretor IS NOT NULL", "")
uniqueEmpreendimentos = GetUniqueValuesAdvanced("Vendas AS V", "E.NomeEmpreendimento", empreendimentoWhereClause, empreendimentoJoinClause)
uniqueEmpresas = GetUniqueValuesAdvanced("Vendas AS V", "E.NomeEmpresa", empresaWhereClause, empreendimentoJoinClause)

' Preparar dados para gráficos
Dim arrMesesNome(12), chartLabels(11), chartData(11)
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

For i = 1 To 12
    chartLabels(i-1) = arrMesesNome(i)
    
    If kpiData("Mes").Exists(CStr(i)) Then
        chartData(i-1) = kpiData("Mes")(CStr(i))
    Else
        chartData(i-1) = 0
    End If
Next
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Relatório de Comissões Completo</title>
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
        .filter-row {
            margin-bottom: 10px;
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        .filter-select {
            width: 100%;
        }
        .filter-btn {
            margin-top: 25px;
        }
        .text-center-v { text-align: center; }
        .text-right-v { text-align: right; }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;"><i class="fas fa-hand-holding-usd"></i> SGVendas - Relatório de Comissões Completo</h2>
        
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row filter-row">
                    <div class="col-md-2">
                        <div class="filter-label">Ano</div>
                        <select class="form-select filter-select" name="ano" id="anoFilter" onchange="this.form.submit()">
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

                    <div class="col-md-2">
                        <div class="filter-label">Semestre</div>
                        <select class="form-select filter-select" name="semestre" id="semestreFilter" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <option value="1" <% If filtroSemestre = "1" Then Response.Write "selected" %>>1º Semestre</option>
                            <option value="2" <% If filtroSemestre = "2" Then Response.Write "selected" %>>2º Semestre</option>
                        </select>
                    </div>

                    <div class="col-md-2">
                        <div class="filter-label">Mês</div>
                        <select class="form-select filter-select" name="mes" id="mesFilter" onchange="this.form.submit()">
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

                    <div class="col-md-2">
                        <div class="filter-label">Trimestre</div>
                        <select class="form-select filter-select" name="trimestre" id="trimestreFilter" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueTrimestres) Then
                                For Each trimestre In uniqueTrimestres
                                    Response.Write "<option value=""" & trimestre & """"
                                    If filtroTrimestre = trimestre Then Response.Write " selected"
                                    Response.Write ">" & trimestre & "</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-2">
                        <div class="filter-label">Diretoria</div>
                        <select class="form-select filter-select" name="diretoria" id="diretoriaFilter" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueDiretorias) Then
                                For Each diretoria In uniqueDiretorias
                                    Response.Write "<option value=""" & diretoria & """"
                                    If filtroDiretoria = diretoria Then Response.Write " selected"
                                    Response.Write ">" & diretoria & "</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-2">
                        <div class="filter-label">Gerência</div>
                        <select class="form-select filter-select" name="gerencia" id="gerenciaFilter" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueGerencias) Then
                                For Each gerencia In uniqueGerencias
                                    Response.Write "<option value=""" & gerencia & """"
                                    If filtroGerencia = gerencia Then Response.Write " selected"
                                    Response.Write ">" & gerencia & "</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                </div>
                <div class="row filter-row mt-3">
                    <div class="col-md-2">
                        <div class="filter-label">Corretor</div>
                        <select class="form-select filter-select" name="corretor" id="corretorFilter" onchange="this.form.submit()">
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
                    
                    <div class="col-md-2">
                        <div class="filter-label">Empreendimento</div>
                        <select class="form-select filter-select" name="empreendimento" id="empreendimentoFilter" onchange="this.form.submit()">
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
                    
                    <div class="col-md-2">
                        <div class="filter-label">Empresa</div>
                        <select class="form-select filter-select" name="empresa" id="empresaFilter" onchange="this.form.submit()">
                            <option value="">Todos</option>
                            <%
                            If IsArray(uniqueEmpresas) Then
                                For Each empresa In uniqueEmpresas
                                    Response.Write "<option value=""" & empresa & """"
                                    If filtroEmpresa = empresa Then Response.Write " selected"
                                    Response.Write ">" & empresa & "</option>"
                                Next
                            End If
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-6 text-end">
                        <button type="button" class="btn btn-secondary filter-btn" onclick="limparFiltros()">
                            <i class="fas fa-times"></i> Limpar Filtros
                        </button>
                    </div>
                </div>
            </form>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-4">
                <div class="kpi-card bg-success-kpi">
                    <i class="fas fa-handshake"></i>
                    <h5>Total de VGV</h5>
                    <p>R$ <%= FormatNumber(totalVGV, 2) %></p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="kpi-card bg-info-kpi">
                    <i class="fas fa-user-tie"></i>
                    <h5>Total de Comissões (Diretoria)</h5>
                    <p>R$ <%= FormatNumber(totalComissoesDiretoria, 2) %></p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="kpi-card bg-warning-kpi">
                    <i class="fas fa-users-cog"></i>
                    <h5>Total de Comissões (Gerência)</h5>
                    <p>R$ <%= FormatNumber(totalComissoesGerencia, 2) %></p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="kpi-card bg-danger-kpi">
                    <i class="fas fa-user-tag"></i>
                    <h5>Total de Comissões (Corretor)</h5>
                    <p>R$ <%= FormatNumber(totalComissoesCorretor, 2) %></p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="kpi-card bg-secondary-kpi">
                    <i class="fas fa-home"></i>
                    <h5>Total de Unidades Vendidas</h5>
                    <p><%= totalUnidadesVendidas %></p>
                </div>
            </div>
        </div>
        
        <!-- -------------------------------------------------------------- -->
        <h2 class="text-white mt-5">Relatório Completo de Comissões</h2>
        <div class="card-kpi p-3 rounded">
            <div class="row">
                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Todas as Diretorias (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-primary-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Diretoria</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topDiretorias, arrDiretorias
                                Set topDiretorias = kpiData("TopDiretorias")
                                arrDiretorias = SortDictionaryByValue(topDiretorias)
                                If IsArray(arrDiretorias) Then
                                    For i = 0 To UBound(arrDiretorias)
                                        comissaoDiretor = topDiretorias(arrDiretorias(i))
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrDiretorias(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(comissaoDiretor, 2) %></td>
                                </tr>
                                <%
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Todas as Gerências (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-success-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Gerência</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topGerencias, arrGerencias
                                Set topGerencias = kpiData("TopGerencias")
                                arrGerencias = SortDictionaryByValue(topGerencias)
                                If IsArray(arrGerencias) Then
                                    For i = 0 To UBound(arrGerencias)
                                        comissaoGerencia = topGerencias(arrGerencias(i))
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrGerencias(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(comissaoGerencia, 2) %></td>
                                </tr>
                                <%
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Todos os Corretores (Comissão)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-info-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Corretor</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topCorretores, arrCorretores
                                Set topCorretores = kpiData("TopCorretores")
                                arrCorretores = SortDictionaryByValue(topCorretores)
                                If IsArray(arrCorretores) Then
                                    For i = 0 To UBound(arrCorretores)
                                        comissaoCorretor = topCorretores(arrCorretores(i))
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrCorretores(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(comissaoCorretor, 2) %></td>
                                </tr>
                                <%
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Todos os Empreendimentos (Comissão da Imobiliária)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-warning-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Empreendimento</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topEmpreendimentos, arrEmpreendimentos
                                Set topEmpreendimentos = kpiData("TopEmpreendimentos")
                                arrEmpreendimentos = SortDictionaryByValue(topEmpreendimentos)
                                If IsArray(arrEmpreendimentos) Then
                                    For i = 0 To UBound(arrEmpreendimentos)
                                        Dim comissaoEmpreendimento
                                        comissaoEmpreendimento = topEmpreendimentos(arrEmpreendimentos(i))
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrEmpreendimentos(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(comissaoEmpreendimento, 2) %></td>
                                </tr>
                                <%
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="row">
                <div class="col-md-6 mb-4">
                    <h4 class="text-dark">Todas as Empresas (Comissão da Imobiliária)</h4>
                    <div class="table-responsive">
                        <table class="table table-striped">
                            <thead class="bg-danger-kpi text-white">
                                <tr>
                                    <th>Posição</th>
                                    <th>Empresa</th>
                                    <th class="text-right-v">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim topEmpresas, arrEmpresas
                                Set topEmpresas = kpiData("TopEmpresas")
                                arrEmpresas = SortDictionaryByValue(topEmpresas)
                                If IsArray(arrEmpresas) Then
                                    For i = 0 To UBound(arrEmpresas)
                                        Dim comissaoEmpresa
                                        comissaoEmpresa = topEmpresas(arrEmpresas(i))
                                %>
                                <tr>
                                    <td><%= i + 1 %></td>
                                    <td><%= arrEmpresas(i) %></td>
                                    <td class="text-right-v"><%= FormatNumber(comissaoEmpresa, 2) %></td>
                                </tr>
                                <%
                                    Next
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        function limparFiltros() {
            window.location.href = window.location.pathname;
        }

        // Dados do VBScript para o gráfico
        const chartLabels = [<% For i=0 To UBound(chartLabels) : Response.Write """" & chartLabels(i) & """" : If i < UBound(chartLabels) Then Response.Write "," : End If : Next %>];
        const chartData = [<% For i=0 To UBound(chartData) : Response.Write chartData(i) : If i < UBound(chartData) Then Response.Write "," : End If : Next %>];

        const chartContainer = document.getElementById('chartContainer'); // Adicionei um contêiner ao canvas
        const oldCanvas = document.getElementById('monthlyCommissionsChart');

        // Remove o canvas antigo
        if (oldCanvas) {
            oldCanvas.remove();
        }
        
        // Cria um novo canvas com o mesmo ID
        const newCanvas = document.createElement('canvas');
        newCanvas.id = 'monthlyCommissionsChart';
        chartContainer.appendChild(newCanvas);

        const ctx = newCanvas.getContext('2d');

        if (ctx) {
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: chartLabels,
                    datasets: [{
                        label: 'Total de Comissões (R$)',
                        data: chartData,
                        backgroundColor: 'rgba(128, 0, 0, 0.7)',
                        borderColor: 'rgb(128, 0, 0)',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Valor de Comissão (R$)'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Mês'
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
                                    if (context.parsed.y !== null) {
                                        label += new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(context.parsed.y);
                                    }
                                    return label;
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