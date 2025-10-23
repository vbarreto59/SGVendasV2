<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->

<%
' ===============================================
' FUNÇÕES E SUB-ROTINAS GLOBAIS - v1
' ===============================================

' Função para remover números e asteriscos de uma string
Function RemoverNumeros(texto)
  Dim regex
  Set regex = New RegExp
  
  regex.Pattern = "[0-9*-]"
  regex.Global = True
  
  RemoverNumeros = Trim(Replace(regex.Replace(texto, ""), " ", " "))
End Function

' Função para processar e agregar dados em um dicionário
' Adaptada para agregar valores de comissão
Sub ProcessaComissoes(mainDict, key, comissaoValor)
  If Not mainDict.Exists(key) Then
    mainDict.Add key, 0
  End If
  mainDict(key) = mainDict(key) + comissaoValor
End Sub

' Função para ordenar as chaves de um dicionário por um valor específico em ordem decrescente
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

' Função para ordenar as chaves de um dicionário (útil para datas)
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

' Função para obter valores únicos de um campo, com base nos filtros
Function GetUniqueValuesWithFilter(conn, fieldName, currentFilterName)
  Dim dict, rs, sqlWhere, sqlQuery, param
  Set dict = Server.CreateObject("Scripting.Dictionary")
  Set rs = Server.CreateObject("ADODB.Recordset")
  
  Dim baseTable, joinClause
  baseTable = "Vendas"
  joinClause = " INNER JOIN (Empreendimento INNER JOIN Empresa ON Empreendimento.Empresa_ID = Empresa.Empresa_ID) ON Vendas.Empreend_ID = Empreendimento.Empreend_ID "
  
  sqlWhere = " WHERE Vendas.Excluido = 0 "
  
  If Request.QueryString("ano") <> "" And currentFilterName <> "Vendas.AnoVenda" Then
    sqlWhere = sqlWhere & " AND Vendas.AnoVenda = " & Request.QueryString("ano")
  End If
  If Request.QueryString("mes") <> "" And currentFilterName <> "Vendas.MesVenda" Then
    sqlWhere = sqlWhere & " AND Vendas.MesVenda = " & Request.QueryString("mes")
  End If
  If Request.QueryString("trimestre") <> "" And currentFilterName <> "Vendas.Trimestre" Then
    sqlWhere = sqlWhere & " AND Vendas.Trimestre = " & Request.QueryString("trimestre")
  End If
  If Request.QueryString("diretoria") <> "" And currentFilterName <> "Vendas.Diretoria" Then
    param = Replace(Request.QueryString("diretoria"), "'", "''")
    sqlWhere = sqlWhere & " AND Vendas.Diretoria = '" & param & "'"
  End If
  If Request.QueryString("gerencia") <> "" And currentFilterName <> "Vendas.Gerencia" Then
    param = Replace(Request.QueryString("gerencia"), "'", "''")
    sqlWhere = sqlWhere & " AND Vendas.Gerencia = '" & param & "'"
  End If
  If Request.QueryString("corretor") <> "" And currentFilterName <> "Vendas.Corretor" Then
    param = Replace(Request.QueryString("corretor"), "'", "''")
    sqlWhere = sqlWhere & " AND Vendas.Corretor = '" & param & "'"
  End If
  If Request.QueryString("empreendimento") <> "" And currentFilterName <> "Empreendimento.NomeEmpreendimento" Then
    param = Replace(Request.QueryString("empreendimento"), "'", "''")
    sqlWhere = sqlWhere & " AND Empreendimento.NomeEmpreendimento = '" & param & "'"
  End If
  If Request.QueryString("empresa") <> "" And currentFilterName <> "Empresa.NomeEmpresa" Then
    param = Replace(Request.QueryString("empresa"), "'", "''")
    sqlWhere = sqlWhere & " AND Empresa.NomeEmpresa = '" & param & "'"
  End If
  
  sqlQuery = "SELECT DISTINCT " & fieldName & " FROM " & baseTable & joinClause & sqlWhere & " ORDER BY " & fieldName & ";"
  
  On Error Resume Next
  rs.Open sqlQuery, conn
  If Err.Number <> 0 Then
    Response.Write "Erro ao obter valores únicos: " & Err.Description & "<br>"
    Response.Write "SQL Query: " & sqlQuery & "<br>"
    Err.Clear
    GetUniqueValuesWithFilter = Array()
    Exit Function
  End If
  On Error Goto 0
  
  Dim cleanFieldName
  cleanFieldName = fieldName
  If InStr(fieldName, ".") > 0 Then
      cleanFieldName = Mid(fieldName, InStr(fieldName, ".") + 1)
  End If

  If Not rs.EOF Then
    Do While Not rs.EOF
      If Not IsNull(rs(cleanFieldName)) Then
        Dim value
        value = CStr(rs(cleanFieldName))
        If Not dict.Exists(value) Then
          dict.Add value, 1
        End If
      End If
      rs.MoveNext
    Loop
  End If
  
  rs.Close
  Set rs = Nothing
  
  GetUniqueValuesWithFilter = dict.Keys
End Function


' ===============================================
' PROCESSAMENTO DE DADOS COM FILTROS
' ===============================================

Dim mensagem
mensagem = Request.QueryString("mensagem")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

' Obter filtros atuais da query string
Dim filtroAno, filtroSemestre, filtroMes, filtroTrimestre, filtroDiretoria, filtroGerencia, filtroCorretor, filtroEmpreendimento, filtroEmpresa
filtroAno = Request.QueryString("ano")
filtroSemestre = Request.QueryString("semestre")
filtroMes = Request.QueryString("mes")
filtroTrimestre = Request.QueryString("trimestre")
filtroDiretoria = Request.QueryString("diretoria")
filtroGerencia = Request.QueryString("gerencia")
filtroCorretor = Request.QueryString("corretor")
filtroEmpreendimento = Request.QueryString("empreendimento")
filtroEmpresa = Request.QueryString("empresa")


' Popula as listas de filtros dinamicamente
Dim uniqueAnos, uniqueMeses, uniqueTrimestres, uniqueDiretorias, uniqueGerencias, uniqueCorretores, uniqueEmpreendimentos, uniqueEmpresas
uniqueAnos = GetUniqueValuesWithFilter(conn, "Vendas.AnoVenda", "Vendas.AnoVenda")
uniqueMeses = GetUniqueValuesWithFilter(conn, "Vendas.MesVenda", "Vendas.MesVenda")
uniqueTrimestres = GetUniqueValuesWithFilter(conn, "Vendas.Trimestre", "Vendas.Trimestre")
uniqueDiretorias = GetUniqueValuesWithFilter(conn, "Vendas.Diretoria", "Vendas.Diretoria")
uniqueGerencias = GetUniqueValuesWithFilter(conn, "Vendas.Gerencia", "Vendas.Gerencia")
uniqueCorretores = GetUniqueValuesWithFilter(conn, "Vendas.Corretor", "Vendas.Corretor")
uniqueEmpreendimentos = GetUniqueValuesWithFilter(conn, "Empreendimento.NomeEmpreendimento", "Empreendimento.NomeEmpreendimento")
uniqueEmpresas = GetUniqueValuesWithFilter(conn, "Empresa.NomeEmpresa", "Empresa.NomeEmpresa")


' Constrói a cláusula WHERE para a query principal
Dim sqlWhere
sqlWhere = " WHERE Vendas.Excluido = 0 "

If filtroAno <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.AnoVenda = " & filtroAno
End If

If filtroSemestre <> "" Then
  If filtroSemestre = "1" Then
    sqlWhere = sqlWhere & " AND Vendas.MesVenda IN (1, 2, 3, 4, 5, 6)"
  ElseIf filtroSemestre = "2" Then
    sqlWhere = sqlWhere & " AND Vendas.MesVenda IN (7, 8, 9, 10, 11, 12)"
  End If
End If

If filtroMes <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.MesVenda = " & filtroMes
End If
If filtroTrimestre <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.Trimestre = " & filtroTrimestre
End If
If filtroDiretoria <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.Diretoria = '" & Replace(filtroDiretoria, "'", "''") & "'"
End If
If filtroGerencia <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.Gerencia = '" & Replace(filtroGerencia, "'", "''") & "'"
End If
If filtroCorretor <> "" Then
  sqlWhere = sqlWhere & " AND Vendas.Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
End If
If filtroEmpreendimento <> "" Then
  sqlWhere = sqlWhere & " AND Empreendimento.NomeEmpreendimento = '" & Replace(filtroEmpreendimento, "'", "''") & "'"
End If
If filtroEmpresa <> "" Then
  sqlWhere = sqlWhere & " AND Empresa.NomeEmpresa = '" & Replace(filtroEmpresa, "'", "''") & "'"
End If


' SQL query para recuperar os dados filtrados com os JOINS, incluindo o campo de comissão
Dim sql
sql = "SELECT Vendas.*, Empreendimento.NomeEmpreendimento AS NomeEmpreendimento, Empreendimento.ComissaoVenda, Empresa.NomeEmpresa AS NomeEmpresa FROM (Vendas INNER JOIN Empreendimento ON Vendas.Empreend_ID = Empreendimento.Empreend_ID) INNER JOIN Empresa ON Empreendimento.Empresa_ID = Empresa.Empresa_ID" & sqlWhere & " ORDER BY Vendas.ID DESC;"

' Cria e abre o Recordset com a nova query
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

' Dicionários para armazenar os dados agregados de comissão
Dim kpiData
Set kpiData = Server.CreateObject("Scripting.Dictionary")
Set kpiData("Mes") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopDiretorias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopGerencias") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopCorretores") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpreendimentos") = Server.CreateObject("Scripting.Dictionary")
Set kpiData("TopEmpresas") = Server.CreateObject("Scripting.Dictionary")
Dim totalComissoesDiretoria, totalComissoesGerencia, totalComissoesCorretor, totalComissoesImobiliaria, totalVGV
totalComissoesDiretoria = 0
totalComissoesGerencia = 0
totalComissoesCorretor = 0
totalComissoesImobiliaria = 0
totalVGV = 0


' Processa os dados do Recordset e calcula as comissões
If Not rs.EOF Then
  Do While Not rs.EOF
    Dim valorUnidade, comissaoVenda, comissaoImobiliaria
    Dim ano, mes, trimestre, semestre, diretoria, gerencia, corretor, empreendimento, empresa
    
    valorUnidade = CDbl(rs("ValorUnidade"))
    
    ' Garantindo que ComissaoVenda não seja nulo ou zero para evitar erros
    comissaoVenda = 0
    If Not IsNull(rs("ComissaoVenda")) Then
        comissaoVenda = CDbl(rs("ComissaoVenda"))
    End If
    
    ' Cálculo das comissões
    comissaoImobiliaria = (valorUnidade * comissaoVenda) / 100
    Dim comissaoDiretor, comissaoGerente, comissaoCorretor
    comissaoDiretor = comissaoImobiliaria * 0.05
    comissaoGerente = comissaoImobiliaria * 0.10
    comissaoCorretor = comissaoImobiliaria * 0.35
    
    mes = CStr(rs("MesVenda"))
    
    diretoria = CStr(rs("Diretoria"))
    gerencia = CStr(rs("Gerencia"))
    corretor = CStr(rs("Corretor"))
    empreendimento = CStr(rs("NomeEmpreendimento"))
    empresa = CStr(rs("NomeEmpresa"))

    ' Popula os dicionários de KPIs de comissão
    Call ProcessaComissoes(kpiData("Mes"), mes, comissaoImobiliaria)
    Call ProcessaComissoes(kpiData("TopCorretores"), corretor, comissaoCorretor)
    Call ProcessaComissoes(kpiData("TopDiretorias"), diretoria, comissaoDiretor)
    Call ProcessaComissoes(kpiData("TopGerencias"), gerencia, comissaoGerente)
    Call ProcessaComissoes(kpiData("TopEmpreendimentos"), empreendimento, comissaoImobiliaria)
    Call ProcessaComissoes(kpiData("TopEmpresas"), empresa, comissaoImobiliaria)

    ' Soma os totais para os cards de KPI
    totalVGV = totalVGV + valorUnidade
    totalComissoesDiretoria = totalComissoesDiretoria + comissaoDiretor
    totalComissoesGerencia = totalComissoesGerencia + comissaoGerente
    totalComissoesCorretor = totalComissoesCorretor + comissaoCorretor
    totalComissoesImobiliaria = totalComissoesImobiliaria + comissaoImobiliaria

    rs.MoveNext
  Loop
End If

rs.Close
Set rs = Nothing

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

' Preparar dados para o gráfico de comissões por mês
Dim chartLabels(11)
Dim chartData(11)
For i = 1 To 12
    chartLabels(i - 1) = arrMesesNome(i)
    chartData(i - 1) = 0 ' Valor padrão se o mês não tiver dados
    If kpiData("Mes").Exists(CStr(i)) Then
        chartData(i - 1) = kpiData("Mes")(CStr(i))
    End If
Next

' Totais para exibição
Dim totalComissoes
totalComissoes = 0
For Each mesKey In kpiData("Mes").Keys
    totalComissoes = totalComissoes + kpiData("Mes")(mesKey)
Next

%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SGVendas - Relatório de Comissões</title>
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
    <h2 class="mt-4 mb-4 text-center" style="color: #800000;"><i class="fas fa-hand-holding-usd"></i> SGVendas - Relatório de Comissões</h2>
    
    <div class="filter-container">
      <form id="filterForm" method="get">
        <div class="row filter-row">
          <div class="col-md-2">
            <div class="filter-label">Ano</div>
            <select class="form-select filter-select" name="ano" id="anoFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each ano In uniqueAnos
                Response.Write "<option value=""" & ano & """"
                If filtroAno = ano Then Response.Write " selected"
                Response.Write ">" & ano & "</option>"
              Next
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
              For Each mes In uniqueMeses
                Dim mesNum : mesNum = CInt(mes)
                Response.Write "<option value=""" & mes & """"
                If CStr(filtroMes) = CStr(mes) Then Response.Write " selected"
                Response.Write ">" & arrMesesNome(mesNum) & "</option>"
              Next
              %>
            </select>
          </div>
          <div class="col-md-2">
            <div class="filter-label">Trimestre</div>
            <select class="form-select filter-select" name="trimestre" id="trimestreFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each trimestre In uniqueTrimestres
                Response.Write "<option value=""" & trimestre & """"
                If filtroTrimestre = trimestre Then Response.Write " selected"
                Response.Write ">" & trimestre & "</option>"
              Next
              %>
            </select>
          </div>
          <div class="col-md-2">
            <div class="filter-label">Diretoria</div>
            <select class="form-select filter-select" name="diretoria" id="diretoriaFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each diretoria In uniqueDiretorias
                Response.Write "<option value=""" & diretoria & """"
                If filtroDiretoria = diretoria Then Response.Write " selected"
                Response.Write ">" & diretoria & "</option>"
              Next
              %>
            </select>
          </div>
          <div class="col-md-2">
            <div class="filter-label">Gerência</div>
            <select class="form-select filter-select" name="gerencia" id="gerenciaFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each gerencia In uniqueGerencias
                Response.Write "<option value=""" & gerencia & """"
                If filtroGerencia = gerencia Then Response.Write " selected"
                Response.Write ">" & gerencia & "</option>"
              Next
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
              For Each corretor In uniqueCorretores
                Response.Write "<option value=""" & corretor & """"
                If filtroCorretor = corretor Then Response.Write " selected"
                Response.Write ">" & corretor & "</option>"
              Next
              %>
            </select>
          </div>
          <div class="col-md-2">
            <div class="filter-label">Empreendimento</div>
            <select class="form-select filter-select" name="empreendimento" id="empreendimentoFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each empreendimento In uniqueEmpreendimentos
                Response.Write "<option value=""" & empreendimento & """"
                If filtroEmpreendimento = empreendimento Then Response.Write " selected"
                Response.Write ">" & empreendimento & "</option>"
              Next
              %>
            </select>
          </div>
          <div class="col-md-2">
            <div class="filter-label">Empresa</div>
            <select class="form-select filter-select" name="empresa" id="empresaFilter" onchange="this.form.submit()">
              <option value="">Todos</option>
              <%
              For Each empresa In uniqueEmpresas
                Response.Write "<option value=""" & empresa & """"
                If filtroEmpresa = empresa Then Response.Write " selected"
                Response.Write ">" & empresa & "</option>"
              Next
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
    <h2 class="text-white mt-5">KPIs das Comissões</h2>
    <div class="row mt-4">
      <!--  <div class="col-md-4">
            <div class="kpi-card bg-primary-kpi">
                <i class="fas fa-dollar-sign"></i>
                <h5>Total de Comissões da Imobiliária</h5>
                <p>R$ <%= FormatNumber(totalComissoesImobiliaria, 2) %></p>
            </div>
        </div> -->
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
    </div>
    <h2 class="text-white mt-5">Ranking das Comissões Recebidas</h2>
    <div class="card-kpi p-3 rounded">
      <div class="row">
        <div class="col-md-6 mb-4">
          <h4 class="text-dark">Top 10 Diretorias (Comissão)</h4>
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
                    If i >= 10 Then Exit For
                    'Dim comissaoDiretor
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
          <h4 class="text-dark">Top 10 Gerências (Comissão)</h4>
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
                    If i >= 10 Then Exit For
                    Dim comissaoGerencia
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
          <h4 class="text-dark">Top 10 Corretores (Comissão)</h4>
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
                    If i >= 10 Then Exit For
                    '' Dim comissaoCorretor
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
          <h4 class="text-dark">Top 10 Empreendimentos (Comissão da Imobiliária)</h4>
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
                    If i >= 10 Then Exit For
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
          <h4 class="text-dark">Top 10 Empresas (Comissão da Imobiliária)</h4>
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
                    If i >= 10 Then Exit For
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
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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