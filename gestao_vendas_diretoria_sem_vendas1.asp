<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: CORRETORES_INATIVOS    -->
<!-- OBS: Relatório de Corretores  -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  
 <!--#include file="usr_acoes_v4GVendas.inc"-->

<%
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    On Error Resume Next 
    set objMail = server.createobject("CDONTS.NewMail")
    if Err.Number <> 0 then 
        set objMail = Nothing ' Garante que a variável seja liberada, mesmo que não criada
    else
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
        objMail.Subject = "SV-DIR-CORRETORES" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
        objMail.MailFormat = 0 ' 0 = Texto Simples
        objMail.Body = "Página Corretores. " & Ucase(Session("Usuario"))
        objMail.Send
        set objMail = Nothing
    end if 
    On Error GoTo 0 
end if%>

<%
' Configuração para evitar cache
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"

' **NOVA LÓGICA DE FILTRO DE DIRETORIA**
diretoriaID = Session("Dir_DiretoriaID")

paginaRedirecionamento1 = "http://www.gabnetweb.com.br/SunnyImob/login_v66a.asp"
paginaRedirecionamento2 = "http://localhost/SunnyImob/login_v66a.asp"
if diretoriaID = "" then
   Response.Write "Erro de processamento! (100)" 
   '*** verificar se é localhost ou no site'

   if (request.ServerVariables("remote_addr") <> "127.0.0.1")  then
      Response.Redirect paginaRedirecionamento1
   else
      Response.Redirect paginaRedirecionamento2
   end if   
  '' Response.end 
end if   

' Conexão com o banco de dados
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

' Calcular data atual para comparação
Dim anoAtual, mesAtual
anoAtual = Year(Date())
mesAtual = Month(Date())

' Declarar variáveis que serão usadas mais tarde
Dim sql, rs, sqlSub, rsSub, sqlGerencia, rsGerencia, sqlGerencias, rsGerencias
Dim corretores(), ultimoAnoArr(), ultimoMesArr(), totalVendasArr(), totalVGVArr()
Dim mesesSemVenderArr(), statusArr(), statusClassArr(), gerenciaAtualArr()
Dim count, i, j, temp, mesesSemVender, statusTexto, statusClass, rowClass
Dim totalCorretores, corretoresAtivos, corretoresInativos
Dim mediaMeses, vgvFormatado, nomeMesUltimaVenda, mesesAtras, widthPercent, barColor
Dim ultimoAno, ultimoMes, corretorNome
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Corretores - Diretoria <%=Session("Dir_Nome")%></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        body { 
            background: #f8f9fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .card { 
            margin-bottom: 20px; 
            border: none;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }
        .card-header {
            border-radius: 10px 10px 0 0 !important;
            font-weight: 600;
            padding: 15px 20px;
        }
        .table th { 
            background-color: #f8f9fa; 
            font-weight: 600;
        }
        .badge-status {
            font-size: 0.8rem;
            padding: 4px 8px;
            border-radius: 12px;
        }
        .badge-vermelho { background-color: #dc3545; color: white; }
        .badge-laranja { background-color: #fd7e14; color: white; }
        .badge-amarelo { background-color: #ffc107; color: black; }
        .badge-verde { background-color: #28a745; color: white; }
        .badge-secondary { background-color: #6c757d; color: white; }
        .badge-info { background-color: #17a2b8; color: white; }
        .btn-filter {
            padding: 10px 20px;
            font-weight: 600;
        }
        
        .status-cell {
            font-weight: bold;
            text-align: center;
        }
        
        .corretor-muito-inativo {
            background-color: #fff5f5 !important;
        }
        .corretor-inativo {
            background-color: #fff9e6 !important;
        }
        
        .meses-indicator {
            height: 6px;
            border-radius: 3px;
            margin-top: 5px;
        }
        
        .header-section {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            color: white;
        }
        
        .metric-card {
            padding: 15px;
            border-radius: 8px;
            color: white;
            text-align: center;
            margin-bottom: 15px;
        }
        
        .metric-value {
            font-size: 1.8rem;
            font-weight: bold;
            margin: 5px 0;
        }
        
        .metric-label {
            font-size: 0.9rem;
            opacity: 0.9;
        }
        
        .alert-custom {
            border-left: 5px solid;
            border-radius: 5px;
        }
        
        .alert-warning-custom {
            border-left-color: #ffc107;
            background-color: #fff9e6;
        }
        
        .alert-danger-custom {
            border-left-color: #dc3545;
            background-color: #fff5f5;
        }
        
        .custom-tooltip {
            position: relative;
            cursor: help;
        }
        
        .custom-tooltip .tooltip-text {
            visibility: hidden;
            width: 200px;
            background-color: #333;
            color: #fff;
            text-align: center;
            border-radius: 6px;
            padding: 5px;
            position: absolute;
            z-index: 1000;
            bottom: 125%;
            left: 50%;
            margin-left: -100px;
            opacity: 0;
            transition: opacity 0.3s;
            font-size: 0.85rem;
        }
        
        .custom-tooltip:hover .tooltip-text {
            visibility: visible;
            opacity: 1;
        }
        
        .table-container {
            max-height: none;
            overflow: visible;
        }
        
        .gerencia-cell {
            text-align: center;
        }
        
        .filter-section {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        
        .table-responsive {
            overflow-x: auto;
        }
        
        .filter-gerencia {
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #dee2e6;
        }
        
        .gerencia-badge {
            cursor: pointer;
            transition: all 0.2s;
        }
        
        .gerencia-badge:hover {
            transform: translateY(-2px);
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }
    </style>
</head>
<body>
    <div class="container mt-4">
        <!-- HEADER -->
        <div class="header-section">
            <div class="row align-items-center">
                <div class="col-md-8">
                    <h1 class="mb-2"><i class="fas fa-users me-2"></i>Relatório de Corretores</h1>
                    <p class="mb-0">Diretoria: <strong><%=Session("Dir_Nome")%></strong></p>
                    <p class="mb-0 mt-2">
                        <small>
                            <i class="fas fa-info-circle me-1"></i>Lista completa de corretores com gerência atual e histórico de vendas
                        </small>
                    </p>
                </div>
                <div class="col-md-4 text-end">
                    <a href="gestao_vendas_diretoria2.asp" class="btn btn-light">
                        <i class="fas fa-chart-bar me-2"></i>Voltar ao Dashboard
                    </a>
                </div>
            </div>
        </div>
        
        <!-- RESULTADOS -->
        <div class="resultados">
        <%
        ' Inicializar arrays
        ReDim corretores(0)
        ReDim ultimoAnoArr(0)
        ReDim ultimoMesArr(0)
        ReDim totalVendasArr(0)
        ReDim totalVGVArr(0)
        ReDim mesesSemVenderArr(0)
        ReDim statusArr(0)
        ReDim statusClassArr(0)
        ReDim gerenciaAtualArr(0)
        
        count = 0
        
        ' Obter lista de todas as gerências únicas para filtros
        Dim gerenciasUnicas()
        ReDim gerenciasUnicas(0)
        Dim gerenciaCount
        gerenciaCount = 0
        
        sqlGerencias = "SELECT DISTINCT Gerencia FROM Vendas "
        sqlGerencias = sqlGerencias & " WHERE Excluido = 0"
        
        If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
            sqlGerencias = sqlGerencias & " AND DiretoriaId = " & CLng(diretoriaID)
        End If
        
        sqlGerencias = sqlGerencias & " AND Gerencia IS NOT NULL AND TRIM(Gerencia) <> ''"
        sqlGerencias = sqlGerencias & " ORDER BY Gerencia"
        
        Set rsGerencias = Server.CreateObject("ADODB.Recordset")
        rsGerencias.Open sqlGerencias, conn
        
        Do While Not rsGerencias.EOF
            If Not IsNull(rsGerencias("Gerencia")) Then
                ReDim Preserve gerenciasUnicas(gerenciaCount)
                gerenciasUnicas(gerenciaCount) = Trim(rsGerencias("Gerencia"))
                gerenciaCount = gerenciaCount + 1
            End If
            rsGerencias.MoveNext
        Loop
        rsGerencias.Close
        Set rsGerencias = Nothing
        
        ' PRIMEIRA CONSULTA: Obter lista de todos os corretores únicos
        sql = "SELECT DISTINCT Corretor FROM Vendas "
        sql = sql & " WHERE Excluido = 0"
        
        If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
            sql = sql & " AND DiretoriaId = " & CLng(diretoriaID)
        End If
        
        sql = sql & " AND Corretor IS NOT NULL AND TRIM(Corretor) <> ''"
        sql = sql & " ORDER BY Corretor"
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn
        
        Do While Not rs.EOF
            corretorNome = Trim(rs("Corretor"))
            If corretorNome <> "" Then
                
                ' CONSULTA 2: Obter dados de vendas do corretor
                sqlSub = "SELECT MAX(AnoVenda) as UltimoAno, MAX(MesVenda) as UltimoMes, "
                sqlSub = sqlSub & " COUNT(*) as TotalVendas, SUM(ValorUnidade) as TotalVGV "
                sqlSub = sqlSub & " FROM Vendas "
                sqlSub = sqlSub & " WHERE Excluido = 0 AND Corretor = '" & Replace(corretorNome, "'", "''") & "'"
                
                If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
                    sqlSub = sqlSub & " AND DiretoriaId = " & CLng(diretoriaID)
                End If
                
                Set rsSub = Server.CreateObject("ADODB.Recordset")
                rsSub.Open sqlSub, conn
                
                If Not rsSub.EOF Then
                    ' Calcular meses sem vender
                    If Not IsNull(rsSub("UltimoAno")) And Not IsNull(rsSub("UltimoMes")) Then
                        ultimoAno = CInt(rsSub("UltimoAno"))
                        ultimoMes = CInt(rsSub("UltimoMes"))
                        mesesSemVender = ((anoAtual - ultimoAno) * 12) + (mesAtual - ultimoMes)
                    Else
                        ultimoAno = 0
                        ultimoMes = 0
                        mesesSemVender = 999 ' Valor alto para corretores sem vendas
                    End If
                    
                    ' CONSULTA 3: Obter a última gerência do corretor
                    sqlGerencia = "SELECT TOP 1 Gerencia FROM Vendas "
                    sqlGerencia = sqlGerencia & " WHERE Excluido = 0 AND Corretor = '" & Replace(corretorNome, "'", "''") & "'"
                    
                    If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
                        sqlGerencia = sqlGerencia & " AND DiretoriaId = " & CLng(diretoriaID)
                    End If
                    
                    sqlGerencia = sqlGerencia & " ORDER BY AnoVenda DESC, MesVenda DESC"
                    
                    Set rsGerencia = Server.CreateObject("ADODB.Recordset")
                    rsGerencia.Open sqlGerencia, conn
                    
                    Dim gerenciaAtual
                    gerenciaAtual = ""
                    If Not rsGerencia.EOF Then
                        If Not IsNull(rsGerencia("Gerencia")) Then
                            gerenciaAtual = Trim(rsGerencia("Gerencia"))
                        End If
                    End If
                    rsGerencia.Close
                    Set rsGerencia = Nothing
                    
                    ' Determinar status
                    If mesesSemVender >= 12 Then
                        statusTexto = "Muito Inativo"
                        statusClass = "badge-vermelho"
                        rowClass = "corretor-muito-inativo"
                    ElseIf mesesSemVender >= 6 Then
                        statusTexto = "Inativo"
                        statusClass = "badge-laranja"
                        rowClass = "corretor-inativo"
                    ElseIf mesesSemVender >= 3 Then
                        statusTexto = "Atenção"
                        statusClass = "badge-amarelo"
                        rowClass = ""
                    ElseIf mesesSemVender <= 2 And mesesSemVender >= 0 Then
                        statusTexto = "Ativo"
                        statusClass = "badge-verde"
                        rowClass = ""
                    Else
                        statusTexto = "Sem Vendas"
                        statusClass = "badge-secondary"
                        rowClass = ""
                    End If
                    
                    ' Armazenar dados nos arrays
                    ReDim Preserve corretores(count)
                    ReDim Preserve ultimoAnoArr(count)
                    ReDim Preserve ultimoMesArr(count)
                    ReDim Preserve totalVendasArr(count)
                    ReDim Preserve totalVGVArr(count)
                    ReDim Preserve mesesSemVenderArr(count)
                    ReDim Preserve statusArr(count)
                    ReDim Preserve statusClassArr(count)
                    ReDim Preserve gerenciaAtualArr(count)
                    
                    corretores(count) = UCase(corretorNome)
                    ultimoAnoArr(count) = ultimoAno
                    ultimoMesArr(count) = ultimoMes
                    totalVendasArr(count) = rsSub("TotalVendas")
                    
                    If Not IsNull(rsSub("TotalVGV")) Then
                        totalVGVArr(count) = rsSub("TotalVGV")
                    Else
                        totalVGVArr(count) = 0
                    End If
                    
                    mesesSemVenderArr(count) = mesesSemVender
                    statusArr(count) = statusTexto
                    statusClassArr(count) = statusClass
                    gerenciaAtualArr(count) = gerenciaAtual
                    
                    count = count + 1
                End If
                
                rsSub.Close
                Set rsSub = Nothing
            End If
            
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        
        ' Ordenar pelo número de meses sem vender (decrescente)
        If count > 1 Then
            For i = 0 To count - 2
                For j = i + 1 To count - 1
                    If mesesSemVenderArr(i) < mesesSemVenderArr(j) Then
                        ' Trocar todos os dados
                        temp = corretores(i)
                        corretores(i) = corretores(j)
                        corretores(j) = temp
                        
                        temp = ultimoAnoArr(i)
                        ultimoAnoArr(i) = ultimoAnoArr(j)
                        ultimoAnoArr(j) = temp
                        
                        temp = ultimoMesArr(i)
                        ultimoMesArr(i) = ultimoMesArr(j)
                        ultimoMesArr(j) = temp
                        
                        temp = totalVendasArr(i)
                        totalVendasArr(i) = totalVendasArr(j)
                        totalVendasArr(j) = temp
                        
                        temp = totalVGVArr(i)
                        totalVGVArr(i) = totalVGVArr(j)
                        totalVGVArr(j) = temp
                        
                        temp = mesesSemVenderArr(i)
                        mesesSemVenderArr(i) = mesesSemVenderArr(j)
                        mesesSemVenderArr(j) = temp
                        
                        temp = statusArr(i)
                        statusArr(i) = statusArr(j)
                        statusArr(j) = temp
                        
                        temp = statusClassArr(i)
                        statusClassArr(i) = statusClassArr(j)
                        statusClassArr(j) = temp
                        
                        temp = gerenciaAtualArr(i)
                        gerenciaAtualArr(i) = gerenciaAtualArr(j)
                        gerenciaAtualArr(j) = temp
                    End If
                Next
            Next
        End If
        
        ' Calcular métricas
        totalCorretores = count
        
        corretoresAtivos = 0
        corretoresInativos = 0
        Dim corretoresAtencao, corretoresMuitoInativos, corretoresSemVendas
        corretoresAtencao = 0
        corretoresMuitoInativos = 0
        corretoresSemVendas = 0
        
        If count > 0 Then
            For i = 0 To count - 1
                Select Case statusArr(i)
                    Case "Ativo"
                        corretoresAtivos = corretoresAtivos + 1
                    Case "Atenção"
                        corretoresAtencao = corretoresAtencao + 1
                        corretoresInativos = corretoresInativos + 1 ' Atenção também é considerado inativo
                    Case "Inativo"
                        corretoresInativos = corretoresInativos + 1
                    Case "Muito Inativo"
                        corretoresMuitoInativos = corretoresMuitoInativos + 1
                        corretoresInativos = corretoresInativos + 1
                    Case "Sem Vendas"
                        corretoresSemVendas = corretoresSemVendas + 1
                End Select
            Next
        End If
        %>
        
        <!-- MÉTRICAS -->
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="metric-card" style="background: #3498db;">
                    <div class="metric-label"><i class="fas fa-users me-2"></i> Total de Corretores</div>
                    <div class="metric-value"><%=totalCorretores%></div>
                    <small>Corretores na diretoria</small>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card" style="background: #2ecc71;">
                    <div class="metric-label"><i class="fas fa-check-circle me-2"></i> Ativos</div>
                    <div class="metric-value"><%=corretoresAtivos%></div>
                    <small>Realizando Vendas</small>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card" style="background: #e74c3c;">
                    <div class="metric-label"><i class="fas fa-user-clock me-2"></i> Inativos/Atenção</div>
                    <div class="metric-value"><%=(totalCorretores-corretoresAtivos)%></div>
                    <small>3+ meses sem vender</small>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card" style="background: #6c757d;">
                    <div class="metric-label"><i class="fas fa-question-circle me-2"></i> Sem Vendas</div>
                    <div class="metric-value"><%=corretoresSemVendas%></div>
                    <small>Nenhuma venda registrada</small>
                </div>
            </div>
        </div>
        
        <!-- FILTROS RÁPIDOS -->
        <div class="filter-section">
            <div class="row">
                <div class="col-md-12">
                    <h6 class="mb-3"><i class="fas fa-filter me-2"></i>Filtros Rápidos - Status</h6>
                    <div class="d-flex flex-wrap gap-2 mb-3">
                        <button class="btn btn-sm btn-outline-primary" onclick="filtrarTabela('todos')">
                            <i class="fas fa-eye me-1"></i>Todos (<%=totalCorretores%>)
                        </button>
                        <button class="btn btn-sm btn-outline-success" onclick="filtrarTabela('Ativo')">
                            <i class="fas fa-check me-1"></i>Ativos (<%=corretoresAtivos%>)
                        </button>

                        <button class="btn btn-sm btn-outline-danger" onclick="filtrarTabela('Inativo')">
                            <i class="fas fa-user-clock me-1"></i>Inativos (<%=corretoresInativos%>)
                        </button>
                        <button class="btn btn-sm btn-outline-secondary" onclick="filtrarTabela('Sem Vendas')">
                            <i class="fas fa-question me-1"></i>Sem Vendas (<%=corretoresSemVendas%>)
                        </button>
                    </div>
                    
                    <!-- FILTROS POR GERÊNCIA -->
                    <div class="filter-gerencia">
                        <h6 class="mb-3"><i class="fas fa-sitemap me-2"></i>Filtros por Gerência</h6>
                        <div class="d-flex flex-wrap gap-2">
                            <button class="btn btn-sm btn-outline-info" onclick="filtrarPorGerencia('todas')">
                                <i class="fas fa-building me-1"></i>Todas Gerências
                            </button>
                            <%
                            ' Contar corretores por gerência para mostrar nos botões
                            Dim gerenciaCountDict
                            Set gerenciaCountDict = Server.CreateObject("Scripting.Dictionary")
                            
                            For i = 0 To count - 1
                                Dim gerenciaForFilter
                                gerenciaForFilter = gerenciaAtualArr(i)
                                If gerenciaForFilter = "" Then gerenciaForFilter = "Não Informada"
                                
                                If Not gerenciaCountDict.Exists(gerenciaForFilter) Then
                                    gerenciaCountDict.Add gerenciaForFilter, 0
                                End If
                                gerenciaCountDict(gerenciaForFilter) = gerenciaCountDict(gerenciaForFilter) + 1
                            Next
                            
                            ' Ordenar gerências por quantidade de corretores (decrescente)
                            Dim gerenciaKeysSorted(), gerenciaCountsSorted()
                            ReDim gerenciaKeysSorted(gerenciaCountDict.Count - 1)
                            ReDim gerenciaCountsSorted(gerenciaCountDict.Count - 1)
                            
                            Dim dictKeys, dictKey, idx
                            dictKeys = gerenciaCountDict.Keys
                            idx = 0
                            For Each dictKey In dictKeys
                                gerenciaKeysSorted(idx) = dictKey
                                gerenciaCountsSorted(idx) = gerenciaCountDict(dictKey)
                                idx = idx + 1
                            Next
                            
                            ' Ordenar por quantidade (bubble sort simples)
                            For i = 0 To UBound(gerenciaKeysSorted) - 1
                                For j = i + 1 To UBound(gerenciaKeysSorted)
                                    If gerenciaCountsSorted(i) < gerenciaCountsSorted(j) Then
                                        temp = gerenciaKeysSorted(i)
                                        gerenciaKeysSorted(i) = gerenciaKeysSorted(j)
                                        gerenciaKeysSorted(j) = temp
                                        
                                        temp = gerenciaCountsSorted(i)
                                        gerenciaCountsSorted(i) = gerenciaCountsSorted(j)
                                        gerenciaCountsSorted(j) = temp
                                    End If
                                Next
                            Next
                            
                            ' Gerar botões para cada gerência
                            For i = 0 To UBound(gerenciaKeysSorted)
                                If gerenciaKeysSorted(i) <> "" Then
                                    Dim corAleatoria, corClasse
                                    ' Gerar cor de fundo aleatória baseada no nome da gerência
                                    Randomize Len(gerenciaKeysSorted(i))
                                    corAleatoria = Int((5 * Rnd) + 1)
                                    
                                    Select Case corAleatoria
                                        Case 1
                                            corClasse = "btn-outline-primary"
                                        Case 2
                                            corClasse = "btn-outline-success"
                                        Case 3
                                            corClasse = "btn-outline-warning"
                                        Case 4
                                            corClasse = "btn-outline-danger"
                                        Case Else
                                            corClasse = "btn-outline-info"
                                    End Select
                            %>
                            <button class="btn btn-sm <%=corClasse%> gerencia-badge" onclick="filtrarPorGerencia('<%=Server.HTMLEncode(gerenciaKeysSorted(i))%>')">
                                <i class="fas fa-building me-1"></i><%=gerenciaKeysSorted(i)%> (<%=gerenciaCountsSorted(i)%>)
                            </button>
                            <%
                                End If
                            Next
                            
                            Set gerenciaCountDict = Nothing
                            %>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- TABELA DE CORRETORES -->
        <% If totalCorretores > 0 Then %>
        <div class="card">
            <div class="card-header text-white d-flex justify-content-between align-items-center" style="background: #f39c12;">
                <div>
                    <i class="fas fa-table me-2"></i>Lista de Corretores
                    <span class="badge bg-light text-dark ms-2" id="contador">Mostrando <%=totalCorretores%> corretores</span>
                    <span class="badge bg-info text-white ms-2" id="filtroAtivo" style="display: none;"></span>
                </div>
                <div class="d-flex align-items-center">
                    <input type="text" id="searchInput" class="form-control form-control-sm me-2" placeholder="Buscar corretor..." style="width: 200px;">
                    <div>
                        <small>Ordenado por: <strong>Meses sem vender (decrescente)</strong></small>
                    </div>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive">
                    <table class="table table-hover mb-0" id="tabelaCorretores">
                        <thead>
                            <tr>
                                <th width="5%">#</th>
                                <th width="20%">Corretor</th>
                                <th width="15%" class="text-center">Última Venda</th>
                                <th width="15%" class="text-center">Meses sem Vender</th>
                                <th width="15%" class="text-center">Gerência Atual</th>
                                <th width="15%" class="text-end">Total VGV</th>
                                <th width="15%" class="text-end">Total Vendas</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            For i = 0 To count - 1
                                ' Determinar classe da linha
                                rowClass = ""
                                If mesesSemVenderArr(i) >= 12 Then
                                    rowClass = "corretor-muito-inativo"
                                ElseIf mesesSemVenderArr(i) >= 6 Then
                                    rowClass = "corretor-inativo"
                                End If
                                
                                ' Formatar valor VGV
                                If totalVGVArr(i) >= 1000000 Then
                                    vgvFormatado = "R$ " & FormatNumber(totalVGVArr(i)/1000000, 2) & " M"
                                ElseIf totalVGVArr(i) >= 1000 Then
                                    vgvFormatado = "R$ " & FormatNumber(totalVGVArr(i)/1000, 0) & " mil"
                                Else
                                    vgvFormatado = "R$ " & FormatNumber(totalVGVArr(i), 0)
                                End If
                                
                                ' Nome do mês da última venda
                                If ultimoMesArr(i) >= 1 And ultimoMesArr(i) <= 12 Then
                                    nomeMesUltimaVenda = MonthName(ultimoMesArr(i), False)
                                Else
                                    nomeMesUltimaVenda = "N/A"
                                End If
                            %>
                            <tr class="<%=rowClass%>" data-status="<%=statusArr(i)%>" data-gerencia="<%=Server.HTMLEncode(gerenciaAtualArr(i))%>">
                                <td class="fw-bold"><%=i+1%></td>
                                <td>
                                    <strong><%=corretores(i)%></strong>
                                    <br>
                                    <small class="text-muted">
                                        <i class="fas fa-building me-1"></i>
                                        <% If gerenciaAtualArr(i) <> "" Then %>
                                            <%=gerenciaAtualArr(i)%>
                                        <% Else %>
                                            Gerência não informada
                                        <% End If %>
                                    </small>
                                </td>
                                <td class="text-center">
                                    <% If ultimoAnoArr(i) > 0 Then %>
                                    <div class="fw-bold">
                                        <%=nomeMesUltimaVenda%>/<%=ultimoAnoArr(i)%>
                                    </div>
                                    <small class="text-muted">
                                        <% 
                                        mesesAtras = mesesSemVenderArr(i)
                                        If mesesAtras = 999 Then
                                            Response.Write "Sem vendas"
                                        ElseIf mesesAtras = 1 Then
                                            Response.Write "Há 1 mês"
                                        ElseIf mesesAtras = 0 Then
                                            Response.Write "Este mês"
                                        Else
                                            Response.Write "Há " & mesesAtras & " meses"
                                        End If
                                        %>
                                    </small>
                                    <% Else %>
                                    <div class="fw-bold text-muted">N/A</div>
                                    <small class="text-muted">Sem vendas registradas</small>
                                    <% End If %>
                                </td>
                                <td class="text-center">
                                    <% If mesesSemVenderArr(i) <> 999 Then %>
                                    <div class="fw-bold <%=statusClassArr(i)%> p-2 rounded">
                                        <% If mesesSemVenderArr(i) = 0 Then %>
                                            Este mês
                                        <% Else %>
                                            <%=mesesSemVenderArr(i)%> meses
                                        <% End If %>
                                    </div>
                                    <!-- Barra de progresso visual -->
                                    <div class="meses-indicator" style="background: #e9ecef;">
                                        <%
                                        If mesesSemVenderArr(i) > 24 Then
                                            widthPercent = 100
                                        ElseIf mesesSemVenderArr(i) = 999 Then
                                            widthPercent = 0
                                        Else
                                            widthPercent = (mesesSemVenderArr(i) / 24) * 100
                                        End If
                                        
                                        If mesesSemVenderArr(i) >= 12 Then
                                            barColor = "#dc3545"
                                        ElseIf mesesSemVenderArr(i) >= 6 Then
                                            barColor = "#fd7e14"
                                        ElseIf mesesSemVenderArr(i) >= 3 Then
                                            barColor = "#ffc107"
                                        ElseIf mesesSemVenderArr(i) >= 0 Then
                                            barColor = "#28a745"
                                        Else
                                            barColor = "#6c757d"
                                        End If
                                        %>
                                        <div class="meses-indicator" style="width: <%=widthPercent%>%; background-color: <%=barColor%>;"></div>
                                    </div>
                                    <% Else %>
                                    <div class="fw-bold badge-secondary p-2 rounded">
                                        Sem vendas
                                    </div>
                                    <% End If %>
                                </td>
                                <td class="text-center gerencia-cell">
                                    <div class="fw-bold">
                                        <% If gerenciaAtualArr(i) <> "" Then %>
                                            <span class="badge bg-info text-white">
                                                <%=gerenciaAtualArr(i)%>
                                            </span>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </div>
                                    <small class="text-muted">
                                        <span class="badge <%=statusClassArr(i)%> badge-status"><%=statusArr(i)%></span>
                                    </small>
                                </td>
                                <td class="text-end">
                                    <div class="fw-bold"><%=vgvFormatado%></div>
                                    <small class="text-muted">
                                        <% 
                                        If totalVendasArr(i) > 0 Then
                                            If totalVGVArr(i)/totalVendasArr(i) >= 1000 Then
                                                Response.Write "R$ " & FormatNumber(totalVGVArr(i)/totalVendasArr(i)/1000, 1) & " mil/venda"
                                            Else
                                                Response.Write "R$ " & FormatNumber(totalVGVArr(i)/totalVendasArr(i), 0) & "/venda"
                                            End If
                                        Else
                                            Response.Write "Sem vendas"
                                        End If
                                        %>
                                    </small>
                                </td>
                                <td class="text-end">
                                    <div class="fw-bold"><%=totalVendasArr(i)%></div>
                                    <small class="text-muted">
                                        <% If totalVendasArr(i) > 0 Then %>
                                            vendas totais
                                        <% Else %>
                                            nenhuma venda
                                        <% End If %>
                                    </small>
                                </td>
                            </tr>
                            <% Next %>
                        </tbody>
                    </table>
                </div>
                
                <!-- LEGENDA E CONTROLES -->
                <div class="p-3 border-top">
                    <div class="row">
                        <div class="col-md-8">
                            <h6 class="mb-2"><i class="fas fa-info-circle me-2"></i>Legenda de Status:</h6>
                            <div class="d-flex flex-wrap gap-2">
                                <div class="d-flex align-items-center me-3 mb-2">
                                    <span class="badge badge-verde badge-status me-2">Ativo</span>
                                    <small class="text-muted">0-2 meses</small>
                                </div>
                                <div class="d-flex align-items-center me-3 mb-2">
                                    <span class="badge badge-amarelo badge-status me-2">Atenção</span>
                                    <small class="text-muted">3-5 meses</small>
                                </div>
                                <div class="d-flex align-items-center me-3 mb-2">
                                    <span class="badge badge-laranja badge-status me-2">Inativo</span>
                                    <small class="text-muted">6-11 meses</small>
                                </div>
                                <div class="d-flex align-items-center me-3 mb-2">
                                    <span class="badge badge-vermelho badge-status me-2">Muito Inativo</span>
                                    <small class="text-muted">12+ meses</small>
                                </div>
                                <div class="d-flex align-items-center me-3 mb-2">
                                    <span class="badge badge-secondary badge-status me-2">Sem Vendas</span>
                                    <small class="text-muted">Nenhuma venda</small>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-4 text-end">
                            <button class="btn btn-sm btn-outline-secondary" onclick="imprimirRelatorio()">
                                <i class="fas fa-print me-1"></i>Imprimir
                            </button>
                            <button class="btn btn-sm btn-outline-primary" onclick="exportarParaExcel()">
                                <i class="fas fa-file-excel me-1"></i>Exportar Excel
                            </button>
                            <button class="btn btn-sm btn-outline-danger" onclick="limparFiltros()">
                                <i class="fas fa-times me-1"></i>Limpar Filtros
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- RESUMO POR GERÊNCIA -->
        <div class="card mt-4">
            <div class="card-header text-white" style="background: #17a2b8;">
                <i class="fas fa-sitemap me-2"></i>Resumo por Gerência
            </div>
            <div class="card-body">
                <%
                ' Calcular resumo por gerência
                Dim gerenciaStats, gerenciaNome
                Set gerenciaStats = Server.CreateObject("Scripting.Dictionary")
                
                For i = 0 To count - 1
                    gerenciaNome = gerenciaAtualArr(i)
                    If gerenciaNome = "" Then gerenciaNome = "Não Informada"
                    
                    If Not gerenciaStats.Exists(gerenciaNome) Then
                        gerenciaStats.Add gerenciaNome, Array(0, 0, 0, 0, 0) ' Total, Ativos, Atenção, Inativos, MuitoInativos
                    End If
                    
                    Dim stats
                    stats = gerenciaStats(gerenciaNome)
                    
                    ' Incrementar total
                    stats(0) = stats(0) + 1
                    
                    ' Classificar por status
                    Select Case statusArr(i)
                        Case "Ativo"
                            stats(1) = stats(1) + 1
                        Case "Atenção"
                            stats(2) = stats(2) + 1
                        Case "Inativo"
                            stats(3) = stats(3) + 1
                        Case "Muito Inativo"
                            stats(4) = stats(4) + 1
                    End Select
                    
                    gerenciaStats(gerenciaNome) = stats
                Next
                
                Dim gerenciaKeysResumo
                gerenciaKeysResumo = gerenciaStats.Keys
                %>
                
                <div class="row">
                    <% 
                    For Each gerenciaNome In gerenciaKeysResumo
                        stats = gerenciaStats(gerenciaNome)
                        Dim percentAtivos
                        If stats(0) > 0 Then
                            percentAtivos = FormatNumber((stats(1) / stats(0)) * 100, 1)
                        Else
                            percentAtivos = 0
                        End If
                    %>
                    <div class="col-md-6 mb-3">
                        <div class="card border-0 h-100" style="background: #f8f9fa;">
                            <div class="card-body">
                                <h6 class="card-title">
                                    <i class="fas fa-building me-2"></i><%=gerenciaNome%>
                                    <span class="badge bg-primary float-end"><%=stats(0)%> corretores</span>
                                </h6>
                                <div class="progress mb-2" style="height: 8px;">
                                    <div class="progress-bar bg-success" style="width: <%=percentAtivos%>%;" title="<%=percentAtivos%>% ativos"></div>
                                </div>
                                <div class="d-flex justify-content-between small">
                                    <span><i class="fas fa-check-circle text-success me-1"></i> <%=stats(1)%> ativos</span>
                                    <span><i class="fas fa-exclamation-triangle text-warning me-1"></i> <%=stats(2)%> atenção</span>
                                    <span><i class="fas fa-user-clock text-danger me-1"></i> <%=stats(3) + stats(4)%> inativos</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    <% Next %>
                </div>
            </div>
        </div>
        <% Else %>
        <div class="alert alert-warning text-center">
            <i class="fas fa-exclamation-triangle fa-2x mb-3"></i>
            <h5>Nenhum corretor encontrado</h5>
            <p class="mb-0">Não há corretores registrados na diretoria.</p>
        </div>
        <% End If %>
        
        <!-- INFORMAÇÕES DO RELATÓRIO -->
        <div class="card mt-4">
            <div class="card-header text-white" style="background: #6c757d;">
                <i class="fas fa-question-circle me-2"></i>Informações do Relatório
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-md-6">
                        <h6><i class="fas fa-database me-2"></i>Fonte de Dados</h6>
                        <ul class="text-muted">
                            <li><strong>Gerência atual:</strong> Última gerência registrada nas vendas do corretor</li>
                            <li><strong>Última venda:</strong> Data da venda mais recente do corretor</li>
                            <li><strong>Meses sem vender:</strong> Calculado com base no mês/ano atual</li>
                            <li><strong>Total VGV:</strong> Soma de todas as vendas (ValorUnidade)</li>
                        </ul>
                    </div>
                    <div class="col-md-6">
                        <h6><i class="fas fa-calendar-alt me-2"></i>Período</h6>
                        <ul class="text-muted">
                            <li><strong>Data atual:</strong> <%=MonthName(mesAtual, True)%> de <%=anoAtual%></li>
                            <li><strong>Atualização:</strong> Dados em tempo real</li>
                            <li><strong>Diretoria:</strong> <%=Session("Dir_Nome")%></li>
                            <li><strong>Total de registros:</strong> <%=totalCorretores%> corretores</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
        <%
        conn.Close
        Set conn = Nothing
        %>
        </div>
    </div>
    
    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
    // Variáveis globais para controle dos filtros
    let filtroStatusAtivo = '';
    let filtroGerenciaAtiva = '';
    
    // Função para filtrar tabela por status
    function filtrarTabela(status) {
        filtroStatusAtivo = status;
        aplicarFiltros();
    }
    
    // Função para filtrar tabela por gerência
    function filtrarPorGerencia(gerencia) {
        filtroGerenciaAtiva = gerencia;
        aplicarFiltros();
    }
    
    // Função para aplicar todos os filtros combinados
    function aplicarFiltros() {
        const rows = document.querySelectorAll('#tabelaCorretores tbody tr');
        let visibleCount = 0;
        
        rows.forEach(function(row) {
            const rowStatus = row.getAttribute('data-status');
            const rowGerencia = row.getAttribute('data-gerencia') || '';
            
            let statusMatch = true;
            let gerenciaMatch = true;
            
            // Aplicar filtro de status
            if (filtroStatusAtivo && filtroStatusAtivo !== 'todos' && filtroStatusAtivo !== '') {
                if (filtroStatusAtivo === 'Inativo') {
                    // Para "Inativos", mostrar Atenção, Inativo e Muito Inativo
                    statusMatch = rowStatus === 'Atenção' || rowStatus === 'Inativo' || rowStatus === 'Muito Inativo';
                } else {
                    statusMatch = rowStatus === filtroStatusAtivo;
                }
            }
            
            // Aplicar filtro de gerência
            if (filtroGerenciaAtiva && filtroGerenciaAtiva !== 'todas' && filtroGerenciaAtiva !== '') {
                gerenciaMatch = rowGerencia === filtroGerenciaAtiva;
            }
            
            if (statusMatch && gerenciaMatch) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });
        
        // Atualizar contador
        document.getElementById('contador').textContent = 'Mostrando ' + visibleCount + ' corretores';
        
        // Atualizar badge do filtro ativo
        const filtroAtivoBadge = document.getElementById('filtroAtivo');
        let filtroTexto = '';
        
        if (filtroStatusAtivo && filtroStatusAtivo !== 'todos' && filtroStatusAtivo !== '') {
            filtroTexto += 'Status: ' + filtroStatusAtivo + ' ';
        }
        
        if (filtroGerenciaAtiva && filtroGerenciaAtiva !== 'todas' && filtroGerenciaAtiva !== '') {
            filtroTexto += 'Gerência: ' + filtroGerenciaAtiva;
        }
        
        if (filtroTexto) {
            filtroAtivoBadge.textContent = filtroTexto;
            filtroAtivoBadge.style.display = 'inline';
        } else {
            filtroAtivoBadge.style.display = 'none';
        }
        
        // Atualizar números nas linhas
        updateRowNumbers();
    }
    
    // Função para limpar todos os filtros
    function limparFiltros() {
        filtroStatusAtivo = '';
        filtroGerenciaAtiva = '';
        document.getElementById('searchInput').value = '';
        
        const rows = document.querySelectorAll('#tabelaCorretores tbody tr');
        rows.forEach(function(row) {
            row.style.display = '';
        });
        
        document.getElementById('contador').textContent = 'Mostrando <%=totalCorretores%> corretores';
        document.getElementById('filtroAtivo').style.display = 'none';
        updateRowNumbers();
    }
    
    // Função para atualizar números das linhas
    function updateRowNumbers() {
        const rows = document.querySelectorAll('#tabelaCorretores tbody tr:visible');
        let count = 1;
        
        rows.forEach(function(row) {
            const td = row.querySelector('td:first-child');
            if (td) {
                td.textContent = count;
            }
            count++;
        });
    }
    
    // Função de busca
    document.getElementById('searchInput').addEventListener('keyup', function() {
        const filter = this.value.toLowerCase();
        const rows = document.querySelectorAll('#tabelaCorretores tbody tr');
        let visibleCount = 0;
        
        rows.forEach(function(row) {
            // Aplicar filtros existentes primeiro
            const rowStatus = row.getAttribute('data-status');
            const rowGerencia = row.getAttribute('data-gerencia') || '';
            
            let statusMatch = true;
            let gerenciaMatch = true;
            
            if (filtroStatusAtivo && filtroStatusAtivo !== 'todos' && filtroStatusAtivo !== '') {
                if (filtroStatusAtivo === 'Inativo') {
                    statusMatch = rowStatus === 'Atenção' || rowStatus === 'Inativo' || rowStatus === 'Muito Inativo';
                } else {
                    statusMatch = rowStatus === filtroStatusAtivo;
                }
            }
            
            if (filtroGerenciaAtiva && filtroGerenciaAtiva !== 'todas' && filtroGerenciaAtiva !== '') {
                gerenciaMatch = rowGerencia === filtroGerenciaAtiva;
            }
            
            // Aplicar filtro de busca
            const text = row.textContent.toLowerCase();
            const searchMatch = text.indexOf(filter) > -1;
            
            if (statusMatch && gerenciaMatch && searchMatch) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });
        
        document.getElementById('contador').textContent = 'Mostrando ' + visibleCount + ' corretores';
        updateRowNumbers();
    });
    
    // Função para exportar para Excel
    function exportarParaExcel() {
        // Criar tabela HTML para exportação
        let tableHTML = '<table border="1">';
        
        // Cabeçalho
        tableHTML += '<tr>';
        document.querySelectorAll('#tabelaCorretores thead th').forEach(th => {
            tableHTML += '<th>' + th.textContent + '</th>';
        });
        tableHTML += '</tr>';
        
        // Dados (apenas linhas visíveis)
        document.querySelectorAll('#tabelaCorretores tbody tr:visible').forEach(row => {
            tableHTML += '<tr>';
            row.querySelectorAll('td').forEach(td => {
                // Remover tags HTML dos badges
                let content = td.innerHTML;
                content = content.replace(/<[^>]*>/g, '');
                content = content.replace(/&nbsp;/g, ' ');
                tableHTML += '<td>' + content + '</td>';
            });
            tableHTML += '</tr>';
        });
        
        tableHTML += '</table>';
        
        // Criar link para download
        const blob = new Blob([tableHTML], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Corretores_' + new Date().toISOString().slice(0,10) + '.xls';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
    
    // Função para imprimir relatório
    function imprimirRelatorio() {
        window.print();
    }
    
    // Inicializar
    document.addEventListener('DOMContentLoaded', function() {
        updateRowNumbers();
    });
    </script>
</body>
</html>