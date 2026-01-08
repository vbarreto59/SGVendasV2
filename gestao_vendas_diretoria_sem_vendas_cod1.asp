<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    On Error Resume Next 
    set objMail = server.createobject("CDONTS.NewMail")
    if Err.Number <> 0 then 
        set objMail = Nothing ' Garante que a variável seja liberada, mesmo que não criada
    else
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
        objMail.Subject = "SV-DIR-CORRET2-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
        objMail.MailFormat = 0 ' 0 = Texto Simples
        objMail.Body = "Página Corretores. " & Ucase(Session("Usuario"))
        objMail.Send
        set objMail = Nothing
    end if 
    On Error GoTo 0 
end if
%>


<%
' Configuração para evitar cache
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"

diretoriaID = Session("Dir_DiretoriaID")

' Conexão com os bancos
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Set connUsuarios = Server.CreateObject("ADODB.Connection")
connUsuarios.Open StrConn

' Data atual para cálculos
Dim anoAtual, mesAtual, diaAtual
anoAtual = Year(Date())
mesAtual = Month(Date())
diaAtual = Day(Date())

' Meses de 2025 e 2026 - CORRIGIDO: 2025 primeiro, depois 2026
Dim meses(24), mesesLabels(24)
Dim m, anoMes, cont
cont = 0

' Primeiro 2025
For ano = 2025 To 2025
    For mes = 1 To 12
        meses(cont) = ano & "-" & Right("0" & mes, 2)
        mesesLabels(cont) = MonthName(mes, True) & "/" & Right(ano, 2)
        cont = cont + 1
    Next
Next

' Depois 2026
For ano = 2026 To 2026
    For mes = 1 To 12
        meses(cont) = ano & "-" & Right("0" & mes, 2)
        mesesLabels(cont) = MonthName(mes, True) & "/" & Right(ano, 2)
        cont = cont + 1
    Next
Next

' Obter usuários ativos
Dim usuariosAtivosDict
Set usuariosAtivosDict = Server.CreateObject("Scripting.Dictionary")

sqlUsuariosAtivos = "SELECT UserId, Nome FROM Usuarios WHERE Ativo = True"
Set rsUsuariosAtivos = Server.CreateObject("ADODB.Recordset")
rsUsuariosAtivos.Open sqlUsuariosAtivos, connUsuarios

Do While Not rsUsuariosAtivos.EOF
    userId = rsUsuariosAtivos("UserId")
    If Not IsNull(userId) Then
        usuariosAtivosDict.Add CStr(userId), rsUsuariosAtivos("Nome")
    End If
    rsUsuariosAtivos.MoveNext
Loop
rsUsuariosAtivos.Close
Set rsUsuariosAtivos = Nothing

' Criar string com IDs de usuários ativos
Dim idsUsuariosAtivos
idsUsuariosAtivos = ""
Dim userIds
userIds = usuariosAtivosDict.Keys
For Each userId In userIds
    If idsUsuariosAtivos <> "" Then idsUsuariosAtivos = idsUsuariosAtivos & ","
    idsUsuariosAtivos = idsUsuariosAtivos & userId
Next

' Arrays para armazenar os dados
Dim corretores(), gerenciaArr(), diasSemVenderArr(), mesesSemVenderArr()
Dim ultimaVendaAno(), ultimaVendaMes(), ultimaVendaDia()
Dim count
count = 0

' Array 2D para vendas por mês - será inicializado depois
Dim vendasPorMes()

' Variável para armazenar todas as gerências únicas
Dim gerenciasUnicasDict
Set gerenciasUnicasDict = Server.CreateObject("Scripting.Dictionary")

If idsUsuariosAtivos <> "" Then
    ' Consulta principal para obter corretores
    sql = "SELECT DISTINCT v.Corretor, v.CorretorId FROM Vendas v "
    sql = sql & " WHERE v.Excluido = 0 "
    sql = sql & " AND v.CorretorId IN (" & idsUsuariosAtivos & ")"
    
    If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
        sql = sql & " AND v.DiretoriaId = " & CLng(diretoriaID)
    End If
    
    sql = sql & " AND v.Corretor IS NOT NULL AND TRIM(v.Corretor) <> ''"
    sql = sql & " ORDER BY v.Corretor"
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, connSales
    
    ' Primeiro, contar quantos corretores temos
    Dim totalCorretores
    totalCorretores = 0
    
    ' Usar um Recordset temporário para contar
    Set rsCount = Server.CreateObject("ADODB.Recordset")
    rsCount.Open sql, connSales, 1, 3 ' adOpenStatic, adLockReadOnly
    rsCount.MoveLast
    totalCorretores = rsCount.RecordCount
    rsCount.Close
    Set rsCount = Nothing
    
    ' Inicializar arrays com o tamanho correto
    If totalCorretores > 0 Then
        ReDim corretores(totalCorretores - 1)
        ReDim gerenciaArr(totalCorretores - 1)
        ReDim diasSemVenderArr(totalCorretores - 1)
        ReDim mesesSemVenderArr(totalCorretores - 1)
        ReDim ultimaVendaAno(totalCorretores - 1)
        ReDim ultimaVendaMes(totalCorretores - 1)
        ReDim ultimaVendaDia(totalCorretores - 1)
        ReDim vendasPorMes(totalCorretores - 1, 23)
        
        ' Inicializar array de vendas com zeros
        For i = 0 To totalCorretores - 1
            For j = 0 To 23
                vendasPorMes(i, j) = 0
            Next
        Next
    End If
    
    ' Reiniciar o recordset para processar os dados
    rs.Close
    rs.Open sql, connSales
    
    Do While Not rs.EOF And count < totalCorretores
        corretorNome = Trim(rs("Corretor"))
        If Not IsNull(rs("CorretorId")) Then
            corretorId = CStr(rs("CorretorId"))
        Else
            corretorId = ""
        End If
        
        If corretorNome <> "" And usuariosAtivosDict.Exists(corretorId) Then
            
            ' Obter última venda para calcular dias sem vender
            sqlUltimaVenda = "SELECT TOP 1 AnoVenda, MesVenda, DiaVenda FROM Vendas "
            sqlUltimaVenda = sqlUltimaVenda & " WHERE Excluido = 0 AND CorretorId = " & corretorId
            sqlUltimaVenda = sqlUltimaVenda & " ORDER BY AnoVenda DESC, MesVenda DESC, DiaVenda DESC"
            
            Set rsUltimaVenda = Server.CreateObject("ADODB.Recordset")
            rsUltimaVenda.Open sqlUltimaVenda, connSales
            
            Dim ultimoAno, ultimoMes, ultimoDia, diasSemVender
            ultimoAno = 0
            ultimoMes = 0
            ultimoDia = 0
            diasSemVender = 9999
            
            If Not rsUltimaVenda.EOF Then
                If Not IsNull(rsUltimaVenda("AnoVenda")) Then ultimoAno = rsUltimaVenda("AnoVenda")
                If Not IsNull(rsUltimaVenda("MesVenda")) Then ultimoMes = rsUltimaVenda("MesVenda")
                If Not IsNull(rsUltimaVenda("DiaVenda")) Then ultimoDia = rsUltimaVenda("DiaVenda")
                
                ' Calcular dias sem vender
                If ultimoAno > 0 And ultimoMes > 0 And ultimoDia > 0 Then
                    Dim dataUltimaVenda, dataAtual, diffDias
                    dataUltimaVenda = DateSerial(ultimoAno, ultimoMes, ultimoDia)
                    dataAtual = Date()
                    diffDias = DateDiff("d", dataUltimaVenda, dataAtual)
                    If diffDias >= 0 Then
                        diasSemVender = diffDias
                    Else
                        diasSemVender = 0
                    End If
                End If
            End If
            rsUltimaVenda.Close
            Set rsUltimaVenda = Nothing
            
            ' Obter gerência atual
            sqlGerencia = "SELECT TOP 1 Gerencia FROM Vendas "
            sqlGerencia = sqlGerencia & " WHERE Excluido = 0 AND CorretorId = " & corretorId
            sqlGerencia = sqlGerencia & " ORDER BY AnoVenda DESC, MesVenda DESC"
            
            Set rsGerencia = Server.CreateObject("ADODB.Recordset")
            rsGerencia.Open sqlGerencia, connSales
            
            Dim gerenciaAtual
            gerenciaAtual = ""
            If Not rsGerencia.EOF Then
                If Not IsNull(rsGerencia("Gerencia")) Then
                    gerenciaAtual = Trim(rsGerencia("Gerencia"))
                End If
            End If
            rsGerencia.Close
            Set rsGerencia = Nothing
            
            ' Adicionar gerência ao dicionário de gerências únicas
            If gerenciaAtual <> "" Then
                If Not gerenciasUnicasDict.Exists(gerenciaAtual) Then
                    gerenciasUnicasDict.Add gerenciaAtual, 0
                End If
            End If
            
            ' Calcular meses sem vender
            Dim mesesSemVender
            If ultimoAno > 0 And ultimoMes > 0 Then
                mesesSemVender = ((anoAtual - ultimoAno) * 12) + (mesAtual - ultimoMes)
                If mesesSemVender < 0 Then mesesSemVender = 0
            Else
                mesesSemVender = 999
            End If
            
            ' Consulta para vendas por mês
            sqlVendasMes = "SELECT AnoVenda, MesVenda, COUNT(*) as QtdVendas FROM Vendas "
            sqlVendasMes = sqlVendasMes & " WHERE Excluido = 0 AND CorretorId = " & corretorId
            sqlVendasMes = sqlVendasMes & " GROUP BY AnoVenda, MesVenda"
            
            Set rsVendasMes = Server.CreateObject("ADODB.Recordset")
            rsVendasMes.Open sqlVendasMes, connSales
            
            Do While Not rsVendasMes.EOF
                anoVenda = rsVendasMes("AnoVenda")
                mesVenda = rsVendasMes("MesVenda")
                qtdVendas = rsVendasMes("QtdVendas")
                
                ' Encontrar índice correspondente
                For j = 0 To 23
                    If meses(j) = anoVenda & "-" & Right("0" & mesVenda, 2) Then
                        vendasPorMes(count, j) = qtdVendas
                        Exit For
                    End If
                Next
                rsVendasMes.MoveNext
            Loop
            rsVendasMes.Close
            Set rsVendasMes = Nothing
            
            ' Armazenar dados nos arrays
            corretores(count) = UCase(corretorNome)
            gerenciaArr(count) = gerenciaAtual
            diasSemVenderArr(count) = diasSemVender
            mesesSemVenderArr(count) = mesesSemVender
            ultimaVendaAno(count) = ultimoAno
            ultimaVendaMes(count) = ultimoMes
            ultimaVendaDia(count) = ultimoDia
            
            count = count + 1
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End If

Set usuariosAtivosDict = Nothing

' ============= ORDENAÇÃO POR DIAS SEM VENDER (DECRESCENTE) =============
If count > 1 Then
    Dim i, j, temp, k
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            ' Ordenar por diasSemVenderArr decrescente (maior primeiro)
            If diasSemVenderArr(i) < diasSemVenderArr(j) Then
                ' Trocar todos os dados entre as posições i e j
                
                ' Trocar corretores
                temp = corretores(i)
                corretores(i) = corretores(j)
                corretores(j) = temp
                
                ' Trocar gerenciaArr
                temp = gerenciaArr(i)
                gerenciaArr(i) = gerenciaArr(j)
                gerenciaArr(j) = temp
                
                ' Trocar diasSemVenderArr
                temp = diasSemVenderArr(i)
                diasSemVenderArr(i) = diasSemVenderArr(j)
                diasSemVenderArr(j) = temp
                
                ' Trocar mesesSemVenderArr
                temp = mesesSemVenderArr(i)
                mesesSemVenderArr(i) = mesesSemVenderArr(j)
                mesesSemVenderArr(j) = temp
                
                ' Trocar ultimaVendaAno
                temp = ultimaVendaAno(i)
                ultimaVendaAno(i) = ultimaVendaAno(j)
                ultimaVendaAno(j) = temp
                
                ' Trocar ultimaVendaMes
                temp = ultimaVendaMes(i)
                ultimaVendaMes(i) = ultimaVendaMes(j)
                ultimaVendaMes(j) = temp
                
                ' Trocar ultimaVendaDia
                temp = ultimaVendaDia(i)
                ultimaVendaDia(i) = ultimaVendaDia(j)
                ultimaVendaDia(j) = temp
                
                ' Trocar linha completa do array 2D vendasPorMes
                For k = 0 To 23
                    temp = vendasPorMes(i, k)
                    vendasPorMes(i, k) = vendasPorMes(j, k)
                    vendasPorMes(j, k) = temp
                Next
            End If
        Next
    Next
End If
' ============= FIM DA ORDENAÇÃO =============
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Simplificado de Corretores</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        body { 
            background: #f8f9fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 14px;
        }
        .table th { 
            background-color: #f8f9fa; 
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        .badge-venda {
            font-size: 0.75rem;
            min-width: 25px;
            height: 25px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border-radius: 12px;
        }
        .header-section {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: white;
        }
        .mes-header {
            background-color: #e9ecef;
            font-weight: bold;
            text-align: center;
            padding: 8px 5px;
            font-size: 0.8rem;
            writing-mode: vertical-rl;
            transform: rotate(180deg);
            white-space: nowrap;
            height: 150px;
            width: 25px;
        }
        .corretor-row:hover {
            background-color: #f5f5f5;
        }
        .dias-sem-vender {
            font-weight: bold;
            padding: 5px 10px;
            border-radius: 4px;
            display: inline-block;
            min-width: 70px;
            text-align: center;
        }
        .dias-vermelho { background-color: #dc3545; color: white; }
        .dias-laranja { background-color: #fd7e14; color: white; }
        .dias-amarelo { background-color: #ffc107; color: black; }
        .dias-verde { background-color: #28a745; color: white; }
        .meses-sem-vender {
            font-weight: bold;
            padding: 5px;
            border-radius: 4px;
            text-align: center;
            min-width: 40px;
            display: inline-block;
        }
        .table-container {
            overflow-x: auto;
        }
        .fixed-column {
            position: sticky;
            left: 0;
            background-color: white;
            z-index: 5;
        }
        .fixed-column-2 {
            position: sticky;
            left: 100px;
            background-color: white;
            z-index: 5;
        }
        .ano-header {
            background-color: #6c757d;
            color: white;
            text-align: center;
            font-weight: bold;
            padding: 5px;
            font-size: 0.9rem;
            writing-mode: vertical-rl;
            transform: rotate(180deg);
            white-space: nowrap;
            height: 150px;
            width: 25px;
        }
        .meses-2025 {
            background-color: #e3f2fd; /* Azul claro para 2025 */
        }
        .meses-2026 {
            background-color: #f3e5f5; /* Roxo claro para 2026 */
        }
        .ordering-note {
            font-size: 0.8rem;
            color: #6c757d;
            font-style: italic;
        }
        .filter-section {
            background-color: #fff;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .filter-badge {
            cursor: pointer;
            transition: all 0.2s;
        }
        .filter-badge:hover {
            transform: translateY(-2px);
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }
        .filter-badge.active {
            background-color: #0d6efd !important;
            color: white !important;
        }
    </style>
</head>
<body>
    <div class="container-fluid mt-3">
        <!-- HEADER -->
        <div class="header-section">
            <div class="row align-items-center">
                <div class="col-md-8">
                    <h4 class="mb-1"><i class="fas fa-users me-2"></i>Relatório Simplificado de Corretores</h4>
                    <p class="mb-0">Diretoria: <strong><%=Session("Dir_Nome")%></strong> | Total: <strong><%=count%></strong> corretores ativos</p>
                    <p class="mb-0 ordering-note">
                        <i class="fas fa-sort-amount-down me-1"></i>Ordenado por: <strong>Dias sem vender (decrescente)</strong>
                    </p>
                </div>
                <div class="col-md-4 text-end">
                    <small>Data: <%=Day(Date())%>/<%=Month(Date())%>/<%=Year(Date())%></small>
                </div>
            </div>
        </div>
        
        <!-- FILTRO POR GERÊNCIA -->
        <div class="filter-section">
            <h6 class="mb-3"><i class="fas fa-filter me-2"></i>Filtrar por Gerência</h6>
            <div class="d-flex flex-wrap gap-2 mb-3">
                <button class="btn btn-sm btn-outline-primary filter-badge active" onclick="filtrarPorGerencia('')" id="filtroTodas">
                    <i class="fas fa-eye me-1"></i>Todas Gerências (<%=count%>)
                </button>
                <%
                ' Obter chaves do dicionário de gerências
                Dim gerenciaKeys, gerenciaKey
                gerenciaKeys = gerenciasUnicasDict.Keys
                
                ' Contar corretores por gerência
                For Each gerenciaKey In gerenciaKeys
                    Dim contadorGerencia
                    contadorGerencia = 0
                    
                    For i = 0 To count - 1
                        If gerenciaArr(i) = gerenciaKey Then
                            contadorGerencia = contadorGerencia + 1
                        End If
                    Next
                    
                    ' Atualizar contagem no dicionário
                    gerenciasUnicasDict(gerenciaKey) = contadorGerencia
                Next
                
                ' Gerar botões para cada gerência
                For Each gerenciaKey In gerenciaKeys
                    If gerenciaKey <> "" Then
                        contadorGerencia = gerenciasUnicasDict(gerenciaKey)
                        
                        ' Gerar cor aleatória baseada no nome da gerência
                        Randomize Len(gerenciaKey)
                        Dim corAleatoria
                        corAleatoria = Int((5 * Rnd) + 1)
                        
                        Dim corClasse
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
                <button class="btn btn-sm <%=corClasse%> filter-badge" onclick="filtrarPorGerencia('<%=Server.HTMLEncode(gerenciaKey)%>')" id="filtro<%=Replace(gerenciaKey, " ", "")%>">
                    <i class="fas fa-building me-1"></i><%=gerenciaKey%> (<%=contadorGerencia%>)
                </button>
                <%
                    End If
                Next
                
                ' Botão para corretores sem gerência
                Dim contadorSemGerencia
                contadorSemGerencia = 0
                For i = 0 To count - 1
                    If gerenciaArr(i) = "" Then
                        contadorSemGerencia = contadorSemGerencia + 1
                    End If
                Next
                
                If contadorSemGerencia > 0 Then
                %>
                <button class="btn btn-sm btn-outline-secondary filter-badge" onclick="filtrarPorGerencia('sem_gerencia')" id="filtroSemGerencia">
                    <i class="fas fa-question-circle me-1"></i>Sem Gerência (<%=contadorSemGerencia%>)
                </button>
                <% End If %>
            </div>
            <div class="mt-2">
                <small class="text-muted">
                    <i class="fas fa-info-circle me-1"></i>Clique em uma gerência para filtrar. Clique em "Todas Gerências" para remover o filtro.
                </small>
            </div>
        </div>
        
        <% If count > 0 Then %>
        <!-- TABELA -->
        <div class="card">
            <div class="card-header text-white d-flex justify-content-between align-items-center" style="background: #f39c12;">
                <div>
                    <i class="fas fa-table me-2"></i>Quantidade de Vendas por Mês (2025-2026)
                    <span class="badge bg-light text-dark ms-2" id="contadorTabela">Mostrando <%=count%> corretores</span>
                    <span class="badge bg-info text-white ms-2" id="filtroAtivo" style="display: none;"></span>
                </div>
                <div>
                    <small>Legenda: <span class="badge bg-success">0-30 dias</span> <span class="badge bg-warning">31-90 dias</span> <span class="badge bg-danger">+90 dias</span></small>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="table-container">
                    <table class="table table-bordered table-hover mb-0" id="tabelaCorretores">
                        <thead>
                            <tr>
                                <th width="50" class="fixed-column">#</th>
                                <th width="200" class="fixed-column-2">Corretor</th>
                                <th width="150">Gerência</th>
                                <th width="120" class="text-center">Dias sem Vender</th>
                                <th width="120" class="text-center">Meses sem Vender</th>
                                
                                <!-- Cabeçalho dos meses - CORRIGIDO: 2025 e 2026 lado a lado -->
                                <!-- Primeiro cabeçalho com os anos -->
                                <%
                                ' Cabeçalho para 2025
                                Response.Write "<th colspan='12' class='ano-header meses-2025'>2025</th>"
                                
                                ' Cabeçalho para 2026
                                Response.Write "<th colspan='12' class='ano-header meses-2026'>2026</th>"
                                %>
                            </tr>
                            <tr>
                                <th colspan="5"></th>
                                <!-- Meses de 2025 (0-11) -->
                                <%
                                For i = 0 To 11
                                    Response.Write "<th class='mes-header meses-2025'>" & mesesLabels(i) & "</th>"
                                Next
                                
                                ' Meses de 2026 (12-23)
                                For i = 12 To 23
                                    Response.Write "<th class='mes-header meses-2026'>" & mesesLabels(i) & "</th>"
                                Next
                                %>
                            </tr>
                        </thead>
                        <tbody id="tbodyCorretores">
                            <% 
                            For i = 0 To count - 1
                                ' Determinar cor para dias sem vender
                                Dim diasClass
                                If diasSemVenderArr(i) = 9999 Then
                                    diasClass = "bg-secondary text-white"
                                ElseIf diasSemVenderArr(i) >= 90 Then
                                    diasClass = "dias-vermelho"
                                ElseIf diasSemVenderArr(i) >= 30 Then
                                    diasClass = "dias-laranja"
                                ElseIf diasSemVenderArr(i) >= 7 Then
                                    diasClass = "dias-amarelo"
                                Else
                                    diasClass = "dias-verde"
                                End If
                                
                                ' Determinar cor para meses sem vender
                                Dim mesesClass
                                If mesesSemVenderArr(i) = 999 Then
                                    mesesClass = "bg-secondary text-white"
                                ElseIf mesesSemVenderArr(i) >= 12 Then
                                    mesesClass = "bg-danger text-white"
                                ElseIf mesesSemVenderArr(i) >= 6 Then
                                    mesesClass = "bg-warning text-white"
                                ElseIf mesesSemVenderArr(i) >= 3 Then
                                    mesesClass = "bg-info text-white"
                                Else
                                    mesesClass = "bg-success text-white"
                                End If
                                
                                ' Determinar ID da gerência para filtro
                                Dim gerenciaId
                                If gerenciaArr(i) = "" Then
                                    gerenciaId = "sem_gerencia"
                                Else
                                    gerenciaId = gerenciaArr(i)
                                End If
                            %>
                            <tr class="corretor-row" data-gerencia="<%=Server.HTMLEncode(gerenciaId)%>">
                                <td class="fw-bold fixed-column"><%=i+1%></td>
                                <td class="fixed-column-2">
                                    <strong><%=corretores(i)%></strong>
                                    <% If ultimaVendaAno(i) > 0 Then %>
                                    <br>
                                    <span class="badge bg-primary"><%=gerenciaArr(i)%></span>
                                    <small class="text-muted">
                                        Última: <%=ultimaVendaDia(i)%>/<%=ultimaVendaMes(i)%>/<%=ultimaVendaAno(i)%>
                                    </small>
                                    <% Else %>
                                    <br>
                                    <small class="text-muted">Nunca vendeu</small>
                                    <% End If %>
                                </td>
                                <td>
                                    <% If gerenciaArr(i) <> "" Then %>
                                        
                                    <% Else %>
                                        <span class="badge bg-secondary">Sem gerência</span>
                                    <% End If %>
                                </td>
                                <td class="text-center">
                                    <span class="dias-sem-vender <%=diasClass%>">
                                        <% If diasSemVenderArr(i) = 9999 Then %>
                                            N/A
                                        <% Else %>
                                            <%=diasSemVenderArr(i)%> dias
                                        <% End If %>
                                    </span>
                                </td>
                                <td class="text-center">
                                    <span class="meses-sem-vender <%=mesesClass%>">
                                        <% If mesesSemVenderArr(i) = 999 Then %>
                                            N/A
                                        <% Else %>
                                            <%=mesesSemVenderArr(i)%>
                                        <% End If %>
                                    </span>
                                </td>
                                
                                <!-- Vendas por mês - 2025 primeiro (0-11) -->
                                <%
                                For j = 0 To 11
                                    qtdVendas = vendasPorMes(i, j)
                                %>
                                <td class="text-center align-middle">
                                    <% If qtdVendas > 0 Then %>
                                        <span class="badge-venda bg-success text-white" 
                                              title="<%=mesesLabels(j)%>: <%=qtdVendas%> vendas">
                                            <%=qtdVendas%>
                                        </span>
                                    <% Else %>
                                        <span class="badge-venda bg-light text-muted" 
                                              title="<%=mesesLabels(j)%>: 0 vendas">
                                            0
                                        </span>
                                    <% End If %>
                                </td>
                                <% Next %>
                                
                                <!-- Vendas por mês - 2026 depois (12-23) -->
                                <%
                                For j = 12 To 23
                                    qtdVendas = vendasPorMes(i, j)
                                %>
                                <td class="text-center align-middle">
                                    <% If qtdVendas > 0 Then %>
                                        <span class="badge-venda bg-success text-white" 
                                              title="<%=mesesLabels(j)%>: <%=qtdVendas%> vendas">
                                            <%=qtdVendas%>
                                        </span>
                                    <% Else %>
                                        <span class="badge-venda bg-light text-muted" 
                                              title="<%=mesesLabels(j)%>: 0 vendas">
                                            0
                                        </span>
                                    <% End If %>
                                </td>
                                <% Next %>
                            </tr>
                            <% Next %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- RESUMO -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="card border-0 shadow-sm">
                    <div class="card-body text-center">
                        <h6 class="text-success"><i class="fas fa-check-circle"></i> Ativos Recentes</h6>
                        <%
                        Dim ativosRecentes
                        ativosRecentes = 0
                        For i = 0 To count - 1
                            If diasSemVenderArr(i) <= 30 And diasSemVenderArr(i) <> 9999 Then
                                ativosRecentes = ativosRecentes + 1
                            End If
                        Next
                        %>
                        <h3 class="text-success"><%=ativosRecentes%></h3>
                        <small class="text-muted">≤ 30 dias sem vender</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card border-0 shadow-sm">
                    <div class="card-body text-center">
                        <h6 class="text-warning"><i class="fas fa-exclamation-triangle"></i> Atenção</h6>
                        <%
                        Dim atencao
                        atencao = 0
                        For i = 0 To count - 1
                            If diasSemVenderArr(i) > 30 And diasSemVenderArr(i) <= 90 Then
                                atencao = atencao + 1
                            End If
                        Next
                        %>
                        <h3 class="text-warning"><%=atencao%></h3>
                        <small class="text-muted">31-90 dias sem vender</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card border-0 shadow-sm">
                    <div class="card-body text-center">
                        <h6 class="text-danger"><i class="fas fa-user-clock"></i> Inativos</h6>
                        <%
                        Dim inativos
                        inativos = 0
                        For i = 0 To count - 1
                            If diasSemVenderArr(i) > 90 And diasSemVenderArr(i) <> 9999 Then
                                inativos = inativos + 1
                            End If
                        Next
                        %>
                        <h3 class="text-danger"><%=inativos%></h3>
                        <small class="text-muted">> 90 dias sem vender</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card border-0 shadow-sm">
                    <div class="card-body text-center">
                        <h6 class="text-secondary"><i class="fas fa-question-circle"></i> Sem Vendas</h6>
                        <%
                        Dim semVendas
                        semVendas = 0
                        For i = 0 To count - 1
                            If diasSemVenderArr(i) = 9999 Then
                                semVendas = semVendas + 1
                            End If
                        Next
                        %>
                        <h3 class="text-secondary"><%=semVendas%></h3>
                        <small class="text-muted">Nenhuma venda registrada</small>
                    </div>
                </div>
            </div>
        </div>
        <% Else %>
        <div class="alert alert-warning text-center">
            <i class="fas fa-exclamation-triangle fa-2x mb-3"></i>
            <h5>Nenhum corretor ativo encontrado</h5>
            <p class="mb-0">Não há corretores com vendas na diretoria.</p>
        </div>
        <% End If %>
    </div>
    
    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
    // Variável global para controlar o filtro ativo
    let filtroGerenciaAtiva = '';
    
    // Função para filtrar por gerência
    function filtrarPorGerencia(gerencia) {
        // Atualizar variável global
        filtroGerenciaAtiva = gerencia;
        
        // Remover classe active de todos os botões
        document.querySelectorAll('.filter-badge').forEach(btn => {
            btn.classList.remove('active');
        });
        
        // Adicionar classe active ao botão selecionado
        if (gerencia === '') {
            document.getElementById('filtroTodas').classList.add('active');
        } else if (gerencia === 'sem_gerencia') {
            document.getElementById('filtroSemGerencia').classList.add('active');
        } else {
            const botaoId = 'filtro' + gerencia.replace(/ /g, '');
            if (document.getElementById(botaoId)) {
                document.getElementById(botaoId).classList.add('active');
            }
        }
        
        // Filtrar linhas da tabela
        const rows = document.querySelectorAll('#tbodyCorretores tr');
        let visibleCount = 0;
        
        rows.forEach((row, index) => {
            const gerenciaRow = row.getAttribute('data-gerencia');
            
            if (filtroGerenciaAtiva === '' || 
                (filtroGerenciaAtiva === 'sem_gerencia' && gerenciaRow === 'sem_gerencia') ||
                gerenciaRow === filtroGerenciaAtiva) {
                row.style.display = '';
                visibleCount++;
                
                // Atualizar número da linha
                const tdNumero = row.querySelector('.fixed-column');
                if (tdNumero) {
                    tdNumero.textContent = visibleCount;
                }
            } else {
                row.style.display = 'none';
            }
        });
        
        // Atualizar contador
        document.getElementById('contadorTabela').textContent = 'Mostrando ' + visibleCount + ' corretores';
        
        // Atualizar badge do filtro ativo
        const filtroAtivoBadge = document.getElementById('filtroAtivo');
        if (filtroGerenciaAtiva !== '') {
            let filtroTexto = 'Filtro: ';
            if (filtroGerenciaAtiva === 'sem_gerencia') {
                filtroTexto += 'Sem Gerência';
            } else {
                filtroTexto += filtroGerenciaAtiva;
            }
            filtroAtivoBadge.textContent = filtroTexto;
            filtroAtivoBadge.style.display = 'inline';
        } else {
            filtroAtivoBadge.style.display = 'none';
        }
    }
    
    // Função para destacar meses com vendas
    document.addEventListener('DOMContentLoaded', function() {
        const badges = document.querySelectorAll('.badge-venda.bg-success');
        badges.forEach(badge => {
            if (parseInt(badge.textContent) > 0) {
                badge.style.boxShadow = '0 0 5px rgba(0,128,0,0.5)';
            }
        });
    });
    
    // Função para limpar filtro (atalho de teclado)
    document.addEventListener('keydown', function(event) {
        if (event.key === 'Escape') {
            filtrarPorGerencia('');
        }
    });
    </script>
</body>
</html>

<%
' Fechar conexões
connSales.Close
Set connSales = Nothing

connUsuarios.Close
Set connUsuarios = Nothing
%>