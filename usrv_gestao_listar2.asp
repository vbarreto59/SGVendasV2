<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: IFWQVSWHFD          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' ==============================================================================
' FUNÇÕES DE UTILIDADE
' ==============================================================================

' Função para tratar valores NULL/Empty do Recordset com um valor seguro (como Nz no Access)
Function SafeValue(ByVal Value, ByVal DefaultValue)
    If IsNull(Value) Or IsEmpty(Value) Then
        SafeValue = DefaultValue
    Else
        SafeValue = Value
    End If
End Function

' ==============================================================================
' PROCESSAMENTO DE FORMULÁRIOS
' ==============================================================================

' Processar ativação/desativação do usuário
If Request.Form("acao") = "toggle_status" Then
    Dim userId, novoStatus
    userId = Request.Form("user_id")
    novoStatus = Request.Form("novo_status")
    
    If userId <> "" And novoStatus <> "" Then
        On Error Resume Next
        
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = StrConn
        cmd.CommandText = "UPDATE Usuarios SET Ativo = ? WHERE UserID = ? AND IdEmp = 2"
        ' Assume que 3 é adInteger (Tipo de dado para Ativo/UserID)
        cmd.Parameters.Append cmd.CreateParameter("Ativo", 3, 1, , novoStatus) 
        cmd.Parameters.Append cmd.CreateParameter("UserID", 3, 1, , userId)
        cmd.Execute
        
        If Err.Number = 0 Then
            Response.Redirect "?success=1&userid=" & userId
        Else
            Response.Redirect "?error=1&msg=" & Server.URLEncode(Err.Description)
        End If
        On Error GoTo 0
        Set cmd = Nothing
    End If
End If

' Processar atualização de diretoria e gerência
If Request.Form("acao") = "update_diretoria_gerencia" Then
    Dim userIdUpdate, diretoriaId, gerenciaId
    userIdUpdate = Request.Form("user_id")
    diretoriaId = Request.Form("diretoria_id")
    gerenciaId = Request.Form("gerencia_id")
    
    If userIdUpdate <> "" And diretoriaId <> "" And gerenciaId <> "" Then
        
        Dim nomeDiretoria, nomeGerencia
        nomeDiretoria = ""
        nomeGerencia = ""
        
        On Error Resume Next
        ' Obter nomes da diretoria e gerência - CORRIGIDO
        Dim rsNomes
        Set rsNomes = Server.CreateObject("ADODB.Recordset")
        ' Consulta corrigida para usar JOIN correto
        rsNomes.Open "SELECT d.NomeDiretoria, g.NomeGerencia " & _
                     "FROM Diretorias d " & _
                     "INNER JOIN Gerencias g ON d.DiretoriaID = g.DiretoriaID " & _
                     "WHERE d.DiretoriaID = " & diretoriaId & " AND g.GerenciaID = " & gerenciaId, StrConn
        
        If Err.Number <> 0 Then
            Response.Redirect "?error=2&msg=" & Server.URLEncode("Erro ao buscar nomes: " & Err.Description)
            Err.Clear
        ElseIf Not rsNomes.EOF Then
            nomeDiretoria = SafeValue(rsNomes("NomeDiretoria"), "")
            nomeGerencia = SafeValue(rsNomes("NomeGerencia"), "")
        Else
            ' Se não encontrou, tentar buscar separadamente
            rsNomes.Close
            ' Buscar nome da diretoria
            rsNomes.Open "SELECT NomeDiretoria FROM Diretorias WHERE DiretoriaID = " & diretoriaId, StrConn
            If Not rsNomes.EOF Then
                nomeDiretoria = SafeValue(rsNomes("NomeDiretoria"), "")
            End If
            rsNomes.Close
            
            ' Buscar nome da gerência
            rsNomes.Open "SELECT NomeGerencia FROM Gerencias WHERE GerenciaID = " & gerenciaId, StrConn
            If Not rsNomes.EOF Then
                nomeGerencia = SafeValue(rsNomes("NomeGerencia"), "")
            End If
        End If
        
        If Not rsNomes Is Nothing Then
             If rsNomes.State = 1 Then rsNomes.Close
             Set rsNomes = Nothing
        End If
        
        ' Atualizar no banco
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = StrConn
        ' Assume 3=adInteger, 200=adVarChar
        cmd.CommandText = "UPDATE Usuarios SET DiretoriaID = ?, Diretoria = ?, GerenciaID = ?, Gerencia = ? WHERE UserID = ? AND IdEmp = 2"
        cmd.Parameters.Append cmd.CreateParameter("DiretoriaID", 3, 1, , CLng(diretoriaId)) 
        cmd.Parameters.Append cmd.CreateParameter("Diretoria", 200, 1, 100, nomeDiretoria) 
        cmd.Parameters.Append cmd.CreateParameter("GerenciaID", 3, 1, , CLng(gerenciaId)) 
        cmd.Parameters.Append cmd.CreateParameter("Gerencia", 200, 1, 100, nomeGerencia) 
        cmd.Parameters.Append cmd.CreateParameter("UserID", 3, 1, , CLng(userIdUpdate)) 
        cmd.Execute
        
        If Err.Number = 0 Then
            Response.Redirect "?success=2&userid=" & userIdUpdate & "&diretoria=" & diretoriaId & "&gerencia=" & gerenciaId
        Else
            Response.Redirect "?error=2&msg=" & Server.URLEncode("Erro ao atualizar banco: " & Err.Description)
        End If
        On Error GoTo 0
        Set cmd = Nothing
    Else
        Response.Redirect "?error=2&msg=" & Server.URLEncode("Diretoria ou Gerência não selecionada.")
    End If
End If

' ==============================================================================
' PREPARAÇÃO DE DADOS (Contagem e Recordsets)
' ==============================================================================

' Exibir mensagens de sucesso/erro (VBScript inalterado)

' --- [ NOVAS CONSULTAS PARA CARDS DE RESUMO ] ---

Dim totalAtivos, totalInativos, totalSemDiretoria, totalSemGerencia
Dim totalAtivosSemDiretoria, totalAtivosSemGerencia

' 1. Total de Usuários ATIVOS (IdEmp = 2)
Set rsTotalAtivos = Server.CreateObject("ADODB.Recordset")
rsTotalAtivos.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND Ativo = -1", StrConn
totalAtivos = SafeValue(rsTotalAtivos("Total"), 0)
If rsTotalAtivos.State = 1 Then rsTotalAtivos.Close
Set rsTotalAtivos = Nothing

' 2. Total de Usuários INATIVOS (IdEmp = 2)
Set rsTotalInativos = Server.CreateObject("ADODB.Recordset")
rsTotalInativos.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND Ativo = 0", StrConn
totalInativos = SafeValue(rsTotalInativos("Total"), 0)
If rsTotalInativos.State = 1 Then rsTotalInativos.Close
Set rsTotalInativos = Nothing

' 3. Total de Usuários SEM DIRETORIA (ativos e inativos)
Set rsTotalSemDiretoria = Server.CreateObject("ADODB.Recordset")
rsTotalSemDiretoria.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND (DiretoriaID IS NULL OR DiretoriaID = 0)", StrConn
totalSemDiretoria = SafeValue(rsTotalSemDiretoria("Total"), 0)
If rsTotalSemDiretoria.State = 1 Then rsTotalSemDiretoria.Close
Set rsTotalSemDiretoria = Nothing

' 4. Total de Usuários SEM GERÊNCIA (ativos e inativos)
Set rsTotalSemGerencia = Server.CreateObject("ADODB.Recordset")
rsTotalSemGerencia.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND (GerenciaID IS NULL OR GerenciaID = 0)", StrConn
totalSemGerencia = SafeValue(rsTotalSemGerencia("Total"), 0)
If rsTotalSemGerencia.State = 1 Then rsTotalSemGerencia.Close
Set rsTotalSemGerencia = Nothing

' 5. Total de Usuários ATIVOS SEM DIRETORIA
Set rsAtivosSemDiretoria = Server.CreateObject("ADODB.Recordset")
rsAtivosSemDiretoria.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND Ativo = -1 AND (DiretoriaID IS NULL OR DiretoriaID = 0)", StrConn
totalAtivosSemDiretoria = SafeValue(rsAtivosSemDiretoria("Total"), 0)
If rsAtivosSemDiretoria.State = 1 Then rsAtivosSemDiretoria.Close
Set rsAtivosSemDiretoria = Nothing

' 6. Total de Usuários ATIVOS SEM GERÊNCIA
Set rsAtivosSemGerencia = Server.CreateObject("ADODB.Recordset")
rsAtivosSemGerencia.Open "SELECT COUNT(UserID) as Total FROM Usuarios WHERE IdEmp = 2 AND Ativo = -1 AND (GerenciaID IS NULL OR GerenciaID = 0)", StrConn
totalAtivosSemGerencia = SafeValue(rsAtivosSemGerencia("Total"), 0)
If rsAtivosSemGerencia.State = 1 Then rsAtivosSemGerencia.Close
Set rsAtivosSemGerencia = Nothing

' 7. Contagem de Ativos por DIRETORIA
Dim totalPorDiretoria
Set rsContDiretoria = Server.CreateObject("ADODB.Recordset")
rsContDiretoria.Open "SELECT d.NomeDiretoria, COUNT(u.UserID) as Total FROM Usuarios u INNER JOIN Diretorias d ON u.DiretoriaID = d.DiretoriaID WHERE u.IdEmp = 2 AND u.Ativo = -1 GROUP BY d.NomeDiretoria ORDER BY COUNT(u.UserID) DESC", StrConn
totalPorDiretoria = ""
Do While Not rsContDiretoria.EOF
    totalPorDiretoria = totalPorDiretoria & Server.HTMLEncode(rsContDiretoria("NomeDiretoria")) & ": <strong>" & rsContDiretoria("Total") & "</strong><br>"
    rsContDiretoria.MoveNext
Loop
If totalPorDiretoria = "" Then totalPorDiretoria = "<span class='text-muted small'>Nenhuma Diretoria com ativos.</span>"
If rsContDiretoria.State = 1 Then rsContDiretoria.Close
Set rsContDiretoria = Nothing

' 8. Contagem de Ativos por GERÊNCIA
Dim totalPorGerencia
Set rsContGerencia = Server.CreateObject("ADODB.Recordset")
rsContGerencia.Open "SELECT g.NomeGerencia, COUNT(u.UserID) as Total FROM Usuarios u INNER JOIN Gerencias g ON u.GerenciaID = g.GerenciaID WHERE u.IdEmp = 2 AND u.Ativo = -1 GROUP BY g.NomeGerencia ORDER BY COUNT(u.UserID) DESC", StrConn
totalPorGerencia = ""
Do While Not rsContGerencia.EOF
    totalPorGerencia = totalPorGerencia & Server.HTMLEncode(rsContGerencia("NomeGerencia")) & ": <strong>" & rsContGerencia("Total") & "</strong><br>"
    rsContGerencia.MoveNext
Loop
If totalPorGerencia = "" Then totalPorGerencia = "<span class='text-muted small'>Nenhuma Gerência com ativos.</span>"
If rsContGerencia.State = 1 Then rsContGerencia.Close
Set rsContGerencia = Nothing
' --- [ FIM DAS NOVAS CONSULTAS ] ---

' Obter todas as diretorias
Set rsDiretorias = Server.CreateObject("ADODB.Recordset")
rsDiretorias.Open "SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria", StrConn

' Obter todos os usuários e os grupos que participam
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open "SELECT * FROM Usuarios WHERE IdEmp = 2 ORDER BY Usuario ASC", StrConn

' Pré-carregar todas as gerências por diretoria para JavaScript
Dim allGerenciasJSON
allGerenciasJSON = "{"
Set rsAllGerencias = Server.CreateObject("ADODB.Recordset")
If Not rsDiretorias.EOF Then
    rsDiretorias.MoveFirst
    firstDiretoria = True
    Do While Not rsDiretorias.EOF
        If Not firstDiretoria Then allGerenciasJSON = allGerenciasJSON & ","
        
        diretoriaId = rsDiretorias("DiretoriaID")
        allGerenciasJSON = allGerenciasJSON & """" & diretoriaId & """: ["
        
        Set rsTempGerencias = Server.CreateObject("ADODB.Recordset")
        rsTempGerencias.Open "SELECT GerenciaID, NomeGerencia FROM Gerencias WHERE DiretoriaID = " & diretoriaId & " ORDER BY NomeGerencia", StrConn
        
        firstGerencia = True
        Do While Not rsTempGerencias.EOF
            If Not firstGerencia Then allGerenciasJSON = allGerenciasJSON & ","
            ' Substituir aspas por aspas escapadas e codificar HTML
            gerenciaNome = Replace(Server.HTMLEncode(SafeValue(rsTempGerencias("NomeGerencia"), "")), """", "\""")
            allGerenciasJSON = allGerenciasJSON & "{""id"": """ & rsTempGerencias("GerenciaID") & """, ""nome"": """ & gerenciaNome & """}"
            firstGerencia = False
            rsTempGerencias.MoveNext
        Loop
        
        If Not rsTempGerencias Is Nothing Then
            If rsTempGerencias.State = 1 Then rsTempGerencias.Close
            Set rsTempGerencias = Nothing
        End If
        
        allGerenciasJSON = allGerenciasJSON & "]"
        firstDiretoria = False
        rsDiretorias.MoveNext
    Loop
End If
allGerenciasJSON = allGerenciasJSON & "}"

' Obter parâmetros da URL para restaurar seleções
Dim successUserId, successDiretoria, successGerencia
successUserId = Request.QueryString("userid")
successDiretoria = Request.QueryString("diretoria")
successGerencia = Request.QueryString("gerencia")
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
  <title>SGVendas - Lista de Usuários</title>
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">
    
    <link rel="stylesheet" href="https://cdn.datatables.net/1.10.22/css/dataTables.bootstrap4.min.css">
    
  <style>
    body {
      background-color: #f8f9fa;
    }
    .table-responsive {
      background-color: white;
      border-radius: 10px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      padding: 20px;
      margin-top: 20px;
    }
    .table-header {
      background-color: #343a40;
      color: white;
      border-radius: 10px 10px 0 0;
      padding: 15px 20px;
      margin-bottom: 0;
    }
    .btn-sm {
      min-width: 70px;
    }
    .table {
      width: 100%;
    }
    .table th {
      white-space: nowrap;
    }
    .badge-permissao {
      font-size: 0.85em;
      padding: 0.35em 0.65em;
    }
    .badge-grupo {
      font-size: 0.8em;
      margin-right: 3px;
      margin-bottom: 3px;
      display: inline-block;
    }
    .grupos-container {
      max-width: 250px;
    }
    .header-actions {
      margin-bottom: 20px;
    }
    .badge-status {
      font-size: 0.85em;
      padding: 0.5em 0.75em;
      border-radius: 50px;
      min-width: 70px;
      display: inline-block;
      text-align: center;
    }
    .badge-ativo {
      background-color: #28a745;
      color: white;
    }
    .badge-inativo {
      background-color: #dc3545;
      color: white;
    }
    .user-inativo {
      opacity: 0.7;
    }
    .btn-toggle {
      width: 80px;
      font-size: 0.8rem;
    }
    .btn-ativo {
      background-color: #28a745;
      border-color: #28a745;
      color: white;
    }
    .btn-inativo {
      background-color: #6c757d;
      border-color: #6c757d;
      color: white;
    }
    .btn-toggle:hover {
      transform: translateY(-1px);
      transition: all 0.2s;
    }
    .toggle-form {
      display: inline;
    }
    .alert-container {
      position: fixed;
      top: 20px;
      right: 20px;
      z-index: 1000;
      min-width: 300px;
    }
    .select-small {
      font-size: 0.8rem;
      padding: 0.25rem 0.5rem;
      height: calc(1.5em + 0.5rem + 2px);
    }
    .diretoria-gerencia-form {
      display: inline;
    }
    .btn-update-dg {
      font-size: 0.7rem;
      padding: 0.2rem 0.4rem;
    }
    .vazio-indicator {
      color: #6c757d;
      font-style: italic;
      font-size: 0.8rem;
      margin-bottom: 5px;
    }
    .select-container {
      position: relative;
    }
    .card-body h5 {
        font-size: 1rem;
    }
    .card-body p.h3 {
        font-size: 1.5rem;
        margin-bottom: 0;
    }
    .card-body p.small strong {
        font-size: 1.1em;
    }
    .card-resumo {
        border-radius: 10px;
        transition: all 0.3s;
        height: 100%;
    }
    .card-resumo:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 20px rgba(0,0,0,0.1);
    }
    .card-title i {
        margin-right: 8px;
    }
    .stat-badge {
        position: absolute;
        top: 10px;
        right: 10px;
        font-size: 0.7rem;
        padding: 3px 8px;
    }
    .badge-warning-light {
        background-color: #fff3cd;
        color: #856404;
    }
    .badge-danger-light {
        background-color: #f8d7da;
        color: #721c24;
    }
    .badge-info-light {
        background-color: #d1ecf1;
        color: #0c5460;
    }
    .badge-success-light {
        background-color: #d4edda;
        color: #155724;
    }
    .small-text {
        font-size: 0.75rem;
        margin-bottom: 0;
    }
    .progress-thin {
        height: 5px;
        margin-top: 5px;
    }
    .card-footer-stats {
        background-color: rgba(0,0,0,0.03);
        border-top: 1px solid rgba(0,0,0,0.125);
        padding: 5px 15px;
        font-size: 0.7rem;
    }
  </style>
<style>
    body {
        transform: scale(0.8); 
        transform-origin: 0 0; 
        width: calc(100% / 0.8); 
    }
</style>  
</head>
<body>

  <div class="container">
        <div class="alert-container">
      <%
      ' Mostrar alertas Bootstrap
      If Request.QueryString("success") = "1" Then
        Response.Write "<div class='alert alert-success alert-dismissible fade show'>" & _
                       "<i class='fas fa-check-circle mr-2'></i>Status do usuário atualizado com sucesso!" & _
                       "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                       "</div>"
      ElseIf Request.QueryString("success") = "2" Then
        Response.Write "<div class='alert alert-success alert-dismissible fade show'>" & _
                       "<i class='fas fa-check-circle mr-2'></i>Diretoria e Gerência atualizadas com sucesso!" & _
                       "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                       "</div>"
      End If
      
      If Request.QueryString("error") = "1" Then
        Dim errorMsgDisplay
        errorMsgDisplay = Request.QueryString("msg")
        If errorMsgDisplay = "" Then errorMsgDisplay = "Erro desconhecido"
        Response.Write "<div class='alert alert-danger alert-dismissible fade show'>" & _
                       "<i class='fas fa-exclamation-circle mr-2'></i>Erro ao atualizar status: " & Server.HTMLEncode(errorMsgDisplay) & _
                       "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                       "</div>"
      ElseIf Request.QueryString("error") = "2" Then
        errorMsgDisplay = Request.QueryString("msg")
        If errorMsgDisplay = "" Then errorMsgDisplay = "Erro desconhecido"
        Response.Write "<div class='alert alert-danger alert-dismissible fade show'>" & _
                       "<i class='fas fa-exclamation-circle mr-2'></i>Erro ao atualizar diretoria/gerência: " & Server.HTMLEncode(errorMsgDisplay) & _
                       "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                       "</div>"
      End If
      %>
    </div>

    <!-- SEÇÃO DE RESUMO ESTENDIDO -->
    <div class="row mb-4">
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo text-white bg-success">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-light"><i class="fas fa-users"></i> Total</span>
                    <h5 class="card-title"><i class="fas fa-user-check"></i> Ativos</h5>
                    <p class="card-text h3"><%= totalAtivos %></p>
                    <div class="progress progress-thin bg-white">
                        <div class="progress-bar bg-success" style="width: 100%"></div>
                    </div>
                </div>
                <div class="card-footer-stats text-white">
                    <span class="small-text"><i class="fas fa-exclamation-circle"></i> Faltam Diretoria: <%= totalAtivosSemDiretoria %></span>
                </div>
            </div>
        </div>
        
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo text-white bg-secondary">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-light"><i class="fas fa-user-slash"></i> Total</span>
                    <h5 class="card-title"><i class="fas fa-user-times"></i> Inativos</h5>
                    <p class="card-text h3"><%= totalInativos %></p>
                    <div class="progress progress-thin bg-white">
                        <div class="progress-bar bg-secondary" style="width: 100%"></div>
                    </div>
                </div>
                <div class="card-footer-stats text-white">
                    <span class="small-text"><i class="fas fa-percentage"></i> <%= FormatPercent(totalInativos/(totalAtivos+totalInativos), 1) %> do total</span>
                </div>
            </div>
        </div>
        
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo bg-warning-light text-dark">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-warning"><i class="fas fa-exclamation-triangle"></i> Crítico</span>
                    <h5 class="card-title"><i class="fas fa-building"></i> Sem Diretoria</h5>
                    <p class="card-text h3"><%= totalSemDiretoria %></p>
                    <div class="progress progress-thin bg-white">
                        <div class="progress-bar bg-warning" style="width: <%= (totalSemDiretoria/(totalAtivos+totalInativos))*100 %>%"></div>
                    </div>
                </div>
                <div class="card-footer-stats text-dark">
                    <span class="small-text"><i class="fas fa-user-check"></i> Ativos: <%= totalAtivosSemDiretoria %></span>
                </div>
            </div>
        </div>
        
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo bg-danger-light text-dark">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-danger"><i class="fas fa-exclamation-circle"></i> Crítico</span>
                    <h5 class="card-title"><i class="fas fa-sitemap"></i> Sem Gerência</h5>
                    <p class="card-text h3"><%= totalSemGerencia %></p>
                    <div class="progress progress-thin bg-white">
                        <div class="progress-bar bg-danger" style="width: <%= (totalSemGerencia/(totalAtivos+totalInativos))*100 %>%"></div>
                    </div>
                </div>
                <div class="card-footer-stats text-dark">
                    <span class="small-text"><i class="fas fa-user-check"></i> Ativos: <%= totalAtivosSemGerencia %></span>
                </div>
            </div>
        </div>
        
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo bg-info-light text-dark">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-info"><i class="fas fa-chart-pie"></i> Distribuição</span>
                    <h5 class="card-title"><i class="fas fa-building"></i> Por Diretoria</h5>
                    <p class="card-text small mb-0"><%= Replace(totalPorDiretoria, "<br>", " | ") %></p>
                </div>
                <div class="card-footer-stats text-dark">
                    <span class="small-text"><i class="fas fa-list-ol"></i> <%= Len(Replace(totalPorDiretoria, "<strong>", "")) - Len(Replace(totalPorDiretoria, ":", "")) %> diretorias</span>
                </div>
            </div>
        </div>
        
        <div class="col-md-2 col-6 mb-3">
            <div class="card card-resumo bg-success-light text-dark">
                <div class="card-body position-relative">
                    <span class="badge stat-badge badge-success"><i class="fas fa-chart-bar"></i> Distribuição</span>
                    <h5 class="card-title"><i class="fas fa-sitemap"></i> Por Gerência</h5>
                    <p class="card-text small mb-0"><%= Replace(totalPorGerencia, "<br>", " | ") %></p>
                </div>
                <div class="card-footer-stats text-dark">
                    <span class="small-text"><i class="fas fa-list-ol"></i> <%= Len(Replace(totalPorGerencia, "<strong>", "")) - Len(Replace(totalPorGerencia, ":", "")) %> gerências</span>
                </div>
            </div>
        </div>
    </div>
    
    <!-- RESUMO RÁPIDO -->
    <div class="alert alert-info mb-3">
        <div class="row">
            <div class="col-md-3">
                <strong><i class="fas fa-users mr-2"></i>Total de Usuários:</strong> 
                <span class="badge badge-primary"><%= totalAtivos + totalInativos %></span>
            </div>
            <div class="col-md-3">
                <strong><i class="fas fa-percentage mr-2"></i>Ativos/Total:</strong> 
                <span class="badge badge-success"><%= totalAtivos %> (<%= FormatPercent(totalAtivos/(totalAtivos+totalInativos), 1) %>)</span>
            </div>
            <div class="col-md-3">
                <strong><i class="fas fa-exclamation-triangle mr-2"></i>Sem D/G:</strong> 
                <span class="badge badge-warning"><%= totalSemDiretoria %> / <%= totalSemGerencia %></span>
            </div>
            <div class="col-md-3">
                <strong><i class="fas fa-clock mr-2"></i>Atualizado:</strong> 
                <span class="badge badge-secondary"><%= Time() %></span>
            </div>
        </div>
    </div>

    <div class="d-flex justify-content-between align-items-center header-actions">
        <!-- Novo Usuário -->
        <a href="usrv_gestao_novo_usuario.asp" class="btn btn-primary">
            <i class="fas fa-user-plus mr-1"></i> Novo Usuário
        </a>

        <div class="d-flex flex-column"> 
            <h4 class="mb-0"><i class="fas fa-users mr-2"></i>Lista de Usuários - Gestão Completa</h4>
            <small class="mb-0 text-danger"><i class="fas fa-exclamation-triangle mr-1"></i>Atualizar um usuário por vez</small>
        </div>
        
        <div>
            <a href="#" class="btn btn-info" onclick="window.close(); return false;">
                <i class="fas fa-times mr-1"></i> Fechar
            </a>
        </div>
    </div>
    
    <div class="table-responsive">
      <table id="tabelaUsuarios" class="table table-striped table-bordered table-hover" style="width:100%">
        <thead class="thead-dark">
          <tr>
            <th>ID</th>
            <th>Usuário</th>
            <th>Status</th>
            <th>Função</th>
            <th>Diretoria</th>
            <th>Gerência</th>
            <th>Grupos</th>
            <th class="text-center">Ações</th>
          </tr>
        </thead>
        <tbody>
          <% 
          While Not rsUsers.EOF 
            userId = rsUsers("UserID")
            
            ' Obter grupos do usuário
            Set rsGrupos = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT g.ID_Grupo, g.Nome_Grupo FROM Grupo g " & _
                          "INNER JOIN Usuario_Grupo ug ON g.ID_Grupo = ug.ID_Grupo " & _
                          "WHERE ug.UserId = " & userId & " ORDER BY g.Nome_Grupo"
                        
            rsGrupos.Open sql, StrConn
            
            grupos = ""
            Do While Not rsGrupos.EOF
              grupos = grupos & "<span class='badge badge-info badge-grupo'>" & Server.HTMLEncode(rsGrupos("Nome_Grupo")) & "</span>"
              rsGrupos.MoveNext
            Loop
            
            If grupos = "" Then
              grupos = "<span class='text-muted'>Nenhum grupo</span>"
            End If
            
            If Not rsGrupos Is Nothing Then
              If rsGrupos.State = 1 Then rsGrupos.Close
              Set rsGrupos = Nothing
            End If
            
            ' Determinar status do usuário
            If CBool(rsUsers("Ativo")) Then
              statusClass = "badge-ativo"
              statusText = "ATIVO"
              btnClass = "btn-ativo"
              btnText = "ATIVO"
              btnIcon = "fas fa-toggle-on"
              novoStatus = "0"
            Else
              statusClass = "badge-inativo"
              statusText = "INATIVO"
              btnClass = "btn-inativo"
              btnText = "INATIVO"
              btnIcon = "fas fa-toggle-off"
              novoStatus = "-1"
            End If
            
            ' Obter dados atuais da diretoria e gerência
            Dim currentDiretoriaID, currentGerenciaID, showVazio
            currentDiretoriaID = SafeValue(rsUsers("DiretoriaID"), "")
            currentGerenciaID = SafeValue(rsUsers("GerenciaID"), "")
            
            ' Se este é o usuário que foi atualizado, usar os valores da URL
            If CStr(userId) = CStr(successUserId) And successDiretoria <> "" Then
                currentDiretoriaID = successDiretoria
                currentGerenciaID = successGerencia
            End If
            
            showVazio = False
            If CBool(rsUsers("Ativo")) And (currentDiretoriaID = "" Or currentGerenciaID = "") Then
                showVazio = True
            End If
          %>
          <tr class="<% If Not CBool(rsUsers("Ativo")) Then Response.Write "user-inativo" %>">
            <td><strong><%=userId%></strong></td>
            <td>
                <strong><%=UCase(rsUsers("Usuario"))%></strong><br>
                <small class="text-muted"><i class="fas fa-user mr-1"></i><%=SafeValue(rsUsers("Nome"), "N/A")%></small><br>
                <small class="text-muted"><i class="fas fa-envelope mr-1"></i><%=SafeValue(rsUsers("Email"), "N/A")%></small><br>
                <small class="text-muted"><i class="fas fa-phone mr-1"></i><%=SafeValue(rsUsers("Telefones"), "N/A")%></small><br>
                <small class="text-muted"><i class="fas fa-id-badge mr-1"></i>CRECI: <%=SafeValue(rsUsers("CRECI"), "N/A")%></small>
            </td>
            <td>
              <span class="badge badge-status <%=statusClass%>">
                <%=statusText%>
              </span>
            </td>
            <td>
              <% 
              Select Case SafeValue(rsUsers("Permissao"), 0)
                Case 1: badgeClass = "badge-danger"
                Case 2: badgeClass = "badge-warning"
                Case 3: badgeClass = "badge-warning"
                Case 4: badgeClass = "badge-info"
                Case 5: badgeClass = "badge-secondary"
                Case 6: badgeClass = "badge-secondary"
                Case Else: badgeClass = "badge-light"
              End Select
              %>
              <span class="badge <%=badgeClass%> badge-permissao"><%=UCase(SafeValue(rsUsers("Funcao"), "N/A"))%></span>
            </td>
            
            <td>
              <form method="post" class="diretoria-gerencia-form" onsubmit="return confirmUpdateDiretoriaGerencia(this, <%=userId%>);">
                <input type="hidden" name="acao" value="update_diretoria_gerencia">
                <input type="hidden" name="user_id" value="<%=userId%>">
                <div class="select-container">
                  <% If showVazio Then %>
                  <div class="vazio-indicator">Atualmente: <strong>Vazio</strong></div>
                  <% End If %>
                  <select name="diretoria_id" class="form-control form-control-sm select-small diretoria-select" 
                    data-userid="<%=userId%>" 
                    data-initial-value="<%=currentDiretoriaID%>"
                    data-initial-gerencia="<%=currentGerenciaID%>"
                    >
                    <option value="">Selecione uma diretoria...</option>
                    <%
                    rsDiretorias.MoveFirst
                    Do While Not rsDiretorias.EOF
                      selected = ""
                      If CStr(rsDiretorias("DiretoriaID")) = CStr(currentDiretoriaID) Then selected = "selected"
                      Response.Write "<option value=""" & rsDiretorias("DiretoriaID") & """ " & selected & ">" & Server.HTMLEncode(rsDiretorias("NomeDiretoria")) & "</option>"
                      rsDiretorias.MoveNext
                    Loop
                    %>
                  </select>
                </div>
            </td>
            
            <td>
                <div class="select-container">
                  <% If showVazio Then %>
                  <div class="vazio-indicator">Atualmente: <strong>Vazio</strong></div>
                  <% End If %>
                  <select name="gerencia_id" class="form-control form-control-sm select-small gerencia-select" 
                    id="gerencia_<%=userId%>"
                    data-initial-value="<%=currentGerenciaID%>"
                    >
                    <option value="">Selecione uma gerência...</option>
                    <%
                    ' Se temos uma diretoria selecionada, carregar as gerências correspondentes
                    If currentDiretoriaID <> "" Then
                        Set rsGerencias = Server.CreateObject("ADODB.Recordset")
                        rsGerencias.Open "SELECT GerenciaID, NomeGerencia FROM Gerencias WHERE DiretoriaID = " & currentDiretoriaID & " ORDER BY NomeGerencia", StrConn
                        
                        Do While Not rsGerencias.EOF
                            selected = ""
                            If CStr(rsGerencias("GerenciaID")) = CStr(currentGerenciaID) Then selected = "selected"
                            Response.Write "<option value=""" & rsGerencias("GerenciaID") & """ " & selected & ">" & Server.HTMLEncode(rsGerencias("NomeGerencia")) & "</option>"
                            rsGerencias.MoveNext
                        Loop
                        
                        If Not rsGerencias Is Nothing Then
                            If rsGerencias.State = 1 Then rsGerencias.Close
                            Set rsGerencias = Nothing
                        End If
                    End If
                    %>
                  </select>
                </div>
            </td>
            
            <td class="grupos-container"><%=grupos%></td>
            <td class="text-center">
              <div class="btn-group-vertical btn-group-sm" role="group">
                                
                <button type="submit" class="btn btn-warning btn-update-dg" title="Atualizar Diretoria e Gerência">
                  <i class="fas fa-sync-alt mr-1"></i>Atualizar
                </button>
              </form> <form method="post" class="toggle-form" onsubmit="return confirmToggle(this);">
                  <input type="hidden" name="acao" value="toggle_status">
                  <input type="hidden" name="user_id" value="<%=userId%>">
                  <input type="hidden" name="novo_status" value="<%=novoStatus%>">
                  <button type="submit" class="btn <%=btnClass%> btn-toggle" title="<% If CBool(rsUsers("Ativo")) Then %>Desativar Usuário<% Else %>Ativar Usuário<% End If %>">
                    <i class="<%=btnIcon%> mr-1"></i><%=btnText%>
                  </button>
                </form>
              </div>
            </td>
          </tr>
          <% 
            rsUsers.MoveNext()
          Wend 
          %>
        </tbody>
      </table>
    </div>
    
    <footer class="text-center text-muted small mb-3">
      <i class="fas fa-chart-line mr-1"></i> Sunny System &copy; <%= Year(Now()) %> | 
      <i class="fas fa-users mr-1"></i> Total: <strong><%= totalAtivos + totalInativos %></strong> | 
      <i class="fas fa-user-check mr-1"></i> Ativos: <strong><%= totalAtivos %></strong> | 
      <i class="fas fa-exclamation-triangle mr-1"></i> Sem D/G: <strong><%= totalSemDiretoria %> / <%= totalSemGerencia %></strong>
    </footer>
  </div>

    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    
    <script src="https://cdn.datatables.net/1.10.22/js/jquery.dataTables.min.js"></script>
  <script src="https://cdn.datatables.net/1.10.22/js/dataTables.bootstrap4.min.js"></script>
    
  <script>
// Dados pré-carregados das gerências
var allGerencias = <%= allGerenciasJSON %>;

// Parâmetros da URL
var urlParams = new URLSearchParams(window.location.search);
var successUserId = urlParams.get('userid');
var successDiretoria = urlParams.get('diretoria');
var successGerencia = urlParams.get('gerencia');

// Função para atualizar gerências usando dados pré-carregados
function updateGerencias(selectElement, userId, selectGerenciaId) {
    var diretoriaId = selectElement.value;
    var gerenciaSelect = document.getElementById('gerencia_' + userId);
    
    if (!gerenciaSelect) return;

    // Limpa as opções anteriores (exceto a primeira)
    gerenciaSelect.innerHTML = '<option value="">Selecione uma gerência...</option>';

    if (diretoriaId !== '' && allGerencias[diretoriaId]) {
        var options = '<option value="">Selecione uma gerência...</option>';
        allGerencias[diretoriaId].forEach(function(gerencia) {
            options += '<option value="' + gerencia.id + '">' + gerencia.nome + '</option>';
        });
        gerenciaSelect.innerHTML = options;
        
        // Selecionar a gerência correta
        var gerenciaToSelect = selectGerenciaId || gerenciaSelect.getAttribute('data-initial-value');
        if (gerenciaToSelect && gerenciaToSelect !== '') {
            setTimeout(function() {
                gerenciaSelect.value = gerenciaToSelect;
                toggleUpdateButtonByForm(selectElement.closest('form'));
            }, 100);
        }
    } 
    
    // Atualiza o estado do botão após mudar as gerências
    toggleUpdateButtonByForm(selectElement.closest('form'));
}

// FUNÇÃO toggleUpdateButton - SIMPLIFICADA
function toggleUpdateButtonByForm(form) {
    if (!form) return; 
    
    var diretoriaSelect = form.querySelector('select[name="diretoria_id"]');
    var gerenciaSelect = form.querySelector('select[name="gerencia_id"]');
    var btnUpdate = form.querySelector('.btn-update-dg');
    
    if (!diretoriaSelect || !gerenciaSelect || !btnUpdate) {
        return; 
    }
    
    var diretoriaValue = diretoriaSelect.value;
    var gerenciaValue = gerenciaSelect.value;
    
    // Habilita o botão APENAS se ambos os selects tiverem um valor selecionado
    btnUpdate.disabled = !(diretoriaValue !== '' && gerenciaValue !== '');
}

// Função para carregar gerências iniciais para usuários com diretoria
function loadInitialGerencias() {
    $('.diretoria-select').each(function() {
        var userId = $(this).data('userid');
        var diretoriaId = $(this).val();
        var initialGerencia = $(this).data('initial-gerencia') || $(this).attr('data-initial-gerencia');
        
        // Se houver uma diretoria selecionada, carrega as gerências correspondentes
        if (diretoriaId && diretoriaId !== '') {
            updateGerencias(this, userId, initialGerencia);
        }
    });
}

// Função para restaurar seleções após atualização
function restoreSelections() {
    if (successUserId && successDiretoria) {
        var diretoriaSelect = document.querySelector('form input[name="user_id"][value="' + successUserId + '"]')
            .closest('form')
            .querySelector('select[name="diretoria_id"]');
        
        if (diretoriaSelect) {
            // Define a diretoria
            diretoriaSelect.value = successDiretoria;
            
            // Carrega e seleciona a gerência
            updateGerencias(diretoriaSelect, successUserId, successGerencia);
        }
    }
}

$(document).ready(function() {
    // Carregar gerências iniciais para todos os usuários que já têm diretoria
    loadInitialGerencias();

    // Restaurar seleções se houver parâmetros na URL
    if (successUserId && successDiretoria) {
        setTimeout(restoreSelections, 500);
    }

    // Inicializar estado dos botões para todos os formulários
    $('.diretoria-gerencia-form').each(function() {
        toggleUpdateButtonByForm(this);
    });

    // Evento para select de diretoria
    $(document).on('change', '.diretoria-select', function() {
        var userId = $(this).data('userid');
        updateGerencias(this, userId);
    });

    // Evento para select de gerência
    $(document).on('change', '.gerencia-select', function() {
        var form = $(this).closest('form')[0];
        toggleUpdateButtonByForm(form);
    });

    $('#tabelaUsuarios').DataTable({
        "order": [[0, "desc"]],
        "pageLength": 100,
        "language": {
            "sEmptyTable": "Nenhum registro encontrado",
            "sInfo": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
            "sInfoEmpty": "Mostrando 0 até 0 de 0 registros",
            "sInfoFiltered": "(Filtrados de _MAX_ registros)",
            "sInfoPostFix": "",
            "sInfoThousands": ".",
            "sLengthMenu": "_MENU_ resultados por página",
            "sLoadingRecords": "Carregando...",
            "sProcessing": "Processando...",
            "sZeroRecords": "Nenhum registro encontrado",
            "sSearch": "Pesquisar:",
            "oPaginate": {
                "sNext": "Próximo",
                "sPrevious": "Anterior",
                "sFirst": "Primeiro",
                "sLast": "Último"
            },
            "oAria": {
                "sSortAscending": ": Ordenar colunas de forma ascendente",
                "sSortDescending": ": Ordenar colunas de forma descendente"
            },
            "select": {
                "rows": {
                    "_": "Selecionado %d linhas",
                    "0": "Nenhuma linha selecionada",
                    "1": "Selecionado 1 linha"
                }
            },
            "decimal": ",",
            "thousands": "."
        },
        "dom": '<"top"lif>rt<"bottom"lip><"clear">',
        "responsive": true,
        "initComplete": function() {
            $('.dataTables_filter input').addClass('form-control').attr('placeholder', 'Pesquisar...');
            $('.dataTables_length select').addClass('form-control');

            // Carregar gerências após o DataTables inicializar
            setTimeout(loadInitialGerencias, 100);

            // Restaurar seleções após o DataTables inicializar
            if (successUserId && successDiretoria) {
                setTimeout(restoreSelections, 1000);
            }
        },
        "columnDefs": [
            { "responsivePriority": 1, "targets": 1 }, 
            { "responsivePriority": 2, "targets": -1 },
            { "responsivePriority": 3, "targets": 6 },
            { "responsivePriority": 4, "targets": 3 },
            { "responsivePriority": 5, "targets": 4 },
            { "responsivePriority": 6, "targets": 5 },
            { "responsivePriority": 7, "targets": 2 },
            { "responsivePriority": 8, "targets": 0 }
        ]
    });

    // Auto-close alerts after 5 seconds
    setTimeout(function() {
        $('.alert').alert('close');
    }, 5000);
});

// Função para confirmar a alteração de status
function confirmToggle(form) {
    var userId = form.user_id.value;
    var novoStatus = form.novo_status.value;
    var acao = (novoStatus == "-1") ? "ativar" : "desativar";
    var nomeUsuario = form.closest('tr').querySelector('td:nth-child(2) strong').textContent;
    
    if (confirm("Tem certeza que deseja " + acao + " o usuário '" + nomeUsuario + "'?")) {
        var btn = form.querySelector('button');
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-1"></i>Processando...';
        btn.disabled = true;
        
        var allButtons = document.querySelectorAll('.btn-toggle, .btn-update-dg');
        allButtons.forEach(function(button) {
            button.disabled = true;
        });
        
        return true;
    }
    return false;
}

// Função para confirmar atualização de diretoria e gerência
function confirmUpdateDiretoriaGerencia(form, userId) {
    // Encontrar os elementos de forma mais robusta
    var diretoriaSelect = form.querySelector('select[name="diretoria_id"]');
    var gerenciaSelect = document.getElementById('gerencia_' + userId);
    
    if (!diretoriaSelect || !gerenciaSelect) {
        alert('Erro: Elementos do formulário não encontrados.');
        return false;
    }

    var diretoriaNome = diretoriaSelect.options[diretoriaSelect.selectedIndex].text;
    var gerenciaNome = gerenciaSelect.options[gerenciaSelect.selectedIndex].text;
    var nomeUsuario = form.closest('tr').querySelector('td:nth-child(2) strong').textContent;
    
    if (diretoriaSelect.value === '' || gerenciaSelect.value === '') {
        alert('Por favor, selecione tanto a Diretoria quanto a Gerência.');
        return false;
    }
    
    if (confirm('Tem certeza que deseja atualizar a Diretoria e Gerência do usuário "' + nomeUsuario + '"?\n\nDiretoria: ' + diretoriaNome + '\nGerência: ' + gerenciaNome)) {
        var btn = form.querySelector('.btn-update-dg');
        if (btn) {
            btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-1"></i>Processando...';
            btn.disabled = true;
        }
        
        var allButtons = document.querySelectorAll('.btn-toggle, .btn-update-dg');
        allButtons.forEach(function(button) {
            button.disabled = true;
        });
        
        return true;
    }
    return false;
}
  </script>
</body>
</html>

<%
' ==============================================================================
' FECHAMENTO DE RECORDSETS
' ==============================================================================
If Not rsDiretorias Is Nothing Then
    If rsDiretorias.State = 1 Then rsDiretorias.Close
    Set rsDiretorias = Nothing
End If
If Not rsUsers Is Nothing Then
    If rsUsers.State = 1 Then rsUsers.Close
    Set rsUsers = Nothing
End If
%>