<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: WVLHTQCGWG          -->
<!-- OBS: Alterado em 12 12 2025 Organização dos cards     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  
<!--#include file="usr_acoes_v4GVendas.inc"-->
<!--#include file="atualizarVendas.asp"-->
<!--#include file="atualizarVendas2.asp"-->
<!--#include file="atualizarVendasTemp.asp"-->
<!--#include file="manutencao_config.asp"-->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if   
Dim mostrarAlertaManutencao
mostrarAlertaManutencao = EstaEmManutencao()
%>

<!-- badge vgv últimos dois anos  -->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if   
'Dim mostrarAlertaManutencao
mostrarAlertaManutencao = EstaEmManutencao()
%>

<!-- =============================================== -->
<!-- PAINEL BOLSA DE VALORES - VGV (NOVO CÓDIGO)    -->
<!-- =============================================== -->
<% Server.Execute("vgv_bolsa_valores.asp") %>
<!-- =============================================== -->
<!--  -->



<%
' =========================================================================
' === FUNÇÃO PARA DETECÇÃO DE DISPOSITIVO MÓVEL (NOVO CÓDIGO) =============
' =========================================================================
Function IsMobile()
    Dim userAgent
    userAgent = Request.ServerVariables("HTTP_USER_AGENT")
    If IsNull(userAgent) Then userAgent = ""

    ' Converte para minúsculas para facilitar a comparação
    userAgent = LCase(userAgent)

    ' Lista de palavras-chave comuns de dispositivos móveis
    ' Você pode adicionar mais palavras-chave conforme necessário.
    Dim mobileKeywords
    mobileKeywords = Array("mobile", "android", "iphone", "ipod", "blackberry", "windows phone", "iemobile", "opera mini", "symbian", "webos")

    Dim keyword
    IsMobile = False ' Assume não ser móvel por padrão

    ' Percorre a lista de palavras-chave
    For Each keyword In mobileKeywords
        If InStr(userAgent, keyword) > 0 Then
            IsMobile = True ' Palavra-chave encontrada, é móvel
            Exit For
        End If
    Next
End Function

Dim vendasFile

' Define o arquivo de vendas com base no resultado da função IsMobile()
If IsMobile() Then
    ' O arquivo para visualização em celular

    vendasFile = "gestao_vendas_list_mob1.asp"

    if Not BloqueioEmail() AND (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
        On Error Resume Next 
        set objMail = server.createobject("CDONTS.NewMail")
        if Err.Number <> 0 then 
            set objMail = Nothing ' Garante que a variável seja liberada, mesmo que não criada
        else
            objMail.From = "sendmail@gabnetweb.com.br"
            objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
            objMail.Subject = "SV-MOB" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
            objMail.MailFormat = 0 ' 0 = Texto Simples
            objMail.Body = "Página Vendas Mobile. " & Ucase(Session("Usuario"))
            objMail.Send
            set objMail = Nothing
        end if 
        On Error GoTo 0 
    end if

    
Else
    ' O arquivo padrão para visualização em desktop
    vendasFile = "gestao_vendas_list3x.asp"
End If



'============================= ATUALIZANDO O BANCO DE DADOS ==================='
Response.Buffer = True
Response.Expires = -1
'On Error Resume Next ' 

' --- CRIAÇÃO DOS OBJETOS ADO DE CONEXÃO ---
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Primeiro UPDATE: Associar Vendas.DiretoriaId com Diretorias.DiretoriaId e atualizar campos
'sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"


'===== Modificado em 25 11 2025'
sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId] WHERE Vendas.NomeDiretor IS NULL;"
'connSales.Execute(sqlUpdate1)

' UPDATE Gerencias -> Vendas
sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId] WHERE Vendas.NomeGerente IS NULL;"
'connSales.Execute(sqlUpdate2)




'Atualizar Nome do Corretor-----------------------------'
sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) " & _
                   "SET Vendas.Corretor = Usuarios.Nome;"
'connSales.Execute(sqlUpdateCorretor)

'=========== Esta é a instrução SQL para atualizar o campo Semestre. ======='
sql = "UPDATE Vendas " & _
      "SET Semestre = SWITCH(" & _
      "    Trimestre IN (1, 2), 1, " & _
      "    Trimestre IN (3, 4), 2" & _
      ") " & _
      "WHERE Trimestre IS NOT NULL;"
On Error Resume Next
connSales.Execute sql

' Verificação de erros.
If Err.Number <> 0 Then
    Response.Write "Ocorreu um erro ao atualizar o campo Semestre: " & Err.Description
Else
   ' Response.Write "O campo Semestre foi atualizado com sucesso para todos os registros."
End If
On Error GoTo 0
' ======================= FINAL ATUALIZAÇÃO DO BANCO DE DADOS ========================'
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="300">
    <title>Menu Administrativo</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
    <style>
        /* Estilo para todos os cards */
        .card {
            border-radius: 15px !important;
            overflow: hidden;
            transition: all 0.3s ease;
            border: none;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            height: 100%;
        }
        
        /* Efeito hover para todos os cards */
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
        }
        
        /* Estilo para os cards do BARRETO (azul) */
        .card-barreto .card-header {
            background: linear-gradient(45deg, #1e3c72, #2a5298, #1c3f95) !important;
            border-radius: 15px 15px 0 0 !important;
            padding: 1rem;
        }
        
        .card-barreto .card-header h5 {
            color: white !important;
            font-weight: 600;
            margin: 0;
        }
        
        .card-barreto .btn-primary {
            background-color: #1e3c72 !important;
            border-color: #1e3c72 !important;
            border-radius: 8px;
            font-weight: 500;
            transition: all 0.3s ease;
        }
        
        .card-barreto .btn-primary:hover {
            background-color: #2a5298 !important;
            border-color: #2a5298 !important;
            transform: scale(1.05);
        }
        
        /* Estilo padrão para outros cards (vermelho) */
        .card-header {
            background: linear-gradient(45deg, #800020, #A52A2A, #4B0012);
            border-radius: 15px 15px 0 0 !important;
            padding: 1rem;
        }
        
        .card-header h5 {
            color: white;
            font-weight: 600;
            margin: 0;
        }
        
        .btn-primary {
            background-color: #800020 !important;
            border-color: #800020 !important;
            border-radius: 8px;
            font-weight: 500;
            transition: all 0.3s ease;
        }
        
        .btn-primary:hover {
            background-color: #A52A2A !important;
            border-color: #A52A2A !important;
            transform: scale(1.05);
        }
        
        .welcome-section, footer .col-md-6:first-child {
            background: linear-gradient(45deg, #800020, #A52A2A, #4B0012);
        }
        
        .welcome-section h1, .card-header h5, footer .col-md-6:first-child h5 {
            color: white;
        }
        
        /* Linha divisória entre os blocos */
        .divider-line {
            border-top: 3px solid #800020;
            margin: 2rem 0;
            opacity: 0.6;
        }
        
        /* Estilo para o corpo dos cards */
        .card-body {
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }
        
        .card-text {
            margin-bottom: 1.5rem;
            color: #555;
            line-height: 1.5;
        }
        
        /* Alerta de manutenção */
        .alert-manutencao {
            z-index: 1000;
            border-radius: 0;
            text-align: center;
            padding: 15px;
            font-weight: bold;
            animation: pulse 2s infinite;
            margin-bottom: 0 !important;
        }
        
        /* Animação pulse para alerta */
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.7; }
            100% { opacity: 1; }
        }
        
        /* Estilo para o cabeçalho da seção BARRETO */
        .barreto-section-header {
            background: linear-gradient(45deg, #1e3c72, #2a5298, #1c3f95);
            color: white;
            padding: 1.5rem;
            border-radius: 15px;
            margin: 2rem 0 1.5rem 0;
            text-align: center;
        }
        
        .barreto-section-header h3 {
            margin: 0;
            font-weight: 600;
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
<% If mostrarAlertaManutencao Then %>
<div class="alert alert-danger alert-manutencao">
    <i class="fas fa-exclamation-triangle"></i> ATENÇÃO: SISTEMA EM MANUTENÇÃO - Algumas funcionalidades podem estar indisponíveis
</div>
<% End If %>    

<%
if not UsuarioGestor() and not UsuarioAdmin() then
     Response.Write("<h3>Função habilitada apenas para Gestores do Sistema.</h3>")
     Response.End
End if
%>
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand" href="#">
                <i class="fas fa-sun me-2"></i>SGVendas - <%=Session("Usuario") & " "%>  <%=Session("EnviaEmail")%>
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link active" href="gestao_painel2.asp"><i class="fas fa-home me-1"></i> Início</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="gestao_logout.asp"><i class="fas fa-sign-out-alt me-1"></i> Sair</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>

    <section class="welcome-section text-center">
        <div class="container">
            <h1 class="display-4 mb-2">SGVendas</h1>
            <p class="lead">Gerencie as vendas e comissões</p>
        </div>
    </section>

    <div class="container mb-5">
        <!-- PRIMEIRA LINHA: VENDAS E SALDOS E COMISSÕES (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-funnel-dollar me-2"></i>A-Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Gerenciamento de Vendas</p>
                        <a href="<%= vendasFile %>" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <a href="gestao_vendas_kpi5.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-line me-2"></i>B-Vendas KPIs</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização do Valor Geral de Vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-funnel-dollar me-2"></i>C-Dashboard Metas x Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Acompanhamento das Metas</p>
                        <a href="gestao_vendas_metas.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- SEGUNDA LINHA (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>D-Dashboard Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as vendas.</p>
                        <a href="dashboard3rand7x.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>E-Comparativo de Metas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as vendas.</p>
                        <a href="dashb_comp_metas3.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <a href="gestao_geomapa_vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-map-marked-alt me-2"></i>F-Geo-Mapa de Vendas</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das regiões com vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Mapa de Vendas
                            </span>
                        </div>
                    </div>
                </a>
            </div>
        </div>

        <!-- TERCEIRA LINHA (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>G-Relat. Geral</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="gestao_vendas_geral.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>H-Comissões Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize os saldos das comissões.</p>
                        <a href="venda_pag_resumo1.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>I-Comissões Mensais</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="gestao_corretores_comissoes.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- QUARTA LINHA (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>J-KPI Comissões</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="gestao_vendas_kpi5comissao.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-sitemap me-2"></i>K-Diretorias</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro e gerenciamento das diretorias da empresa.</p>
                        <a href="diretoria_list.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>L-Gerências</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro e acompanhamento dos gerentes de departamento.</p>
                        <a href="gerencia_list.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- QUINTA LINHA (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>M-Usuários</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro de usuários.</p>
                        <a href="usrv_gestao_listar2.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>N-Metas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro de Metas da Tocca.</p>
                        <a href="gestao_metasEmpresa.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>



        <!-- SEÇÃO BARRETO - APENAS PARA USUÁRIO BARRETO -->
        <% If Session("Usuario") = "BARRETO" then %>
           <div class="divider-line"></div>     


        <!-- LINHA 1 BARRETO (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S1-Ficha Corretor</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Informações dos Corretores.</p>
                        <a href="ficha_corretor5.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S2-$ Individual</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="gestao_ganhos_individuais1.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>  

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S2-$ Individual Mensal</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Média mensal de ganhos.</p>
                        <a href="gestao_ganhos_individ_mensal3.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div> 


            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S3-QTD Semanas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="vendas_semana.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S4-QTD Mensais</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as comissões.</p>
                        <a href="vendas_valores_mensais.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- LINHA 2 BARRETO (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-file-alt me-2"></i>S5-Relatórios</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Relatórios gerenciais e consolidados.</p>
                        <a href="menu_relatorios.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S6-Bloquear Emails</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Bloqueia o Envio de Emails.</p>
                        <a href="bloqueiaEmail.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>**S7-Corretores Sem Venda</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Corretores sem Vendas.</p>
                        <a href="gestao_corretores_sem_vendas2.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- LINHA 3 BARRETO (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>S8-Visualizar Log</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualizar Log.</p>
                        <a href="tool_visualizar_log.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>**S9-Saldo das Comissões</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualizar Saldo das Comissões</p>
                        <a href="gestao_vendas_comissao_saldo3.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>**S10-2 Vendas Anuais</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualizar Saldo das Comissões</p>
                        <a href="gestao_vendas_metas_2anos.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- LINHA 4 BARRETO (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>**S11-Metas Por Gerências</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Define as Metas Por Gerência</p>
                        <a href="meta_gerenciamento.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>S12-Manutenção</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro de Metas da Tocca.</p>
                        <a href="manut_menu.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>**S13-Vendas para Diretorias</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Permitir Diretor ver Relatório de Vendas.</p>
                        <a href="manut_vendas_diretorias.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <!-- LINHA 5 BARRETO (3 cards por linha) -->
        <div class="row g-4 mb-4">
            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>S14-Backup para Email</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Backup para Email.</p>
                        <a href="backupToccaMDBEmail.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>S15-Gerar JSON</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Backup para Email.</p>
                        <a href="tool_venda_criar_json.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card card-barreto">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>**S16-Painel Tipo Bolsa</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Painel Tipo Bolsa.</p>
                        <a href="painel_bolsa_valores3.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>


        </div>
        <% End if %>
        
        
    </div>

    <!--#include file="footer.inc"-->    

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>