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
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if   
%>

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
sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute(sqlUpdate1)

' UPDATE Gerencias -> Vendas
sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId];"
connSales.Execute(sqlUpdate2)

'Atualizar Nome do Corretor-----------------------------'
sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) " & _
                   "SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute(sqlUpdateCorretor)

' Esta é a instrução SQL para atualizar o campo Semestre.
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
        /* Gradient background for title containers with white text */
        .welcome-section, .card-header, footer .col-md-6:first-child {
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
    </style>
</head>
<body>

<%
if not UsuarioGestor() and not UsuarioAdmin() then
     Response.Write("<h3>Função habilitada apenas para Gestores do Sistema.</h3>")
     Response.End
End if
%>
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand" href="#">
                <i class="fas fa-sun me-2"></i>SGVendas - <%=Session("Usuario")%>
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
        <div class="row g-4">
            <!-- PRIMEIRA LINHA: VENDAS E SALDOS E COMISSÕES -->
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-funnel-dollar me-2"></i>Vendas</h5>
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
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Saldos Comissões 1</h5>
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
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Saldos Comissões</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize os saldos das comissões.</p>
                        <a href="gestao_vendas_comissao_saldo3.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>


        </div>

        <!-- LINHA DIVISÓRIA -->
        <div class="divider-line"></div>

        <div class="row g-4">
            <!-- SEGUNDA LINHA: DEMAIS OPÇÕES -->
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Dashboard Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Visualize as vendas.</p>
                        <a href="dashboard3rand1.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-funnel-dollar me-2"></i>Dashboard Metas x Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Acompanhamento das Metas</p>
                        <a href="gestao_vendas_metas.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_geomapa_vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-map-marked-alt me-2"></i>Geo-Mapa de Vendas</h5>
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

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-file-alt me-2"></i>Relatórios</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Relatórios gerenciais e consolidados.</p>
                        <a href="menu_relatorios.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
                    <!-- LINHA DIVISÓRIA -->
        <div class="divider-line"></div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-sitemap me-2"></i>Diretorias</h5>
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
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Gerências</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro e acompanhamento dos gerentes de departamento.</p>
                        <a href="gerencia_list.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Usuários</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro de usuários.</p>
                        <a href="usrv_gestao_listar.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Metas</h5>
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
    </div>

    <footer class="text-center mt-auto">
        <div class="container">
            <div class="row">
                <div class="col-md-12">
                    <p><small>Valter Barreto</p>
                    <p>&copy; 2025 Todos os direitos reservados</p></small>
                    <div class="social-icons">
                    </div>
                </div>
            </div>
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>