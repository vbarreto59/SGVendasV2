<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"  
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Menu de Relatórios</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
    <style>
        /* CORREÇÃO CRÍTICA: AFUNDA O CONTEÚDO ABAIXO DA NAVBAR FIXA */
        body {
            /* 70px é uma estimativa; ajuste este valor (ex: 80px) se a barra for mais alta. */
            padding-top: 70px; 
        }
        
        /* ESTILOS DA NAVBAR: Fundo Bordô e Texto Branco */
        nav.navbar {
            background-color: #800000 !important; /* Bordô */
        }
        
        .navbar-brand, .nav-link, nav.navbar .fa-sun, nav.navbar .fa-times, nav.navbar .fa-sign-out-alt {
            color: #ffffff !important; /* Branco */
        }
        
        .nav-link:hover {
            color: #cccccc !important; /* Branco mais claro no hover */
        }
        
        /* ESTILOS DOS CARDS E RESTANTE DO CONTEÚDO */
        .group-header {
            background: rgba(255, 255, 255, 0.95);
            color: #8B0000;
            padding: 15px;
            border-radius: 10px;
            margin: 30px 0 20px 0;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.3);
            border-left: 5px solid #8B0000;
        }
        .group-header h3 {
            margin: 0;
            font-size: 1.5rem;
            font-weight: 700;
        }
        .group-divider {
            border-bottom: 2px solid #8B0000;
            margin: 25px 0;
            opacity: 0.3;
        }
        .card {
            transition: transform 0.3s ease;
            border: none;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            background: rgba(255, 255, 255, 0.95);
            border-radius: 15px;
            border: 1px solid rgba(255, 255, 255, 0.3);
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(139, 0, 0, 0.2);
        }
        
        /* CORREÇÃO PARA OS TÍTULOS DOS CARDS */
        .card-header {
            background: #8B0000; /* Cor bordô principal, igual ao botão */
            border-bottom: 2px solid #8B0000; /* Borda na cor principal */
            font-weight: 600;
            border-radius: 15px 15px 0 0 !important;
            padding: 15px;
            color: white !important; /* Texto branco para contraste */
        }
        
        .btn-primary {
            background: #8B0000;
            border: none;
            border-radius: 8px;
            padding: 10px;
            font-weight: 600;
            font-size: 12px;
            color: white;
        }
        .btn-primary:hover {
            background: #660000;
        }
        .card-text {
            color: #7f8c8d;
            font-size: 14px;
        }
        
        .card-header h5 {
            color: white !important; 
            font-weight: 700;
        }
        
        .display-4, h2, footer h5 {
            color: #8B0000;
        }
        
        /* Estilo para a seção de boas-vindas */
        .welcome-section {
            padding-top: 20px;
            padding-bottom: 20px;
        }

    </style>
</head>
<body>
    

    <div class="container py-5">
        
        <h2 class="text-center mb-4">Menu de Relatórios</h2>
        
        <div class="row g-4">
            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_kpi3.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-line me-2"></i>A - Relat. Vendas KPIs</h5>
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

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_top10.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-trophy me-2"></i>B - Relat. TOP 10</h5>
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

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_kpi5comissao.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-money-bill-wave me-2"></i>C - Relat. Comissões Gerais</h5>
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

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_geral.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>D - Relat. Geral</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Detalhes completos sobre as vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>      

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_mapa_vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-map-marked-alt me-2"></i>E - Corretor - Mapa de Vendas (QTD)</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização da quantidade de unidades vendidas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar
                            </span>
                        </div>
                    </div>
                </a>
            </div>  

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_localidade1.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-map-marker-alt me-2"></i>F - Venda por Localidade</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das vendas por localidades.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar
                            </span>
                        </div>
                    </div>
                </a>
            </div>
        </div>

        <div class="group-divider"></div>

        <div class="row g-4">
            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_corretores.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>G - Vendas Individual Corretor</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das vendas do corretor.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Extrato
                            </span>
                        </div>
                    </div>
                </a>
            </div>  

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_comissoes.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-money-check me-2"></i>H - Comissão Ano/Mês</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>      

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_extrato_comissoes.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-file-invoice-dollar me-2"></i>I - Comissão Tabela Anual</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>              

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_comissoes_anual.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-table me-2"></i>J - Vendas Tabela Anual</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>
        </div>

        <div class="group-divider"></div>

        <div class="row g-4">
            <div class="col-12 col-md-6 col-lg-4">
                <a href="diretoria_list.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>L - Listagem de Diretores</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Lista e detalhes dos diretores da empresa.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gerencia_list.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-users me-2"></i>M - Listagem de Gerentes</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Lista e detalhes dos gerentes de departamento.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_geomapa_vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-globe-americas me-2"></i>Geo - Mapa de Vendas</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das regiões com vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Mapa
                            </span>
                        </div>
                    </div>
                </a>
            </div>  
        </div>
    </div>
    
    <footer class="text-center mt-auto py-3">
        <div class="container">
            <div class="row">
                <div class="col-md-6">
                    <h5><i class="fas fa-sun me-2"></i>SGVendas</h5>
                    <p>Valter Barreto</p>
                </div>
                <div class="col-md-6">
                    <p>&copy; 2025 Todos os direitos reservados</p>
                    <div class="social-icons">
                        <a href="#" class="me-2 text-decoration-none"><i class="fab fa-facebook-f text-dark"></i></a>
                        <a href="#" class="me-2 text-decoration-none"><i class="fab fa-twitter text-dark"></i></a>
                        <a href="#" class="me-2 text-decoration-none"><i class="fab fa-linkedin-in text-dark"></i></a>
                        <a href="#" class="text-decoration-none"><i class="fab fa-instagram text-dark"></i></a>
                    </div>
                </div>
            </div>
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>