<header style="
    /* Fundo Bordô FORÇADO */
    background-color: #800000 !important;
    
    /* Cor de texto padrão (garantia) */
    color: #ffffff !important;
    
    /* Altura Reduzida (5px topo/base) */
    padding: 5px 20px; 
    
    /* Margem Inferior para separar do conteúdo abaixo */
    margin-bottom: 8px !important; 
    
    font-size: 20px;
    font-weight: bold;
    text-align: left;
    box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    
    /* Ações que COLAM NO TOPO: */
    position: fixed;
    top: 0; 
    /* ------------------------ */
    
    width: 100%;
    z-index: 1000;
">
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand text-white" style="color: #ffffff !important;" href="#">
                <i class="fas fa-sun me-2"></i>SGVendas - <%=Session("Usuario")%>
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link active text-white" 
                           style="color: #ffffff !important;" 
                           href="#" 
                           onclick="window.close();">
                            <i class="fas fa-times me-1"></i> Fechar
                        </a>
                    </li>

                    <li class="nav-item">
                        <a class="nav-link text-white" style="color: #ffffff !important;" href="gestao_logout.asp"><i class="fas fa-sign-out-alt me-1"></i> Sair</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>
</header>