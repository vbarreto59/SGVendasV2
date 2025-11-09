<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->

<% Response.Charset = "UTF-8" %>
<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn
    sql = "INSERT INTO Gerencias (DiretoriaID, NomeGerencia, NomeGerente, TelefoneGerente) VALUES (?, ?, ?, ?)"
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1
        .Parameters.Append .CreateParameter(, 3, 1, , Request.Form("DiretoriaID"))
        .Parameters.Append .CreateParameter(, 200, 1, 100, Request.Form("NomeGerencia"))
        .Parameters.Append .CreateParameter(, 200, 1, 100, Request.Form("NomeGerente"))
        .Parameters.Append .CreateParameter(, 200, 1, 20, Request.Form("TelefoneGerente"))
        .Execute
    End With
    conn.Close
    Set conn = Nothing

    Response.Redirect "gerencia_list.asp"
End If
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cadastrar Gerência</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
</head>
<body>
    <!-- Navbar from gestao_painel2.asp -->
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand" href="#">
                <i class="fas fa-sun me-2"></i>SunnyImob.
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link" href="gestao_painel2.asp"><i class="fas fa-home me-1"></i> Início</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="#"><i class="fas fa-cog me-1"></i> Configurações</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="gestao_logout.asp"><i class="fas fa-sign-out-alt me-1"></i> Sair</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>
    
    <!-- Welcome section from gestao_painel2.asp -->
    <section class="welcome-section text-center">
        <div class="container">
            <h1 class="display-4 mb-2">SGVendas</h1>
            <p class="lead">Gerencie as operações de gestão e vendas</p>
        </div>
    </section>

    <div class="container py-5">
        <h2 class="text-center mb-4">Nova Gerência</h2>
        
        <div class="card p-4 mx-auto" style="max-width: 600px;">
            <form method="post" action="gerencia_create.asp">
                <div class="mb-3">
                    <label for="DiretoriaID" class="form-label">Diretoria</label>
                    <select name="DiretoriaID" id="DiretoriaID" required class="form-control">
                        <option value="">Selecione uma diretoria</option>
                        <% 
                        Set conn = Server.CreateObject("ADODB.Connection")
                        conn.Open strConn
                        Set rs = conn.Execute("SELECT * FROM Diretorias ORDER BY NomeDiretoria")
                        Do While Not rs.EOF
                        %>
                        <option value="<%=rs("DiretoriaID")%>"><%=rs("NomeDiretoria")%></option>
                        <% rs.MoveNext: Loop: rs.Close: conn.Close %>
                    </select>
                </div>

                <div class="mb-3">
                    <label for="NomeGerencia" class="form-label">Nome da Gerência</label>
                    <input type="text" name="NomeGerencia" id="NomeGerencia" required class="form-control">
                </div>

                <div class="mb-3">
                    <label for="NomeGerente" class="form-label">Nome do Gerente</label>
                    <input type="text" name="NomeGerente" id="NomeGerente" required class="form-control">
                </div>

                <div class="mb-3">
                    <label for="TelefoneGerente" class="form-label">Telefone do Gerente</label>
                    <input type="text" name="TelefoneGerente" id="TelefoneGerente" class="form-control">
                </div>

                <div class="d-grid gap-2 d-md-flex justify-content-md-end mt-4">
                    <button type="submit" class="btn btn-primary me-md-2">
                        <i class="fas fa-save me-1"></i> Cadastrar
                    </button>
                    <a href="javascript:window.close()" class="btn btn-secondary">
                        <i class="fas fa-arrow-left me-1"></i> Voltar
                    </a>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Footer from gestao_painel2.asp -->
    <footer class="text-center mt-auto">
        <div class="container">
            <div class="row">
                <div class="col-md-6">
                    <h5><i class="fas fa-sun me-2"></i>SunnyImob</h5>
                    <p>Valter Barreto</p>
                </div>
                <div class="col-md-6">
                    <p>&copy; 2023 Todos os direitos reservados</p>
                    <div class="social-icons">
                        <a href="#" class="me-2"><i class="fab fa-facebook-f"></i></a>
                        <a href="#" class="me-2"><i class="fab fa-twitter"></i></a>
                        <a href="#" class="me-2"><i class="fab fa-linkedin-in"></i></a>
                        <a href="#"><i class="fab fa-instagram"></i></a>
                    </div>
                </div>
            </div>
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>