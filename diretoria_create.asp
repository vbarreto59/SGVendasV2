<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim NomeDiretoria, NomeDiretor, TelefoneDiretor, EstadoDiretoria, UserId
    Dim Usuario, NomeUsuario, rsUsuario, sqlUsuario
    
    ' Pega os dados do formulário
    NomeDiretoria = Request.Form("NomeDiretoria")
    NomeDiretor = Request.Form("NomeDiretor")
    TelefoneDiretor = Request.Form("TelefoneDiretor")
    EstadoDiretoria = Request.Form("EstadoDiretoria")
    UserId = Request.Form("UserId") ' Pega o UserId do formulário
    
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn

    ' Prepara as variáveis para a inserção (evita erros com valores nulos)
    Dim sUserId, sUsuario, sNomeUsuario
    sUserId = "NULL"
    sUsuario = "NULL"
    sNomeUsuario = "NULL"

    ' Se um UserId foi selecionado, busca os dados do usuário na tabela Usuarios
    If UserId <> "" Then
        sqlUsuario = "SELECT Usuario, Nome FROM Usuarios WHERE UserId=" & UserId
        Set rsUsuario = conn.Execute(sqlUsuario)

        If Not rsUsuario.EOF Then
            Usuario = rsUsuario("Usuario")
            NomeUsuario = rsUsuario("Nome")
            sUserId = UserId
            sUsuario = "'" & Replace(Usuario, "'", "''") & "'"
            sNomeUsuario = "'" & Replace(NomeUsuario, "'", "''") & "'"
        End If
        rsUsuario.Close
        Set rsUsuario = Nothing
    End If

    ' Constroi o comando SQL de INSERT de forma mais limpa
    Dim sql
    sql = "INSERT INTO Diretorias (NomeDiretoria, NomeDiretor, TelefoneDiretor, EstadoDiretoria, UserId, Usuario, Nome) VALUES (" & _
          "'" & Replace(NomeDiretoria, "'", "''") & "', " & _
          "'" & Replace(NomeDiretor, "'", "''") & "', " & _
          "'" & Replace(TelefoneDiretor, "'", "''") & "', " & _
          "'" & Replace(EstadoDiretoria, "'", "''") & "', " & _
          sUserId & ", " & _
          sUsuario & ", " & _
          sNomeUsuario & ")"

    ' Use esta linha para depurar o SQL antes de executá-lo
    ' Response.Write sql & "<hr>"
    ' Response.End

    conn.Execute sql
    conn.Close
    Set conn = Nothing

    Response.Redirect "diretoria_list.asp"
End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cadastrar Diretoria</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
</head>
<body>
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
    
    <section class="welcome-section text-center">
        <div class="container">
            <h1 class="display-4 mb-2">SGVendas</h1>
            <p class="lead">Gerencie as operações de gestão e vendas</p>
        </div>
    </section>

    <div class="container">
        <h2 class="text-center mb-4">Cadastrar Nova Diretoria</h2>
        
        <div class="card p-4">
            <form method="post" action="diretoria_create.asp">
                <div class="mb-3">
                    <label for="NomeDiretoria" class="form-label">Nome da Diretoria</label>
                    <input type="text" name="NomeDiretoria" id="NomeDiretoria" required class="form-control">
                </div>
                <div class="mb-3">
                    <label for="NomeDiretor" class="form-label">Nome do Diretor</label>
                    <input type="text" name="NomeDiretor" id="NomeDiretor" required class="form-control">
                </div>
                <div class="mb-3">
                    <label for="TelefoneDiretor" class="form-label">Telefone do Diretor</label>
                    <input type="text" name="TelefoneDiretor" id="TelefoneDiretor" required class="form-control">
                </div>
                <div class="mb-3">
                    <label for="EstadoDiretoria" class="form-label">Estado da Diretoria</label>
                    <select name="EstadoDiretoria" id="EstadoDiretoria" required class="form-control">
                        <option value="">Selecione...</option>
                        <option value="Ativo">Ativo</option>
                        <option value="Inativo">Inativo</option>
                    </select>
                </div>
                
                <div class="mb-3">
                    <label for="searchUserInput" class="form-label">Buscar Usuário</label>
                    <input type="text" id="searchUserInput" class="form-control" placeholder="Digite para buscar...">
                </div>

                <div class="mb-3">
                    <label for="UserId" class="form-label">Usuário</label>
                    <%
                   '' Dim conn, rsUsuarios, sqlUsuarios
                    Set conn = Server.CreateObject("ADODB.Connection")
                    conn.Open strConn
                    sqlUsuarios = "SELECT UserId, Usuario, Nome FROM Usuarios ORDER BY Nome"
                    Set rsUsuarios = conn.Execute(sqlUsuarios)
                    %>
                    <select name="UserId" id="UserId" class="form-control">
                        <option value="">Selecione um Usuário</option>
                        <%
                        Do While Not rsUsuarios.EOF
                        %>
                        <option value="<%=rsUsuarios("UserId")%>" data-search-text="<%=LCase(rsUsuarios("Nome") & " " & rsUsuarios("Usuario"))%>"><%=rsUsuarios("Nome") & " (" & rsUsuarios("Usuario") & ")"%></option>
                        <%
                            rsUsuarios.MoveNext
                        Loop
                        rsUsuarios.Close
                        Set rsUsuarios = Nothing
                        conn.Close
                        Set conn = Nothing
                        %>
                    </select>
                </div>

                <div class="d-grid gap-2 d-md-flex justify-content-md-end mt-4">
                    <button type="submit" class="btn btn-primary me-md-2">
                        <i class="fas fa-save me-1"></i> Salvar
                    </button>
                    <a href="javascript:window.close()" class="btn btn-secondary">
                        <i class="fas fa-arrow-left me-1"></i> Voltar
                    </a>
                </div>
            </form>
        </div>
    </div>
    
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
    <script>
        document.getElementById('searchUserInput').addEventListener('keyup', function() {
            var filter = this.value.toLowerCase();
            var options = document.getElementById('UserId').options;
            for (var i = 0; i < options.length; i++) {
                var optionText = options[i].getAttribute('data-search-text');
                if (optionText) {
                    if (optionText.indexOf(filter) > -1) {
                        options[i].style.display = '';
                    } else {
                        options[i].style.display = 'none';
                    }
                }
            }
        });
    </script>
</body>
</html>