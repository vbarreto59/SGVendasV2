<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->

<% Response.Charset = "UTF-8" %>
<%
Dim id, conn, rs, rs2, sql, sqlUsuario, rsUsuario
id = Request.QueryString("id")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConn

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim DiretoriaID, NomeGerencia, NomeGerente, TelefoneGerente, UserId, sUserId, sUsuario, sNomeUsuario
    
    ' Pega os dados do formulário
    DiretoriaID = Request.Form("DiretoriaID")
    NomeGerencia = Request.Form("NomeGerencia")
    NomeGerente = Request.Form("NomeGerente")
    TelefoneGerente = Request.Form("TelefoneGerente")
    UserId = Request.Form("UserId") ' Pega o UserId do formulário

    ' Prepara as variáveis para a inserção (evita erros com valores nulos)
    sUserId = "NULL"
    sUsuario = "NULL"
    sNomeUsuario = "NULL"

    ' Se um UserId foi selecionado, busca os dados do usuário na tabela Usuarios
    If UserId <> "" Then
        sqlUsuario = "SELECT Usuario, Nome FROM Usuarios WHERE UserId=" & UserId
        Set rsUsuario = conn.Execute(sqlUsuario)

        If Not rsUsuario.EOF Then
            ' Para garantir que as variáveis Usuario e NomeUsuario estão declaradas
            Dim Usuario, NomeUsuario
            Usuario = rsUsuario("Usuario")
            NomeUsuario = rsUsuario("Nome")
            sUserId = UserId
            sUsuario = "'" & Replace(Usuario, "'", "''") & "'"
            sNomeUsuario = "'" & Replace(NomeUsuario, "'", "''") & "'"
        End If
        rsUsuario.Close
        Set rsUsuario = Nothing
    End If
    
    ' Atualiza o comando SQL e os parâmetros para incluir os novos campos
    sql = "UPDATE Gerencias SET DiretoriaID=?, NomeGerencia=?, NomeGerente=?, TelefoneGerente=?, UserId=?, Usuario=?, Nome=? WHERE GerenciaID=?"
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1
        .Parameters.Append .CreateParameter(, 3, 1, , DiretoriaID)
        .Parameters.Append .CreateParameter(, 200, 1, 100, NomeGerencia)
        .Parameters.Append .CreateParameter(, 200, 1, 100, NomeGerente)
        .Parameters.Append .CreateParameter(, 200, 1, 20, TelefoneGerente)
        
        ' Adiciona os novos parâmetros, tratando os casos de valores nulos
        If sUserId = "NULL" Then
            .Parameters.Append .CreateParameter(, 3, 1, , Empty)
            .Parameters.Append .CreateParameter(, 200, 1, 100, Empty)
            .Parameters.Append .CreateParameter(, 200, 1, 100, Empty)
        Else
            .Parameters.Append .CreateParameter(, 3, 1, , sUserId)
            ' As variáveis 'Usuario' e 'NomeUsuario' precisam estar definidas aqui
            .Parameters.Append .CreateParameter(, 200, 1, 100, Replace(Usuario, "'", "''"))
            .Parameters.Append .CreateParameter(, 200, 1, 100, Replace(NomeUsuario, "'", "''"))
        End If
        
        .Parameters.Append .CreateParameter(, 3, 1, , id)
        .Execute
    End With
    conn.Close
    Set conn = Nothing

    Response.Redirect "gerencia_list.asp"
    Response.End
End If

Set rs = conn.Execute("SELECT * FROM Gerencias WHERE GerenciaID=" & id)

If Not rs.EOF Then
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Editar Gerência</title>
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

    <div class="container py-5">
        <h2 class="text-center mb-4">Editar Gerência</h2>

        <div class="card p-4 mx-auto" style="max-width: 600px;">
            <form method="post" action="gerencia_update.asp?id=<%=id%>">
                <div class="mb-3">
                    <label for="DiretoriaID" class="form-label">Diretoria</label>
                    <select name="DiretoriaID" id="DiretoriaID" required class="form-control">
                        <%
                        Set rs2 = conn.Execute("SELECT * FROM Diretorias ORDER BY NomeDiretoria")
                        Do While Not rs2.EOF
                        Dim sel: sel = ""
                        If rs2("DiretoriaID") = rs("DiretoriaID") Then sel = "selected"
                        %>
                        <option value="<%=rs2("DiretoriaID")%>" <%=sel%>><%=rs2("NomeDiretoria")%></option>
                        <%
                        rs2.MoveNext
                        Loop
                        rs2.Close
                        Set rs2 = Nothing
                        %>
                    </select>
                </div>

                <div class="mb-3">
                    <label for="NomeGerencia" class="form-label">Nome da Gerência</label>
                    <input type="text" name="NomeGerencia" id="NomeGerencia" value="<%=rs("NomeGerencia")%>" required class="form-control">
                </div>

                <div class="mb-3">
                    <label for="NomeGerente" class="form-label">Nome do Gerente</label>
                    <input type="text" name="NomeGerente" id="NomeGerente" value="<%=rs("NomeGerente")%>" required class="form-control">
                </div>

                <div class="mb-3">
                    <label for="TelefoneGerente" class="form-label">Telefone do Gerente</label>
                    <input type="text" name="TelefoneGerente" id="TelefoneGerente" value="<%=rs("TelefoneGerente")%>" class="form-control">
                </div>
                
                <div class="mb-3">
                    <label for="searchUserInput" class="form-label">Buscar Usuário</label>
                    <input type="text" id="searchUserInput" class="form-control" placeholder="Digite para buscar...">
                </div>

                <div class="mb-3">
                    <label for="UserId" class="form-label">Usuário</label>
                    <%
                    Dim rsUsuarios, sqlUsuarios
                    sqlUsuarios = "SELECT UserId, Usuario, Nome FROM Usuarios ORDER BY Nome"
                    Set rsUsuarios = conn.Execute(sqlUsuarios)
                    
                    Dim currentUserId
                    If Not IsNull(rs("UserId")) Then
                        currentUserId = CStr(rs("UserId"))
                    Else
                        currentUserId = ""
                    End If
                    %>
                    <select name="UserId" id="UserId" class="form-control">
                        <option value="">Selecione um Usuário</option>
                        <%
                        Do While Not rsUsuarios.EOF
                            Dim isSelected
                            isSelected = ""
                            If CStr(rsUsuarios("UserId")) = currentUserId Then
                                isSelected = "selected"
                            End If
                        %>
                        <option value="<%=rsUsuarios("UserId")%>" <%=isSelected%> data-search-text="<%=LCase(rsUsuarios("Nome") & " " & rsUsuarios("Usuario"))%>"><%=rsUsuarios("Nome") & " (" & rsUsuarios("Usuario") & ")"%></option>
                        <%
                            rsUsuarios.MoveNext
                        Loop
                        rsUsuarios.Close
                        Set rsUsuarios = Nothing
                        %>
                    </select>
                </div>

                <div class="d-grid gap-2 d-md-flex justify-content-md-end mt-4">
                    <button type="submit" class="btn btn-primary me-md-2">
                        <i class="fas fa-save me-1"></i> Atualizar
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
<%
End If
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>