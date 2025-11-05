<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes.inc"-->

<!--#include file="registra_log.asp"-->

<%Session("redirecionar") = "gestao_painel2.asp"%>

<%
' Verificar se o formulário foi submetido
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim usuario, senha
    usuario = UCase(Request.Form("usuario"))
    senha = Request.Form("senha")
    
    ' Verificar credenciais
    If VerificarLogin(usuario, senha) Then
        ' Registrar login bem-sucedido
        RegistrarLog usuario, True
        Session("usuario") = usuario
        Session("Acoes") = AcoesPermitidas()
        Session("Funcao") = UserFuncao()

        'registrar no log'
        Call InserirLog ("Sistema", "LOGIN", "Usuário autenticado com sucesso: " & usuario) 


'============================= EMAIL ============================================'
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    set objMail = server.createobject("CDONTS.NewMail")
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    objMail.Subject = "SV-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    objMail.MailFormat = 0
    objMail.Body = "Login SGVendas"
    objMail.Send
    set objMail = Nothing
end if 
'----------- fim envio de email'      



        

        Response.Redirect(Session("redirecionar"))
    Else
        ' Registrar login mal-sucedido
        RegistrarLog usuario, False
        Response.Write("<script>alert('Usuário ou senha incorretos!');</script>")
    End If
End If

Function VerificarLogin(usuario, senha)
    Dim conn, rs, sql
    VerificarLogin = False
    
    ' Incluir arquivo de conexão
    Server.Execute("conexao.asp")
    
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn ' strConn deve ser definida no conexao.asp
    
    sql = "SELECT * FROM usuarios WHERE usuario = '" & Replace(usuario, "'", "''") & "' AND senhaAtual = '" & Replace(senha, "'", "''") & "'"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        VerificarLogin = True
        Session("UserId") = rs("UserId")
    End If
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function

Sub RegistrarLog(usuario, sucesso)
    Dim conn, rs, sql, resultado, idUsuario
    
    On Error Resume Next

    ' Primeiro vamos obter o idUser do usuário
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn
    
    ' Buscar o ID do usuário
    sql = "SELECT UserId FROM usuarios WHERE usuario = '" & Replace(usuario, "'", "''") & "'"

    Set rs = conn.Execute(sql)

  
    If Not rs.EOF Then
        idUsuario = rs("Userid")
    Else
        idUsuario = 0 ' Usuário não encontrado
    End If
    rs.Close

    ' Determinar o texto do histórico
    If sucesso Then
        resultado = UCase(usuario)&": "&"Login Gestão realizado com sucesso"
    Else
        resultado = "Tentativa de login falhou"
    End If
    
    ' Inserir registro na tabela HistAtuLog
    If idUsuario > 0 Then
        sql = "INSERT INTO HistAtuLog (idUser, Usuario, Historico) VALUES (" & _
              idUsuario & ", '" & Usuario &"'," & _
              "'" & Replace(resultado, "'", "''") & "')"
        conn.Execute sql
    'response.Write(sql)
    'response.End() 
    End If
    
    ' Depuração - exibe o erro se houver
    If Err.Number <> 0 Then
        Response.Write "<script>console.error('Erro ao registrar log: " & Server.HTMLEncode(Err.Description) & "');</script>"
        Response.Write "<script>console.error('SQL: " & Server.HTMLEncode(sql) & "');</script>"
        Err.Clear
    End If
    
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema - Login</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
    :root {
        --bordo-principal: #390f0a;
        --bordo-hover: #300909;
        --fundo-claro: #390f0a;
        --texto-principal: #2c1e1e;
    }

    body {
        margin: 0;
        padding: 0;
        font-family: 'Segoe UI', Arial, sans-serif;
        background-color: var(--fundo-claro);
        height: 100vh;
        display: flex;
        justify-content: center;
        align-items: center;
        overflow: hidden;
    }

    .container {
        background-color: #CCCCCC;
        padding: 2.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 8px 30px rgba(75, 14, 14, 0.15);
        width: 100%;
        max-width: 450px;
    }

    .logo-container {
        width: 100%;
        height: 200px;
        margin-bottom: 25px;
        overflow: hidden;
        border-radius: 8px;
        box-shadow: 0 4px 15px rgba(75, 14, 14, 0.15);
    }

    .logo-animado {
        width: 100%;
        height: 100%;
        object-fit: cover;
        opacity: 0;
        transform: scale(0.8);
        animation: fadeInZoom 1.5s ease-out forwards;
        animation-delay: 0.3s;
    }

    @keyframes fadeInZoom {
        0% { opacity: 0; transform: scale(0.8); }
        100% { opacity: 1; transform: scale(1); }
    }

    .btn-login {
        background-color: var(--bordo-principal);
        color: white !important;
        border: none;
        padding: 12px 30px;
        font-size: 16px;
        font-weight: 600;
        cursor: pointer;
        border-radius: 8px;
        transition: all 0.3s;
        box-shadow: 0 4px 10px rgba(75, 14, 14, 0.3);
        width: 100%;
        max-width: 200px;
    }

    .btn-login:hover {
        background-color: var(--bordo-hover);
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(75, 14, 14, 0.4);
    }

    .modal {
        display: none;
        position: fixed;
        z-index: 1000;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        overflow: auto;
        background-color: rgba(0,0,0,0.5);
        backdrop-filter: blur(5px);
    }

    .modal-content {
        background-color: white;
        margin: 10% auto;
        padding: 2rem;
        border-radius: 12px;
        max-width: 400px;
        box-shadow: 0 10px 30px rgba(75, 14, 14, 0.2);
        animation: modalFadeIn 0.4s ease;
    }

    @keyframes modalFadeIn {
        from { opacity: 0; transform: translateY(-30px) scale(0.95); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }

    .modal-title {
        color: var(--bordo-principal);
        font-size: 1.5rem;
        font-weight: 700;
        text-align: center;
    }

    .close {
        color: #ccc;
        position: absolute;
        right: 0;
        top: 0;
        font-size: 1.75rem;
        font-weight: 300;
        cursor: pointer;
    }

    .close:hover {
        color: var(--bordo-hover);
    }

    .form-group label {
        font-weight: 600;
        color: var(--texto-principal);
    }

    .form-control {
        width: 100%;
        padding: 0.75rem 1rem;
        font-size: 1rem;
        border: 1px solid #ddd;
        border-radius: 8px;
        background-color: #fefcfc;
        transition: border-color 0.3s, box-shadow 0.3s;
    }

    .form-control:focus {
        border-color: var(--bordo-principal);
        box-shadow: 0 0 0 3px rgba(75, 14, 14, 0.2);
        outline: none;
        background-color: white;
    }

    .btn-submit {
        background-color: var(--bordo-principal);
        color: white;
        padding: 0.75rem;
        font-size: 1rem;
        font-weight: 600;
        cursor: pointer;
        border-radius: 8px;
        transition: all 0.3s;
        width: 100%;
        margin-top: 0.5rem;
        box-shadow: 0 4px 10px rgba(75, 14, 14, 0.3);
    }

    .btn-submit:hover {
        background-color: var(--bordo-hover);
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(75, 14, 14, 0.4);
    }

    .btn-submit:active {
        transform: translateY(0);
    }
</style>

</head>
<body>
    <div class="container">
        <h3>Tocca Onze - SGVendas - v2</h3>
        <div class="logo-container">
            <img src="img/cupe.jpg" alt="Logo" width="550" height="287" class="logo-animado">
        </div>
        <button class="btn-login" onclick="abrirModal()">Acessar Sistema</button>
    </div>
    
    <!-- Modal de Login reformulado -->
    <div id="loginModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 class="modal-title">Login SGVendas</h2>
                <span class="close" onclick="fecharModal()">&times;</span>
            </div>
            <form method="post" action="">
                <div class="form-group">
                    <label for="usuario">Usuário</label>
                    <input type="text" id="usuario" name="usuario" class="form-control" placeholder="Digite seu usuário" required>
                </div>
                <div class="form-group">
                    <label for="senha">Senha</label>
                    <input type="password" id="senha" name="senha" class="form-control" placeholder="Digite sua senha" required>
                </div>
                <button type="submit" class="btn-submit">Entrar no Sistema</button>
            </form>
        </div>
    </div>
    
    <script>
        function abrirModal() {
            document.getElementById('loginModal').style.display = 'block';
        }
        
        function fecharModal() {
            document.getElementById('loginModal').style.display = 'none';
        }
        
        // Fechar o modal se clicar fora dele
        window.onclick = function(event) {
            var modal = document.getElementById('loginModal');
            if (event.target == modal) {
                fecharModal();
            }
        }
        
        // Fechar com ESC
        document.onkeydown = function(evt) {
            evt = evt || window.event;
            if (evt.keyCode == 27) {
                fecharModal();
            }
        };
    </script>
</body>
</html>