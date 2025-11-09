<%@ Language="VBSCRIPT" %>
<% 
' Definir o tipo de conteúdo e codificação de caracteres
Response.Charset = "UTF-8"
Response.CodePage = 65001

' Função para verificar se está em manutenção
Function EstaEmManutencao()
    Dim fs, arquivoManutencao
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    arquivoManutencao = Server.MapPath("manut.txt")
    
    EstaEmManutencao = fs.FileExists(arquivoManutencao)
    
    Set fs = Nothing
End Function

' Função para verificar se é o usuário admin
Function EAdmin()
    EAdmin = (LCase(Session("Usuario")) = "barreto")
End Function
%>

<!DOCTYPE html>
<html>
<head>
    <title>Site em Manutenção</title>
    <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        h1 { color: #d9534f; }
        .container { max-width: 600px; margin: 0 auto; }
        .admin-panel { margin-top: 30px; padding: 15px; background: #f8f9fa; border-radius: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Site em Manutenção</h1>
        <p>Nosso site está passando por manutenção programada.</p>
        <p>Por favor, volte mais tarde. Agradecemos sua compreensão.</p>
        <img src="img/manutencao_on.jpg" alt="Em Manutenção" class="status-img">
        
        <% If EAdmin() Then %>
        <div class="admin-panel">
            <h3>Painel de Administrador</h3>
            <p>Você está visualizando esta página como administrador.</p>
            <p>
                <a href="ManutencaoOff.asp" class="btn btn-success">Desativar Manutenção</a>
            </p>
        </div>
        <% End If %>
    </div>
Usuário: <%=UCase(usuarioLogado)%><button class="logout-btn" onclick="logout()">(Sair)</button>
    <script>
        function logout() {
                window.location.href = 'gestao_login.asp';
            }
    </script>

</body>
</html>