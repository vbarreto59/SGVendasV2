<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<% 'V1
'------------envio de email de login'
Function SendMail(sql)
     if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
       
        'Envia email --------------------------------------------------------------------------------
         set objMail = server.createobject("CDONTS.NewMail")

             objMail.From = "sendmail@gabnetweb.com.br"
             objMail.To  = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"

         objMail.Subject = "SI-USR.NOVO-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time

         objMail.MailFormat = 0
         objMail.Body = sql
         objMail.Send
         'Response.Write "Mensagem Enviada"
         set objMail = Nothing

     end if 'if ip...
     SendMail = ""
end Function
'----------- fim envio de email'
%>

<%
' Valter Barreto - 14 07 2025'    
' ====================================================================================
' BLOC0 VBSCRIPT: Processamento do formulário (executado ao submeter o POST)
' ====================================================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Declaração de variáveis para os dados do formulário
    Dim Usuario, Senha, Email, Telefones, Acao, Permissao, Funcao, Ativo, ID_Grupo
    
    ' Coleta dos dados do formulário
    Usuario = Request.Form("Usuario")
    Senha = Request.Form("Senha")
    Email = Request.Form("Email")
    Telefones = Request.Form("Telefones")
    Acao = Request.Form("Acao")
    Permissao = Request.Form("Permissao")
    Funcao = Request.Form("Funcao")
    Ativo = Request.Form("Ativo")
    ID_Grupo = 2 ' ID fixo para TOCCA_ONZE
    Nome = Request.Form("Nome")
    CRECI = Request.Form("CRECI")

    ' Converte o valor do checkbox 'Ativo' para o formato booleano do Access
    If Ativo = "1" Then
        Ativo = -1 ' True para Access
    Else
        Ativo = 0  ' False para Access
    End If

    ' Força a permissão para CORRETOR (5) e função para CORRETOR
    Permissao = 5
    Funcao = "CORRETOR"
    Acao = "/CONSULTAR/"

    Dim cmd, rs, newUserId
    Set cmd = Server.CreateObject("ADODB.Command")
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open StrConn

    ' 1. Inserir o novo usuário na tabela Usuarios
    cmd.ActiveConnection = objConn
    cmd.CommandText = "INSERT INTO Usuarios (Usuario, SenhaAtual, Nome, Creci, Email, Telefones, Acoes, Permissao, Funcao, Ativo, UsuarioCadastro, GrupoId, Grupo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?)"
    
    cmd.Parameters.Append cmd.CreateParameter("Usuario", 202, 1, 50, Usuario)
    cmd.Parameters.Append cmd.CreateParameter("Senha", 202, 1, 50, Senha)
    cmd.Parameters.Append cmd.CreateParameter("Nome", 202, 1, 255, Nome)
    cmd.Parameters.Append cmd.CreateParameter("CRECI", 202, 1, 50, CRECI)
    cmd.Parameters.Append cmd.CreateParameter("Email", 202, 1, 255, Email)
    cmd.Parameters.Append cmd.CreateParameter("Telefones", 202, 1, 255, Telefones)
    cmd.Parameters.Append cmd.CreateParameter("Acao", 202, 1, 255, Acao)
    cmd.Parameters.Append cmd.CreateParameter("Permissao", 3, 1, , Permissao)
    cmd.Parameters.Append cmd.CreateParameter("Funcao", 202, 1, 50, Funcao)
    cmd.Parameters.Append cmd.CreateParameter("Ativo", 11, 1, , CBool(Ativo))
    cmd.Parameters.Append cmd.CreateParameter("UsuarioCadastro", 202, 1, 50, Session("Usuario"))
    cmd.Parameters.Append cmd.CreateParameter("GrupoId", 3, 1, 4, 1)
    cmd.Parameters.Append cmd.CreateParameter("Grupo", 202, 1, 255, "TOCCA_ONZE")

    
    cmd.Execute
    Set cmd = Nothing

    ' 2. Obter o UserId do usuário recém-inserido
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT MAX(UserId) AS LastID FROM Usuarios", objConn, 1, 3
    
    If Not rs.EOF Then
        newUserId = rs("LastID")
    Else
        Response.Write "Erro: Não foi possível obter o ID do novo usuário."
        objConn.Close
        Set objConn = Nothing
        Response.End
    End If
    rs.Close
    Set rs = Nothing

    ' 3. Inserir a relação na tabela Usuario_Grupo
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandText = "INSERT INTO Usuario_Grupo (UserId, ID_Grupo) VALUES (?, ?)"
    
    cmd.Parameters.Append cmd.CreateParameter("UserIdParam", 3, 1, , newUserId)
    cmd.Parameters.Append cmd.CreateParameter("IDGrupoParam", 3, 1, , 1)   ' GRUPO TOCCA_ONZE'
    
    cmd.Execute
    Set cmd = Nothing
    
    objConn.Close
    Set objConn = Nothing

    ' Registrar Log
    SendMail("Novo usuário: " & Usuario & " - Senha: " & Senha & " - Email: " & Email & " - Telefones: " & Telefones & " - Grupo: TOCCA_ONZE")

    Response.Redirect("usrv_gestao_listar.asp")
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Sunny - Novo Usuário</title>
    
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">  
    
    <style>
        body {
            background-color: #f8f9fa;
        }
        .card-form {
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .form-header {
            background-color: #343a40;
            color: white;
            border-radius: 10px 10px 0 0 !important;
        }
        .btn-custom {
            min-width: 100px;
        }
        .required-field::after {
            content: " *";
            color: red;
        }
        .switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 34px;
        }
        .switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            transition: .4s;
            border-radius: 34px;
        }
        .slider:before {
            position: absolute;
            content: "";
            height: 26px;
            width: 26px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        input:checked + .slider {
            background-color: #28a745;
        }
        input:checked + .slider:before {
            transform: translateX(26px);
        }
        .status-label {
            margin-left: 10px;
            font-weight: 500;
        }
        .status-ativo {
            color: #28a745;
        }
        .status-inativo {
            color: #dc3545;
        }
        .input-group.password-toggle {
            position: relative;
        }
        .input-group.password-toggle .form-control {
            padding-right: 40px;
        }
        .input-group.password-toggle .input-group-append {
            position: absolute;
            right: 0;
            top: 0;
            bottom: 0;
            z-index: 5;
            display: flex;
            align-items: center;
        }
        .input-group.password-toggle .input-group-text {
            background-color: transparent;
            border: none;
            padding: 0.375rem 0.75rem;
            cursor: pointer;
            color: #888;
        }
        .input-group.password-toggle .input-group-text:hover {
            color: #333;
        }
        .fixed-value {
            background-color: #e9ecef;
            cursor: not-allowed;
        }
    </style>
</head>
<body>
    <div class="container py-5">
        <div class="row justify-content-center">
            <div class="col-md-8 col-lg-6">
                <div class="card card-form mb-4">
                    <div class="card-header form-header">
                        <h4 class="mb-0"><i class="fas fa-user-plus mr-2"></i>Novo Usuário</h4>
                    </div>
                    
                    <div class="card-body">
                        <form method="post" action="usrv_novo.asp">
                            <div class="form-group">
                                <label for="Nome" class="required-field">Nome Completo</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-signature"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Nome" name="Nome" placeholder="Digite o nome completo do usuário" required>
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="Usuario" class="required-field">Usuário</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-user"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Usuario" name="Usuario" placeholder="Usuário será gerado automaticamente">
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="Senha" class="required-field">Senha</label>
                                <div class="input-group password-toggle">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-lock"></i></span>
                                    </div>
                                    <input type="password" class="form-control" id="Senha" name="Senha" placeholder="Senha será gerada automaticamente">
                                    <div class="input-group-append">
                                        <span class="input-group-text" id="togglePassword">
                                            <i class="fas fa-eye"></i>
                                        </span>
                                    </div>
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="CRECI">CRECI</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-id-card-alt"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="CRECI" name="CRECI" placeholder="Digite o CRECI (opcional)">
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="Email" class="required-field">Email</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-envelope"></i></span>
                                    </div>
                                    <input type="email" class="form-control" id="Email" name="Email" placeholder="Digite o email" required>
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="Telefones">Telefones</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-phone"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Telefones" name="Telefones" placeholder="Ex: (XX) XXXX-XXXX, (XX) XXXX-XXXX">
                                </div>
                            </div>

                            <div class="form-group">
                                <label class="required-field">Grupo</label>
                                <input type="text" class="form-control fixed-value" value="TOCCA_ONZE" readonly>
                                <input type="hidden" name="ID_Grupo" value="2">
                            </div>

                            <div class="form-group">
                                <label class="required-field">Nível de Permissão</label>
                                <input type="text" class="form-control fixed-value" value="CORRETOR" readonly>
                                <input type="hidden" name="Permissao" value="5">
                            </div>

                            <div class="form-group">
                                <label class="required-field">Função</label>
                                <input type="text" class="form-control fixed-value" id="Funcao" name="Funcao" value="CORRETOR" readonly>
                            </div>

                            <div class="form-group">
                                <label class="required-field">Ações Permitidas</label>
                                <textarea class="form-control fixed-value" id="Acao" name="Acao" rows="3" readonly>/CONSULTAR/</textarea>
                            </div>

                            <div class="form-group">
                                <label>Status do Usuário</label>
                                <div class="d-flex align-items-center">
                                    <label class="switch">
                                        <input type="checkbox" id="Ativo" name="Ativo" value="1" checked>
                                        <span class="slider"></span>
                                    </label>
                                    <span class="status-label status-ativo">ATIVO</span>
                                </div>
                            </div>

                            <div class="form-group text-right mt-4">
                                <button type="submit" class="btn btn-primary btn-custom">
                                    <i class="fas fa-save mr-1"></i> Salvar
                                </button>
                                <a href="usr_listar.asp" class="btn btn-secondary btn-custom">
                                    <i class="fas fa-times mr-1"></i> Cancelar
                                </a>
                            </div>
                        </form>
                    </div>
                </div>
                
                <div class="text-center text-muted small">
                    Sunny &copy; <%= Year(Now()) %>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    
    <script>
        // Função para remover acentos e caracteres especiais
        function removerAcentos(str) {
            return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        }

        // Função para gerar o usuário
        function gerarUsuario(nomeCompleto) {
            if (!nomeCompleto) return "";
            const partesNome = removerAcentos(nomeCompleto).toLowerCase().split(" ");
            let usuario = partesNome[0];
            for (let i = 1; i < partesNome.length; i++) {
                if (partesNome[i].length > 0) {
                    usuario += partesNome[i].charAt(0);
                }
            }
            return usuario;
        }

        // Função para gerar a senha
        function gerarSenha(primeiroNome) {
            if (!primeiroNome) return "";
            const nomeBase = removerAcentos(primeiroNome).toLowerCase();
            const digitosAleatorios = Math.floor(Math.random() * 99) + 1;
            const digitosFormatados = digitosAleatorios < 10 ? "0" + digitosAleatorios : "" + digitosAleatorios;
            return `${nomeBase}@tv${digitosFormatados}`;
        }

        // Executa quando o DOM estiver completamente carregado
        $(document).ready(function() {
            // Lógica para o switch de Ativo/Inativo
            $('#Ativo').change(function() {
                if($(this).is(':checked')) {
                    $('.status-label').removeClass('status-inativo').addClass('status-ativo').text('ATIVO');
                } else {
                    $('.status-label').removeClass('status-ativo').addClass('status-inativo').text('INATIVO');
                }
            });

            // Lógica para o "olho" de mostrar/esconder senha
            $('#togglePassword').click(function() {
                const type = $('#Senha').attr('type') === 'password' ? 'text' : 'password';
                $('#Senha').attr('type', type);
                $(this).find('i').toggleClass('fa-eye fa-eye-slash');
            });

            // Gera usuário e senha automaticamente ao digitar o nome
            $('#Nome').on('input', function() {
                const nomeCompleto = $(this).val().trim();
                const primeiroNome = nomeCompleto.split(" ")[0] || "";
                $('#Usuario').val(gerarUsuario(nomeCompleto));
                $('#Senha').val(gerarSenha(primeiroNome));
            });
        });
    </script>
</body>
</html>