
<%@ Language="VBSCRIPT" Codepage="65001" %>
<% 
' Definir o tipo de conteúdo e codificação de caracteres
Response.Charset = "UTF-8"
Response.CodePage = 65001
%>
<!--#include file="manutencao_config.asp"-->
<%
' Verificar se é admin para acessar este menu
If Not EAdmin() Then
    Response.Status = "403 Forbidden"
    Response.Write "Acesso negado - Somente administradores"
    Response.End
End If

' Verificar se o sistema está em manutenção
Dim emManutencao
emManutencao = EstaEmManutencao()
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta charset="UTF-8">
    <title>Controle de Manutenção</title>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            text-align: center; 
            padding: 20px; 
            background-color: #f8f9fa;
        }
        .container { 
            max-width: 500px; 
            margin: 0 auto; 
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .status-panel { 
            padding: 20px; 
            margin: 20px 0; 
            border-radius: 5px; 
        }
        .manutencao-on { 
            background-color: #f8d7da; 
            border: 1px solid #f5c6cb; 
            color: #721c24;
        }
        .manutencao-off { 
            background-color: #d4edda; 
            border: 1px solid #c3e6cb; 
            color: #155724;
        }
        .btn { 
            padding: 12px 20px; 
            margin: 10px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 16px;
            text-decoration: none;
            display: inline-block;
            transition: all 0.3s ease;
        }
        .btn:hover {
            opacity: 0.9;
            transform: translateY(-2px);
        }
        .btn-on { 
            background-color: #dc3545; 
            color: white; 
        }
        .btn-off { 
            background-color: #28a745; 
            color: white; 
        }
        .status-img { 
            max-width: 120px; 
            margin: 20px auto; 
            filter: drop-shadow(0 2px 4px rgba(0,0,0,0.1));
        }
        h2 {
            color: #343a40;
            margin-bottom: 25px;
        }
        .footer-info {
            margin-top: 30px; 
            font-size: 13px; 
            color: #6c757d;
            border-top: 1px solid #e9ecef;
            padding-top: 15px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Controle de Manutenção do Sistema</h2>
        
        <div class="status-panel <% If emManutencao Then %>manutencao-on<% Else %>manutencao-off<% End If %>">
            <% If emManutencao Then %>
                <img src="img/manutencao_on.jpg" alt="Em Manutenção" class="status-img">
                <h3 style="color: #dc3545;">SISTEMA EM MANUTENÇÃO</h3>
                <p>Os usuários comuns estão sendo redirecionados para a página de manutenção.</p>
            <% Else %>
                <img src="img/manutencao_off.jpg" alt="Sistema Operacional" class="status-img">
                <h3 style="color: #28a745;">SISTEMA OPERACIONAL</h3>
                <p>Todos os usuários podem acessar o sistema normalmente.</p>
            <% End If %>
        </div>
        
        <div class="action-buttons">
            <a href="ManutencaoOn.asp" class="btn btn-on">
                <i class="fas fa-power-off"></i> Ativar Modo Manutenção
            </a>
            
            <a href="ManutencaoOff.asp" class="btn btn-off">
                <i class="fas fa-check-circle"></i> Desativar Modo Manutenção
            </a>
            <br>
            <a href="menu.asp" class="btn btn-warning">
                <i class="fas fa-check-circle"></i> Voltar
            </a>
        </div>
        
        <div class="footer-info">
            <p>Usuário logado: <strong><% =Server.HTMLEncode(Session("Usuario")) %></strong></p>
            <p>IP: <% =Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR")) %></p>
            <p>Última verificação: <% =FormatDateTime(Now(), vbLongDate) & " às " & FormatDateTime(Now(), vbLongTime) %></p>
        </div>
    </div>
    
    <!-- Ícones do Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
</body>
</html>