<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: DIRETORIAS_VENDAS_CONTROL -->
<!-- OBS: Controle de acesso ao diretório de vendas -->
<!-- ###################################### -->

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

' Definir constantes
Const ARQUIVO_DIRETORIAS = "DIRETORIAS_VENDAS.TXT"
Const MENSAGEM_BLOQUEIO = "O acesso ao diretório de vendas está temporariamente bloqueado para manutenção."

' Função para verificar se o diretório de vendas está BLOQUEADO
' RETORNA TRUE se estiver BLOQUEADO (arquivo existe)
Function DiretorioVendasBloqueado()
    Dim fs, arquivoPath
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    arquivoPath = Server.MapPath(ARQUIVO_DIRETORIAS)
    DiretorioVendasBloqueado = fs.FileExists(arquivoPath)
    Set fs = Nothing
End Function

' Função para BLOQUEAR diretório de vendas (CRIA arquivo)
Sub BloquearDiretorioVendas()
    Dim fs, arquivo
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    
    Dim arquivoPath
    arquivoPath = Server.MapPath(ARQUIVO_DIRETORIAS)
    
    Set arquivo = fs.CreateTextFile(arquivoPath, True)
    arquivo.WriteLine("DIRETORIO_VENDAS_BLOQUEADO")
    arquivo.WriteLine("Data/Hora: " & Now())
    arquivo.WriteLine("Bloqueado por: " & Session("Usuario"))
    arquivo.WriteLine("IP: " & Request.ServerVariables("REMOTE_ADDR"))
    arquivo.WriteLine("Mensagem: " & MENSAGEM_BLOQUEIO)
    arquivo.WriteLine("")
    arquivo.WriteLine("Este arquivo bloqueia o acesso ao diretório de vendas.")
    arquivo.WriteLine("Para liberar o acesso, exclua este arquivo.")
    arquivo.Close
    Set arquivo = Nothing
    
    Set fs = Nothing
End Sub

' Função para LIBERAR diretório de vendas (EXCLUI arquivo)
Sub LiberarDiretorioVendas()
    Dim fs, arquivoPath
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    arquivoPath = Server.MapPath(ARQUIVO_DIRETORIAS)
    
    If fs.FileExists(arquivoPath) Then
        fs.DeleteFile arquivoPath, True
    End If
    
    Set fs = Nothing
End Sub

' Processar ações se houver parâmetro na URL
Dim acao
acao = Request.QueryString("acao")

If acao = "liberar" Then
    LiberarDiretorioVendas()
    Response.Redirect "?msg=liberado"
ElseIf acao = "bloquear" Then
    BloquearDiretorioVendas()
    Response.Redirect "?msg=bloqueado"
End If

' Verificar status atual - TRUE = BLOQUEADO, FALSE = LIBERADO
Dim diretorioBloqueado
diretorioBloqueado = DiretorioVendasBloqueado()
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta charset="UTF-8">
    <title>Controle de Diretório de Vendas</title>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            text-align: center; 
            padding: 20px; 
            background-color: #f8f9fa;
        }
        .container { 
            max-width: 600px; 
            margin: 0 auto; 
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .status-panel { 
            padding: 25px; 
            margin: 20px 0; 
            border-radius: 8px; 
            text-align: left;
        }
        .diretorio-off { 
            background-color: #d4edda; 
            border: 1px solid #c3e6cb; 
            color: #155724;
        }
        .diretorio-on { 
            background-color: #f8d7da; 
            border: 1px solid #f5c6cb; 
            color: #721c24;
        }
        .btn { 
            padding: 12px 25px; 
            margin: 10px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 16px;
            text-decoration: none;
            display: inline-block;
            transition: all 0.3s ease;
            font-weight: 500;
        }
        .btn:hover {
            opacity: 0.9;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        .btn-on { 
            background-color: #28a745; 
            color: white; 
        }
        .btn-off { 
            background-color: #dc3545; 
            color: white; 
        }
        .btn-warning {
            background-color: #ffc107;
            color: #212529;
        }
        .status-img { 
            max-width: 100px; 
            margin: 0 auto 20px; 
            display: block;
        }
        h2 {
            color: #343a40;
            margin-bottom: 25px;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 15px;
        }
        .footer-info {
            margin-top: 30px; 
            font-size: 13px; 
            color: #6c757d;
            border-top: 1px solid #e9ecef;
            padding-top: 15px;
            text-align: left;
        }
        .alert {
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        .alert-success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .alert-danger {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        .info-box {
            background-color: #e9ecef;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
            text-align: left;
            font-size: 14px;
        }
        .arquivo-info {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
            font-family: monospace;
            font-size: 12px;
            text-align: left;
        }
        .status-indicator {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            margin-right: 5px;
        }
        .status-ok {
            background-color: #28a745;
        }
        .status-blocked {
            background-color: #dc3545;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2><i class="fas fa-folder"></i> Controle de Diretório de Vendas</h2>
        
        <% ' Exibir mensagens de confirmação
        Dim msg
        msg = Request.QueryString("msg")
        If msg = "liberado" Then
            Response.Write "<div class='alert alert-success'>"
            Response.Write "<i class='fas fa-check-circle'></i> Diretório de vendas LIBERADO com sucesso!"
            Response.Write "</div>"
        ElseIf msg = "bloqueado" Then
            Response.Write "<div class='alert alert-danger'>"
            Response.Write "<i class='fas fa-exclamation-circle'></i> Diretório de vendas BLOQUEADO com sucesso!"
            Response.Write "</div>"
        End If
        %>
        
        <div class="status-panel <% If diretorioBloqueado Then %>diretorio-on<% Else %>diretorio-off<% End If %>">
            <% If diretorioBloqueado Then %>
                <img src="img/warning.png" alt="Diretório Bloqueado" class="status-img">
                <h3 style="color: #dc3545;"><span class="status-indicator status-blocked"></span> DIRETÓRIO DE VENDAS BLOQUEADO</h3>
                <p><strong>Status:</strong> Acesso RESTRITO aos usuários</p>
                <p><strong>Arquivo de controle:</strong> <code><%= ARQUIVO_DIRETORIAS %></code> (EXISTE)</p>
                <p><strong>Mensagem exibida aos usuários:</strong> "<%= MENSAGEM_BLOQUEIO %>"</p>
            <% Else %>
                <img src="img/check_circle.png" alt="Diretório Liberado" class="status-img">
                <h3 style="color: #28a745;"><span class="status-indicator status-ok"></span> DIRETÓRIO DE VENDAS LIBERADO</h3>
                <p><strong>Status:</strong> Acesso PERMITIDO aos usuários</p>
                <p><strong>Arquivo de controle:</strong> <code><%= ARQUIVO_DIRETORIAS %></code> (NÃO EXISTE)</p>
                <p>Os usuários podem acessar o diretório de vendas normalmente.</p>
            <% End If %>
        </div>
        
        <div class="info-box">
            <h4><i class="fas fa-info-circle"></i> Como funciona:</h4>
            <p><strong>LÓGICA DE CONTROLE:</strong></p>
            <ul>
                <li><span class="status-indicator status-blocked"></span> <strong>BLOQUEAR</strong> = CRIAR arquivo <code><%= ARQUIVO_DIRETORIAS %></code></li>
                <li><span class="status-indicator status-ok"></span> <strong>LIBERAR</strong> = EXCLUIR arquivo <code><%= ARQUIVO_DIRETORIAS %></code></li>
            </ul>
            <p>Quando o arquivo existe, o sistema interpreta que o diretório está bloqueado.</p>
        </div>
        
        <% If diretorioBloqueado Then %>
        <div class="arquivo-info">
            <strong><i class="fas fa-file-alt"></i> Conteúdo do arquivo atual:</strong><br>
            <%
            Dim fsArq, arquivoPath, conteudo
            arquivoPath = Server.MapPath(ARQUIVO_DIRETORIAS)
            Set fsArq = Server.CreateObject("Scripting.FileSystemObject")
            
            If fsArq.FileExists(arquivoPath) Then
                Dim objArquivo
                Set objArquivo = fsArq.OpenTextFile(arquivoPath, 1)
                conteudo = Replace(objArquivo.ReadAll(), vbCrLf, "<br>")
                objArquivo.Close
                Set objArquivo = Nothing
                Response.Write conteudo
            End If
            
            Set fsArq = Nothing
            %>
        </div>
        <% End If %>
        
        <div class="action-buttons">
            <% If diretorioBloqueado Then %>
                <a href="?acao=liberar" class="btn btn-on" onclick="return confirm('Tem certeza que deseja LIBERAR o diretório de vendas?\n\nIsso irá EXCLUIR o arquivo de bloqueio.')">
                    <i class="fas fa-unlock"></i> Liberar Acesso ao Diretório
                </a>
                <p class="info-text"><small>Irá EXCLUIR o arquivo <%= ARQUIVO_DIRETORIAS %></small></p>
            <% Else %>
                <a href="?acao=bloquear" class="btn btn-off" onclick="return confirm('Tem certeza que deseja BLOQUEAR o diretório de vendas?\n\nIsso irá CRIAR um arquivo de bloqueio.')">
                    <i class="fas fa-lock"></i> Bloquear Acesso ao Diretório
                </a>
                <p class="info-text"><small>Irá CRIAR o arquivo <%= ARQUIVO_DIRETORIAS %></small></p>
            <% End If %>
            
            <br><br>
            <a href="menu.asp" class="btn btn-warning">
                <i class="fas fa-arrow-left"></i> Voltar ao Menu
            </a>
        </div>
        
        <div class="footer-info">
            <p><strong>Status atual:</strong> 
                <% If diretorioBloqueado Then %>
                    <span style="color: #dc3545;"><span class="status-indicator status-blocked"></span> BLOQUEADO (arquivo existe)</span>
                <% Else %>
                    <span style="color: #28a745;"><span class="status-indicator status-ok"></span> LIBERADO (arquivo não existe)</span>
                <% End If %>
            </p>
            <p><strong>Usuário administrador:</strong> <% =Server.HTMLEncode(Session("Usuario")) %></p>
            <p><strong>IP de acesso:</strong> <% =Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR")) %></p>
            <p><strong>Data/Hora da consulta:</strong> <% =FormatDateTime(Now(), vbLongDate) & " às " & FormatDateTime(Now(), vbLongTime) %></p>
            <p><strong>Arquivo de controle:</strong> <% =ARQUIVO_DIRETORIAS %></p>
            <p><strong>Caminho físico no servidor:</strong> <br><small><% =Server.MapPath(ARQUIVO_DIRETORIAS) %></small></p>
        </div>
    </div>
    
    <!-- Ícones do Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    
    <script>
        // Confirmar ações importantes
        function confirmarAcao(acao) {
            if (acao === 'bloquear') {
                return confirm('ATENÇÃO: BLOQUEAR o diretório de vendas impedirá o acesso dos usuários.\n\nEsta ação irá CRIAR o arquivo de bloqueio.\n\nContinuar?');
            } else if (acao === 'liberar') {
                return confirm('Deseja LIBERAR o acesso ao diretório de vendas?\n\nEsta ação irá EXCLUIR o arquivo de bloqueio.\n\nContinuar?');
            }
            return true;
        }
        
        // Atualizar a página a cada 30 segundos para mostrar status atualizado
        setTimeout(function() {
            window.location.reload();
        }, 30000);
    </script>
</body>
</html>