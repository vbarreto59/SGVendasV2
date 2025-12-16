<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: ZOSIDTLKPU          -->
<!-- MODIFICAÇÃO: Adicionado funcionalidade de exclusão -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<% 
Response.Charset = "UTF-8" 

' ============================
' PROCESSAMENTO DE EXCLUSÃO
' ============================
Dim mensagem, mensagem_tipo
mensagem = ""
mensagem_tipo = "" ' success, danger, warning

' Verificar se há ação de exclusão via GET
If Request.QueryString("acao") = "excluir" And Request.QueryString("id") <> "" Then
    Dim id_excluir
    id_excluir = Request.QueryString("id")
    
    If IsNumeric(id_excluir) Then
        ' Verificar se a gerência está sendo usada em vendas
        Set connCheck = Server.CreateObject("ADODB.Connection")
        connCheck.Open strConn
        
        ' Verificar na tabela Vendas
        Set rsCheckVendas = connCheck.Execute("SELECT COUNT(*) as total FROM Vendas WHERE GerenciaID = " & id_excluir & " AND EXCLUIDO = 0")
        Dim total_vendas
        total_vendas = 0
        If Not rsCheckVendas.EOF Then
            total_vendas = rsCheckVendas("total")
        End If
        rsCheckVendas.Close
        Set rsCheckVendas = Nothing
        
        If total_vendas > 0 Then
            ' Gerência está em uso, não pode excluir
            mensagem = "Não é possível excluir esta gerência pois ela está vinculada a " & total_vendas & " venda(s)."
            mensagem_tipo = "warning"
        Else
            ' Tentar excluir
            On Error Resume Next
            connCheck.Execute("DELETE FROM Gerencias WHERE GerenciaID = " & id_excluir)
            
            If Err.Number = 0 Then
                mensagem = "Gerência excluída com sucesso!"
                mensagem_tipo = "success"
                
                ' Registrar log
                Call InserirLog("GERENCIAS", "DELETE", "Gerência ID " & id_excluir & " excluída")
            Else
                mensagem = "Erro ao excluir gerência: " & Err.Description
                mensagem_tipo = "danger"
            End If
            On Error GoTo 0
        End If
        
        connCheck.Close
        Set connCheck = Nothing
    Else
        mensagem = "ID inválido!"
        mensagem_tipo = "danger"
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lista de Gerências</title>
    
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
    
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">
    
    <style>
        body {
            /* Define a escala de 0.8 (80%) */
            transform: scale(0.8); 
            
            /* Define o ponto de origem para o canto superior esquerdo */
            transform-origin: 0 0; 
            
            /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
            /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
            width: calc(100% / 0.8); 
        }
        
        .btn-excluir {
            background-color: #dc3545;
            border-color: #dc3545;
            color: white;
        }
        
        .btn-excluir:hover {
            background-color: #c82333;
            border-color: #bd2130;
        }
        
        .modal-confirmacao .modal-header {
            background-color: #dc3545;
            color: white;
        }
        
        .modal-aviso .modal-header {
            background-color: #ffc107;
            color: #212529;
        }
    </style>    
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
        <h2 class="text-center mb-4">Lista de Gerências</h2>
        
        <% If mensagem <> "" Then %>
        <div class="alert alert-<%=mensagem_tipo%> alert-dismissible fade show" role="alert">
            <% 
            Select Case mensagem_tipo
                Case "success"
                    Response.Write "<i class='fas fa-check-circle me-2'></i>"
                Case "warning"
                    Response.Write "<i class='fas fa-exclamation-triangle me-2'></i>"
                Case "danger"
                    Response.Write "<i class='fas fa-times-circle me-2'></i>"
                Case Else
                    Response.Write "<i class='fas fa-info-circle me-2'></i>"
            End Select
            %>
            <%=mensagem%>
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        </div>
        <% End If %>
        
        <p>
            <a class="btn btn-primary" href="gerencia_create.asp" target="_blank">
                <i class="fas fa-plus me-1"></i> Nova Gerência
            </a>
        </p>
        
        <div class="table-responsive">
            <table id="tabelaGerencias" class="display compact nowrap" style="width:100%">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Diretoria</th>
                        <th>Gerência</th>
                        <th>Nome Gerente</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConn
' A consulta agora faz um LEFT JOIN apenas com a tabela Diretorias
sql = "SELECT G.*, D.NomeDiretoria FROM Gerencias G LEFT JOIN Diretorias D ON G.DiretoriaID = D.DiretoriaID ORDER BY G.NomeGerencia"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
%>
                    <tr id="gerencia-<%=rs("GerenciaID")%>">
                        <td><%=rs("GerenciaID")%></td>
                        <td class="name-column"><span title="<%=rs("NomeDiretoria")%>"><%=rs("NomeDiretoria")%></span></td>
                        <td class="name-column"><span title="<%=rs("NomeGerencia")%>"><%=rs("NomeGerencia")%></span></td>
                        <td><%=rs("Nome")%></td>
                        <td>
                            <a class="btn btn-sm btn-info" href="gerencia_update.asp?id=<%=rs("GerenciaID")%>" target="_blank">
                                <i class="fas fa-edit me-1"></i> Editar
                            </a>
                            <%if Trim(Session("Usuario")) = "BARRETO" then %>
                                <button class="btn btn-sm btn-excluir" onclick="confirmarExclusao(<%=rs("GerenciaID")%>, '<%=Replace(Server.HTMLEncode(rs("NomeGerencia")), "'", "\'")%>')">
                                    <i class="fas fa-trash-alt me-1"></i> Excluir
                                </button>
                            <%end if%>                                
                        </td>
                    </tr>
<%
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- Modal de Confirmação de Exclusão -->
    <div class="modal fade modal-confirmacao" id="modalExclusao" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="fas fa-exclamation-triangle me-2"></i>Confirmar Exclusão
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <p>Tem certeza que deseja excluir a gerência <strong id="nomeGerenciaExcluir"></strong>?</p>
                    <p class="text-danger"><strong>Atenção:</strong> Esta ação não pode ser desfeita!</p>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <a href="#" id="btnConfirmarExclusao" class="btn btn-danger">
                        <i class="fas fa-trash-alt me-1"></i> Excluir
                    </a>
                </div>
            </div>
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

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    
    <script>
        $(document).ready(function() {
            $('#tabelaGerencias').DataTable({
                "language": {
                    "url": "//cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
                },
                "pageLength": 100,
                "paging": true,
                "lengthChange": true,
                "searching": true,
                "ordering": true,
                "info": true,
                "autoWidth": true,
                "responsive": true,
                "columnDefs": [
                    {
                        "targets": [4], // Coluna de ações
                        "orderable": false,
                        "searchable": false
                    }
                ]
            });
            
            // Auto-fechar alertas após 5 segundos
            setTimeout(function() {
                $('.alert').alert('close');
            }, 5000);
        });
        
        function confirmarExclusao(id, nome) {
            // Configurar modal
            document.getElementById('nomeGerenciaExcluir').textContent = nome;
            
            // Configurar link de exclusão
            var linkExclusao = document.getElementById('btnConfirmarExclusao');
            linkExclusao.href = '?acao=excluir&id=' + id;
            
            // Mostrar modal
            var modal = new bootstrap.Modal(document.getElementById('modalExclusao'));
            modal.show();
        }
        
        // Versão alternativa com SweetAlert2 (opcional)
        function confirmarExclusaoSweetAlert(id, nome) {
            Swal.fire({
                title: 'Confirmar Exclusão',
                html: 'Tem certeza que deseja excluir a gerência <strong>' + nome + '</strong>?<br><br><span class="text-danger">Esta ação não pode ser desfeita!</span>',
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#d33',
                cancelButtonColor: '#3085d6',
                confirmButtonText: 'Sim, excluir!',
                cancelButtonText: 'Cancelar'
            }).then((result) => {
                if (result.isConfirmed) {
                    window.location.href = '?acao=excluir&id=' + id;
                }
            });
        }
    </script>
</body>
</html>