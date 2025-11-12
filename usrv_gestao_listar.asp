<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->
<%
' Processar ativação/desativação do usuário
If Request.Form("acao") = "toggle_status" Then
    Dim userId, novoStatus
    userId = Request.Form("user_id")
    novoStatus = Request.Form("novo_status")
    
    If userId <> "" And novoStatus <> "" Then
        On Error Resume Next
        ' Criar objeto Command para melhor controle
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = StrConn
        cmd.CommandText = "UPDATE Usuarios SET Ativo = ? WHERE UserID = ? AND IdEmp = 2"
        cmd.Parameters.Append cmd.CreateParameter("Ativo", 3, 1, , novoStatus)
        cmd.Parameters.Append cmd.CreateParameter("UserID", 3, 1, , userId)
        cmd.Execute
        
        If Err.Number = 0 Then
            ' Redirecionar para evitar reenvio do formulário
            Response.Redirect "?success=1&userid=" & userId
        Else
            Response.Redirect "?error=1&msg=" & Server.URLEncode(Err.Description)
        End If
        On Error GoTo 0
        Set cmd = Nothing
    End If
End If

' Mostrar mensagens de sucesso/erro via QueryString
If Request.QueryString("success") = "1" Then
    Response.Write "<script>alert('Status do usuário atualizado com sucesso!');</script>"
End If

If Request.QueryString("error") = "1" Then
    Dim errorMsg
    errorMsg = Request.QueryString("msg")
    If errorMsg = "" Then errorMsg = "Erro desconhecido"
    Response.Write "<script>alert('Erro ao atualizar status do usuário: " & Replace(errorMsg, "'", "`") & "');</script>"
End If

' Obter todos os usuários e os grupos que participam
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open "SELECT * FROM Usuarios WHERE IdEmp = 2 ORDER BY Usuario ASC", StrConn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>SGVendas - Lista de Usuários</title>
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  
  <!-- Bootstrap CSS -->
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
  
  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">
  
  <!-- DataTables CSS -->
  <link rel="stylesheet" href="https://cdn.datatables.net/1.10.22/css/dataTables.bootstrap4.min.css">
  
  <style>
    body {
      background-color: #f8f9fa;
    }
    .table-responsive {
      background-color: white;
      border-radius: 10px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      padding: 20px;
      margin-top: 20px;
    }
    .table-header {
      background-color: #343a40;
      color: white;
      border-radius: 10px 10px 0 0;
      padding: 15px 20px;
      margin-bottom: 0;
    }
    .btn-sm {
      min-width: 70px;
    }
    .table {
      width: 100%;
    }
    .table th {
      white-space: nowrap;
    }
    .badge-permissao {
      font-size: 0.85em;
      padding: 0.35em 0.65em;
    }
    .badge-grupo {
      font-size: 0.8em;
      margin-right: 3px;
      margin-bottom: 3px;
      display: inline-block;
    }
    .grupos-container {
      max-width: 250px;
    }
    .header-actions {
      margin-bottom: 20px;
    }
    .badge-status {
      font-size: 0.85em;
      padding: 0.5em 0.75em;
      border-radius: 50px;
      min-width: 70px;
      display: inline-block;
      text-align: center;
    }
    .badge-ativo {
      background-color: #28a745;
      color: white;
    }
    .badge-inativo {
      background-color: #dc3545;
      color: white;
    }
    .user-inativo {
      opacity: 0.7;
    }
    .btn-toggle {
      width: 80px;
      font-size: 0.8rem;
    }
    .btn-ativo {
      background-color: #28a745;
      border-color: #28a745;
      color: white;
    }
    .btn-inativo {
      background-color: #6c757d;
      border-color: #6c757d;
      color: white;
    }
    .btn-toggle:hover {
      transform: translateY(-1px);
      transition: all 0.2s;
    }
    .toggle-form {
      display: inline;
    }
    .alert-container {
      position: fixed;
      top: 20px;
      right: 20px;
      z-index: 1000;
      min-width: 300px;
    }
  </style>
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
</style>  
</head>
<body>

  <div class="container">
    <!-- Container para alertas -->
    <div class="alert-container">
      <%
      ' Mostrar alertas Bootstrap em vez de JavaScript
      If Request.QueryString("success") = "1" Then
        Response.Write "<div class='alert alert-success alert-dismissible fade show'>" & _
                      "<i class='fas fa-check-circle mr-2'></i>Status do usuário atualizado com sucesso!" & _
                      "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                      "</div>"
      End If
      
      If Request.QueryString("error") = "1" Then
        Dim errorMsgDisplay
        errorMsgDisplay = Request.QueryString("msg")
        If errorMsgDisplay = "" Then errorMsgDisplay = "Erro desconhecido"
        Response.Write "<div class='alert alert-danger alert-dismissible fade show'>" & _
                      "<i class='fas fa-exclamation-circle mr-2'></i>Erro ao atualizar status: " & Server.HTMLEncode(errorMsgDisplay) & _
                      "<button type='button' class='close' data-dismiss='alert'><span>&times;</span></button>" & _
                      "</div>"
      End If
      %>
    </div>

    <div class="d-flex justify-content-between align-items-center header-actions">
      <h4 class="mb-0"><i class="fas fa-users mr-2"></i>Lista de Usuários</h4>
      <div>
        <a href="#" class="btn btn-info" onclick="window.close(); return false;">
          <i class="fas fa-times mr-1"></i> Fechar
        </a>
      </div>
    </div>
    
    <div class="table-responsive">
      <table id="tabelaUsuarios" class="table table-striped table-bordered table-hover" style="width:100%">
        <thead class="thead-dark">
          <tr>
            <th>ID</th>
            <th>Usuário</th>
            <th>Status</th>
            <th>Função</th>
            <th>Grupos</th>
            <th class="text-center">Ações</th>
          </tr>
        </thead>
        <tbody>
          <% 
          While Not rsUsers.EOF 
            userId = rsUsers("UserID")
            
            ' Obter grupos do usuário
            Set rsGrupos = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT g.ID_Grupo, g.Nome_Grupo FROM Grupo g " & _
                         "INNER JOIN Usuario_Grupo ug ON g.ID_Grupo = ug.ID_Grupo " & _
                         "WHERE ug.UserId = " & userId & " ORDER BY g.Nome_Grupo"
                       
            rsGrupos.Open sql, StrConn
            
            grupos = ""
            Do While Not rsGrupos.EOF
              grupos = grupos & "<span class='badge badge-info badge-grupo'>" & Server.HTMLEncode(rsGrupos("Nome_Grupo")) & "</span>"
              rsGrupos.MoveNext
            Loop
            
            If grupos = "" Then
              grupos = "<span class='text-muted'>Nenhum grupo</span>"
            End If
            
            rsGrupos.Close
            Set rsGrupos = Nothing
            
            ' Determinar status do usuário
            If CBool(rsUsers("Ativo")) Then
              statusClass = "badge-ativo"
              statusText = "ATIVO"
              btnClass = "btn-ativo"
              btnText = "ATIVO"
              btnIcon = "fas fa-toggle-on"
              novoStatus = "0"
            Else
              statusClass = "badge-inativo"
              statusText = "INATIVO"
              btnClass = "btn-inativo"
              btnText = "INATIVO"
              btnIcon = "fas fa-toggle-off"
              novoStatus = "-1"
            End If
          %>
          <tr class="<% If Not CBool(rsUsers("Ativo")) Then Response.Write "user-inativo" %>">
            <td><strong><%=userId%></strong></td>
            <td>
                <strong><%=UCase(rsUsers("Usuario"))%></strong><br>
                <small class="text-muted"><i class="fas fa-user mr-1"></i><%=rsUsers("Nome")%></small><br>
                <small class="text-muted"><i class="fas fa-envelope mr-1"></i><%=rsUsers("Email")%></small><br>
                <small class="text-muted"><i class="fas fa-phone mr-1"></i><%=rsUsers("Telefones")%></small><br>
                <small class="text-muted"><i class="fas fa-id-badge mr-1"></i>CRECI: <%=rsUsers("CRECI")%></small>
            </td>
            <td>
              <span class="badge badge-status <%=statusClass%>">
                <%=statusText%>
              </span>
            </td>
            <td>
              <% 
              Select Case rsUsers("Permissao")
                Case 1: badgeClass = "badge-danger"
                Case 2: badgeClass = "badge-warning"
                Case 3: badgeClass = "badge-warning"
                Case 4: badgeClass = "badge-info"
                Case 5: badgeClass = "badge-secondary"
                Case 6: badgeClass = "badge-secondary"
                Case Else: badgeClass = "badge-light"
              End Select
              %>
              <span class="badge <%=badgeClass%> badge-permissao"><%=UCase(rsUsers("Funcao"))%></span>
            </td>
            
            <td class="grupos-container"><%=grupos%></td>
            <td class="text-center">
              <div class="btn-group btn-group-sm" role="group">
                <!-- Botão Liga/Desliga -->
                <form method="post" class="toggle-form" onsubmit="return confirmToggle(this);">
                  <input type="hidden" name="acao" value="toggle_status">
                  <input type="hidden" name="user_id" value="<%=userId%>">
                  <input type="hidden" name="novo_status" value="<%=novoStatus%>">
                  <button type="submit" class="btn <%=btnClass%> btn-toggle" title="<% If CBool(rsUsers("Ativo")) Then %>Desativar Usuário<% Else %>Ativar Usuário<% End If %>">
                    <i class="<%=btnIcon%> mr-1"></i><%=btnText%>
                  </button>
                </form>
              </div>
            </td>
          </tr>
          <% 
            rsUsers.MoveNext()
          Wend 
          %>
        </tbody>
      </table>
    </div>
    
    <footer class="text-center text-muted small mb-3">
      Sunny System &copy; <%= Year(Now()) %>
    </footer>
  </div>

  <!-- jQuery first, then Popper.js, then Bootstrap JS -->
  <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
  
  <!-- DataTables JS -->
  <script src="https://cdn.datatables.net/1.10.22/js/jquery.dataTables.min.js"></script>
  <script src="https://cdn.datatables.net/1.10.22/js/dataTables.bootstrap4.min.js"></script>
  
  <script>
$(document).ready(function() {
    $('#tabelaUsuarios').DataTable({
        "order": [[0, "desc"]],
        "pageLength": 100,
        "language": {
            "sEmptyTable": "Nenhum registro encontrado",
            "sInfo": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
            "sInfoEmpty": "Mostrando 0 até 0 de 0 registros",
            "sInfoFiltered": "(Filtrados de _MAX_ registros)",
            "sInfoPostFix": "",
            "sInfoThousands": ".",
            "sLengthMenu": "_MENU_ resultados por página",
            "sLoadingRecords": "Carregando...",
            "sProcessing": "Processando...",
            "sZeroRecords": "Nenhum registro encontrado",
            "sSearch": "Pesquisar:",
            "oPaginate": {
                "sNext": "Próximo",
                "sPrevious": "Anterior",
                "sFirst": "Primeiro",
                "sLast": "Último"
            },
            "oAria": {
                "sSortAscending": ": Ordenar colunas de forma ascendente",
                "sSortDescending": ": Ordenar colunas de forma descendente"
            },
            "select": {
                "rows": {
                    "_": "Selecionado %d linhas",
                    "0": "Nenhuma linha selecionada",
                    "1": "Selecionado 1 linha"
                }
            },
            "decimal": ",",
            "thousands": "."
        },
        "dom": '<"top"lif>rt<"bottom"lip><"clear">',
        "responsive": true,
        "initComplete": function() {
            $('.dataTables_filter input').addClass('form-control').attr('placeholder', 'Pesquisar...');
            $('.dataTables_length select').addClass('form-control');
        },
        "columnDefs": [
            { "responsivePriority": 1, "targets": 1 }, // Usuário
            { "responsivePriority": 2, "targets": -1 }, // Ações
            { "responsivePriority": 3, "targets": 5 }, // Grupos
            { "responsivePriority": 4, "targets": 3 }, // Função
            { "responsivePriority": 5, "targets": 2 }, // Status
            { "responsivePriority": 6, "targets": 0 }  // ID
        ]
    });

    // Auto-close alerts after 5 seconds
    setTimeout(function() {
        $('.alert').alert('close');
    }, 5000);
});

// Função para confirmar a alteração de status
function confirmToggle(form) {
    var userId = form.user_id.value;
    var novoStatus = form.novo_status.value;
    var acao = (novoStatus == "-1") ? "ativar" : "desativar";
    var nomeUsuario = form.closest('tr').querySelector('td:nth-child(2) strong').textContent;
    
    if (confirm("Tem certeza que deseja " + acao + " o usuário '" + nomeUsuario + "'?")) {
        // Mostrar loading no botão
        var btn = form.querySelector('button');
        var originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-1"></i>Processando...';
        btn.disabled = true;
        
        // Desabilitar todos os botões para evitar múltiplos cliques
        var allButtons = document.querySelectorAll('.btn-toggle');
        allButtons.forEach(function(button) {
            button.disabled = true;
        });
        
        // Enviar o formulário
        return true;
    }
    return false;
}
  </script>
</body>
</html>

<%
rsUsers.Close()
Set rsUsers = Nothing
%>