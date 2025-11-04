<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->


<%'funcional sem botao ativar'
' Obter todos os usuários e os grupos que participam
Dim rsUsers, rsGrupos, userId, grupos
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
  </style>
</head>
<body>

  <div class="container">
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
            Dim statusClass, statusText
            If CBool(rsUsers("Ativo")) Then
              statusClass = "badge-ativo"
              statusText = "ATIVO"
            Else
              statusClass = "badge-inativo"
              statusText = "INATIVO"
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
            <!--  -->
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
                <!-- <a href="usr_edit.asp?id=<%=userId%>" class="btn btn-warning" title="Editar"> -->
                  <!-- <i class="fas fa-edit"></i> -->
                </a>


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
});         
  </script>
</body>
</html>

<%
rsUsers.Close()
Set rsUsers = Nothing
%>