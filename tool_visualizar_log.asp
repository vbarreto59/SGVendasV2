<%@ Language=VBScript CodePage=65001 %>
<% Response.CodePage = 65001 %>
<% Response.Charset = "UTF-8" %>
<!--#include file="conSunSales.asp"-->
<%
' Verificar se o usuário está logado (opcional)
If Session("Usuario") = "" Then
    Response.Redirect "login.asp"
End If
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Visualizar Logs do Sistema</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.bootstrap5.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        .table-container {
            background: white;
            border-radius: 10px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            padding: 20px;
            margin-top: 20px;
        }
        .page-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px 0;
            margin-bottom: 30px;
            border-radius: 10px;
        }
        .badge-success { background-color: #28a745; }
        .badge-warning { background-color: #ffc107; color: #000; }
        .badge-danger { background-color: #dc3545; }
        .badge-info { background-color: #17a2b8; }
        .badge-secondary { background-color: #6c757d; }
        .btn-export {
            margin-right: 5px;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <!-- Cabeçalho -->
        <div class="page-header">
            <div class="container">
                <div class="row">
                    <div class="col-md-8">
                        <h1><i class="fas fa-clipboard-list"></i> Logs do Sistema</h1>
                        <p class="lead">Registro de todas as atividades do sistema</p>
                    </div>
                    <div class="col-md-4 text-end">
                        <p><strong>Usuário:</strong> <%=Server.HTMLEncode(Session("Usuario"))%></p>
                        <a href="javascript:history.back()" class="btn btn-light">Voltar</a>
                        <a href="logout.asp" class="btn btn-outline-light">Sair</a>
                    </div>
                </div>
            </div>
        </div>

        <!-- Filtros -->
        <div class="card mb-4">
            <div class="card-header">
                <h5 class="mb-0"><i class="fas fa-filter"></i> Filtros</h5>
            </div>
            <div class="card-body">
                <div class="row g-3">
                    <div class="col-md-3">
                        <label class="form-label">Ação</label>
                        <select id="filterAcao" class="form-select">
                            <option value="">Todas as ações</option>
                            <option value="INSERT">INSERT</option>
                            <option value="UPDATE">UPDATE</option>
                            <option value="DELETE">DELETE</option>
                            <option value="LOGIN">LOGIN</option>
                            <option value="LOGOUT">LOGOUT</option>
                            <option value="ERRO">ERRO</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Tabela</label>
                        <select id="filterTabela" class="form-select">
                            <option value="">Todas as tabelas</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Usuário</label>
                        <select id="filterUsuario" class="form-select">
                            <option value="">Todos os usuários</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label">Data</label>
                        <input type="date" id="filterData" class="form-control">
                    </div>
                </div>
                <div class="row mt-3">
                    <div class="col-12">
                        <button id="btnLimparFiltros" class="btn btn-secondary">Limpar Filtros</button>
                        <button id="btnExportarExcel" class="btn btn-success btn-export">
                            <i class="fas fa-file-excel"></i> Exportar Excel
                        </button>
                        <button id="btnImprimir" class="btn btn-info btn-export">
                            <i class="fas fa-print"></i> Imprimir
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tabela -->
        <div class="table-container">
            <table id="tabelaLogs" class="table table-striped table-hover" style="width:100%">
                <thead>
                    <tr>
                        <th>Data/Hora</th>
                        <th>Usuário</th>
                        <th>Tabela</th>
                        <th>Ação</th>
                        <th>Descrição</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    ' Conexão com o banco e consulta dos logs
                    Set conn = Server.CreateObject("ADODB.Connection")
                    conn.Open StrConnSales
                    
                    sql = "SELECT * FROM log_operations ORDER BY DataHora DESC"
                    
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    rs.CursorLocation = 3 ' adUseClient
                    rs.Open sql, conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
                    
                    Do While Not rs.EOF
                        dataHora = rs("DataHora")
                        usuario = rs("Usuario")
                        tabela = rs("TabelaAfetada")
                        acao = rs("Acao")
                        descricao = rs("Descricao")
                        
                        ' Formatar data/hora
                        If Not IsNull(dataHora) Then
                            dataFormatada = FormatDateTime(dataHora, vbLongDate) & " " & FormatDateTime(dataHora, vbLongTime)
                        Else
                            dataFormatada = "N/A"
                        End If
                    %>
                    <tr>
                        <td><%=dataFormatada%></td>
                        <td><span class="badge bg-primary"><%=Server.HTMLEncode(usuario)%></span></td>
                        <td><%=Server.HTMLEncode(tabela)%></td>
                        <td>
                            <%
                            Select Case acao
                                Case "INSERT"
                                    badgeClass = "badge-success"
                                Case "UPDATE"
                                    badgeClass = "badge-warning"
                                Case "DELETE"
                                    badgeClass = "badge-danger"
                                Case "LOGIN", "LOGOUT"
                                    badgeClass = "badge-info"
                                Case "ERRO"
                                    badgeClass = "badge-danger"
                                Case Else
                                    badgeClass = "badge-secondary"
                            End Select
                            %>
                            <span class="badge <%=badgeClass%>"><%=Server.HTMLEncode(acao)%></span>
                        </td>
                        <td><%=Server.HTMLEncode(descricao)%></td>
                    </tr>
                    <%
                        rs.MoveNext
                    Loop
                    
                    If rs.State = 1 Then rs.Close
                    Set rs = Nothing
                    If conn.State = 1 Then conn.Close
                    Set conn = Nothing
                    %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
    
    <!-- DataTables JS -->
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.bootstrap5.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.print.min.js"></script>

    <script>
    $(document).ready(function() {
        // Inicializar DataTable
        var table = $('#tabelaLogs').DataTable({
            language: {
                url: '//cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json'
            },
            dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6"f>>rt<"row"<"col-sm-12 col-md-6"i><"col-sm-12 col-md-6"p>>',
            pageLength: 25,
            order: [[0, 'desc']],
            responsive: true,
            columnDefs: [
                { responsivePriority: 1, targets: 0 }, // Data/Hora
                { responsivePriority: 2, targets: 4 }, // Descrição
                { responsivePriority: 3, targets: 3 }, // Ação
                { responsivePriority: 4, targets: 2 }, // Tabela
                { responsivePriority: 5, targets: 1 }  // Usuário
            ]
        });

        // Preencher filtros dinamicamente
        function popularFiltros() {
            // Limpar filtros existentes
            $('#filterTabela').find('option:not(:first)').remove();
            $('#filterUsuario').find('option:not(:first)').remove();
            
            // Popular filtro de tabelas
            var tabelas = [];
            table.column(2).data().each(function(value, index) {
                if (value && tabelas.indexOf(value) === -1) {
                    tabelas.push(value);
                    $('#filterTabela').append('<option value="' + value + '">' + value + '</option>');
                }
            });
            
            // Popular filtro de usuários
            var usuarios = [];
            table.column(1).data().each(function(value, index) {
                if (value && usuarios.indexOf(value) === -1) {
                    usuarios.push(value);
                    $('#filterUsuario').append('<option value="' + value + '">' + value + '</option>');
                }
            });
        }

        // Chamar após a inicialização da tabela
        setTimeout(popularFiltros, 1000);

        // Aplicar filtros
        $('#filterAcao').on('change', function() {
            table.column(3).search(this.value).draw();
        });

        $('#filterTabela').on('change', function() {
            table.column(2).search(this.value).draw();
        });

        $('#filterUsuario').on('change', function() {
            table.column(1).search(this.value).draw();
        });

        $('#filterData').on('change', function() {
            if (this.value) {
                table.column(0).search(this.value).draw();
            }
        });

        // Limpar filtros
        $('#btnLimparFiltros').on('click', function() {
            $('#filterAcao').val('');
            $('#filterTabela').val('');
            $('#filterUsuario').val('');
            $('#filterData').val('');
            table.search('').columns().search('').draw();
        });

        // Exportar para Excel
        $('#btnExportarExcel').on('click', function() {
            var excelButton = $('.buttons-excel');
            if (excelButton.length) {
                excelButton.click();
            } else {
                // Fallback - criar botão de exportação temporário
                table.button().add(0, {
                    extend: 'excel',
                    text: 'Excel',
                    className: 'btn btn-success'
                });
                table.button(0).trigger();
            }
        });

        // Imprimir
        $('#btnImprimir').on('click', function() {
            var printButton = $('.buttons-print');
            if (printButton.length) {
                printButton.click();
            } else {
                table.button().add(1, {
                    extend: 'print',
                    text: 'Imprimir',
                    className: 'btn btn-info'
                });
                table.button(1).trigger();
            }
        });
    });
    </script>
</body>
</html>