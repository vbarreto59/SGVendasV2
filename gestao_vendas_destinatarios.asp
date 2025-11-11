<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"
%>

<%
' Verifica se foi passado o ID da venda
Dim vendaId
vendaId = Request.QueryString("id")
If Not IsNumeric(vendaId) Or vendaId = "" Then
    Response.Redirect "gestao_vendas_list3x.asp"
End If

' Cria as conexões
Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Busca os dados da venda para preencher o formulário
Dim rsVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
rsVenda.Open "SELECT * FROM Vendas WHERE ID = " & vendaId, connSales

If rsVenda.EOF Then
    Response.Redirect "gestao_vendas_list3x.asp"
End If

' ====================================================================
' PROCESSAMENTO DO FORMULÁRIO (POST)
' ====================================================================

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, diretoriaId, gerenciaId
    Dim userIdDiretoria, userIdGerencia, corretorId
    
    action = Request.Form("action")
    diretoriaId = Request.Form("diretoriaId")
    gerenciaId = Request.Form("gerenciaId")
    userIdDiretoria = Request.Form("userIdDiretoria")
    userIdGerencia = Request.Form("userIdGerencia")
    corretorId = Request.Form("corretorId")
    
    If action = "updateDiretoriaGerencia" Then
        ' Busca os nomes da diretoria, gerência e usuários
        Dim nomeDiretoria, nomeGerencia, nomeDiretor, nomeGerente, nomeCorretor
        
        ' Busca nome da diretoria
        Set rsNomes = Server.CreateObject("ADODB.Recordset")
        rsNomes.Open "SELECT NomeDiretoria FROM Diretorias WHERE DiretoriaID = " & CInt(diretoriaId), conn
        If Not rsNomes.EOF Then
            nomeDiretoria = rsNomes("NomeDiretoria")
        Else
            nomeDiretoria = ""
        End If
        rsNomes.Close
        
        ' Busca nome da gerência
        rsNomes.Open "SELECT NomeGerencia FROM Gerencias WHERE GerenciaID = " & CInt(gerenciaId), conn
        If Not rsNomes.EOF Then
            nomeGerencia = rsNomes("NomeGerencia")
        Else
            nomeGerencia = ""
        End If
        rsNomes.Close
        
        ' Busca nome do diretor
        rsNomes.Open "SELECT Nome FROM Usuarios WHERE UserId = " & CInt(userIdDiretoria), conn
        If Not rsNomes.EOF Then
            nomeDiretor = rsNomes("Nome")
        Else
            nomeDiretor = ""
        End If
        rsNomes.Close
        
        ' Busca nome do gerente
        rsNomes.Open "SELECT Nome FROM Usuarios WHERE UserId = " & CInt(userIdGerencia), conn
        If Not rsNomes.EOF Then
            nomeGerente = rsNomes("Nome")
        Else
            nomeGerente = ""
        End If
        rsNomes.Close
        
        ' Busca nome do corretor
        rsNomes.Open "SELECT Nome FROM Usuarios WHERE UserId = " & CInt(corretorId), conn
        If Not rsNomes.EOF Then
            nomeCorretor = rsNomes("Nome")
        Else
            nomeCorretor = ""
        End If
        rsNomes.Close
        Set rsNomes = Nothing
        
        ' Atualiza a venda com os novos dados
        sql = "UPDATE Vendas SET " & _
              "DiretoriaId = " & CInt(diretoriaId) & ", " & _
              "Diretoria = '" & Replace(nomeDiretoria, "'", "''") & "', " & _
              "GerenciaId = " & CInt(gerenciaId) & ", " & _
              "Gerencia = '" & Replace(nomeGerencia, "'", "''") & "', " & _
              "UserIdDiretoria = " & CInt(userIdDiretoria) & ", " & _
              "NomeDiretor = '" & Replace(nomeDiretor, "'", "''") & "', " & _
              "UserIdGerencia = " & CInt(userIdGerencia) & ", " & _
              "NomeGerente = '" & Replace(nomeGerente, "'", "''") & "', " & _
              "CorretorId = " & CInt(corretorId) & ", " & _
              "Corretor = '" & Replace(nomeCorretor, "'", "''") & "', " & _
              "Usuario = '" & Session("Usuario") & "' " & _
              "WHERE ID = " & CInt(vendaId)
        
        On Error Resume Next
        connSales.Execute(sql)
        
        If Err.Number = 0 Then
            Response.Redirect "gestao_vendas_list3x.asp?mensagem=Diretoria, gerência e usuários atualizados com sucesso!"
        Else
            Response.Write "<script>alert('Erro ao atualizar: " & Replace(Err.Description, "'", "\'") & "');</script>"
        End If
    End If
End If

' ====================================================================
' BUSCA DE DADOS PARA DROPDOWNS
' ====================================================================

' Busca diretorias para o dropdown
Dim rsDiretorias
Set rsDiretorias = Server.CreateObject("ADODB.Recordset")
rsDiretorias.Open "SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria", conn

' Busca gerencias para o dropdown (todas inicialmente)
Dim rsGerencias
Set rsGerencias = Server.CreateObject("ADODB.Recordset")
rsGerencias.Open "SELECT GerenciaID, NomeGerencia, DiretoriaID FROM Gerencias ORDER BY NomeGerencia", conn

' Busca todos os usuários para os selects
Dim rsUsuarios
Set rsUsuarios = Server.CreateObject("ADODB.Recordset")
rsUsuarios.Open "SELECT UserId, Nome, Funcao FROM Usuarios WHERE Nome <> '' ORDER BY Nome", conn

' Busca corretores (usuários com função Corretor)
Dim rsCorretores
Set rsCorretores = Server.CreateObject("ADODB.Recordset")
rsCorretores.Open "SELECT UserId, Nome FROM Usuarios WHERE Funcao = 'Corretor' AND Nome <> '' ORDER BY Nome", conn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atualizar Diretoria e Gerência</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <!-- Select2 para selects com busca -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    
    <style>
        body {
            background-color: #807777;
            color: #fff;
            padding: 20px;
        }
        .card {
            background-color: #fff;
            color: #000;
            margin-bottom: 20px;
            border: none;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .card-header {
            background-color: #800000;
            color: white;
            font-weight: bold;
            border-radius: 10px 10px 0 0 !important;
        }
        .btn-maroon {
            background-color: #800000;
            color: white;
            border: none;
        }
        .btn-maroon:hover {
            background-color: #a00;
            color: white;
        }
        .current-data {
            background-color: #41464A;
            border-left: 4px solid #800000;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 5px;
        }
        .form-section {
            margin-bottom: 30px;
        }
        .form-section h5 {
            color: #800000;
            border-bottom: 2px solid #800000;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        
        /* Estilos para Select2 */
        .select2-container--default .select2-selection--single,
        .select2-container--default .select2-selection--multiple {
            background-color: #fff;
            color: #000;
            border: 1px solid #ced4da;
            height: 38px;
        }
        .select2-container--default .select2-selection--single .select2-selection__rendered {
            color: #000;
            line-height: 36px;
        }
        .select2-container--default .select2-selection--single .select2-selection__arrow {
            height: 36px;
        }
        .select2-dropdown {
            background-color: #fff;
            color: #000;
        }
        .select2-container--default .select2-results__option[aria-selected=true] {
            background-color: #f8f9fa;
            color: #000;
        }
        .select2-container--default .select2-results__option--highlighted[aria-selected] {
            background-color: #800000;
            color: #fff;
        }
    </style>
</head>
<body>
    <div class="container" style="padding-top: 70px;">
        <button type="button" onclick="window.close();" class="btn btn-success mb-3">
            <i class="fas fa-times me-2"></i>Fechar
        </button>
        
        <h2 class="mt-4 mb-4 text-white">
            <i class="fas fa-users me-2"></i>Atualizar Diretoria e Gerência - Venda ID: <%=vendaId%>
        </h2>
        
        <!-- Dados Atuais -->
        <div class="current-data">
            <h5 class="text-maroon"><i class="fas fa-info-circle me-2"></i>Dados Atuais da Venda</h5>
            <div class="row">
                <div class="col-md-4">
                    <strong>Diretoria:</strong> <%=rsVenda("Diretoria")%><br>
                    <strong>Gerência:</strong> <%=rsVenda("Gerencia")%>
                </div>
                <div class="col-md-4">
                    <strong>Diretor:</strong> <%=rsVenda("NomeDiretor")%><br>
                    <strong>Gerente:</strong> <%=rsVenda("NomeGerente")%>
                </div>
                <div class="col-md-4">
                    <strong>Corretor:</strong> <%=rsVenda("Corretor")%>
                </div>
            </div>
        </div>

        <form method="post" id="formDiretoriaGerencia">
            <input type="hidden" name="vendaId" value="<%=vendaId%>">
            
            <!-- Card Diretoria e Gerência -->
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-building me-2"></i>Selecionar Diretoria e Gerência
                </div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label for="diretoriaId" class="form-label">Diretoria *</label>
                            <select class="form-select select2" id="diretoriaId" name="diretoriaId" required>
                                <option value="">Selecione a Diretoria...</option>
                                <%
                                If Not rsDiretorias.EOF Then
                                    rsDiretorias.MoveFirst
                                    Do While Not rsDiretorias.EOF
                                %>
                                    <option value="<%= rsDiretorias("DiretoriaID") %>"
                                        <% If rsDiretorias("DiretoriaID") = rsVenda("DiretoriaId") Then Response.Write "selected" %>>
                                        <%= rsDiretorias("NomeDiretoria") %>
                                    </option>
                                <%
                                    rsDiretorias.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label for="gerenciaId" class="form-label">Gerência *</label>
                            <select class="form-select select2" id="gerenciaId" name="gerenciaId" required>
                                <option value="">Selecione a Gerência...</option>
                                <%
                                If Not rsGerencias.EOF Then
                                    rsGerencias.MoveFirst
                                    Do While Not rsGerencias.EOF
                                %>
                                    <option value="<%= rsGerencias("GerenciaID") %>" 
                                        data-diretoria="<%= rsGerencias("DiretoriaID") %>"
                                        <% If rsGerencias("GerenciaID") = rsVenda("GerenciaId") Then Response.Write "selected" %>>
                                        <%= rsGerencias("NomeGerencia") %>
                                    </option>
                                <%
                                    rsGerencias.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Usuários da Diretoria -->
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-user-tie me-2"></i>Usuário da Diretoria
                </div>
                <div class="card-body">
                    <div class="mb-3">
                        <label for="userIdDiretoria" class="form-label">A Comissão da Diretoria vai para: *</label>
                        <select class="form-select select2" id="userIdDiretoria" name="userIdDiretoria" required>
                            <option value="">Selecione o usuário da diretoria...</option>
                            <%
                            If Not rsUsuarios.EOF Then
                                rsUsuarios.MoveFirst
                                Do While Not rsUsuarios.EOF
                            %>
                                <option value="<%= rsUsuarios("UserId") %>"
                                    <% If rsUsuarios("UserId") = rsVenda("UserIdDiretoria") Then Response.Write "selected" %>>
                                    <%= rsUsuarios("Nome") %> 
                                    <% If rsUsuarios("Funcao") <> "" Then Response.Write "(" & rsUsuarios("Funcao") & ")" %>
                                </option>
                            <%
                                rsUsuarios.MoveNext
                                Loop
                            End If
                            %>
                        </select>
                    </div>
                </div>
            </div>
            
            <!-- Card Usuários da Gerência -->
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-user-shield me-2"></i>Usuário da Gerência
                </div>
                <div class="card-body">
                    <div class="mb-3">
                        <label for="userIdGerencia" class="form-label">A Comissão da Gerência vai para: *</label>
                        <select class="form-select select2" id="userIdGerencia" name="userIdGerencia" required>
                            <option value="">Selecione o usuário da gerência...</option>
                            <%
                            ' Reposiciona o recordset para reutilizar
                            If Not rsUsuarios.BOF Then
                                rsUsuarios.MoveFirst
                                Do While Not rsUsuarios.EOF
                            %>
                                <option value="<%= rsUsuarios("UserId") %>"
                                    <% If rsUsuarios("UserId") = rsVenda("UserIdGerencia") Then Response.Write "selected" %>>
                                    <%= rsUsuarios("Nome") %> 
                                    <% If rsUsuarios("Funcao") <> "" Then Response.Write "(" & rsUsuarios("Funcao") & ")" %>
                                </option>
                            <%
                                rsUsuarios.MoveNext
                                Loop
                            End If
                            %>
                        </select>
                    </div>
                </div>
            </div>
            
            <!-- Card Corretor -->
            <div class="card">
                <div class="card-header">
                    <i class="fas fa-user me-2"></i>Corretor
                </div>
                <div class="card-body">
                    <div class="mb-3">
                        <label for="corretorId" class="form-label">A Comissão do Corretor vai para: *</label>
                        <select class="form-select select2" id="corretorId" name="corretorId" required>
                            <option value="">Selecione o corretor...</option>
                            <%
                            If Not rsCorretores.EOF Then
                                rsCorretores.MoveFirst
                                Do While Not rsCorretores.EOF
                            %>
                                <option value="<%= rsCorretores("UserId") %>"
                                    <% If rsCorretores("UserId") = rsVenda("CorretorId") Then Response.Write "selected" %>>
                                    <%= rsCorretores("Nome") %>
                                </option>
                            <%
                                rsCorretores.MoveNext
                                Loop
                            End If
                            %>
                        </select>
                    </div>
                </div>
            </div>
            
            <!-- Botões de Ação -->
            <div class="card">
                <div class="card-body">
                    <div class="d-grid gap-2 d-md-flex justify-content-md-end">
                        <button type="submit" name="action" value="updateDiretoriaGerencia" class="btn btn-maroon btn-lg">
                            <i class="fas fa-save me-2"></i>Atualizar Diretoria, Gerência e Usuários
                        </button>
                    </div>
                </div>
            </div>
        </form>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    
    <!-- Select2 -->
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/i18n/pt-BR.js"></script>
    
    <script>
        $(document).ready(function() {
            // Inicializa select2 nos selects
            $('.select2').select2({
                language: "pt-BR",
                placeholder: "Selecione...",
                allowClear: true,
                width: '100%'
            });
            
            // Filtra gerencias quando seleciona diretoria
            $('#diretoriaId').change(function() {
                var diretoriaId = $(this).val();
                var $gerenciaSelect = $('#gerenciaId');
                
                if (diretoriaId) {
                    // Habilita o select de gerência
                    $gerenciaSelect.prop('disabled', false);
                    
                    // Filtra as opções mostrando apenas as gerencias da diretoria selecionada
                    $gerenciaSelect.find('option').each(function() {
                        var $option = $(this);
                        if ($option.val() === '') {
                            $option.show(); // Mostra a opção vazia
                        } else {
                            var optionDiretoria = $option.data('diretoria');
                            if (optionDiretoria == diretoriaId) {
                                $option.show();
                            } else {
                                $option.hide();
                                if ($option.is(':selected')) {
                                    $option.prop('selected', false);
                                }
                            }
                        }
                    });
                    
                    // Atualiza o Select2
                    $gerenciaSelect.trigger('change.select2');
                } else {
                    // Se nenhuma diretoria selecionada, mostra todas as gerencias
                    $gerenciaSelect.find('option').show();
                    $gerenciaSelect.trigger('change.select2');
                }
            });
            
            // Dispara o change na carga inicial para filtrar corretamente
            $('#diretoriaId').trigger('change');
            
            // Validação do formulário
            $('#formDiretoriaGerencia').submit(function(e) {
                var isValid = true;
                var requiredFields = ['#diretoriaId', '#gerenciaId', '#userIdDiretoria', '#userIdGerencia', '#corretorId'];
                
                requiredFields.forEach(function(field) {
                    if (!$(field).val()) {
                        isValid = false;
                        $(field).addClass('is-invalid');
                    } else {
                        $(field).removeClass('is-invalid');
                    }
                });
                
                if (!isValid) {
                    e.preventDefault();
                    alert('Por favor, preencha todos os campos obrigatórios.');
                    return false;
                }
                
                return true;
            });
        });
    </script>
</body>
</html>
<%
' Fecha conexões
If IsObject(rsVenda) Then
    rsVenda.Close
    Set rsVenda = Nothing
End If

If IsObject(rsDiretorias) Then
    rsDiretorias.Close
    Set rsDiretorias = Nothing
End If

If IsObject(rsGerencias) Then
    rsGerencias.Close
    Set rsGerencias = Nothing
End If

If IsObject(rsUsuarios) Then
    rsUsuarios.Close
    Set rsUsuarios = Nothing
End If

If IsObject(rsCorretores) Then
    rsCorretores.Close
    Set rsCorretores = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>