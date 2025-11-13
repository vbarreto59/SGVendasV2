<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="registra_log.asp"-->


<%
' Verificar se o usuário está logado
If Session("Usuario") = "" Then
    Response.Redirect "login.asp"
End If

' -----------------------------------------------------------------------------------
' INICIALIZAÇÃO E CONEXÃO COM BANCOS DE DADOS
' -----------------------------------------------------------------------------------
' Verifica se as strings de conexão estão configuradas.
If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "Erro: Conexões com bancos de dados não configuradas"
    Response.End
End If

' Cria e abre as conexões com os bancos de dados.
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' #################### Processar alterações se for POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim vendaId, novaDiretoriaId, novaGerenciaId, novoCorretorId
    Dim userIdDiretoria, userIdGerencia
    Dim mensagem
    
    vendaId = Request.Form("vendaId")
    novaDiretoriaId = Request.Form("diretoriaId")
    novaGerenciaId = Request.Form("gerenciaId")
    novoCorretorId = Request.Form("corretorId")
    userIdDiretoria = Request.Form("userIdDiretoria")
    userIdGerencia = Request.Form("userIdGerencia")
    
    If vendaId <> "" And IsNumeric(vendaId) Then
        ' Buscar os novos nomes
        Dim novaDiretoriaNome, novaGerenciaNome, novoCorretorNome
        Dim nomeDiretor, nomeGerente
        
        ' Nomes das estruturas organizacionais
        novaDiretoriaNome = GetDataFromDB(conn, "Diretorias", "NomeDiretoria", "DiretoriaID", novaDiretoriaId)
        
        If novaGerenciaId <> "" And IsNumeric(novaGerenciaId) Then
            novaGerenciaNome = GetDataFromDB(conn, "Gerencias", "NomeGerencia", "GerenciaID", novaGerenciaId)
        Else
            novaGerenciaNome = "Não aplicável"
            novaGerenciaId = 0
        End If
        
        ' Nomes das pessoas (usuários)
        If userIdDiretoria <> "" And IsNumeric(userIdDiretoria) Then
            nomeDiretor = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", userIdDiretoria)
        Else
            nomeDiretor = "Não definido"
            userIdDiretoria = 0
        End If
        
        If userIdGerencia <> "" And IsNumeric(userIdGerencia) Then
            nomeGerente = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", userIdGerencia)
        Else
            nomeGerente = "Não definido"
            userIdGerencia = 0
        End If
        
        novoCorretorNome = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", novoCorretorId)
        
        ' Atualizar a venda
        sqlUpdate = "UPDATE Vendas SET " & _
                   "DiretoriaId = " & novaDiretoriaId & ", " & _
                   "Diretoria = '" & SanitizeSQL(novaDiretoriaNome) & "', " & _
                   "UserIdDiretoria = " & userIdDiretoria & ", " & _
                   "NomeDiretor = '" & SanitizeSQL(nomeDiretor) & "', " & _
                   "GerenciaId = " & novaGerenciaId & ", " & _
                   "Gerencia = '" & SanitizeSQL(novaGerenciaNome) & "', " & _
                   "UserIdGerencia = " & userIdGerencia & ", " & _
                   "NomeGerente = '" & SanitizeSQL(nomeGerente) & "', " & _
                   "CorretorId = " & novoCorretorId & ", " & _
                   "Corretor = '" & SanitizeSQL(novoCorretorNome) & "', " & _
                   "Usuario = '" & SanitizeSQL(Session("Usuario")) & "' " & _
                   "WHERE ID = " & vendaId
        
        On Error Resume Next
        connSales.Execute(sqlUpdate)

'==========================Enviar email======================='

' ##################### PREPARAR DADOS PARA O EMAIL
Dim alteracoesDetalhes
alteracoesDetalhes = ""

' Montar detalhes das pessoas atuais
alteracoesDetalhes = alteracoesDetalhes & "PESSOAS DEFINIDAS NA VENDA:" & vbCrLf & vbCrLf
alteracoesDetalhes = alteracoesDetalhes & "DIRETORIA: " & nomeDiretor & vbCrLf
alteracoesDetalhes = alteracoesDetalhes & "GERÊNCIA: " & nomeGerente & vbCrLf
alteracoesDetalhes = alteracoesDetalhes & "CORRETOR: " & novoCorretorNome & vbCrLf

' ##################### ENVIAR EMAIL COM TODAS AS INFORMAÇÕES
If (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") Then
    
    Set objMail = Server.CreateObject("CDONTS.NewMail")
    objMail.From = "sendmail@gabnetweb.com.br"
    objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    
    objMail.Subject = "SV-ALT. PESSOAS - " & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    
    objMail.MailFormat = 0 ' Texto Simples
    
    Dim emailBody
    emailBody = "ALTERAÇÃO DE PESSOAS NA VENDA - SISTEMA DE VENDAS" & vbCrLf & vbCrLf
    emailBody = emailBody & "Venda ID: " & vendaId & vbCrLf
    emailBody = emailBody & "Usuário: " & Ucase(Session("Usuario")) & vbCrLf
    emailBody = emailBody & "IP: " & request.serverVariables("REMOTE_ADDR") & vbCrLf
    emailBody = emailBody & "Data/Hora: " & Now() & vbCrLf & vbCrLf
    emailBody = emailBody & alteracoesDetalhes & vbCrLf
    emailBody = emailBody & "=====================================" & vbCrLf
    emailBody = emailBody & "Este é um email automático do sistema de gestão de vendas." & vbCrLf
    
    objMail.Body = emailBody
    
    objMail.Send
    Set objMail = Nothing

End If
'============================================================='




' ##################### ATUALIZAR TABELA PAGAMENTOS_COMISSOES
' Atualizar cada tipo de recebedor individualmente

' ###### Atualizar Diretoria
If userIdDiretoria <> "" And IsNumeric(userIdDiretoria) And userIdDiretoria <> "0" Then
    sqlUpdateDiretoria = "UPDATE PAGAMENTOS_COMISSOES SET " & _
                        "UsuariosUserId = " & userIdDiretoria & ", " & _
                        "UsuariosNome = '" & SanitizeSQL(nomeDiretor) & "' " & _
                        "WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'diretoria'"
    connSales.Execute(sqlUpdateDiretoria)
Else
    ' Se não tem diretoria, remove o registro se existir
    'sqlDeleteDiretoria = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'diretoria'"
    'connSales.Execute(sqlDeleteDiretoria)
End If

' ###### Atualizar Gerência
If userIdGerencia <> "" And IsNumeric(userIdGerencia) And userIdGerencia <> "0" Then
    sqlUpdateGerencia = "UPDATE PAGAMENTOS_COMISSOES SET " & _
                       "UsuariosUserId = " & userIdGerencia & ", " & _
                       "UsuariosNome = '" & SanitizeSQL(nomeGerente) & "' " & _
                       "WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'gerencia'"
    connSales.Execute(sqlUpdateGerencia)
Else
    ' Se não tem gerência, remove o registro se existir
   '' sqlDeleteGerencia = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'gerencia'"
    'connSales.Execute(sqlDeleteGerencia)
End If

' Atualizar Corretor
If novoCorretorId <> "" And IsNumeric(novoCorretorId) Then
    sqlUpdateCorretor = "UPDATE PAGAMENTOS_COMISSOES SET " & _
                       "UsuariosUserId = " & novoCorretorId & ", " & _
                       "UsuariosNome = '" & SanitizeSQL(novoCorretorNome) & "' " & _
                       "WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'corretor'"
    connSales.Execute(sqlUpdateCorretor)
Else
    ' Se não tem corretor, remove o registro se existir
    'sqlDeleteCorretor = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & vendaId & " AND TipoRecebedor = 'corretor'"
    'connSales.Execute(sqlDeleteCorretor)
End If
' #####################################################

        
        If Err.Number <> 0 Then
            mensagem = "Erro ao atualizar venda: " & Err.Description
        Else
            ' Atualizar a tabela COMISSOES_A_PAGAR
            sqlUpdateComissoes = "UPDATE COMISSOES_A_PAGAR SET " & _
                               "UserIdDiretoria = " & userIdDiretoria & ", " & _
                               "NomeDiretor = '" & SanitizeSQL(nomeDiretor) & "', " & _
                               "UserIdGerencia = " & userIdGerencia & ", " & _
                               "NomeGerente = '" & SanitizeSQL(nomeGerente) & "', " & _
                               "UserIdCorretor = " & novoCorretorId & ", " & _
                               "NomeCorretor = '" & SanitizeSQL(novoCorretorNome) & "', " & _
                               "Usuario = '" & SanitizeSQL(Session("Usuario")) & "' " & _
                               "WHERE ID_Venda = " & vendaId
            
            connSales.Execute(sqlUpdateComissoes)
            
            If Err.Number <> 0 Then
                mensagem = "Venda atualizada, mas erro nas comissões: " & Err.Description
            Else
                ' Registrar log
                Call InserirLog("VENDAS", "UPDATE", "Pessoas atualizadas na venda ID: " & vendaId)
                mensagem = "Pessoas atualizadas com sucesso!"
            End If
        End If
        On Error GoTo 0
    Else
        mensagem = "ID da venda inválido!"
    End If
End If
' ################################################################
' Buscar todas as vendas
Set rsVendas = connSales.Execute("SELECT ID, NomeEmpreendimento, Unidade, NomeCliente, DataVenda, " & _
                                "DiretoriaId, Diretoria, UserIdDiretoria, NomeDiretor, " & _
                                "GerenciaId, Gerencia, UserIdGerencia, NomeGerente, " & _
                                "CorretorId, Corretor " & _
                                "FROM Vendas WHERE EXCLUIDO=0 ORDER BY DataVenda DESC ")

' Buscar dados para os dropdowns - CRIANDO RECORDSETS SEPARADOS PARA CADA SELECT
Set rsDiretorias = conn.Execute("SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria")
Set rsGerencias = conn.Execute("SELECT GerenciaID, NomeGerencia FROM Gerencias ORDER BY NomeGerencia")

' RecordSets SEPARADOS para cada select de usuários
Set rsUsuariosDiretoria = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Nome <> '' ORDER BY Nome")
Set rsUsuariosGerencia = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Nome <> '' ORDER BY Nome")
Set rsCorretores = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Funcao = 'Corretor' AND Nome <> '' ORDER BY Nome")

' Função auxiliar
Function GetDataFromDB(oConn, sTable, sField, sWhereField, sWhereValue)
    Dim sResult
    On Error Resume Next
    Set rs = oConn.Execute("SELECT " & sField & " FROM " & sTable & " WHERE " & sWhereField & " = " & sWhereValue)
    If Err.Number = 0 And Not rs.EOF Then
        sResult = rs(sField)
    Else
        sResult = "Desconhecido"
    End If
    If IsObject(rs) Then rs.Close
    Set rs = Nothing
    GetDataFromDB = sResult
    On Error GoTo 0
End Function

Function SanitizeSQL(sValue)
    If IsNull(sValue) Then
        SanitizeSQL = ""
    Else
        SanitizeSQL = Replace(sValue, "'", "''")
    End If
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Pessoas nas Vendas | Sistema</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --success: #27ae60;
            --warning: #f39c12;
            --danger: #e74c3c;
        }
        
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #2c3e50;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 20px;
        }
        
        .app-container {
            max-width: 1800px;
            margin: 0 auto;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1.5rem;
            border-radius: 12px 12px 0 0;
            margin-bottom: 0;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .app-title {
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.8rem;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
            background: rgba(255, 255, 255, 0.95);
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1.2rem 1.5rem;
            font-weight: 600;
        }
        
        .table th {
            background-color: var(--primary);
            color: white;
            border: none;
        }
        
        .btn-sm {
            border-radius: 6px;
            font-weight: 600;
        }
        
        .modal-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
        }
        
        .form-label {
            font-weight: 600;
            color: var(--primary);
        }
        
        .info-badge {
            background: linear-gradient(135deg, var(--secondary), #2980b9);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        
        .alert {
            border-radius: 8px;
            border: none;
        }
        
        .section-title {
            color: var(--primary);
            font-weight: 600;
            font-size: 1.1rem;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #e9ecef;
        }
    </style>
</head>
<body>
    <div class="app-container">
        <!-- Header -->
        <div class="app-header">
            <div class="d-flex justify-content-between align-items-center">
                <h1 class="app-title">
                    <i class="fas fa-users-cog me-2"></i>Gestão de Pessoas nas Vendas
                </h1>
                <div class="info-badge">
                    <i class="fas fa-user me-1"></i><%= Session("Usuario") %>
                </div>
            </div>
        </div>

        <!-- Conteúdo Principal -->
        <div class="card">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <button type="button" onclick="window.close();" class="btn btn-secondary">
                        <i class="fas fa-arrow-left me-2"></i>Voltar
                    </button>
                    <div class="d-flex gap-2">
                        <a href="gestao_vendas_list2x.asp" class="btn btn-secondary">
                            <i class="fas fa-list me-2"></i>Lista de Vendas
                        </a>
                        <a href="gestao_vendas.asp" class="btn btn-primary">
                            <i class="fas fa-plus me-2"></i>Nova Venda
                        </a>
                    </div>
                </div>

                <% If mensagem <> "" Then %>
                <div class="alert alert-success alert-dismissible fade show" role="alert">
                    <i class="fas fa-check-circle me-2"></i><%= mensagem %>
                    <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                </div>
                <% End If %>

                <!-- Tabela de Vendas -->
                <div class="table-responsive">
                  <table id="tabelaVendas" class="table table-striped table-hover" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Empreendimento</th>
                                <th>Unidade</th>
                                <th>Cliente</th>
                                <th>Data Venda</th>
                                <th>Diretoria</th>
                                <th>Pessoa Diretoria</th>
                                <th>Gerência</th>
                                <th>Pessoa Gerência</th>
                                <th>Corretor</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            If Not rsVendas.EOF Then
                                rsVendas.MoveFirst
                                Do While Not rsVendas.EOF 
                            %>
                            <tr>
                                <td><strong><%= rsVendas("ID") %></strong></td>
                                <td><%= rsVendas("NomeEmpreendimento") %></td>
                                <td><%= rsVendas("Unidade") %></td>
                                <td><%= rsVendas("NomeCliente") %></td>
                                <td><%= FormatDateTime(rsVendas("DataVenda"), 2) %></td>
                                <td><%= rsVendas("Diretoria") %></td>
                                <td><%= rsVendas("NomeDiretor") %></td>
                                <td><%= rsVendas("Gerencia") %></td>
                                <td><%= rsVendas("NomeGerente") %></td>
                                <td><%= rsVendas("Corretor") %></td>
                                <td>
                                    <button type="button" class="btn btn-warning btn-sm" 
                                            onclick="abrirModal(<%= rsVendas("ID") %>, 
                                            '<%= rsVendas("DiretoriaId") %>', 
                                            '<%= rsVendas("UserIdDiretoria") %>', 
                                            '<%= rsVendas("GerenciaId") %>', 
                                            '<%= rsVendas("UserIdGerencia") %>', 
                                            '<%= rsVendas("CorretorId") %>')">
                                        <i class="fas fa-edit me-1"></i>Alterar
                                    </button>
                                </td>
                            </tr>
                            <%
                                    rsVendas.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="11" class="text-center">Nenhuma venda encontrada</td>
                            </tr>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal para Alteração -->
    <div class="modal fade" id="modalAlteracao" tabindex="-1">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="fas fa-edit me-2"></i>Alterar Pessoas da Venda
                    </h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                </div>
                <form method="post" id="formAlteracao">
                    <div class="modal-body">
                        <input type="hidden" id="vendaId" name="vendaId">
                        
                        <!-- Seção Diretoria -->
                        <div class="row mb-4">
                            <div class="col-12">
                                <h6 class="section-title">
                                    <i class="fas fa-user-tie me-2"></i>Diretoria
                                </h6>
                            </div>
                            <div class="col-md-6">
                                <label for="diretoriaId" class="form-label">Nome da Diretoria</label>
                                <select class="form-select" id="diretoriaId" name="diretoriaId" required>
                                    <option value="">Selecione a diretoria...</option>
                                    <% 
                                    If Not rsDiretorias.EOF Then
                                        rsDiretorias.MoveFirst
                                        Do While Not rsDiretorias.EOF 
                                    %>
                                        <option value="<%= rsDiretorias("DiretoriaID") %>"><%= rsDiretorias("NomeDiretoria") %></option>
                                    <%
                                            rsDiretorias.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                            <div class="col-md-6">
                                <label for="userIdDiretoria" class="form-label">Pessoa da Diretoria</label>
                                <select class="form-select" id="userIdDiretoria" name="userIdDiretoria">
                                    <option value="">Selecione a pessoa...</option>
                                    <% 
                                    If Not rsUsuariosDiretoria.EOF Then
                                        rsUsuariosDiretoria.MoveFirst
                                        Do While Not rsUsuariosDiretoria.EOF 
                                    %>
                                        <option value="<%= rsUsuariosDiretoria("UserId") %>"><%= rsUsuariosDiretoria("Nome") %></option>
                                    <%
                                            rsUsuariosDiretoria.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                        
                        <!-- Seção Gerência -->
                        <div class="row mb-4">
                            <div class="col-12">
                                <h6 class="section-title">
                                    <i class="fas fa-user-tie me-2"></i>Gerência
                                </h6>
                            </div>
                            <div class="col-md-6">
                                <label for="gerenciaId" class="form-label">Nome da Gerência</label>
                                <select class="form-select" id="gerenciaId" name="gerenciaId">
                                    <option value="">Selecione a gerência...</option>
                                    <% 
                                    If Not rsGerencias.EOF Then
                                        rsGerencias.MoveFirst
                                        Do While Not rsGerencias.EOF 
                                    %>
                                        <option value="<%= rsGerencias("GerenciaID") %>"><%= rsGerencias("NomeGerencia") %></option>
                                    <%
                                            rsGerencias.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                            <div class="col-md-6">
                                <label for="userIdGerencia" class="form-label">Pessoa da Gerência</label>
                                <select class="form-select" id="userIdGerencia" name="userIdGerencia">
                                    <option value="">Selecione a pessoa...</option>
                                    <% 
                                    If Not rsUsuariosGerencia.EOF Then
                                        rsUsuariosGerencia.MoveFirst
                                        Do While Not rsUsuariosGerencia.EOF 
                                    %>
                                        <option value="<%= rsUsuariosGerencia("UserId") %>"><%= rsUsuariosGerencia("Nome") %></option>
                                    <%
                                            rsUsuariosGerencia.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                        
                        <!-- Seção Corretor -->
                        <div class="row mb-4">
                            <div class="col-12">
                                <h6 class="section-title">
                                    <i class="fas fa-user me-2"></i>Corretor
                                </h6>
                            </div>
                            <div class="col-md-12">
                                <label for="corretorId" class="form-label">Corretor</label>
                                <select class="form-select" id="corretorId" name="corretorId" required>
                                    <option value="">Selecione o corretor...</option>
                                    <% 
                                    If Not rsCorretores.EOF Then
                                        rsCorretores.MoveFirst
                                        Do While Not rsCorretores.EOF 
                                    %>
                                        <option value="<%= rsCorretores("UserId") %>"><%= rsCorretores("Nome") %></option>
                                    <%
                                            rsCorretores.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                        
                        <div class="alert alert-info mt-3">
                            <i class="fas fa-info-circle me-2"></i>
                            <strong>Informação:</strong> As alterações serão aplicadas tanto na venda quanto nas comissões a pagar.
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="submit" class="btn btn-success">
                            <i class="fas fa-save me-2"></i>Salvar Alterações
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>


    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    
    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- DataTables JS -->
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.4.1/js/dataTables.responsive.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.4.1/js/responsive.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.print.min.js"></script>
    



<script>
        // Inicializar DataTable
        $(document).ready(function() {
            $('#tabelaVendas').DataTable({
                language: {
                    url: '//cdn.datatables.net/plug-ins/1.13.4/i18n/pt-BR.json'
                },
                responsive: true,
                dom: '<"row"<"col-md-6"B><"col-md-6"f>>rtip',
                buttons: [
                    {
                        extend: 'copy',
                        text: '<i class="fas fa-copy"></i> Copiar',
                        className: 'btn btn-secondary'
                    },
                    {
                        extend: 'excel',
                        text: '<i class="fas fa-file-excel"></i> Excel',
                        className: 'btn btn-success'
                    },
                    {
                        extend: 'pdf',
                        text: '<i class="fas fa-file-pdf"></i> PDF',
                        className: 'btn btn-danger'
                    },
                    {
                        extend: 'print',
                        text: '<i class="fas fa-print"></i> Imprimir',
                        className: 'btn btn-info'
                    }
                ],
                pageLength: 25,
                order: [[0, 'desc']], // Ordenar pela coluna ID decrescente
                columnDefs: [
                    {
                        targets: [10], // Coluna de ações
                        orderable: false,
                        searchable: false
                    }
                ]
            });
        });
</script>

    
    <script>
        function abrirModal(vendaId, diretoriaId, userIdDiretoria, gerenciaId, userIdGerencia, corretorId) {
            document.getElementById('vendaId').value = vendaId;
            document.getElementById('diretoriaId').value = diretoriaId;
            document.getElementById('userIdDiretoria').value = userIdDiretoria;
            document.getElementById('gerenciaId').value = gerenciaId;
            document.getElementById('userIdGerencia').value = userIdGerencia;
            document.getElementById('corretorId').value = corretorId;
            
            var modal = new bootstrap.Modal(document.getElementById('modalAlteracao'));
            modal.show();
        }
        
        // Fechar modal após sucesso
        <% If mensagem <> "" Then %>
            var modal = bootstrap.Modal.getInstance(document.getElementById('modalAlteracao'));
            if (modal) modal.hide();
        <% End If %>
    </script>
</body>
</html>

<%
' Fechar conexões
If IsObject(rsVendas) Then
    rsVendas.Close
    Set rsVendas = Nothing
End If

If IsObject(rsDiretorias) Then
    rsDiretorias.Close
    Set rsDiretorias = Nothing
End If

If IsObject(rsGerencias) Then
    rsGerencias.Close
    Set rsGerencias = Nothing
End If

If IsObject(rsUsuariosDiretoria) Then
    rsUsuariosDiretoria.Close
    Set rsUsuariosDiretoria = Nothing
End If

If IsObject(rsUsuariosGerencia) Then
    rsUsuariosGerencia.Close
    Set rsUsuariosGerencia = Nothing
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