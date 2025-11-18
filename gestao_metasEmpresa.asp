<%@ Language=VBScript %>

<% Response.Buffer = True %>

<!--#include file="conSunSales.asp"-->

<%
    if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
        On Error Resume Next 
        set objMail = server.createobject("CDONTS.NewMail")
        if Err.Number <> 0 then 
            set objMail = Nothing ' Garante que a variável seja liberada, mesmo que não criada
        else
            objMail.From = "sendmail@gabnetweb.com.br"
            objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
            objMail.Subject = "SV-RELMETAS-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
            objMail.MailFormat = 0 ' 0 = Texto Simples
            objMail.Body = "Página Relatório com Metas. " & Ucase(Session("Usuario"))
            objMail.Send
            set objMail = Nothing
        end if 
        On Error GoTo 0 
    end if
%>


<%
' Declarar variáveis
Dim acao, id, ano, mes, meta, mensagem, mensagemTipo
Dim rs, sql, connSales

' Inicializar variáveis
acao = Request.Form("acao")
id = Request.Form("id")
ano = Trim(Request.Form("ano"))
mes = Trim(Request.Form("mes"))
meta = Request.Form("meta")

meta = meta/1
' O Access (Jet) não suporta meta = meta/1 para conversão. 
' Vamos confiar no Replace para o formato correto na inserção/atualização.
mensagem = ""
mensagemTipo = ""

' Função para criar tabela se não existir
Sub CriarTabelaSeNaoExistir(conn)
    On Error Resume Next
    
    Dim sqlCheck, rsCheck
    
    ' 1. MUDANÇA: No Access/Jet, a forma mais fácil de verificar a existência da tabela 
    ' é tentando executar uma query simples e verificando se dá erro.
    ' O INFORMATION_SCHEMA.TABLES é específico do SQL Server.
    sqlCheck = "SELECT TOP 1 ID FROM MetaEmpresa" ' Query simples para verificar
    conn.Execute(sqlCheck)
    
    If Err.Number <> 0 Then
        ' Tabela não existe (ou deu erro na conexão/permissão), tentar criar
        Err.Clear ' Limpa o erro do SELECT
        
        Dim sqlCreate
        ' 2. MUDANÇA: Tipos de dados e sintaxe do Access/Jet SQL
        '   - IDENTITY(1,1) (SQL Server) torna-se COUNTER (Access) para chave primária autonumeração.
        '   - DECIMAL(18,2) torna-se CURRENCY (ideal para valores monetários) ou DOUBLE.
        '   - DATETIME DEFAULT GETDATE() torna-se DATETIME DEFAULT NOW()
        sqlCreate = "CREATE TABLE MetaEmpresa (" & _
                    "ID COUNTER PRIMARY KEY, " & _
                    "Ano INT NOT NULL, " & _
                    "Mes INT NOT NULL, " & _
                    "Meta CURRENCY NOT NULL, " & _
                    "DataCriacao DATETIME DEFAULT NOW())"
        
        conn.Execute(sqlCreate)
        
        ' 3. MUDANÇA: Criar índice único após a criação da tabela
        sqlCreate = "CREATE UNIQUE INDEX IX_MetaEmpresa_AnoMes ON MetaEmpresa (Ano, Mes)"
        conn.Execute(sqlCreate)
        
        If Err.Number = 0 Then
            Response.Write "<div class='message info'>Tabela MetaEmpresa criada automaticamente!</div>"
        End If
    End If
    
    If Err.Number <> 0 Then
        Response.Write "<div class='message erro'>Erro ao verificar/criar tabela: " & Err.Description & "</div>"
    End If
    
    ' Não precisamos de rsCheck, pois usamos conn.Execute.
    Err.Clear ' Limpa qualquer erro residual antes de continuar o código
End Sub

' Função para validar e converter valores (manter, embora o uso seja limitado no código)
Function ValidarNumero(valor, padrao)
    If IsNumeric(valor) Then
        ValidarNumero = valor
    Else
        ValidarNumero = padrao
    End If
End Function

' Processar ações do formulário
If acao <> "" Then
    On Error Resume Next
    
    ' Abrir conexão
    Set connSales = Server.CreateObject("ADODB.Connection")
    ' A string de conexão StrConnSales em conSunSales.asp deve ser um DSN ou 
    ' uma string OLEDB para MDB, por exemplo:
    ' StrConnSales = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\caminho\seu_banco.mdb;"
    connSales.Open StrConnSales
    
    If Err.Number <> 0 Then
        mensagem = "Erro ao conectar ao banco: " & Err.Description
        mensagemTipo = "erro"
    Else
        ' Criar tabela se necessário
        Call CriarTabelaSeNaoExistir(connSales)
        
        Select Case acao
            Case "cadastrar"
                If ano <> "" And mes <> "" And meta <> "" Then
                    ' Verificar se já existe registro para o mesmo ano/mês
                    ' Usar CLng para garantir que são números inteiros
                    sql = "SELECT ID FROM MetaEmpresa WHERE Ano = " & CLng(ano) & " AND Mes = " & CLng(mes)
                    Set rs = connSales.Execute(sql)
                    
                    If Not rs.EOF Then
                        mensagem = "Já existe uma meta cadastrada para este ano e mês!"
                        mensagemTipo = "erro"
                    Else
                        ' 4. MUDANÇA: Simplificar a formatação do valor para Access (Jet)
                        ' O código original tinha 2 Replace, que pode ser redundante/complicado.
                        ' Vamos garantir que o valor use PONTO como separador decimal, 
                        ' que é o padrão da maioria dos SGBDs (incluindo Jet/Access).
                        Dim metaFormatada
                        metaFormatada = Replace(meta, ",", ".") ' Troca vírgula por ponto
                        metaFormatada = CDbl(metaFormatada)     ' Converte para Double (se falhar, dá erro)
                        
                        sql = "INSERT INTO MetaEmpresa (Ano, Mes, Meta) VALUES (" & _
                              CLng(ano) & ", " & CLng(mes) & ", " & _
                              metaFormatada & ")" ' Access aceita o número formatado com ponto
                        connSales.Execute(sql)
                        
                        If Err.Number = 0 Then
                            mensagem = "Meta cadastrada com sucesso!"
                            mensagemTipo = "sucesso"
                            ano = "" : mes = "" : meta = ""
                        Else
                            mensagem = "Erro ao cadastrar: " & Err.Description
                            mensagemTipo = "erro"
                        End If
                    End If
                    If IsObject(rs) Then rs.Close
                Else
                    mensagem = "Preencha todos os campos!"
                    mensagemTipo = "erro"
                End If
                
            Case "editar"
                If id <> "" And ano <> "" And mes <> "" And meta <> "" Then
                    ' 5. MUDANÇA: Simplificar a formatação do valor para Access (Jet)
                    Dim metaFormatadaEdit
                    metaFormatadaEdit = Replace(meta, ",", ".")
                    metaFormatadaEdit = CDbl(metaFormatadaEdit)
                    
                    sql = "UPDATE MetaEmpresa SET " & _
                          "Ano = " & CLng(ano) & ", " & _
                          "Mes = " & CLng(mes) & ", " & _
                          "Meta = " & metaFormatadaEdit & " " & _
                          "WHERE ID = " & CLng(id)
                          
                    connSales.Execute(sql)
                    
                    If Err.Number = 0 Then
                        mensagem = "Meta atualizada com sucesso!"
                        mensagemTipo = "sucesso"
                        id = "" ' Sair do modo edição
                    Else
                        mensagem = "Erro ao atualizar: " & Err.Description
                        mensagemTipo = "erro"
                    End If
                End If
                
            Case "excluir"
                ' Nenhuma alteração necessária, DELETE FROM funciona no Access
                If id <> "" Then
                    sql = "DELETE FROM MetaEmpresa WHERE ID = " & CLng(id)
                    connSales.Execute(sql)
                    
                    If Err.Number = 0 Then
                        mensagem = "Meta excluída com sucesso!"
                        mensagemTipo = "sucesso"
                    Else
                        mensagem = "Erro ao excluir: " & Err.Description
                        mensagemTipo = "erro"
                    End If
                End If
                
            Case "carregar"
                ' Nenhuma alteração necessária, SELECT funciona no Access
                If id <> "" Then
                    sql = "SELECT * FROM MetaEmpresa WHERE ID = " & CLng(id)
                    Set rs = connSales.Execute(sql)
                    If Not rs.EOF Then
                        ano = rs("Ano")
                        mes = rs("Mes")
                        meta = FormatNumber(rs("Meta"), 2, -1, -1, -1) ' Mantido para formatação de exibição
                    End If
                    If IsObject(rs) Then rs.Close
                End If
        End Select
        
        ' Fechar conexão
        connSales.Close
        Set connSales = Nothing
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Metas da Empresa</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: #f0f2f5; padding: 20px; }
        .container { max-width: 1000px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: #2c3e50; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; }
        .form-section { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #e9ecef; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #495057; }
        input { width: 100%; max-width: 300px; padding: 8px 12px; border: 1px solid #ced4da; border-radius: 4px; font-size: 14px; }
        .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-right: 10px; transition: background 0.3s; }
        .btn-primary { background: #3498db; color: white; }
        .btn-success { background: #27ae60; color: white; }
        .btn-warning { background: #f39c12; color: white; }
        .btn-danger { background: #e74c3c; color: white; }
        .btn-secondary { background: #95a5a6; color: white; }
        .btn:hover { opacity: 0.9; }
        .message { padding: 12px; margin: 10px 0; border-radius: 4px; text-align: center; }
        .sucesso { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .erro { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; background: white; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #dee2e6; }
        th { background: #34495e; color: white; font-weight: bold; }
        tr:hover { background: #f8f9fa; }
        .actions { white-space: nowrap; }
        .actions form { display: inline; }
        h2 { color: #2c3e50; margin-bottom: 15px; }
        .current-year { font-size: 12px; color: #7f8c8d; margin-top: 5px; }
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
        <div class="header">
            <h1>📊 Gestão de Metas da Empresa</h1>
        </div>
        
        <div class="content">
            <% If mensagem <> "" Then %>
                <div class="message <%= mensagemTipo %>">
                    <%= mensagem %>
                </div>
            <% End If %>
            
            <div class="form-section">
                <h2><% If id = "" Then %>➕ Cadastrar Nova Meta<% Else %>✏️ Editar Meta<% End If %></h2>
                <form method="post" action="">
                    <input type="hidden" name="id" value="<%= Server.HTMLEncode(id) %>">
                    
                    <div class="form-group">
                        <label for="ano">📅 Ano:</label>
                        <input type="number" id="ano" name="ano" value="<%= Server.HTMLEncode(ano) %>" 
                               min="2020" max="2030" required>
                        <div class="current-year">Ano atual: <%= Year(Now) %></div>
                    </div>
                    
                    <div class="form-group">
                        <label for="mes">📋 Mês (1-12):</label>
                        <input type="number" id="mes" name="mes" value="<%= Server.HTMLEncode(mes) %>" 
                               min="1" max="12" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="meta">💰 Meta (R$):</label>
                        <input type="text" id="meta" name="meta" value="<%= Server.HTMLEncode(meta) %>" 
                               placeholder="Ex: 100000,00" required>
                    </div>
                    
                    <div class="form-group">
                        <% If id = "" Then %>
                            <button type="submit" name="acao" value="cadastrar" class="btn btn-primary">
                                ✅ Cadastrar Meta
                            </button>
                        <% Else %>
                            <button type="submit" name="acao" value="editar" class="btn btn-success">
                                💾 Atualizar
                            </button>
                            <button type="button" onclick="limparFormulario()" class="btn btn-secondary">
                                ❌ Cancelar
                            </button>
                        <% End If %>
                    </div>
                </form>
            </div>
            
            <h2>📋 Metas Cadastradas</h2>
            <%
            On Error Resume Next
            
            ' Abrir conexão para listar registros
            Set connSales = Server.CreateObject("ADODB.Connection")
            connSales.Open StrConnSales
            
            If Err.Number = 0 Then
                ' Nenhuma alteração necessária, SELECT funciona no Access
                sql = "SELECT * FROM MetaEmpresa ORDER BY Ano DESC, Mes DESC"
                Set rs = connSales.Execute(sql)
                
                If Err.Number = 0 Then
                    If rs.EOF Then
                        Response.Write "<p style='text-align: center; color: #7f8c8d; padding: 20px;'>Nenhuma meta cadastrada ainda.</p>"
                    Else
            %>
                        <table>
                            <thead>
                                <tr>
                                    <th>Ano</th>
                                    <th>Mês</th>
                                    <th>Meta (R$)</th>
                                    <th>Ações</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% Do While Not rs.EOF %>
                                <tr>
                                    <td><%= rs("Ano") %></td>
                                    <td><%= rs("Mes") %></td>
                                    <td>R$ <%= FormatNumber(rs("Meta"), 2, -1, -1, -1) %></td>
                                    <td class="actions">
                                        <form method="post" style="display: inline;">
                                            <input type="hidden" name="id" value="<%= rs("ID") %>">
                                            <input type="hidden" name="acao" value="carregar">
                                            <button type="submit" class="btn btn-warning">✏️ Editar</button>
                                        </form>
                                        <form method="post" style="display: inline;" onsubmit="return confirm('Tem certeza que deseja excluir esta meta?');">
                                            <input type="hidden" name="id" value="<%= rs("ID") %>">
                                            <input type="hidden" name="acao" value="excluir">
                                            <button type="submit" class="btn btn-danger">🗑️ Excluir</button>
                                        </form>
                                    </td>
                                </tr>
                                <% 
                                rs.MoveNext
                                Loop 
                                %>
                            </tbody>
                        </table>
            <%
                    End If
                    If IsObject(rs) Then rs.Close
                Else
                    Response.Write "<div class='message erro'>Erro ao carregar metas: " & Err.Description & "</div>"
                End If
                
                connSales.Close
            Else
                Response.Write "<div class='message erro'>Erro de conexão ao carregar metas: " & Err.Description & "</div>"
            End If
            
            Set connSales = Nothing
            Set rs = Nothing
            %>
        </div>
    </div>

    <script>
        function limparFormulario() {
            window.location.href = '<%= Request.ServerVariables("SCRIPT_NAME") %>';
        }
        
        // Formatação automática do campo de meta
        document.getElementById('meta')?.addEventListener('blur', function(e) {
            let valor = e.target.value.replace(/[^\d,]/g, '').replace(',', '.');
            valor = parseFloat(valor);
            if (!isNaN(valor)) {
                e.target.value = valor.toLocaleString('pt-BR', {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                });
            }
        });
    </script>
</body>
</html>