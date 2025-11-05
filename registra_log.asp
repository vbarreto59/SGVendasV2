<%
' Função para inserir registros na tabela log_operations
Function InserirLog(tabelaAfetada, acao, descricao)
    
    
    Dim conn, sql, cmd, usuario
    
    ' Obter usuário da Session
    usuario = Session("Usuario")
    
    ' Se não houver usuário na session, usar "Sistema" como padrão
    If usuario = "" Or IsEmpty(usuario) Then
        usuario = "Sistema"
    End If
    

    ' Criar conexão usando a string StrConnSales
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open StrConnSales
    
    If Err.Number <> 0 Then
        InserirLog = "Erro na conexão: " & Err.Description
        Exit Function
    End If
    
    ' Criar comando SQL
    sql = "INSERT INTO log_operations (Usuario, TabelaAfetada, Acao, Descricao) " & _
          "VALUES (?, ?, ?, ?)"
    
    ' Criar command object
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1 ' adCmdText
        
        ' Adicionar parâmetros (evita SQL Injection)
        .Parameters.Append .CreateParameter("@Usuario", 200, 1, 255, usuario) ' adVarChar
        .Parameters.Append .CreateParameter("@TabelaAfetada", 200, 1, 255, tabelaAfetada) ' adVarChar
        .Parameters.Append .CreateParameter("@Acao", 200, 1, 255, acao) ' adVarChar
        .Parameters.Append .CreateParameter("@Descricao", 200, 1, 4000, descricao) ' adVarChar
    End With
    
    ' Executar inserção
    cmd.Execute
    
    If Err.Number = 0 Then
        InserirLog = "Log inserido com sucesso!"
    Else
        InserirLog = "Erro ao inserir log: " & Err.Description
    End If
    
    ' Limpar objetos
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    
    On Error GoTo 0
End Function
%>