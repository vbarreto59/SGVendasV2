


<%
' Configurar variáveis
Dim conn, success, message
success = True
message = ""

Response.ContentType = "text/html"
Response.Charset = "utf-8"

' Iniciar conexão
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open StrConnSales

If Err.Number <> 0 Then
    success = False
    message = "Erro na conexão: " & Err.Description
Else
    ' Executar os comandos em sequência
    On Error Resume Next
    
    ' 1. Limpar tabela VENDA_TEMP
    conn.Execute "DELETE FROM VENDA_TEMP"
    If Err.Number <> 0 Then
        success = False
        message = message & "Erro ao limpar VENDA_TEMP: " & Err.Description & "<br>"
    End If
    
    ' 2. Inserir dados da Diretoria
    If success Then
        conn.Execute "INSERT INTO VENDA_TEMP ( ID_Venda, Diretoria, UserId, Gerencia, Nome, Cargo, VUnid, VBruto, [Desc], VLiq, Premio, VTotal ) " & _
                     "SELECT qryComissaoDiretor.ID_Venda, qryComissaoDiretor.Diretoria, qryComissaoDiretor.UserId, qryComissaoDiretor.Gerencia, " & _
                     "qryComissaoDiretor.Nome, qryComissaoDiretor.Cargo, qryComissaoDiretor.VUnid, qryComissaoDiretor.VBruto, qryComissaoDiretor.Desc, " & _
                     "qryComissaoDiretor.VLiq, qryComissaoDiretor.Premio, qryComissaoDiretor.VTotal " & _
                     "FROM qryComissaoDiretor"
        
        If Err.Number <> 0 Then
            success = False
            message = message & "Erro ao inserir dados da Diretoria: " & Err.Description & "<br>"
        End If
    End If
    
    ' 3. Inserir dados da Gerencia
    If success Then
        conn.Execute "INSERT INTO VENDA_TEMP ( ID_Venda, Diretoria, UserId, Gerencia, Nome, Cargo, VUnid, VBruto, [Desc], VLiq, Premio, VTotal ) " & _
                     "SELECT qryComissaoGerente.ID_Venda, qryComissaoGerente.Diretoria, qryComissaoGerente.UserId, qryComissaoGerente.Gerencia, " & _
                     "qryComissaoGerente.Nome, qryComissaoGerente.Cargo, qryComissaoGerente.VUnid, qryComissaoGerente.VBruto, qryComissaoGerente.Desc, " & _
                     "qryComissaoGerente.VLiq, qryComissaoGerente.Premio, qryComissaoGerente.VTotal " & _
                     "FROM qryComissaoGerente"
        
        If Err.Number <> 0 Then
            success = False
            message = message & "Erro ao inserir dados da Gerencia: " & Err.Description & "<br>"
        End If
    End If
    
    ' 4. Inserir dados do Corretor
    If success Then
        conn.Execute "INSERT INTO VENDA_TEMP ( ID_Venda, Diretoria, UserId, Gerencia, Nome, Cargo, VUnid, VBruto, [Desc], VLiq, Premio, VTotal ) " & _
                     "SELECT qryComissaoCorretor.ID_Venda, qryComissaoCorretor.Diretoria, qryComissaoCorretor.UserId, qryComissaoCorretor.Gerencia, " & _
                     "qryComissaoCorretor.Nome, qryComissaoCorretor.Cargo, qryComissaoCorretor.VUnid, qryComissaoCorretor.VBruto, qryComissaoCorretor.Desc, " & _
                     "qryComissaoCorretor.VLiq, qryComissaoCorretor.Premio, qryComissaoCorretor.VTotal " & _
                     "FROM qryComissaoCorretor"
        
        If Err.Number <> 0 Then
            success = False
            message = message & "Erro ao inserir dados do Corretor: " & Err.Description & "<br>"
        End If
    End If
    
    ' 5. Verificar total de registros inseridos
    If success Then
        Dim rsCount
        Set rsCount = conn.Execute("SELECT COUNT(*) as Total FROM VENDA_TEMP")
        If Not rsCount.EOF Then
            message = message & "Total de registros em VENDA_TEMP: " & rsCount("Total")
        End If
        rsCount.Close
        Set rsCount = Nothing
    End If
    
End If

' Fechar conexão
If IsObject(conn) Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If

' Exibir resultado simples
If success Then
    'Response.Write "SUCESSO: Comandos executados com sucesso. " & message
Else
   '' Response.Write "ERRO: " & message
End If
%>