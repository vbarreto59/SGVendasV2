
<%
Response.Buffer = True
Response.Expires = -1
On Error Resume Next ' Liga a checagem de erros para tratamento

Dim conn, connSales, sqlUpdate, sqlSelect, rsVendas
Dim dbEmpreendimentoPathClean
Dim lErro

lErro = False

' =========================================================================================
' === EXTRAÇÃO DO PATH DO BANCO DE DADOS DE EMPREENDIMENTO ================================
' =========================================================================================
' Extrai o caminho físico do arquivo MDB da connection string StrConn.
If InStr(1, StrConn, "Data Source=", 1) > 0 Then
    dbEmpreendimentoPathClean = Trim(Mid(StrConn, InStr(1, StrConn, "Data Source=", 1) + Len("Data Source=")))
    If Right(dbEmpreendimentoPathClean, 1) = ";" Then
        dbEmpreendimentoPathClean = Left(dbEmpreendimentoPathClean, Len(dbEmpreendimentoPathClean) - 1)
    End If
Else
    ' Se não encontrar 'Data Source=', assume que StrConn é o próprio path (menos comum para conexões DSN-less)
    dbEmpreendimentoPathClean = StrConn
End If

' =========================================================================================
' === ESTABELECENDO CONEXÕES ===============================================================
' =========================================================================================
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

' Tenta conectar
conn.Open StrConn
If Err.Number <> 0 Then
    Response.Write ""
    lErro = True
End If
On Error GoTo 0

If Not lErro Then
    On Error Resume Next
    connSales.Open StrConnSales
    If Err.Number <> 0 Then
        Response.Write ""
        lErro = True
    End If
    On Error GoTo 0
End If


' =========================================================================================
' === EXECUÇÃO DA ATUALIZAÇÃO =============================================================
' =========================================================================================
%>

<% If dbEmpreendimentoPathClean = "" Or InStr(dbEmpreendimentoPathClean, "=") > 0 Then %>

<% ElseIf lErro Then %>

<% Else %>
    

    <%
    ' SQL para fazer o UPDATE com JOIN (Vendas.Empreend_id = Empreendimento.Empreend_id)
    sqlUpdate = "UPDATE ([;DATABASE=" & dbEmpreendimentoPathClean & "].Empreendimento AS Emp " & _
                "INNER JOIN Vendas ON Emp.Empreend_id = Vendas.Empreend_id) " & _
                "SET Vendas.Localizacao = [Emp].[Localizacao];"
    


    ' Executa a instrução SQL no banco de Vendas (connSales)
    On Error Resume Next
    connSales.Execute(sqlUpdate)
    
    If Err.Number <> 0 Then

    Else

        
        ' =================================================================
        ' === ROTINA DE VERIFICAÇÃO (SELECT) ==============================
        ' =================================================================
     
        
        ' Seleciona os IDs e a Localização para verificação
        sqlSelect = "SELECT TOP 100 id, Localizacao, NomeEmpreendimento FROM Vendas WHERE Localizacao IS NOT NULL ORDER BY id DESC"
        
        Set rsVendas = Server.CreateObject("ADODB.Recordset")
        rsVendas.Open sqlSelect, connSales, 1, 3 ' Usando adOpenKeyset, adLockOptimistic
        
        If rsVendas.EOF And rsVendas.BOF Then
           
        Else
 
            
            Do While Not rsVendas.EOF

                rsVendas.MoveNext
            Loop

        End If
        
        ' Limpa o Recordset
        rsVendas.Close
        Set rsVendas = Nothing
    End If
    On Error GoTo 0 ' Desliga a checagem de erro

    ' Fecha as conexões
    If Not conn Is Nothing Then conn.Close
    If Not connSales Is Nothing Then connSales.Close
End If

Set conn = Nothing
Set connSales = Nothing
%>
<hr>


</body>
</html>