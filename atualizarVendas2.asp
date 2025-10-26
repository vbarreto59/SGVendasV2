
<%
' ===============================================
' ROTINA PARA ATUALIZAR Empresa_ID e NomeEmpresa NA TABELA Vendas
' ===============================================

' Configuração dos caminhos dos bancos de dadosDim connSales, connSunnyc
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunSalesPath = Split(StrConnSales, "Data Source=")(1)

If dbSunnyPath = "" Or dbSunSalesPath = "" Then
    Response.Write "Erro: Não foi possível extrair caminhos dos bancos de dados<br>"
    Response.End
End If

' Inicializar conexões
'Dim connSales, connSunny
Set connSales = Server.CreateObject("ADODB.Connection")
Set connSunny = Server.CreateObject("ADODB.Connection")

On Error Resume Next
connSales.Open StrConnSales
connSunny.Open StrConn

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar aos bancos de dados: " & Err.Description & "<br>"
    Response.End
End If
On Error GoTo 0

' Selecionar todos os registros de Vendas
'Dim rsVendas
Set rsVendas = Server.CreateObject("ADODB.Recordset")
rsVendas.Open "SELECT ID, Empreend_ID FROM Vendas", connSales, 3, 3 ' adOpenStatic, adLockOptimistic

If Err.Number <> 0 Then
    Response.Write "Erro ao selecionar registros de Vendas: " & Err.Description & "<br>"
    Response.End
End If

' Criar recordset para consultar Empreendimento
Dim rsEmpreendimento
Set rsEmpreendimento = Server.CreateObject("ADODB.Recordset")

If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        'Dim vendaID, empreendID, empresaID, nomeEmpresa, sqlUpdate
        vendaID = rsVendas("ID")
        empreendID = rsVendas("Empreend_ID")
        
        ' Verificar se empreendID é nulo
        If IsNull(empreendID) Then
            ' Atualizar com valores padrão para registros nulos
            sqlUpdate = "UPDATE Vendas SET Empresa_ID = NULL, NomeEmpresa = 'Desconhecida' WHERE ID = " & vendaID
            
            On Error Resume Next
            connSales.Execute sqlUpdate
            If Err.Number <> 0 Then
                Response.Write "Erro ao atualizar Venda ID " & vendaID & ": " & Err.Description & "<br>"
                Response.Write "SQL: " & Server.HTMLEncode(sqlUpdate) & "<br>"
            End If
            On Error GoTo 0
        Else
            ' Consultar Empresa_ID e NomeEmpresa na tabela Empreendimento
            Dim sqlEmpreendimento
            sqlEmpreendimento = "SELECT Empresa_ID, Localidade, NomeEmpresa FROM [;DATABASE=" & dbSunnyPath & "].Empreendimento WHERE Empreend_ID = " & empreendID

            On Error Resume Next
            rsEmpreendimento.Open sqlEmpreendimento, connSunny
            If Err.Number <> 0 Then
                Response.Write "Erro ao consultar Empreendimento para Venda ID " & vendaID & ": " & Err.Description & "<br>"
                Response.Write "SQL: " & Server.HTMLEncode(sqlEmpreendimento) & "<br>"
            Else
                ' Definir valores padrão para quando não há correspondência
                empresaID = "NULL"
                nomeEmpresa = "NULL" ' Alterado para NULL ao invés de string

                ' Se encontrou correspondência, obter os valores
                If Not rsEmpreendimento.EOF Then
                    If Not IsNull(rsEmpreendimento("Empresa_ID")) Then
                        empresaID = rsEmpreendimento("Empresa_ID")
                        Localidade = rsEmpreendimento("Localidade")
                    End If
                    If Not IsNull(rsEmpreendimento("NomeEmpresa")) Then
                        ' Usar aspas simples corretamente formatadas
                        nomeEmpresa = "'" & Replace(rsEmpreendimento("NomeEmpresa"), "'", "''") & "'"
                    End If
                End If

                rsEmpreendimento.Close

                nomeEmpresa = UCase(nomeEmpresa)

                ' Construir a instrução UPDATE corretamente
                If nomeEmpresa = "NULL" Then
                    sqlUpdate = "UPDATE Vendas SET Empresa_ID = " & empresaID & ", NomeEmpresa = NULL WHERE ID = " & vendaID
                Else
                    sqlUpdate = "UPDATE Vendas SET Empresa_ID = " & empresaID & ", NomeEmpresa = " & nomeEmpresa & ", Localidade = '" & localidade &  "' WHERE ID = " & vendaID
                End If
'Response.Write sqlUpdate
'Response.end 
                
                connSales.Execute sqlUpdate
                If Err.Number <> 0 Then
                    Response.Write "Erro ao atualizar Venda ID " & vendaID & ": " & Err.Description & "<br>"
                    Response.Write "SQL: " & Server.HTMLEncode(sqlUpdate) & "<br>"
                    ' Debug adicional para ver os valores
                    Response.Write "Empresa_ID: " & empresaID & "<br>"
                    Response.Write "NomeEmpresa: " & nomeEmpresa & "<br><br>"
                End If
                On Error GoTo 0
            End If
            On Error GoTo 0
        End If

        rsVendas.MoveNext
    Loop
End If

' Fechar recordsets e conexões
rsVendas.Close
Set rsVendas = Nothing
If rsEmpreendimento.State = 1 Then rsEmpreendimento.Close
Set rsEmpreendimento = Nothing
connSales.Close
Set connSales = Nothing
connSunny.Close
Set connSunny = Nothing

'Response.Write "<br><strong>Atualização concluída.</strong><br>"
%>