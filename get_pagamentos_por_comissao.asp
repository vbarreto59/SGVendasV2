<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

' Função para formatar valores numéricos para JSON (usando ponto como separador decimal)
Function SafeFormatNumberForJson(value)
    If IsNumeric(value) Then
        'SafeFormatNumberForJson = CDbl(Replace(Replace(value, ",", "."), "R$", ""))
        SafeFormatNumberForJson = (Replace(Replace(value, ",", "."), "R$", ""))
    Else
        SafeFormatNumberForJson = 0.0
    End If
End Function

' Função para escapar strings para JSON
Function EscapeJsonString(str)
    If IsNull(str) Or str = "" Then
        EscapeJsonString = ""
    Else
        EscapeJsonString = Replace(Replace(Replace(str, "\", "\\"), """", "\"""), vbCrLf, "\n")
    End If
End Function

' Obtém o ID da venda
Dim idVenda
idVenda = Request.QueryString("idVenda")

' Valida o ID da venda
If Not IsNumeric(idVenda) Or idVenda = "" Then
    Response.Write "{""success"": false, ""error"": ""ID de venda inválido.""}"
    Response.End
End If

' Cria a conexão
Dim connSales, rs, sql
Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales
If Err.Number <> 0 Then
    Response.Write "{""success"": false, ""error"": ""Erro ao conectar ao banco de dados: " & EscapeJsonString(Err.Description) & """}"
    Response.End
End If
On Error GoTo 0

' Cria o recordset
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 ' adUseClient
rs.CursorType = 0 ' adOpenForwardOnly
rs.LockType = 1 ' adLockReadOnly

' Monta a query
sql = "SELECT ID_Pagamento, DataPagamento, ValorPago, Status, UsuariosNome, TipoRecebedor, TipoPagamento, ID_Venda, Obs, UsuariosUserId " & _
      "FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & CInt(idVenda) & " ORDER BY DataPagamento DESC"

On Error Resume Next
rs.Open sql, connSales
If Err.Number <> 0 Then
    Response.Write "{""success"": false, ""error"": ""Erro ao executar query: " & EscapeJsonString(Err.Description) & """}"
    rs.Close
    Set rs = Nothing
    connSales.Close
    Set connSales = Nothing
    Response.End
End If
On Error GoTo 0

' Monta o JSON de resposta
Dim json, recordCount
recordCount = 0
json = "{""success"": true, ""data"": ["

If Not rs.EOF Then
    Dim first
    first = True
    Do While Not rs.EOF
        recordCount = recordCount + 1
        If Not first Then json = json & ","
        json = json & "{" & _
            """ID_Pagamento"": " & rs("ID_Pagamento") & "," & _
            """DataPagamento"": """ & EscapeJsonString(FormatDateTime(rs("DataPagamento"), 2)) & """," & _
            """TipoPagamento"": """ & EscapeJsonString(rs("TipoPagamento")) & """," & _
            """ValorPago"": " & SafeFormatNumberForJson(rs("ValorPago")) & "," & _
            """Status"": """ & EscapeJsonString(rs("Status")) & """," & _
            """UsuariosNome"": """ & EscapeJsonString(rs("UsuariosNome")) & """," & _
            """TipoRecebedor"": """ & EscapeJsonString(rs("TipoRecebedor")) & """," & _
            """Obs"": """ & EscapeJsonString(rs("Obs") & "") & """," & _
            """UsuariosUserId"": """ & EscapeJsonString(rs("UsuariosUserId") & "") & """," & _
            """ID_Venda"": " & rs("ID_Venda") & "}"
        first = False
        rs.MoveNext
    Loop
End If

json = json & "]}"

' Fecha o recordset e a conexão
rs.Close
Set rs = Nothing
connSales.Close
Set connSales = Nothing

' Retorna o JSON
Response.Write json
%>