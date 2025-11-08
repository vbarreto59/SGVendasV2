<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

' Obtém os parâmetros
Dim idVenda, nomeUsuario
idVenda = Request.Form("id_venda")
nomeUsuario = Request.Form("nome")

' Valida o ID da venda
If Not IsNumeric(idVenda) Or idVenda = "" Then
    Response.Write "<div class='alert alert-danger'>ID de venda inválido.</div>"
    Response.End
End If

' Cria a conexão
Dim connSales, rs, sql
Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>Erro ao conectar ao banco de dados: " & Err.Description & "</div>"
    Response.End
End If
On Error GoTo 0

' Cria o recordset
Set rs = Server.CreateObject("ADODB.Recordset")

' Monta a query baseada na estrutura do seu código
sql = "SELECT ID_Pagamento, DataPagamento, ValorPago, Status, UsuariosNome, TipoRecebedor, TipoPagamento, ID_Venda, Obs, UsuariosUserId " & _
      "FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & CInt(idVenda)

' Se foi passado nome do usuário, filtrar por ele
If nomeUsuario <> "" Then
    sql = sql & " AND UsuariosNome LIKE '%" & Replace(nomeUsuario, "'", "''") & "%'"
End If

sql = sql & " ORDER BY DataPagamento DESC"

On Error Resume Next
rs.Open sql, connSales
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>Erro ao executar consulta: " & Err.Description & "</div>"
    Response.Write "<!-- SQL: " & sql & " -->"
    rs.Close
    Set rs = Nothing
    connSales.Close
    Set connSales = Nothing
    Response.End
End If
On Error GoTo 0

' Verifica se há registros
If Not rs.EOF Then
    Dim totalPagos
    totalPagos = 0
    
    Response.Write "<div class='table-responsive'>"
    Response.Write "<table class='table table-sm table-striped table-bordered'>"
    Response.Write "<thead class='table-light'>"
    Response.Write "<tr>"
    Response.Write "<th>ID Pagamento</th>"
    Response.Write "<th>Usuário</th>"
    Response.Write "<th class='text-end'>Valor Pago</th>"
    Response.Write "<th>Data Pagamento</th>"
    Response.Write "<th>Tipo Recebedor</th>"
    Response.Write "<th>Tipo Pagamento</th>"
    Response.Write "<th>Status</th>"
    Response.Write "<th>Observações</th>"
    Response.Write "</tr>"
    Response.Write "</thead>"
    Response.Write "<tbody>"
    
    Do While Not rs.EOF
        Dim valorPago, dataPagamentoFormatada
        valorPago = 0
        
        ' Formata o valor pago
        If Not IsNull(rs("ValorPago")) Then 
            valorPago = CDbl(rs("ValorPago"))
        End If
        totalPagos = totalPagos + valorPago
        
        ' Formata a data
        If Not IsNull(rs("DataPagamento")) Then
            dataPagamentoFormatada = FormatDateTime(rs("DataPagamento"), 2) 
        Else
            dataPagamentoFormatada = "N/A"
        End If
        
        ' Determina a classe do badge baseado no status
        Dim badgeClass
        Select Case LCase(rs("Status"))
            Case "realizado", "pago"
                badgeClass = "bg-success"
            Case "pendente"
                badgeClass = "bg-warning text-dark"
            Case "cancelado"
                badgeClass = "bg-danger"
            Case Else
                badgeClass = "bg-secondary"
        End Select
        
        Response.Write "<tr>"
        Response.Write "<td>" & rs("ID_Pagamento") & "</td>"
        Response.Write "<td><strong>" & rs("UsuariosNome") & "</strong></td>"
        Response.Write "<td class='text-end'><strong> " & FormatNumber(valorPago, 2) & "</strong></td>"
        Response.Write "<td>" & dataPagamentoFormatada & "</td>"
        Response.Write "<td>" & rs("TipoRecebedor") & "</td>"
        Response.Write "<td>" & rs("TipoPagamento") & "</td>"
        Response.Write "<td><span class='badge " & badgeClass & "'>" & rs("Status") & "</span></td>"
        Response.Write "<td>" & Server.HTMLEncode(rs("Obs") & "") & "</td>"
        Response.Write "</tr>"
        
        rs.MoveNext
    Loop
    
    Response.Write "</tbody>"
    Response.Write "<tfoot class='table-primary'>"
    Response.Write "<tr>"
    Response.Write "<td colspan='2'><strong>Total Pago:</strong></td>"
    Response.Write "<td class='text-end'><strong> " & FormatNumber(totalPagos, 2) & "</strong></td>"
    Response.Write "<td colspan='5'></td>"
    Response.Write "</tr>"
    Response.Write "</tfoot>"
    Response.Write "</table>"
    Response.Write "</div>"
    
    ' Adiciona informações resumidas
    Response.Write "<div class='row mt-3'>"
    Response.Write "<div class='col-md-6'>"
    Response.Write "<div class='alert alert-info'>"
    Response.Write "<i class='fas fa-info-circle me-2'></i>"
    Response.Write "<strong>Resumo:</strong> " & recordCount & " pagamento(s) encontrado(s) para a venda " & idVenda
    Response.Write "</div>"
    Response.Write "</div>"
    Response.Write "</div>"
    
Else
    Response.Write "<div class='alert alert-info text-center'>"
    Response.Write "<i class='fas fa-info-circle me-2'></i>"
    Response.Write "Nenhum pagamento encontrado para a venda " & idVenda
    If nomeUsuario <> "" Then
        Response.Write " do usuário " & nomeUsuario
    End If
    Response.Write "</div>"
End If

' Fecha o recordset e a conexão
rs.Close
Set rs = Nothing
connSales.Close
Set connSales = Nothing
%>