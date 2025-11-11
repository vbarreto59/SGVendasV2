<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% 'Execução da duplicação'
    If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<%
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Função para sanitizar strings SQL
Function SanitizeSQL(texto)
    If IsNull(texto) Or texto = "" Then
        SanitizeSQL = ""
    Else
        SanitizeSQL = Replace(texto, "'", "''")
    End If
End Function

' Função para formatar valores para SQL
Function GetSQLValue(valor, tipo)
    If IsNull(valor) Or valor = "" Then
        Select Case tipo
            Case "number": GetSQLValue = "0"
            Case "text": GetSQLValue = "''"
            Case "date": GetSQLValue = "NULL"
            Case Else: GetSQLValue = "NULL"
        End Select
    Else
        Select Case tipo
            Case "number"
                If IsNumeric(valor) Then
                    GetSQLValue = Replace(Replace(valor, ",", "."), " ", "")
                Else
                    GetSQLValue = "0"
                End If
            Case "text"
                GetSQLValue = "'" & SanitizeSQL(valor) & "'"
            Case "date"
                If IsDate(valor) Then
                    GetSQLValue = "#" & FormatDateTime(valor, 2) & "#"
                Else
                    GetSQLValue = "NULL"
                End If
        End Select
    End If
End Function

Dim idVenda
idVenda = Request.QueryString("id")

If idVenda = "" Then
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=ID da venda não informado&tipo=erro"
End If

If Not IsNumeric(idVenda) Then
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=ID inválido&tipo=erro"
End If

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

On Error Resume Next

' Buscar os dados da venda original
Dim sqlBuscar, rsVenda
sqlBuscar = "SELECT * FROM Vendas WHERE ID = " & idVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
rsVenda.Open sqlBuscar, connSales

If rsVenda.EOF Then
    rsVenda.Close
    Set rsVenda = Nothing
    connSales.Close
    Set connSales = Nothing
    Response.Redirect "gestao_vendas_list3x.asp?mensagem=Venda não encontrada&tipo=erro"
End If

' Calcular valor líquido geral
Dim valorLiqGeral
valorLiqGeral = 0
If Not IsNull(rsVenda("ValorLiqDiretoria")) And IsNumeric(rsVenda("ValorLiqDiretoria")) Then 
    valorLiqGeral = valorLiqGeral + CDbl(rsVenda("ValorLiqDiretoria"))
End If
If Not IsNull(rsVenda("ValorLiqGerencia")) And IsNumeric(rsVenda("ValorLiqGerencia")) Then 
    valorLiqGeral = valorLiqGeral + CDbl(rsVenda("ValorLiqGerencia"))
End If
If Not IsNull(rsVenda("ValorLiqCorretor")) And IsNumeric(rsVenda("ValorLiqCorretor")) Then 
    valorLiqGeral = valorLiqGeral + CDbl(rsVenda("ValorLiqCorretor"))
End If

' Criar INSERT usando a mesma estrutura do seu sistema
Dim sqlInsert
sqlInsert = "INSERT INTO Vendas (" & _
    "Empreend_ID, NomeEmpreendimento, Unidade, UnidadeM2, Corretor, CorretorId, " & _
    "ValorUnidade, ComissaoPercentual, ValorComissaoGeral, DataVenda, " & _
    "DiaVenda, MesVenda, AnoVenda, Trimestre, Obs, Usuario, " & _
    "DiretoriaId, Diretoria, GerenciaId, Gerencia, " & _
    "ComissaoDiretoria, ValorDiretoria, " & _
    "ComissaoGerencia, ValorGerencia, " & _
    "ComissaoCorretor, ValorCorretor, " & _
    "PremioDiretoria, PremioGerencia, PremioCorretor, " & _
    "DescontoPerc, DescontoBruto, DescontoDescricao, " & _
    "DescontoDiretoria, DescontoGerencia, DescontoCorretor, " & _
    "ValorLiqDiretoria, ValorLiqGerencia, ValorLiqCorretor, ValorLiqGeral, " & _
    "NomeDiretor, UserIdDiretoria, NomeGerente, UserIdGerencia, " & _
    "Localidade, Semestre, DataRegistro, Excluido" & _
    ") VALUES (" & _
    GetSQLValue(rsVenda("Empreend_ID"), "number") & ", " & _
    GetSQLValue(rsVenda("NomeEmpreendimento"), "text") & ", " & _
    GetSQLValue(rsVenda("Unidade"), "text") & ", " & _
    GetSQLValue(rsVenda("UnidadeM2"), "number") & ", " & _
    GetSQLValue(rsVenda("Corretor"), "text") & ", " & _
    GetSQLValue(rsVenda("CorretorId"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorUnidade"), "number") & ", " & _
    GetSQLValue(rsVenda("ComissaoPercentual"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorComissaoGeral"), "number") & ", " & _
    GetSQLValue(rsVenda("DataVenda"), "date") & ", " & _
    GetSQLValue(rsVenda("DiaVenda"), "number") & ", " & _
    GetSQLValue(rsVenda("MesVenda"), "number") & ", " & _
    GetSQLValue(rsVenda("AnoVenda"), "number") & ", " & _
    GetSQLValue(rsVenda("Trimestre"), "number") & ", " & _
    GetSQLValue(rsVenda("Obs"), "text") & ", " & _
    "'" & SanitizeSQL(Session("Usuario")) & "', " & _
    GetSQLValue(rsVenda("DiretoriaId"), "number") & ", " & _
    GetSQLValue(rsVenda("Diretoria"), "text") & ", " & _
    GetSQLValue(rsVenda("GerenciaId"), "number") & ", " & _
    GetSQLValue(rsVenda("Gerencia"), "text") & ", " & _
    GetSQLValue(rsVenda("ComissaoDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("ComissaoGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("ComissaoCorretor"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorCorretor"), "number") & ", " & _
    GetSQLValue(rsVenda("PremioDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("PremioGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("PremioCorretor"), "number") & ", " & _
    GetSQLValue(rsVenda("DescontoPerc"), "number") & ", " & _
    GetSQLValue(rsVenda("DescontoBruto"), "number") & ", " & _
    GetSQLValue(rsVenda("DescontoDescricao"), "text") & ", " & _
    GetSQLValue(rsVenda("DescontoDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("DescontoGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("DescontoCorretor"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorLiqDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorLiqGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("ValorLiqCorretor"), "number") & ", " & _
    valorLiqGeral & ", " & _
    GetSQLValue(rsVenda("NomeDiretor"), "text") & ", " & _
    GetSQLValue(rsVenda("UserIdDiretoria"), "number") & ", " & _
    GetSQLValue(rsVenda("NomeGerente"), "text") & ", " & _
    GetSQLValue(rsVenda("UserIdGerencia"), "number") & ", " & _
    GetSQLValue(rsVenda("Localidade"), "text") & ", " & _
    GetSQLValue(rsVenda("Semestre"), "number") & ", " & _
    "NOW(), " & _
    "0)"

' Executar a inserção na tabela Vendas
connSales.Execute sqlInsert

Dim mensagem, tipoMensagem, novoID

If Err.Number <> 0 Then
    mensagem = "Erro ao duplicar venda: " & Err.Description
    tipoMensagem = "erro"
Else
    ' Buscar o ID da nova venda criada
    Dim rsNovaVenda
    Set rsNovaVenda = connSales.Execute("SELECT @@IDENTITY AS NewID")
    If Not rsNovaVenda.EOF Then
        novoID = rsNovaVenda("NewID")
    End If
    rsNovaVenda.Close
    Set rsNovaVenda = Nothing
    
    ' AGORA INSERIR NA TABELA COMISSOES_A_PAGAR
    Dim sqlComissoes
    sqlComissoes = "INSERT INTO COMISSOES_A_PAGAR (" & _
        "ID_Venda, Empreendimento, Unidade, DataVenda, UserIdDiretoria, NomeDiretor, " & _
        "UserIdGerencia, NomeGerente, UserIdCorretor, NomeCorretor, PercDiretoria, ValorDiretoria, " & _
        "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, StatusPagamento, Usuario, " & _
        "PremioDiretoria, PremioGerencia, PremioCorretor, " & _
        "DescontoPerc, DescontoBruto, DescontoDescricao, " & _
        "DescontoDiretoria, DescontoGerencia, DescontoCorretor, " & _
        "ValorLiqDiretoria, ValorLiqGerencia, ValorLiqCorretor) " & _
        "VALUES (" & _
        novoID & ", " & _
        GetSQLValue(rsVenda("NomeEmpreendimento"), "text") & ", " & _
        GetSQLValue(rsVenda("Unidade"), "text") & ", " & _
        GetSQLValue(rsVenda("DataVenda"), "date") & ", " & _
        GetSQLValue(rsVenda("DiretoriaId"), "number") & ", " & _
        GetSQLValue(rsVenda("NomeDiretor"), "text") & ", " & _
        GetSQLValue(rsVenda("GerenciaId"), "number") & ", " & _
        GetSQLValue(rsVenda("NomeGerente"), "text") & ", " & _
        GetSQLValue(rsVenda("CorretorId"), "number") & ", " & _
        GetSQLValue(rsVenda("Corretor"), "text") & ", " & _
        GetSQLValue(rsVenda("ComissaoDiretoria"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorDiretoria"), "number") & ", " & _
        GetSQLValue(rsVenda("ComissaoGerencia"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorGerencia"), "number") & ", " & _
        GetSQLValue(rsVenda("ComissaoCorretor"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorCorretor"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorComissaoGeral"), "number") & ", " & _
        "'Pendente', " & _
        "'" & SanitizeSQL(Session("Usuario")) & "', " & _
        GetSQLValue(rsVenda("PremioDiretoria"), "number") & ", " & _
        GetSQLValue(rsVenda("PremioGerencia"), "number") & ", " & _
        GetSQLValue(rsVenda("PremioCorretor"), "number") & ", " & _
        GetSQLValue(rsVenda("DescontoPerc"), "number") & ", " & _
        GetSQLValue(rsVenda("DescontoBruto"), "number") & ", " & _
        GetSQLValue(rsVenda("DescontoDescricao"), "text") & ", " & _
        GetSQLValue(rsVenda("DescontoDiretoria"), "number") & ", " & _
        GetSQLValue(rsVenda("DescontoGerencia"), "number") & ", " & _
        GetSQLValue(rsVenda("DescontoCorretor"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorLiqDiretoria"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorLiqGerencia"), "number") & ", " & _
        GetSQLValue(rsVenda("ValorLiqCorretor"), "number") & ")"
    
    ' Executar inserção na tabela COMISSOES_A_PAGAR
    connSales.Execute sqlComissoes
    
    If Err.Number <> 0 Then
        mensagem = "Venda duplicada (ID: " & novoID & "), mas erro ao inserir comissões: " & Err.Description
        tipoMensagem = "aviso"
    Else
        mensagem = "Venda duplicada com sucesso! Nova venda ID: " & novoID
        tipoMensagem = "sucesso"
    End If
    
    ' Registrar no log
    Call RegistrarLog("Venda " & idVenda & " duplicada para " & novoID, Session("Usuario"))
End If

On Error GoTo 0

' Fechar conexões
rsVenda.Close
Set rsVenda = Nothing
connSales.Close
Set connSales = Nothing

' Função para registrar log
Sub RegistrarLog(acao, usuario)
    On Error Resume Next
    Dim connLog, sqlLog
    Set connLog = Server.CreateObject("ADODB.Connection")
    connLog.Open StrConnSales
    
    sqlLog = "INSERT INTO LogSistema (DataHora, Usuario, Acao, Descricao, IP) " & _
             "VALUES (NOW(), '" & usuario & "', 'DUPLICAR_VENDA', '" & Replace(acao, "'", "''") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    
    connLog.Execute sqlLog
    
    connLog.Close
    Set connLog = Nothing
    On Error GoTo 0
End Sub

' Redirecionar de volta para a lista com mensagem
Response.Redirect "gestao_vendas_list3x.asp?mensagem=" & Server.URLEncode(mensagem) & "&tipo=" & tipoMensagem
%>