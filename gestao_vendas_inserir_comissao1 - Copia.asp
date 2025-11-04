<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

' Função auxiliar para formatar valores
Function FormatarValor(valor)
    valor = Replace(valor, ".", ",")
    valor = Replace(valor, ",", ".")
    FormatarValor = valor
End Function

' Função para converter valores monetários corretamente
Function ParseCurrency(value)
    On Error Resume Next
    If IsNumeric(value) Then
        ParseCurrency = CDbl(value)
        Exit Function
    End If
    ParseCurrency = CDbl(Replace(Replace(Replace(value, ".", ""), ",", ".")))
    If Err.Number <> 0 Then ParseCurrency = 0
    On Error GoTo 0
End Function

' Obtém o ID da venda do parâmetro (QueryString)
Dim vendaId
vendaId = Request.QueryString("id")
If Not IsNumeric(vendaId) Or vendaId = "" Then
    Response.Write "<script>alert('Erro: ID da venda inválido.');window.location.href='gestao_vendas_list2x.asp';</script>"
    Response.End
End If

' Cria as conexões
Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Busca os dados da venda na tabela Vendas
Dim rsVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
rsVenda.Open "SELECT * FROM Vendas WHERE ID = " & CInt(vendaId), connSales

If rsVenda.EOF Then
    Response.Write "<script>alert('Erro: Venda não encontrada.');window.location.href='gestao_vendas_list2x.asp';</script>"
    rsVenda.Close
    Set rsVenda = Nothing
    Response.End
End If

' Obtém os dados da tabela Vendas
Dim empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
Dim dataVenda, obs, m2, diretoriaId, gerenciaId, trimestre
Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor
Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
Dim nomeDiretor, nomeGerente, nomeCorretor, nomeEmpreendimento

empreend_id = rsVenda("Empreend_ID")
unidade = Server.HTMLEncode(rsVenda("Unidade"))
corretorId = rsVenda("CorretorId")
diretoriaId = rsVenda("DiretoriaId")
gerenciaId = rsVenda("GerenciaId")
trimestre = rsVenda("Trimestre")
dataVenda = rsVenda("DataVenda")
obs = Server.HTMLEncode(rsVenda("Obs"))
valorUnidade = ParseCurrency(rsVenda("ValorUnidade"))
m2 = ParseCurrency(rsVenda("UnidadeM2"))
comissaoPercentual = ParseCurrency(rsVenda("ComissaoPercentual"))
comissaoDiretoria = ParseCurrency(rsVenda("ComissaoDiretoria"))
comissaoGerencia = ParseCurrency(rsVenda("ComissaoGerencia"))
comissaoCorretor = ParseCurrency(rsVenda("ComissaoCorretor"))

' Fecha o recordset da venda
rsVenda.Close
Set rsVenda = Nothing

' Cálculo das comissões
valorComissaoGeral = valorUnidade * (comissaoPercentual / 100)
valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / 100)
valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / 100)
valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / 100)

' Lógica para INSERIR na tabela COMISSOES_A_PAGAR, com VERIFICAÇÃO de duplicidade
Dim rsCheck
Set rsCheck = Server.CreateObject("ADODB.Recordset")

' Consulta para verificar se a comissão já existe para esta venda
rsCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(vendaId), connSales

If Not rsCheck.EOF Then
    ' Se a comissão já existe, exibe uma mensagem e não insere
    Response.Write "<script>alert('A comissão para esta venda já foi gerada e não pode ser criada novamente.');window.location.href='gestao_vendas_list2x.asp';</script>"
    rsCheck.Close
    Set rsCheck = Nothing
    Response.End
Else
    ' Se a comissão não existe, insere o novo registro
    rsCheck.Close
    Set rsCheck = Nothing

    ' Validações
    If IsEmpty(vendaId) Or IsNull(vendaId) Or vendaId = "" Then
        Response.Write "<script>alert('Erro: ID da venda inválido.');window.location.href='gestao_vendas_list2x.asp';</script>"
        Response.End
    End If
    If IsEmpty(diretoriaId) Or IsNull(diretoriaId) Or diretoriaId = "" Then
        diretoriaId = 0
    End If
    If IsEmpty(gerenciaId) Or IsNull(gerenciaId) Or gerenciaId = "" Then
        gerenciaId = 0
    End If
    If IsEmpty(corretorId) Or IsNull(corretorId) Or corretorId = "" Then
        Response.Write "<script>alert('Erro: ID do corretor inválido.');window.location.href='gestao_vendas_list2x.asp';</script>"
        Response.End
    End If
    If IsEmpty(dataVenda) Or IsNull(dataVenda) Or dataVenda = "" Then
        Response.Write "<script>alert('Erro: Data de venda inválida.');window.location.href='gestao_vendas_list2x.asp';</script>"
        Response.End
    End If
    If IsEmpty(unidade) Or IsNull(unidade) Or unidade = "" Then
        Response.Write "<script>alert('Erro: Unidade inválida.');window.location.href='gestao_vendas_list2x.asp';</script>"
        Response.End
    End If

    ' Arredondar valores decimais
    comissaoDiretoria = FormatarValor(comissaoDiretoria)
    comissaoGerencia = FormatarValor(comissaoGerencia)
    comissaoCorretor = FormatarValor(comissaoCorretor)
    valorComissaoDiretoria = FormatarValor(valorComissaoDiretoria)
    valorComissaoGerencia = FormatarValor(valorComissaoGerencia)
    valorComissaoCorretor = FormatarValor(valorComissaoCorretor)
    valorComissaoGeral = FormatarValor(valorComissaoGeral)

    ' Busca os nomes do diretor, gerente, corretor e empreendimento
    Dim rsNomes
    Set rsNomes = Server.CreateObject("ADODB.Recordset")
    
    ' Busca nome do diretor
    rsNomes.Open "SELECT u.Nome FROM Usuarios u INNER JOIN Diretorias d ON u.UserId = d.UserId WHERE d.DiretoriaID = " & CInt(diretoriaId), conn
    If Not rsNomes.EOF Then
        nomeDiretor = rsNomes("Nome")
        If IsNull(nomeDiretor) Then nomeDiretor = ""
    Else
        nomeDiretor = ""
    End If
    rsNomes.Close
    
    ' Busca nome do gerente
    rsNomes.Open "SELECT u.Nome FROM Usuarios u INNER JOIN Gerencias g ON u.UserId = g.UserId WHERE g.GerenciaID = " & CInt(gerenciaId), conn
    If Not rsNomes.EOF Then
        nomeGerente = rsNomes("Nome")
        If IsNull(nomeGerente) Then nomeGerente = ""
    Else
        nomeGerente = ""
    End If
    rsNomes.Close
    
    ' Busca nome do corretor
    rsNomes.Open "SELECT Nome FROM Usuarios WHERE UserId = " & CInt(corretorId), conn
    If Not rsNomes.EOF Then
        nomeCorretor = rsNomes("Nome")
        If IsNull(nomeCorretor) Then nomeCorretor = ""
    Else
        nomeCorretor = ""
    End If
    rsNomes.Close
    
    ' Busca nome do empreendimento
    Dim rsEmp
    Set rsEmp = Server.CreateObject("ADODB.Recordset")
    rsEmp.Open "SELECT NomeEmpreendimento FROM Empreendimento WHERE Empreend_ID = " & empreend_id, conn
    If Not rsEmp.EOF Then
        nomeEmpreendimento = rsEmp("NomeEmpreendimento")
        If IsNull(nomeEmpreendimento) Then nomeEmpreendimento = ""
    Else
        nomeEmpreendimento = ""
        Response.Write "<script>alert('Erro: Empreendimento não encontrado.');window.location.href='gestao_vendas_list2x.asp';</script>"
        rsEmp.Close
        Set rsEmp = Nothing
        Response.End
    End If
    rsEmp.Close
    Set rsEmp = Nothing
    Set rsNomes = Nothing

    ' Insere na tabela COMISSOES_A_PAGAR
    Dim sql
    sql = "INSERT INTO COMISSOES_A_PAGAR (ID_Venda, Empreend_ID, Empreendimento, Unidade, DataVenda, " & _
          "UserIdDiretoria, UserIdGerencia, UserIdCorretor, PercDiretoria, ValorDiretoria, " & _
          "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, " & _
          "NomeDiretor, NomeGerente, NomeCorretor) " & _
          "VALUES (" & CInt(vendaId) & ", " & CInt(empreend_id) & ", '" & Replace(nomeEmpreendimento, "'", "''") & "', '" & Replace(unidade, "'", "''") & "', '" & Replace(dataVenda, "'", "''") & "', " & _
          CInt(diretoriaId) & ", " & CInt(gerenciaId) & ", " & CInt(corretorId) & ", " & _
          Replace(CStr(comissaoDiretoria), ",", ".") & ", " & Replace(CStr(valorComissaoDiretoria), ",", ".") & ", " & _
          Replace(CStr(comissaoGerencia), ",", ".") & ", " & Replace(CStr(valorComissaoGerencia), ",", ".") & ", " & _
          Replace(CStr(comissaoCorretor), ",", ".") & ", " & Replace(CStr(valorComissaoCorretor), ",", ".") & ", " & _
          Replace(CStr(valorComissaoGeral), ",", ".") & ", " & _
          "'" & Replace(nomeDiretor, "'", "''") & "', " & _
          "'" & Replace(nomeGerente, "'", "''") & "', " & _
          "'" & Replace(nomeCorretor, "'", "''") & "')"

    On Error Resume Next
    connSales.Execute(sql)
    If Err.Number <> 0 Then
        Response.Write "<script>alert('Erro ao gerar comissão: " & Replace(Err.Description, "'", "\'") & "');window.location.href='gestao_vendas_list2x.asp';</script>"
        Response.End
    End If
    On Error GoTo 0

    ' Fecha conexões
    If IsObject(conn) Then
        conn.Close
        Set conn = Nothing
    End If
    If IsObject(connSales) Then
        connSales.Close
        Set connSales = Nothing
    End If

    ' Redireciona com mensagem de sucesso
    Response.Redirect "gestao_vendas_list2x.asp?mensagem=Comissão gerada com sucesso!"
End If
%>