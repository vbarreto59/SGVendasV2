<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<% ' funcional em 04 11 2025'
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"
%>

<%
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
%>

<%
' original do site ====================================================================
' FUNÇÕES AUXILIARES
' ====================================================================

' Função para remover números e asteriscos de uma string
Function RemoverNumeros(texto)
    Dim regex
    Set regex = New RegExp
    
    ' Remove números (0-9) e asteriscos (*)
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    
    RemoverNumeros = regex.Replace(texto, "")
    
    ' Remove espaços extras
    RemoverNumeros = Trim(Replace(RemoverNumeros, "  ", " "))
End Function

Function FormatarValor(valor)
    valor = Replace(valor, "." , ",")
    valor = Replace(valor, "," , ".")
    FormatarValor = valor
End Function    

' ====================================================================
' INICIALIZAÇÃO E BUSCA DE DADOS
' ====================================================================

' Verifica se foi passado o ID da venda
Dim vendaId
vendaId = Request.QueryString("id")
If Not IsNumeric(vendaId) Or vendaId = "" Then
    Response.Redirect "gestao_vendas_list3x.asp"
End If

' Cria as conexões
Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Busca os dados da venda para preencher o formulário (do banco de vendas)
Dim rsVenda
Set rsVenda = Server.CreateObject("ADODB.Recordset")
rsVenda.Open "SELECT * FROM Vendas WHERE ID = " & vendaId, connSales

If rsVenda.EOF Then
    Response.Redirect "gestao_vendas_list3x.asp"
End If

' ====================================================================
' PROCESSAMENTO DO FORMULÁRIO (POST)
' ====================================================================

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Declaração de variáveis
    Dim empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
    Dim dataVenda, obs, m2, diretoriaId, gerenciaId, trimestre
    Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor
    Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
    Dim nomeDiretor, nomeGerente, nomeCorretor, nomeEmpreendimento, nomeDiretoria, nomeGerencia
    
    ' CAMPOS DE PREMIAÇÃO ADICIONADOS - 04 11 2025
    Dim premioDiretoria, premioGerencia, premioCorretor

    ' Obtenção segura dos valores do formulário
    vendaId = Request.Form("vendaId")
    empreend_id = Request.Form("empreend_id")
    unidade = Server.HTMLEncode(Request.Form("unidade"))
    corretorId = Request.Form("corretorId")
    diretoriaId = Request.Form("diretoriaId")
    gerenciaId = Request.Form("gerenciaId")
    trimestre = Request.Form("trimestre")
    dataVenda = Request.Form("dataVenda")
    obs = Server.HTMLEncode(Request.Form("obs"))
    
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
    
    ' Conversão segura dos valores numéricos
    valorUnidade = ParseCurrency(Request.Form("valorUnidade"))
    m2 = ParseCurrency(Request.Form("m2"))
    m2 = Replace(m2, ",", ".")
    comissaoPercentual = ParseCurrency(Request.Form("comissaoPercentual"))
    comissaoDiretoria = ParseCurrency(Request.Form("comissaoDiretoria"))
    comissaoGerencia = ParseCurrency(Request.Form("comissaoGerencia"))
    comissaoCorretor = ParseCurrency(Request.Form("comissaoCorretor"))
    
    ' PREMIAÇÕES - Coleta e formatação
    premioDiretoria = ParseCurrency(Request.Form("premioDiretoria"))
    premioGerencia = ParseCurrency(Request.Form("premioGerencia"))
    premioCorretor = ParseCurrency(Request.Form("premioCorretor"))
    
    ' Validação dos valores obrigatórios
    If Not IsNumeric(valorUnidade) Or Not IsNumeric(comissaoPercentual) Or _
       Not IsNumeric(comissaoDiretoria) Or Not IsNumeric(comissaoGerencia) Or _
       Not IsNumeric(comissaoCorretor) Or Not IsNumeric(premioDiretoria) Or _
       Not IsNumeric(premioGerencia) Or Not IsNumeric(premioCorretor) Then
        Response.Write "<script>alert('Valores numéricos inválidos!');history.back();</script>"
        Response.End
    End If
    
    ' Validação dos percentuais (devem estar entre 0 e 100)
    If comissaoPercentual > 100 Or comissaoDiretoria > 100 Or _
       comissaoGerencia > 100 Or comissaoCorretor > 100 Then
        Response.Write "<script>alert('Os percentuais devem ser valores entre 0 e 100!');history.back();</script>"
        Response.End
    End If
    
    ' Cálculo das comissões com validação
    valorComissaoGeral = FormatarValor(valorComissaoGeral)
    valorComissaoGeral = valorUnidade * (comissaoPercentual / 100)
    
    ' Cálculo dos valores parciais
    valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / 100)
    valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / 100)
    valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / 100)
    
    ' Determinação do trimestre
    If dataVenda <> "" Then
        If trimestre = "" Then
            trimestre = Int((Month(dataVenda) - 1) / 3) + 1
        End If
    Else
        trimestre = Int((Month(Now()) - 1) / 3) + 1
    End If
    
    ' Lógica de ação
    action = Request.Form("action")
    
    If action = "updateVenda" Then
        ' Busca os nomes atuais da diretoria e gerência (do banco principal)
        Set rsNomes = Server.CreateObject("ADODB.Recordset")
        
        ' Busca nome da diretoria
        rsNomes.Open "SELECT NomeDiretoria FROM Diretorias WHERE DiretoriaID = " & CInt(diretoriaId), conn
        If Not rsNomes.EOF Then
            nomeDiretoria = rsNomes("NomeDiretoria")
        Else
            nomeDiretoria = ""
        End If
        rsNomes.Close
        
        ' Busca nome da gerência
        rsNomes.Open "SELECT NomeGerencia FROM Gerencias WHERE GerenciaID = " & CInt(gerenciaId), conn
        If Not rsNomes.EOF Then
            nomeGerencia = rsNomes("NomeGerencia")
        Else
            nomeGerencia = ""
        End If
        rsNomes.Close
        
        'Busca Empreendimento'
        rsNomes.Open "SELECT NomeEmpreendimento FROM Empreendimento WHERE Empreend_ID =" & empreend_id , conn
        If Not rsNomes.EOF Then
            nomeEmpreend = rsNomes("NomeEmpreendimento")
        Else
            nomeEmpreend = ""
        End If
        rsNomes.Close
        Set rsNomes = Nothing

        empreend_id = Request.Form("empreend_id")
        valorComissaoGeral = FormatarValor(valorComissaoGeral)
        valorComissaoDiretoria = FormatarValor(valorComissaoDiretoria)
        valorComissaoGerencia = FormatarValor(valorComissaoGerencia)
        valorComissaoCorretor = FormatarValor(valorComissaoCorretor)
        valorUnidade = FormatarValor(valorUnidade)
        
        ' Atualização segura da tabela Vendas incluindo os nomes E PREMIAÇÕES
        sql = "UPDATE Vendas SET " & _
              "Empreend_ID = " & empreend_id & ", " & _
              "NomeEmpreendimento = '" & Replace(nomeEmpreend, "'", "''") & "', " & _
              "Unidade = '" & Replace(unidade, "'", "''") & "', " & _
              "UnidadeM2 = " & m2 & ", " & _
              "CorretorId = " & CInt(corretorId) & ", " & _
              "ValorUnidade = " & valorUnidade & ", " & _
              "ComissaoPercentual = " & comissaoPercentual & ", " & _
              "ValorComissaoGeral = " & valorComissaoGeral & ", " & _
              "DataVenda = '" & dataVenda & "', " & _
              "DiaVenda = " & Day(dataVenda) & ", " & _
              "MesVenda = " & Month(dataVenda) & ", " & _
              "AnoVenda = " & Year(dataVenda) & ", " & _
              "Trimestre = " & CInt(trimestre) & ", " & _
              "Obs = '" & Replace(obs, "'", "''") & "', " & _
              "DiretoriaId = " & CInt(diretoriaId) & ", " & _
              "Diretoria = '" & Replace(nomeDiretoria, "'", "''") & "', " & _
              "GerenciaId = " & CInt(gerenciaId) & ", " & _
              "Gerencia = '" & Replace(nomeGerencia, "'", "''") & "', " & _
              "ComissaoDiretoria = " & comissaoDiretoria & ", " & _
              "ValorDiretoria = " & valorComissaoDiretoria & ", " & _
              "ComissaoGerencia = " & comissaoGerencia & ", " & _
              "ValorGerencia = " & valorComissaoGerencia & ", " & _
              "ComissaoCorretor = " & comissaoCorretor & ", " & _
              "ValorCorretor = " & valorComissaoCorretor & ", " & _
              "PremioDiretoria = " & premioDiretoria & ", " & _
              "PremioGerencia = " & premioGerencia & ", " & _
              "PremioCorretor = " & premioCorretor & ", " & _
              "Usuario = '" & Session("Usuario") & "' " & _
              "WHERE ID = " & CInt(vendaId)

        On Error Resume Next
        connSales.Execute(sql)
        
        Response.Redirect "gestao_vendas_list3x.asp?mensagem=Venda atualizada com sucesso!"
        
    ElseIf action = "gerarComissoes" Then
        ' Lógica para INSERIR na tabela COMISSOES_A_PAGAR, com VERIFICAÇÃO de duplicidade
        Dim rsCheck
        Set rsCheck = Server.CreateObject("ADODB.Recordset")
        
        ' Consulta para verificar se a comissão já existe para esta venda (no banco de vendas)
        rsCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(vendaId), connSales
        
        If Not rsCheck.EOF Then
            ' Se a comissão já existe, exibe uma mensagem e não insere
            Response.Write "<script>alert('A comissão para esta venda já foi gerada e não pode ser criada novamente.');history.back();</script>"
            rsCheck.Close
            Set rsCheck = Nothing
            Response.End
        Else
            ' Se a comissão não existe, insere o novo registro
            rsCheck.Close
            Set rsCheck = Nothing

            ' Validações
            If IsEmpty(vendaId) Or IsNull(vendaId) Or vendaId = "" Then
                Response.Write "<script>alert('Erro: ID da venda inválido.');history.back();</script>"
                Response.End
            End If
            If IsEmpty(diretoriaId) Or IsNull(diretoriaId) Or diretoriaId = "" Then
                diretoriaId = 0
            End If
            If IsEmpty(gerenciaId) Or IsNull(gerenciaId) Or gerenciaId = "" Then
                gerenciaId = 0
            End If
            If IsEmpty(corretorId) Or IsNull(corretorId) Or corretorId = "" Then
                Response.Write "<script>alert('Erro: ID do corretor inválido.');history.back();</script>"
                Response.End
            End If
            If IsEmpty(dataVenda) Or IsNull(dataVenda) Or dataVenda = "" Then
                Response.Write "<script>alert('Erro: Data de venda inválida.');history.back();</script>"
                Response.End
            End If
            If IsEmpty(unidade) Or IsNull(unidade) Or unidade = "" Then
                Response.Write "<script>alert('Erro: Unidade inválida.');history.back();</script>"
                Response.End
            End If

            ' Arredondar valores decimais para duas casas
            comissaoDiretoria = FormatarValor(comissaoDiretoria)
            comissaoGerencia = FormatarValor(comissaoGerencia)
            comissaoCorretor = FormatarValor(comissaoCorretor)
            valorComissaoDiretoria = FormatarValor(valorComissaoDiretoria)
            valorComissaoGerencia = FormatarValor(valorComissaoGerencia)
            valorComissaoCorretor = FormatarValor(valorComissaoCorretor)
            valorComissaoGeral = FormatarValor(valorComissaoGeral)

            ' Busca os nomes do diretor, gerente, corretor e empreendimento
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
                Response.Write "<script>alert('Erro: Empreendimento não encontrado.');history.back();</script>"
                rsEmp.Close
                Set rsEmp = Nothing
                Response.End
            End If
            rsEmp.Close
            Set rsEmp = Nothing
            Set rsNomes = Nothing

            ' Insere na tabela COMISSOES_A_PAGAR incluindo PREMIAÇÕES
            sql = "INSERT INTO COMISSOES_A_PAGAR (ID_Venda, Empreendimento, Unidade, DataVenda, " & _
                  "UserIdDiretoria, UserIdGerencia, UserIdCorretor, PercDiretoria, ValorDiretoria, " & _
                  "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, " & _
                  "NomeDiretor, NomeGerente, NomeCorretor, PremioDiretoria, PremioGerencia, PremioCorretor) " & _
                  "VALUES (" & CInt(vendaId) & ", '" & Replace(nomeEmpreendimento, "'", "''") & "', '" & Replace(unidade, "'", "''") & "', '" & Replace(dataVenda, "'", "''") & "', " & _
                  CInt(diretoriaId) & ", " & CInt(gerenciaId) & ", " & CInt(corretorId) & ", " & _
                  Replace(CStr(comissaoDiretoria), ",", ".") & ", " & Replace(CStr(valorComissaoDiretoria), ",", ".") & ", " & _
                  Replace(CStr(comissaoGerencia), ",", ".") & ", " & Replace(CStr(valorComissaoGerencia), ",", ".") & ", " & _
                  Replace(CStr(comissaoCorretor), ",", ".") & ", " & Replace(CStr(valorComissaoCorretor), ",", ".") & ", " & _
                  Replace(CStr(valorComissaoGeral), ",", ".") & ", " & _
                  "'" & Replace(nomeDiretor, "'", "''") & "', " & _
                  "'" & Replace(nomeGerente, "'", "''") & "', " & _
                  "'" & Replace(nomeCorretor, "'", "''") & "', " & _
                  Replace(CStr(premioDiretoria), ",", ".") & ", " & _
                  Replace(CStr(premioGerencia), ",", ".") & ", " & _
                  Replace(CStr(premioCorretor), ",", ".") & ")"

            On Error Resume Next
            connSales.Execute(sql)
            If Err.Number <> 0 Then
                Response.Write "<script>alert('Erro ao gerar comissão: " & Replace(Err.Description, "'", "\'") & "');history.back();</script>"
                Response.End
            End If
            On Error GoTo 0

            Response.Redirect "gestao_vendas_list3x.asp?mensagem=Comissão gerada com sucesso!"
        End If
    End If
    
    ' Fecha o recordset de verificação, se ainda estiver aberto
    If IsObject(rsCheck) Then
        rsCheck.Close
        Set rsCheck = Nothing
    End If
End If

' ====================================================================
' BUSCA DE DADOS PARA DROPDOWNS (EXISTENTE)
' ====================================================================
' Busca empreendimentos para o dropdown (do banco principal)
Dim rsEmpreend
Set rsEmpreend = Server.CreateObject("ADODB.Recordset")
rsEmpreend.Open "SELECT Empreend_ID, NomeEmpreendimento, ComissaoVenda FROM Empreendimento ORDER BY NomeEmpreendimento", conn

' Busca diretorias para o dropdown (do banco principal)
Dim rsDiretorias
Set rsDiretorias = Server.CreateObject("ADODB.Recordset")
rsDiretorias.Open "SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria", conn

' Busca gerencias da diretoria selecionada (do banco principal)
Dim rsGerencias
Set rsGerencias = Server.CreateObject("ADODB.Recordset")
rsGerencias.Open "SELECT GerenciaID, NomeGerencia FROM Gerencias WHERE DiretoriaID = " & rsVenda("DiretoriaId") & " ORDER BY NomeGerencia", conn

' Busca corretores para o dropdown (apenas com função "Corretor") (do banco principal)
Dim rsCorretores
Set rsCorretores = Server.CreateObject("ADODB.Recordset")
rsCorretores.Open "SELECT UserId, Nome FROM Usuarios WHERE Funcao = 'Corretor' AND Nome <> '' ORDER BY Nome", conn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Editar Venda</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <!-- Select2 para selects com busca -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    
    <style>
        body {
            background-color: #807777;
            color: #fff;
            padding: 20px;
        }
        .card {
            background-color: #fff;
            color: #000;
            margin-bottom: 20px;
        }
        .card-header {
            background-color: #f8f9fa;
            font-weight: bold;
        }
        .btn-maroon {
            background-color: #800000;
            color: white;
        }
        .btn-maroon:hover {
            background-color: #a00;
            color: white;
        }
        .comissao-result {
            font-weight: bold;
            color: #17a2b8;
        }
        .comissao-dist {
            font-size: 0.9rem;
            color: #6c757d;
        }
        .error-message {
            color: #dc3545;
            font-size: 0.875em;
        }
        
        /* CORREÇÃO PARA OS SELECTS */
        .select2-container--default .select2-selection--single,
        .select2-container--default .select2-selection--multiple {
            background-color: #fff;
            color: #000;
            border: 1px solid #ced4da;
        }
        .select2-container--default .select2-selection--single .select2-selection__rendered {
            color: #000;
        }
        .select2-container--default .select2-selection--single .select2-selection__placeholder {
            color: #6c757d;
        }
        .select2-dropdown {
            background-color: #fff;
            color: #000;
        }
        .select2-container--default .select2-results__option[aria-selected=true] {
            background-color: #f8f9fa;
            color: #000;
        }
        .select2-container--default .select2-results__option--highlighted[aria-selected] {
            background-color: #007bff;
            color: #fff;
        }
        
        /* Estilo adicional para a seção de premiações */
        .premio-card {
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            border-left: 4px solid #ffc107;
        }
        
        .premio-card .card-header {
            background: linear-gradient(to right, #ffc107, #ffb300);
            color: #000;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container" style="padding-top: 70px;">
        <button type="button" onclick="window.close();" class="btn btn-success">
            <i class="fas fa-times me-2"></i>Fechar
        </button>
        <h2 class="mt-4 mb-4"><i class="fas fa-edit"></i> Editar Venda - ID:<%=vendaId%></h2>
        
        <form method="post" id="formVenda">
            <!-- Campos hidden para dia, mês e ano -->
            <input type="hidden" id="diaVenda" name="diaVenda">
            <input type="hidden" id="mesVenda" name="mesVenda">
            <input type="hidden" id="anoVenda" name="anoVenda">
            <input type="hidden" id="vendaId" name="vendaId" value="<%=vendaId%>">
            
            <!-- Card Empreendimento -->
            <div class="card">
                <div class="card-header">Empreendimento</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label for="empreend_id" class="form-label">Empreendimento *</label>
                            <select class="form-select select2" id="empreend_id" name="empreend_id" required style="color: black !important;">
                                <option value="" style="color: black;">Selecione...</option>
                                <%
                                If Not rsEmpreend.EOF Then
                                    rsEmpreend.MoveFirst
                                    Do While Not rsEmpreend.EOF
                                %>
                                    <option value="<%= rsEmpreend("Empreend_ID") %>"
                                        <% If rsEmpreend("Empreend_ID") = rsVenda("Empreend_ID") Then Response.Write "selected" %>
                                        style="color: black;"
                                    >
                                        <%= RemoverNumeros(rsEmpreend("NomeEmpreendimento")) & "-" &rsEmpreend("Empreend_ID") %>
                                    </option>
                                <%
                                    rsEmpreend.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>

                        <div class="col-md-3">
                            <label for="unidade" class="form-label">Unidade *</label>
                            <input type="text" class="form-control" id="unidade" name="unidade" 
                                value="<%= rsVenda("Unidade") %>">
                        </div>
                        <div class="col-md-3">
                            <label for="m2" class="form-label">M² *</label>
                            <input type="text" class="form-control" id="m2" name="m2" 
                                value="<%= FormatNumber(rsVenda("UnidadeM2"), 2) %>">
                        </div>
                    </div>
                    
                    <div class="row">
                        <div class="col-md-6">
                            <label for="valorUnidade" class="form-label">Valor da Unidade (R$) *</label>
                            <input type="text" class="form-control" id="valorUnidade" name="valorUnidade" 
                                value="<%= FormatNumber(rsVenda("ValorUnidade"), 2) %>" required>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoPercentual" class="form-label">% Comissão *</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoPercentual" name="comissaoPercentual" 
                                    value="<%= FormatNumber(rsVenda("ComissaoPercentual"), 2) %>" required>
                                <span class="input-group-text">%</span>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Valor Comissão</label>
                            <div class="form-control comissao-result" id="valorComissaoText">
                                R$ <%= FormatNumber(rsVenda("ValorComissaoGeral"), 2) %>
                            </div>
                            <input type="hidden" id="valorComissaoHidden" name="valorComissao" 
                                value="<%= rsVenda("ValorComissaoGeral") %>">
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Vendido Por -->
            <div class="card">
                <div class="card-header">Vendido Por</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-4">
                            <label for="diretoriaId" class="form-label">Diretoria *</label>
                            <select class="form-select" id="diretoriaId" name="diretoriaId" required>
                                <option value="">Selecione...</option>
                                <%
                                If Not rsDiretorias.EOF Then
                                    rsDiretorias.MoveFirst
                                    Do While Not rsDiretorias.EOF
                                %>
                                    <option value="<%= rsDiretorias("DiretoriaID") %>"
                                        <% If rsDiretorias("DiretoriaID") = rsVenda("DiretoriaID") Then Response.Write "selected" %>
                                    >
                                        <%= rsDiretorias("NomeDiretoria") %>
                                    </option>
                                <%
                                    rsDiretorias.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label for="gerenciaId" class="form-label">Gerência *</label>
                            <select class="form-select" id="gerenciaId" name="gerenciaId" required>
                                <option value="">Selecione...</option>
                                <%
                                If Not rsGerencias.EOF Then
                                    rsGerencias.MoveFirst
                                    Do While Not rsGerencias.EOF
                                %>
                                    <option value="<%= rsGerencias("GerenciaID") %>"
                                        <% If rsGerencias("GerenciaID") = rsVenda("GerenciaID") Then Response.Write "selected" %>
                                    >
                                        <%= rsGerencias("NomeGerencia") %>
                                    </option>
                                <%
                                    rsGerencias.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label for="corretorId" class="form-label">Corretor *</label>
                            <select class="form-select select2" id="corretorId" name="corretorId" required>
                                <option value="">Selecione...</option>
                                <%
                                If Not rsCorretores.EOF Then
                                    rsCorretores.MoveFirst
                                    Do While Not rsCorretores.EOF
                                %>
                                    <option value="<%= rsCorretores("UserId") %>"
                                        <% If rsCorretores("UserId") = rsVenda("CorretorId") Then Response.Write "selected" %>
                                    >
                                        <%= rsCorretores("Nome") %>
                                    </option>
                                <%
                                    rsCorretores.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Comissão -->
            <div class="card">
                <div class="card-header">Distribuição da Comissão</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-3">
                            <label for="comissaoDiretoria" class="form-label">% Diretoria</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoDiretoria" name="comissaoDiretoria" 
                                    value="<%= FormatNumber(rsVenda("ComissaoDiretoria"), 2) %>">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoDiretoria" 
                                value="R$ <%= FormatNumber(rsVenda("ValorDiretoria"), 2) %>" readonly>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoGerencia" class="form-label">% Gerência</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoGerencia" name="comissaoGerencia" 
                                    value="<%= FormatNumber(rsVenda("ComissaoGerencia"), 2) %>">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoGerencia" 
                                value="R$ <%= FormatNumber(rsVenda("ValorGerencia"), 2) %>" readonly>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoCorretor" class="form-label">% Corretor</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoCorretor" name="comissaoCorretor" 
                                    value="<%= FormatNumber(rsVenda("ComissaoCorretor"), 2) %>">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoCorretor" 
                                value="R$ <%= FormatNumber(rsVenda("ValorCorretor"), 2) %>" readonly>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Total Comissão</label>
                            <input type="text" class="form-control" id="valorComissaoSoma" 
                                value="R$ <%= FormatNumber(rsVenda("ValorComissaoGeral"), 2) %>" readonly>
                            <div id="comissaoError" class="error-message mt-2"></div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Premiações - NOVA SEÇÃO -->
            <div class="card premio-card">
                <div class="card-header">Premiações</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-4">
                            <label for="premioDiretoria" class="form-label">Prêmio Diretoria (R$)</label>

                            <%
                                
                                Dim dblPremioDiretoria
                                
                                If IsNull(rsVenda("PremioDiretoria")) Then
                                    dblPremioDiretoria = 0
                                Else
                                    If IsNumeric(rsVenda("PremioDiretoria")) Then
                                        dblPremioDiretoria = CDbl(rsVenda("PremioDiretoria"))
                                    Else
                                        dblPremioDiretoria = 0 ' Trata caso a string não seja numérica
                                    End If
                                End If
                            %>                            
                            <div class="input-group">
                                <span class="input-group-text">R$</span>
                                <input type="text" class="form-control" id="premioDiretoria" name="premioDiretoria" 
                                    value="<%= FormatNumber(dblPremioDiretoria, 2) %>">
                            </div>
                        </div>
                        <%
                        

                        ' 1. Prêmio Gerência
                        Dim dblPremioGerencia
                        If IsNull(rsVenda("PremioGerencia")) Then
                            dblPremioGerencia = 0
                        Else
                            ' Garantia de que o valor é numérico (evita Type Mismatch)
                            If IsNumeric(rsVenda("PremioGerencia")) Then
                                dblPremioGerencia = CDbl(rsVenda("PremioGerencia"))
                            Else
                                dblPremioGerencia = 0
                            End If
                        End If

                        ' 2. Prêmio Corretor
                        Dim dblPremioCorretor
                        If IsNull(rsVenda("PremioCorretor")) Then
                            dblPremioCorretor = 0
                        Else
                            ' Garantia de que o valor é numérico (evita Type Mismatch)
                            If IsNumeric(rsVenda("PremioCorretor")) Then
                                dblPremioCorretor = CDbl(rsVenda("PremioCorretor"))
                            Else
                                dblPremioCorretor = 0
                            End If
                        End If

                        %>

                        <div class="col-md-4">
                            <label for="premioGerencia" class="form-label">Prêmio Gerência (R$)</label>
                            <div class="input-group">
                                <span class="input-group-text">R$</span>
                                <input type="text" class="form-control" id="premioGerencia" name="premioGerencia" 
                                    value="<%= FormatNumber(dblPremioGerencia, 2) %>">
                            </div>
                        </div>
                        <div class="col-md-4">
                            <label for="premioCorretor" class="form-label">Prêmio Corretor (R$)</label>
                            <div class="input-group">
                                <span class="input-group-text">R$</span>
                                <input type="text" class="form-control" id="premioCorretor" name="premioCorretor" 
                                    value="<%= FormatNumber(dblPremioCorretor, 2) %>">
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Outras Informações -->
            <div class="card">
                <div class="card-header">Outras Informações</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-3">
                            <label for="dataVenda" class="form-label">Data da Venda *</label>
                            <input type="date" class="form-control" id="dataVenda" name="dataVenda"
                                value="<%= Year(rsVenda("DataVenda")) %>-<%= Right("0" & Month(rsVenda("DataVenda")), 2) %>-<%= Right("0" & Day(rsVenda("DataVenda")), 2) %>" required>
                        </div>
                        <div class="col-md-3">
                            <label for="trimestre" class="form-label">Trimestre</label>
                            <select class="form-select" id="trimestre" name="trimestre">
                                <option value="">Selecione...</option>
                                <option value="1" <% If rsVenda("Trimestre") = 1 Then Response.Write "selected" %> >1º Trimestre</option>
                                <option value="2" <% If rsVenda("Trimestre") = 2 Then Response.Write "selected" %> >2º Trimestre</option>
                                <option value="3" <% If rsVenda("Trimestre") = 3 Then Response.Write "selected" %> >3º Trimestre</option>
                                <option value="4" <% If rsVenda("Trimestre") = 4 Then Response.Write "selected" %> >4º Trimestre</option>
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label for="obs" class="form-label">Observações</label>
                            <textarea class="form-control" id="obs" name="obs" rows="3"><%= rsVenda("Obs") %></textarea>
                        </div>
                    </div>
                    
                    <div class="d-grid gap-2 d-md-flex justify-content-md-end">
                        <button type="submit" name="action" value="updateVenda" class="btn btn-success">
                            <i class="fas fa-save"></i> Atualizar Venda
                        </button>
                    </div>
                </div>
            </div>
        </form>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- jQuery e jQuery Mask -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.16/jquery.mask.min.js"></script>
    
    <!-- Select2 -->
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/i18n/pt-BR.js"></script>
    
    <script>
        $(document).ready(function() {
            // Inicializa select2 nos selects
            $('.select2').select2({
                language: "pt-BR",
                placeholder: "Selecione...",
                allowClear: true
            });
            
            // Máscaras para os campos - INCLUINDO PREMIAÇÕES
            $('#valorUnidade').mask('#.##0,00', {reverse: true});
            $('#comissaoPercentual, #comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').mask('##0,00', {reverse: true});
            $('#m2').mask('#0,00', {reverse: true});
            
            // NOVAS MÁSCARAS PARA PREMIAÇÕES
            $('#premioDiretoria, #premioGerencia, #premioCorretor').mask('#.##0,00', {reverse: true});
            
            // Formata campos monetários como somente leitura
            $('#valorComissaoDiretoria, #valorComissaoGerencia, #valorComissaoCorretor').mask('#.##0,00', {reverse: true, readonly: true});
            
            // Carrega gerencias quando seleciona diretoria
            $('#diretoriaId').change(function() {
                var diretoriaId = $(this).val();
                if (diretoriaId) {
                    $('#gerenciaId').prop('disabled', false);
                    $.getJSON('get_gerencias.asp', {diretoriaId: diretoriaId}, function(data) {
                        var options = '<option value="">Selecione...</option>';
                        $.each(data, function(key, val) {
                            options += '<option value="' + val.GerenciaID + '">' + val.NomeGerencia + '</option>';
                        });
                        $('#gerenciaId').html(options);
                    });
                } else {
                    $('#gerenciaId').prop('disabled', true).html('<option value="">Selecione uma diretoria primeiro</option>');
                }
            });
            
            // Preenche comissão padrão quando seleciona empreendimento
            $('#empreend_id').change(function() {
                var selected = $(this).find('option:selected');
                var comissao = selected.data('comissao');
                if (comissao) {
                    $('#comissaoPercentual').val(comissao.toString().replace('.', ',')).trigger('input');
                }
            });
            
            // Atualiza dia, mês, ano e trimestre quando seleciona data
            $('#dataVenda').change(function() {
                var data = new Date($(this).val());
                if (!isNaN(data.getTime())) {
                    $('#diaVenda').val(data.getDate());
                    $('#mesVenda').val(data.getMonth() + 1);
                    $('#anoVenda').val(data.getFullYear());
                    
                    // Calcula o trimestre
                    var mes = data.getMonth() + 1;
                    var trimestre = Math.floor((mes - 1) / 3) + 1;
                    $('#trimestre').val(trimestre);
                }
            });
            
            // Função para validar números
            function validarNumero(valor) {
                valor = valor.replace(/\./g, '').replace(',', '.');
                return !isNaN(parseFloat(valor)) && isFinite(valor);
            }
            
            // Calcula a comissão
            function calcularComissoes() {
                try {
                    // Valores principais
                    var valorInput = $('#valorUnidade').val();
                    var percentualInput = $('#comissaoPercentual').val();
                    
                    // Remove pontos e substitui vírgula por ponto para cálculo
                    var valor = parseFloat(valorInput.replace(/\./g, '').replace(',', '.')) || 0;
                    var percentual = parseFloat(percentualInput.replace(',', '.')) || 0;
                    
                    // Cálculo CORRETO da comissão total (5% de 100.000 = 5.000)
                    var comissaoTotal = valor * (percentual / 100);
                    
                    // Valores das comissões parciais
                    var percDiretoria = parseFloat($('#comissaoDiretoria').val().replace(',', '.')) || 0;
                    var percGerencia = parseFloat($('#comissaoGerencia').val().replace(',', '.')) || 0;
                    var percCorretor = parseFloat($('#comissaoCorretor').val().replace(',', '.')) || 0;
                    
                    // Cálculo dos valores parciais
                    var comissaoDiretoria = comissaoTotal * (percDiretoria / 100);
                    var comissaoGerencia = comissaoTotal * (percGerencia / 100);
                    var comissaoCorretor = comissaoTotal * (percCorretor / 100);
                    
                    // Soma das comissões parciais
                    var totalDistribuido = comissaoDiretoria + comissaoGerencia + comissaoCorretor;
                    
                    // Validação do total distribuído
                    var diferenca = Math.abs(comissaoTotal - totalDistribuido);
                    if (diferenca > 0.01) {
                        $('#comissaoError').text('');
                    } else {
                        $('#comissaoError').text('');
                    }
                    
                    // Formata os valores para exibição
                    $('#valorComissaoText').text('R$ ' + comissaoTotal.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoHidden').val(comissaoTotal.toFixed(2));
                    
                    $('#valorComissaoDiretoria').val(
                        'R$ ' + comissaoDiretoria.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoGerencia').val(
                        'R$ ' + comissaoGerencia.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoCorretor').val(
                        'R$ ' + comissaoCorretor.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoSoma').val(
                        'R$ ' + totalDistribuido.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );

                } catch(e) {
                    console.error("Erro no cálculo:", e);
                }
            }
            
            // Configura os eventos para cálculo automático
            $('#valorUnidade, #comissaoPercentual').on('input change', calcularComissoes);
            $('#comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').on('input change', calcularComissoes);
            
            // Calcula a comissão inicial
            calcularComissoes();
        });
    </script>
</body>
</html>
<%
' Fecha conexões
If IsObject(rsVenda) Then
    rsVenda.Close
    Set rsVenda = Nothing
End If

If IsObject(rsEmpreend) Then
    rsEmpreend.Close
    Set rsEmpreend = Nothing
End If

If IsObject(rsDiretorias) Then
    rsDiretorias.Close
    Set rsDiretorias = Nothing
End If

If IsObject(rsGerencias) Then
    rsGerencias.Close
    Set rsGerencias = Nothing
End If

If IsObject(rsCorretores) Then
    rsCorretores.Close
    Set rsCorretores = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If
%>