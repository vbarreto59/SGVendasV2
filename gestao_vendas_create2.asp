<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="registra_log.asp"-->

<% ' funcional - incluir desconto em 06 11 2025
    Function RemoverNumeros(texto)
        Dim regex
        Set regex = New RegExp
        
        ' Remove números (0-9) e asteriscos (*)
        regex.Pattern = "[0-9*-]"
        regex.Global = True
        
        RemoverNumerosEAsteriscos = regex.Replace(texto, "")
        
        ' Remove espaços extras (opcional)
        RemoverNumeros = Trim(Replace(RemoverNumerosEAsteriscos, "  ", " "))
    End Function    

Function FormatNumberForSQL(sValue)
    ' Remove o separador de milhares (o ponto)
    sValue = Replace(sValue, ".", "")
    ' Substitui o separador decimal (a vírgula) por um ponto
    sValue = Replace(sValue, ",", ".")
    FormatNumberForSQL = sValue
End Function    
%>

<%
' -----------------------------------------------------------------------------------
' INICIALIZAÇÃO E CONEXÃO COM BANCOS DE DADOS
' -----------------------------------------------------------------------------------
' Verifica se as strings de conexão estão configuradas.
If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "Erro: Conexões com bancos de dados não configuradas"
    Response.End
End If

' Cria e abre as conexões com os bancos de dados.
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' -----------------------------------------------------------------------------------
' PROCESSAMENTO DO FORMULÁRIO (MÉTODO POST)
' -----------------------------------------------------------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Declaração das variáveis.
    Dim empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
    Dim dataVenda, obs, usuario, m2, diretoriaId, gerenciaId
    Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor, trimestre
    Dim nomeEmpreendimento, corretorNome, diretoriaNome, gerenciaNome
    Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
    Dim sqlVendas, sqlComissoes, vendaId
    
    ' CAMPOS DE PREMIAÇÃO ADICIONADOS
    Dim premioDiretoria, premioGerencia, premioCorretor
    
    ' Coleta e formatação dos dados do formulário.
    ' A função `GetFormattedNumber` centraliza a lógica de formatação.
    empreend_id = Request.Form("empreend_id")
    unidade = Request.Form("unidade")
    corretorId = Request.Form("corretorId")
    diretoriaId = Request.Form("diretoriaId")
    gerenciaId = Request.Form("gerenciaId")
    trimestre = Request.Form("trimestre")
    dataVenda = Request.Form("dataVenda")
    obs = Request.Form("obs")
    m2 = GetFormattedNumber(Request.Form("m2"))
    valorUnidade = GetFormattedNumber(Request.Form("valorUnidade"))
    comissaoPercentual = GetFormattedNumber(Request.Form("comissaoPercentual"))
    comissaoDiretoria = GetFormattedNumber(Request.Form("comissaoDiretoria"))
    comissaoGerencia = GetFormattedNumber(Request.Form("comissaoGerencia"))
    comissaoCorretor = GetFormattedNumber(Request.Form("comissaoCorretor"))

    ' Premiação em 04 11 2025 -----------------------------
    ' Coleta dos valores monetários da premiação. A divisão por 100 foi removida,
    ' assumindo que GetFormattedNumber já entrega o valor numérico (ex: 10000.00).
    premioDiretoria = GetFormattedNumber(Request.Form("premioDiretoria"))    
    premioGerencia  = GetFormattedNumber(Request.Form("premioGerencia"))    
    premioCorretor  = GetFormattedNumber(Request.Form("premioCorretor"))    
    '----------------------'

    usuario = Session("Usuario")
    
    ' Validação de dados numéricos essenciais.
    If Not IsNumeric(valorUnidade) Or Not IsNumeric(comissaoPercentual) Then
        Response.Write "<script>alert('Valores inválidos!');history.back();</script>"
        Response.End
    End If

    ' A função `GetDataFromDB` centraliza a busca de dados no banco,
    ' evitando a repetição de código para cada Recordset.
    nomeEmpreendimento = GetDataFromDB(conn, "Empreendimento", "NomeEmpreendimento", "Empreend_ID", empreend_id)
    corretorNome = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", corretorId)
    diretoriaNome = GetDataFromDB(conn, "Diretorias", "NomeDiretoria", "DiretoriaID", diretoriaId)
    
    ' Trata o caso onde a gerência pode não ser selecionada.
    If gerenciaId <> "" And IsNumeric(gerenciaId) Then
        gerenciaNome = GetDataFromDB(conn, "Gerencias", "NomeGerencia", "GerenciaID", gerenciaId)
    Else
        gerenciaNome = "Não aplicável"
        gerenciaId = 0
    End If
    
    ' Extrai ano, mês, dia e calcula o trimestre da data de venda.
    Dim anoVenda, mesVenda, diaVenda
    If Trim(dataVenda) <> "" Then
        anoVenda = Year(dataVenda)
        mesVenda = Month(dataVenda)
        diaVenda = Day(dataVenda)
        If Trim(trimestre) = "" Then trimestre = Int((mesVenda - 1) / 3) + 1
    Else
        ' Se a data de venda está vazia, usa a data e hora atuais.
        dataVenda = Now()
        anoVenda = Year(dataVenda)
        mesVenda = Month(dataVenda)
        diaVenda = Day(dataVenda)
        trimestre = Int((mesVenda - 1) / 3) + 1
    End If

    ' Formatação da data para o SQL.
    dataVendaSQL = FormatDateTimeForSQL(dataVenda)
    dataRegistroSQL = FormatDateTimeForSQL(Now())

    ' Cálculo das comissões.
    vFatorDivisao = 10000
    valorComissaoGeral = valorUnidade * (comissaoPercentual / vFatorDivisao)
    valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / vFatorDivisao)
    valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / vFatorDivisao)
    valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / vFatorDivisao)

    ' -----------------------------------------------------------------------------------
    ' INSERÇÃO NO BANCO DE DADOS
    ' -----------------------------------------------------------------------------------
    ' Inserção na tabela Vendas.
    sqlVendas = "INSERT INTO Vendas (" & _
    "Empreend_ID, NomeEmpreendimento, Unidade, UnidadeM2, Corretor, CorretorId, " & _
    "ValorUnidade, ComissaoPercentual, ValorComissaoGeral, DataVenda, " & _
    "DiaVenda, MesVenda, AnoVenda, Trimestre, Obs, Usuario, " & _
    "DiretoriaId, Diretoria, GerenciaId, Gerencia, " & _
    "ComissaoDiretoria, ValorDiretoria, " & _
    "ComissaoGerencia, ValorGerencia, " & _
    "ComissaoCorretor, ValorCorretor, " & _
    "PremioDiretoria, PremioGerencia, PremioCorretor) " & _
    "VALUES (" & empreend_id & ", '" & SanitizeSQL(nomeEmpreendimento) & "', " & _
    "'" & SanitizeSQL(unidade) & "', " & m2 & ", " & _
    "'" & SanitizeSQL(corretorNome) & "', " & corretorId & ", " & _
    valorUnidade & ", " & comissaoPercentual & ", " & valorComissaoGeral & ", " & dataVendaSQL & ", " & _
    diaVenda & ", " & mesVenda & ", " & anoVenda & ", " & trimestre & ", " & _
    "'" & SanitizeSQL(obs) & "', '" & SanitizeSQL(usuario) & "', " & _
    diretoriaId & ", '" & SanitizeSQL(diretoriaNome) & "', " & gerenciaId & ", " & _
    "'" & SanitizeSQL(gerenciaNome) & "', " & comissaoDiretoria & ", " & valorComissaoDiretoria & ", " & _
    comissaoGerencia & ", " & valorComissaoGerencia & ", " & comissaoCorretor & ", " & valorComissaoCorretor & ", " & _
    premioDiretoria & ", " & premioGerencia & ", " & premioCorretor & ")"

    connSales.Execute(sqlVendas)

    ' Obtém o ID da venda recém-inserida.
    Set rsLastID = connSales.Execute("SELECT MAX(ID) AS NewID FROM Vendas")
    If Not rsLastID.EOF Then vendaId = rsLastID("NewID")
    rsLastID.Close
    
    '-------- Inserção na tabela COMISSOES_A_PAGAR. (COM PRÊMIOS INCLUÍDOS)
    sqlComissoes = "INSERT INTO COMISSOES_A_PAGAR (" & _
    "ID_Venda, Empreendimento, Unidade, DataVenda, UserIdDiretoria, NomeDiretor, " & _
    "UserIdGerencia, NomeGerente, UserIdCorretor, NomeCorretor, PercDiretoria, ValorDiretoria, " & _
    "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, StatusPagamento, Usuario, " & _
    "PremioDiretoria, PremioGerencia, PremioCorretor) " & _
    "VALUES (" & vendaId & ", '" & SanitizeSQL(nomeEmpreendimento) & "', '" & SanitizeSQL(unidade) & "', " & _
    dataVendaSQL & ", " & diretoriaId & ", '" & SanitizeSQL(diretoriaNome) & "', " & gerenciaId & ", " & _
    "'" & SanitizeSQL(gerenciaNome) & "', " & corretorId & ", '" & SanitizeSQL(corretorNome) & "', " & _
    comissaoDiretoria & ", " & valorComissaoDiretoria & ", " & comissaoGerencia & ", " & valorComissaoGerencia & ", " & _
    comissaoCorretor & ", " & valorComissaoCorretor & ", " & valorComissaoGeral & ", 'PENDENTE', '" & SanitizeSQL(usuario) & "', " & _
    premioDiretoria & ", " & premioGerencia & ", " & premioCorretor & ")"    

    connSales.Execute(sqlComissoes)
    
    '============================= LOG ============================================'
    if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
        set objMail = server.createobject("CDONTS.NewMail")
            objMail.From = "sendmail@gabnetweb.com.br"
            objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
        objMail.Subject = "SV-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
        objMail.MailFormat = 0
        objMail.Body = "Nova venda. " & sqlVendas
        objMail.Send
        set objMail = Nothing
    end if 
    '----------- fim envio de email'
    ' registrar log'
    Call InserirLog ("VENDAS", "INSERT", "Nova venda inserida ID: " & vendaId )
    
    ' Redireciona para a página de sucesso após a inserção.
    Response.Redirect "gestao_vendas_list2x.asp?mensagem=Venda cadastrada com sucesso!"
End If


' -----------------------------------------------------------------------------------
' BUSCA DE DADOS PARA DROPDOWNS (MÉTODO GET)
' -----------------------------------------------------------------------------------
' Busca e popula os recordsets para os dropdowns na página.
Set rsEmpreend = conn.Execute("SELECT Empreend_ID, NomeEmpreendimento, ComissaoVenda FROM Empreendimento ORDER BY NomeEmpreendimento")
Set rsDiretorias = conn.Execute("SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria")
Set rsCorretores = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Funcao = 'Corretor' AND Nome <> '' ORDER BY Nome")
%>

<% ' -----------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES
' ----------------------------------------------------------------------------------- %>
<%
' Função para formatar números, removendo pontos e substituindo vírgulas por pontos.
Function GetFormattedNumber(sValue)
    GetFormattedNumber = Replace(Replace(sValue, ".", ""), ",", ".")
End Function

' Função para buscar dados de uma tabela com base em um critério.
Function GetDataFromDB(oConn, sTable, sField, sWhereField, sWhereValue)
    Dim sResult
    Set rs = oConn.Execute("SELECT " & sField & " FROM " & sTable & " WHERE " & sWhereField & " = " & sWhereValue)
    If Not rs.EOF Then
        sResult = rs(sField)
    Else
        sResult = "Desconhecido"
    End If
    rs.Close
    Set rs = Nothing
    GetDataFromDB = sResult
End Function

' Função para formatar a data para o formato SQL.
Function FormatDateTimeForSQL(dDate)
    If Trim(dDate & "") = "" Then
        FormatDateTimeForSQL = "NULL"
    Else
        FormatDateTimeForSQL = "'" & FormatDateTime(dDate, 2) & "'"
    End If
End Function

' Função para sanitizar strings, escapando aspas simples.
Function SanitizeSQL(sValue)
    SanitizeSQL = Replace(sValue, "'", "''")
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="300">
    <title>Nova Venda | Sistema</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <!-- Select2 para selects com busca -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --success: #27ae60;
            --warning: #f39c12;
            --danger: #e74c3c;
            --light-bg: #f8f9fa;
            --dark-text: #2c3e50;
            --card-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            --hover-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
        }
        
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: var(--dark-text);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 20px;
        }
        
        .app-container {
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1.5rem;
            border-radius: 12px 12px 0 0;
            margin-bottom: 0;
            box-shadow: var(--card-shadow);
        }
        
        .app-title {
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.8rem;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: var(--card-shadow);
            transition: transform 0.3s, box-shadow 0.3s;
            margin-bottom: 1.5rem;
            overflow: hidden;
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
        }
        
        .card:hover {
            transform: translateY(-2px);
            box-shadow: var(--hover-shadow);
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1.2rem 1.5rem;
            font-weight: 600;
            font-size: 1.1rem;
        }
        
        .card-body {
            padding: 2rem;
        }
        
        .form-label {
            font-weight: 600;
            color: var(--primary);
            margin-bottom: 0.5rem;
        }
        
        .form-control, .form-select {
            border: 2px solid #e9ecef;
            border-radius: 8px;
            padding: 0.75rem 1rem;
            font-size: 0.95rem;
            transition: all 0.3s ease;
        }
        
        .form-control:focus, .form-select:focus {
            border-color: var(--secondary);
            box-shadow: 0 0 0 0.2rem rgba(52, 152, 219, 0.25);
        }
        
        .input-group-text {
            background-color: var(--primary);
            color: white;
            border: 2px solid var(--primary);
            font-weight: 600;
        }
        
        .comissao-result {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 700;
            font-size: 1.1rem;
            border-radius: 8px;
            padding: 0.75rem;
            text-align: center;
        }
        
        .comissao-dist {
            font-size: 0.9rem;
            color: #6c757d;
            font-weight: 500;
        }
        
        .error-message {
            color: var(--danger);
            font-size: 0.875em;
            font-weight: 500;
        }
        
        .btn {
            border-radius: 8px;
            padding: 0.75rem 1.5rem;
            font-weight: 600;
            transition: all 0.3s ease;
            border: none;
        }
        
        .btn-success {
            background: linear-gradient(135deg, var(--success), #2ecc71);
            box-shadow: 0 4px 15px rgba(39, 174, 96, 0.3);
        }
        
        .btn-success:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(39, 174, 96, 0.4);
        }
        
        .btn-secondary {
            background: linear-gradient(135deg, #6c757d, #868e96);
            box-shadow: 0 4px 15px rgba(108, 117, 125, 0.3);
        }
        
        .btn-secondary:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(108, 117, 125, 0.4);
        }
        
        .required-field::after {
            content: " *";
            color: var(--danger);
        }
        
        .comissao-card {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border-left: 4px solid var(--secondary);
        }
        
        .comissao-value {
            font-weight: 700;
            color: var(--primary);
            font-size: 1.1rem;
        }
        
        /* Select2 Custom Styles */
        .select2-container--default .select2-selection--single,
        .select2-container--default .select2-selection--multiple {
            border: 2px solid #e9ecef;
            border-radius: 8px;
            padding: 0.5rem;
            background-color: #fff;
            color: var(--dark-text);
            transition: all 0.3s ease;
        }
        
        .select2-container--default .select2-selection--single:focus,
        .select2-container--default .select2-selection--multiple:focus {
            border-color: var(--secondary);
            box-shadow: 0 0 0 0.2rem rgba(52, 152, 219, 0.25);
        }
        
        .select2-container--default .select2-selection--single .select2-selection__rendered {
            color: var(--dark-text);
            font-size: 0.95rem;
        }
        
        .select2-container--default .select2-selection--single .select2-selection__placeholder {
            color: #6c757d;
        }
        
        .select2-dropdown {
            border: 2px solid var(--secondary);
            border-radius: 8px;
            box-shadow: var(--hover-shadow);
        }
        
        .select2-container--default .select2-results__option[aria-selected=true] {
            background-color: #e3f2fd;
            color: var(--primary);
        }
        
        .select2-container--default .select2-results__option--highlighted[aria-selected] {
            background-color: var(--secondary);
            color: white;
        }
        
        .form-section {
            margin-bottom: 2rem;
        }
        
        .section-title {
            color: var(--primary);
            font-weight: 600;
            margin-bottom: 1.5rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #e9ecef;
        }
        
        @media (max-width: 768px) {
            .card-body {
                padding: 1.5rem;
            }
            
            .app-title {
                font-size: 1.4rem;
            }
            
            .btn {
                width: 100%;
                margin-bottom: 0.5rem;
            }
        }
        
        .floating-action {
            position: fixed;
            bottom: 2rem;
            right: 2rem;
            z-index: 1000;
        }
        
        .info-badge {
            background: linear-gradient(135deg, var(--secondary), #2980b9);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
    </style>
</head>
<body>
    <div class="app-container">
        <!-- Header -->
        <div class="app-header">
            <div class="d-flex justify-content-between align-items-center">
                <h1 class="app-title">
                    <i class="fas fa-plus-circle"></i> Nova Venda
                </h1>
                <div class="info-badge">
                    <i class="fas fa-user me-1"></i><%= Session("Usuario") %>
                </div>
            </div>
        </div>

        <!-- Conteúdo Principal -->
        <div class="card">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <button type="button" onclick="window.close();" class="btn btn-secondary">
                        <i class="fas fa-arrow-left me-2"></i>Voltar
                    </button>
                    <div class="d-flex gap-2">
                        <a href="gestao_vendas_list2x.asp" class="btn btn-secondary">
                            <i class="fas fa-list me-2"></i>Lista de Vendas
                        </a>
                    </div>
                </div>

                <form method="post" id="formVenda">
                    <!-- Campos hidden para dia, mês e ano -->
                    <input type="hidden" id="diaVenda" name="diaVenda">
                    <input type="hidden" id="mesVenda" name="mesVenda">
                    <input type="hidden" id="anoVenda" name="anoVenda">
                    
                    <!-- Seção Empreendimento -->
                    <div class="form-section">
                        <h3 class="section-title">
                            <i class="fas fa-building me-2"></i>Informações do Empreendimento
                        </h3>
                        <div class="row g-3">
                            <div class="col-md-6">
                                <label for="empreend_id" class="form-label required-field">Empreendimento</label>
                                <select class="form-select select2" id="empreend_id" name="empreend_id" required>
                                    <option value="">Selecione o empreendimento...</option>
                                    <% 
                                    If Not rsEmpreend.EOF Then
                                        rsEmpreend.MoveFirst
                                        Do While Not rsEmpreend.EOF 
                                    %>
                                        <option value="<%= rsEmpreend("Empreend_ID") %>" data-comissao="<%= rsEmpreend("ComissaoVenda") %>">
                                            <%= RemoverNumeros(rsEmpreend("NomeEmpreendimento")) %>
                                        </option>
                                    <%
                                            rsEmpreend.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                            <div class="col-md-3">
                                <label for="unidade" class="form-label required-field">Unidade</label>
                                <input type="text" class="form-control" id="unidade" name="unidade" placeholder="Ex: 101A">
                            </div>
                            <div class="col-md-3">
                                <label for="m2" class="form-label required-field">Metragem (m²)</label>
                                <input type="text" class="form-control" id="m2" name="m2" value="0" placeholder="Ex: 75,00">
                            </div>
                        </div>
                        
                        <div class="row g-3 mt-2">
                            <div class="col-md-6">
                                <label for="valorUnidade" class="form-label required-field">Valor da Unidade</label>
                                <input type="text" class="form-control" id="valorUnidade" name="valorUnidade" placeholder="R$ 0,00" required>
                            </div>
                            <div class="col-md-3">
                                <label for="comissaoPercentual" class="form-label required-field">Percentual de Comissão</label>
                                <div class="input-group">
                                    <input type="text" class="form-control" id="comissaoPercentual" name="comissaoPercentual" placeholder="0,00" required>
                                    <span class="input-group-text">%</span>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">Valor da Comissão</label>
                                <div class="comissao-result" id="valorComissaoText">R$ 0,00</div>
                                <input type="hidden" id="valorComissaoHidden" name="valorComissao">
                            </div>
                        </div>
                    </div>
                    
                    <!-- Seção Equipe de Vendas -->
                    <div class="form-section">
                        <h3 class="section-title">
                            <i class="fas fa-users me-2"></i>Equipe de Vendas
                        </h3>
                        <div class="row g-3">
                            <div class="col-md-4">
                                <label for="diretoriaId" class="form-label required-field">Diretoria</label>
                                <select class="form-select" id="diretoriaId" name="diretoriaId" required>
                                    <option value="">Selecione a diretoria...</option>
                                    <% 
                                    If Not rsDiretorias.EOF Then
                                        rsDiretorias.MoveFirst
                                        Do While Not rsDiretorias.EOF 
                                    %>
                                        <option value="<%= rsDiretorias("DiretoriaID") %>"><%= rsDiretorias("NomeDiretoria") %></option>
                                    <%
                                            rsDiretorias.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                            <div class="col-md-4">
                                <label for="gerenciaId" class="form-label required-field">Gerência</label>
                                <select class="form-select" id="gerenciaId" name="gerenciaId" required disabled>
                                    <option value="">Selecione uma diretoria primeiro</option>
                                </select>
                            </div>
                            <div class="col-md-4">
                                <label for="corretorId" class="form-label required-field">Corretor</label>
                                <select class="form-select select2" id="corretorId" name="corretorId" required>
                                    <option value="">Selecione o corretor...</option>
                                    <% 
                                    If Not rsCorretores.EOF Then
                                        rsCorretores.MoveFirst
                                        Do While Not rsCorretores.EOF 
                                    %>
                                        <option value="<%= rsCorretores("UserId") %>"><%= rsCorretores("Nome") %></option>
                                    <%
                                            rsCorretores.MoveNext
                                        Loop
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Seção Distribuição de Comissões -->
                    <div class="form-section">
                        <h3 class="section-title">
                            <i class="fas fa-chart-pie me-2"></i>Distribuição de Comissões
                        </h3>
                        <div class="card comissao-card">
                            <div class="card-body">
                                <div class="row g-3">
                                    <div class="col-md-3">
                                        <label for="comissaoDiretoria" class="form-label">Diretoria</label>
                                        <div class="input-group">
                                            <input type="text" class="form-control" id="comissaoDiretoria" name="comissaoDiretoria" value="5,00">
                                            <span class="input-group-text">%</span>
                                        </div>
                                        <div class="comissao-value mt-2" id="valorComissaoDiretoriaText">R$ 0,00</div>
                                        <input type="hidden" id="valorComissaoDiretoria" name="valorComissaoDiretoria">
                                    </div>
                                    <div class="col-md-3">
                                        <label for="comissaoGerencia" class="form-label">Gerência</label>
                                        <div class="input-group">
                                            <input type="text" class="form-control" id="comissaoGerencia" name="comissaoGerencia" value="10,00">
                                            <span class="input-group-text">%</span>
                                        </div>
                                        <div class="comissao-value mt-2" id="valorComissaoGerenciaText">R$ 0,00</div>
                                        <input type="hidden" id="valorComissaoGerencia" name="valorComissaoGerencia">
                                    </div>
                                    <div class="col-md-3">
                                        <label for="comissaoCorretor" class="form-label">Corretor</label>
                                        <div class="input-group">
                                            <input type="text" class="form-control" id="comissaoCorretor" name="comissaoCorretor" value="35,00">
                                            <span class="input-group-text">%</span>
                                        </div>
                                        <div class="comissao-value mt-2" id="valorComissaoCorretorText">R$ 0,00</div>
                                        <input type="hidden" id="valorComissaoCorretor" name="valorComissaoCorretor">
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Total Distribuído</label>
                                        <div class="comissao-result" id="valorComissaoSomaText">R$ 0,00</div>
                                        <input type="hidden" id="valorComissaoSoma" name="valorComissaoSoma">
                                        <div id="comissaoError" class="error-message mt-2"></div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- #### Distribuicao de Premios 04 11 2025 -->
                    <div class="form-section">
                        <h3 class="section-title">
                            <i class="fas fa-trophy me-2"></i>Premiações
                        </h3>
                        <div class="card premio-card">
                            <div class="card-body">
                                <div class="row g-3">
                                    <div class="col-md-4">
                                        <label for="premioDiretoria" class="form-label">Prêmio Diretoria</label>
                                        <div class="input-group">
                                            <span class="input-group-text">R$</span>
                                            <input type="text" class="form-control" id="premioDiretoria" name="premioDiretoria" value="0,00">
                                        </div>
                                    </div>
                                    <div class="col-md-4">
                                        <label for="premioGerencia" class="form-label">Prêmio Gerência</label>
                                        <div class="input-group">
                                            <span class="input-group-text">R$</span>
                                            <input type="text" class="form-control" id="premioGerencia" name="premioGerencia" value="0,00">
                                        </div>
                                    </div>
                                    <div class="col-md-4">
                                        <label for="premioCorretor" class="form-label">Prêmio Corretor</label>
                                        <div class="input-group">
                                            <span class="input-group-text">R$</span>
                                            <input type="text" class="form-control" id="premioCorretor" name="premioCorretor" value="0,00">
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <!-- ################## -->
                    
                    <!-- Seção Informações Adicionais -->
                    <div class="form-section">
                        <h3 class="section-title">
                            <i class="fas fa-info-circle me-2"></i>Informações Adicionais
                        </h3>
                        <div class="row g-3">
                            <div class="col-md-3">
                                <label for="dataVenda" class="form-label required-field">Data da Venda</label>
                                <input type="date" class="form-control" id="dataVenda" name="dataVenda" required>
                            </div>
                            <div class="col-md-3">
                                <label for="trimestre" class="form-label">Trimestre</label>
                                <select class="form-select" id="trimestre" name="trimestre">
                                    <option value="">Selecione o trimestre...</option>
                                    <option value="1">1º Trimestre</option>
                                    <option value="2">2º Trimestre</option>
                                    <option value="3">3º Trimestre</option>
                                    <option value="4">4º Trimestre</option>
                                </select>
                            </div>
                            <div class="col-md-6">
                                <label for="obs" class="form-label">Observações</label>
                                <textarea class="form-control" id="obs" name="obs" rows="3" placeholder="Observações adicionais sobre a venda..."></textarea>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Botões de Ação -->
                    <div class="d-grid gap-2 d-md-flex justify-content-md-end mt-4">
                        <a href="gestao_vendas_list2x.asp" class="btn btn-secondary me-md-2">
                            <i class="fas fa-times me-2"></i>Cancelar
                        </a>
                        <button type="submit" class="btn btn-success">
                            <i class="fas fa-save me-2"></i>Salvar Venda
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- jQuery e jQuery Mask -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.16/jquery.mask.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-maskmoney/3.0.2/jquery.maskMoney.min.js"></script>    
    
    <!-- Select2 -->
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/i18n/pt-BR.js"></script>
    


<script>
    // 3. Aplicação da máscara
    $(document).ready(function() {
        $('#premioDiretoria, #premioGerencia, #premioCorretor').maskMoney({
            allowNegative: false,
            thousands: '.',
            decimal: ',',
            affixesStay: true
        });
        
        // Importante: para que a máscara funcione corretamente no primeiro carregamento, 
        // é uma boa prática chamar o trigger no carregamento.
        $('#premioDiretoria').trigger('mask.maskMoney');
        $('#premioGerencia').trigger('mask.maskMoney');
        $('#premioCorretor').trigger('mask.maskMoney');
    });
</script>

    <script>
        $(document).ready(function() {
            // Inicializa select2 nos selects
            $('.select2').select2({
                language: "pt-BR",
                placeholder: "Selecione...",
                allowClear: true,
                width: '100%'
            });
            
            // Máscaras para os campos
            $('#valorUnidade').mask('#.##0,00', {reverse: true});
            $('#comissaoPercentual, #comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').mask('##0,00', {reverse: true});
            $('#m2').mask('#0,00', {reverse: true});
            
            // Carrega gerencias quando seleciona diretoria
            $('#diretoriaId').change(function() {
                var diretoriaId = $(this).val();
                if (diretoriaId) {
                    $('#gerenciaId').prop('disabled', false);
                    $.getJSON('get_gerencias.asp', {diretoriaId: diretoriaId}, function(data) {
                        var options = '<option value="">Selecione a gerência...</option>';
                        $.each(data, function(key, val) {
                            options += '<option value="' + val.GerenciaID + '">' + val.NomeGerencia + '</option>';
                        });
                        $('#gerenciaId').html(options);
                    }).fail(function() {
                        $('#gerenciaId').html('<option value="">Erro ao carregar gerencias</option>');
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
                    
                    // Cálculo da comissão total
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
                    
                    $('#valorComissaoDiretoriaText').text('R$ ' + comissaoDiretoria.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoDiretoria').val(comissaoDiretoria.toFixed(2));
                    
                    $('#valorComissaoGerenciaText').text('R$ ' + comissaoGerencia.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoGerencia').val(comissaoGerencia.toFixed(2));
                    
                    $('#valorComissaoCorretorText').text('R$ ' + comissaoCorretor.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoCorretor').val(comissaoCorretor.toFixed(2));
                    
                    $('#valorComissaoSomaText').text('R$ ' + totalDistribuido.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoSoma').val(totalDistribuido.toFixed(2));

                } catch(e) {
                    console.error("Erro no cálculo:", e);
                }
            }
            
            // Configura os eventos para cálculo automático
            $('#valorUnidade, #comissaoPercentual').on('input change', calcularComissoes);
            $('#comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').on('input change', calcularComissoes);
            
            // Calcula a comissão inicial
            calcularComissoes();

            // Define a data atual como padrão
            var today = new Date().toISOString().split('T')[0];
            $('#dataVenda').val(today).trigger('change');
        });
    </script>
</body>
</html>
<%
' Fecha conexões e recordsets
If IsObject(rsEmpreend) Then
    rsEmpreend.Close
    Set rsEmpreend = Nothing
End If

If IsObject(rsDiretorias) Then
    rsDiretorias.Close
    Set rsDiretorias = Nothing
End If

If IsObject(rsCorretores) Then
    rsCorretores.Close
    Set rsCorretores = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>