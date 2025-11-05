<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if   
%>
<%
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

' Função auxiliar para formatar valores
Function FormatarValor(valor)
    If IsNull(valor) Or valor = "" Then
        FormatarValor = "0"
        Exit Function
    End If
    valor = Replace(valor, ".", ",")
    valor = Replace(valor, ",", ".")
    FormatarValor = valor
End Function

' Função para converter valores monetários corretamente
Function ParseCurrency(value)
    On Error Resume Next
    If IsNull(value) Or value = "" Then
        ParseCurrency = 0
        Exit Function
    End If
    If IsNumeric(value) Then
        ParseCurrency = CDbl(value)
        Exit Function
    End If
    ParseCurrency = CDbl(Replace(Replace(Replace(value, ".", ""), ",", ".")))
    If Err.Number <> 0 Then ParseCurrency = 0
    On Error GoTo 0
End Function

' Cria as conexões
Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Buscar todas as vendas que ainda não têm comissões geradas
Dim sqlVendasPendentes, rsVendasPendentes
sqlVendasPendentes = "SELECT v.* " & _
                    "FROM Vendas v " & _
                    "LEFT JOIN COMISSOES_A_PAGAR c ON v.ID = c.ID_Venda " & _
                    "WHERE c.ID_Venda IS NULL " & _
                    "AND v.excluido = 0 " & _
                    "ORDER BY v.DataVenda DESC"

Set rsVendasPendentes = connSales.Execute(sqlVendasPendentes)

Dim comissoesInseridas, comissoesComErro
comissoesInseridas = 0
comissoesComErro = 0
Dim detalhesInseridas, detalhesErros
detalhesInseridas = ""
detalhesErros = ""

If Not rsVendasPendentes.EOF Then
    Do While Not rsVendasPendentes.EOF
        Dim vendaId, empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
        Dim dataVenda, obs, m2, diretoriaId, gerenciaId, trimestre
        Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor
        Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
        Dim nomeDiretor, nomeGerente, nomeCorretor, nomeEmpreendimento
        Dim premioDiretoria, premioGerencia, premioCorretor
        
        ' Obter dados da venda
        vendaId = rsVendasPendentes("ID")
        empreend_id = rsVendasPendentes("Empreend_ID")
        unidade = Server.HTMLEncode(rsVendasPendentes("Unidade"))
        corretorId = rsVendasPendentes("CorretorId")
        diretoriaId = rsVendasPendentes("DiretoriaId")
        gerenciaId = rsVendasPendentes("GerenciaId")
        trimestre = rsVendasPendentes("Trimestre")
        dataVenda = rsVendasPendentes("DataVenda")
        obs = Server.HTMLEncode(rsVendasPendentes("Obs"))
        valorUnidade = ParseCurrency(rsVendasPendentes("ValorUnidade"))
        m2 = ParseCurrency(rsVendasPendentes("UnidadeM2"))
        
        comissaoPercentual = ParseCurrency(rsVendasPendentes("ComissaoPercentual"))
        comissaoDiretoria = ParseCurrency(rsVendasPendentes("ComissaoDiretoria"))
        comissaoGerencia = ParseCurrency(rsVendasPendentes("ComissaoGerencia"))
        comissaoCorretor = ParseCurrency(rsVendasPendentes("ComissaoCorretor"))
        
        ' Obter valores de premiação
        premioDiretoria = ParseCurrency(rsVendasPendentes("PremioDiretoria"))
        premioGerencia = ParseCurrency(rsVendasPendentes("PremioGerencia"))
        premioCorretor = ParseCurrency(rsVendasPendentes("PremioCorretor"))
        
        ' Cálculo das comissões
        valorComissaoGeral = valorUnidade * (comissaoPercentual / 100)
        valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / 100)
        valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / 100)
        valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / 100)
        
        ' Validações básicas - usando estrutura If/Else
        Dim podeProcessar
        podeProcessar = True
        
        If IsNull(vendaId) Or vendaId = "" Then
            detalhesErros = detalhesErros & "<li>Venda ID inválido</li>"
            comissoesComErro = comissoesComErro + 1
            podeProcessar = False
        ElseIf IsNull(corretorId) Or corretorId = "" Then
            detalhesErros = detalhesErros & "<li>Venda " & vendaId & ": ID do corretor inválido</li>"
            comissoesComErro = comissoesComErro + 1
            podeProcessar = False
        ElseIf IsNull(dataVenda) Or dataVenda = "" Then
            detalhesErros = detalhesErros & "<li>Venda " & vendaId & ": Data de venda inválida</li>"
            comissoesComErro = comissoesComErro + 1
            podeProcessar = False
        End If
        
        ' Se não pode processar, vai para próxima venda
        If Not podeProcessar Then
            rsVendasPendentes.MoveNext
        Else
            ' Definir valores padrão para IDs nulos
            If IsNull(diretoriaId) Or diretoriaId = "" Then diretoriaId = 0
            If IsNull(gerenciaId) Or gerenciaId = "" Then gerenciaId = 0
            
            ' Arredondar valores decimais
            comissaoDiretoria = FormatarValor(comissaoDiretoria)
            comissaoGerencia = FormatarValor(comissaoGerencia)
            comissaoCorretor = FormatarValor(comissaoCorretor)
            valorComissaoDiretoria = FormatarValor(valorComissaoDiretoria)
            valorComissaoGerencia = FormatarValor(valorComissaoGerencia)
            valorComissaoCorretor = FormatarValor(valorComissaoCorretor)
            valorComissaoGeral = FormatarValor(valorComissaoGeral)
            premioDiretoria = FormatarValor(premioDiretoria)
            premioGerencia = FormatarValor(premioGerencia)
            premioCorretor = FormatarValor(premioCorretor)
            
            ' Buscar nomes
            Dim rsNomes
            Set rsNomes = Server.CreateObject("ADODB.Recordset")
            
            ' Busca nome do diretor
            nomeDiretor = ""
            If diretoriaId > 0 Then
                rsNomes.Open "SELECT u.Nome FROM Usuarios u INNER JOIN Diretorias d ON u.UserId = d.UserId WHERE d.DiretoriaID = " & CInt(diretoriaId), conn
                If Not rsNomes.EOF Then
                    nomeDiretor = rsNomes("Nome")
                    If IsNull(nomeDiretor) Then nomeDiretor = ""
                End If
                rsNomes.Close
            End If
            
            ' Busca nome do gerente
            nomeGerente = ""
            If gerenciaId > 0 Then
                rsNomes.Open "SELECT u.Nome FROM Usuarios u INNER JOIN Gerencias g ON u.UserId = g.UserId WHERE g.GerenciaID = " & CInt(gerenciaId), conn
                If Not rsNomes.EOF Then
                    nomeGerente = rsNomes("Nome")
                    If IsNull(nomeGerente) Then nomeGerente = ""
                End If
                rsNomes.Close
            End If
            
            ' Busca nome do corretor
            nomeCorretor = ""
            rsNomes.Open "SELECT Nome FROM Usuarios WHERE UserId = " & CInt(corretorId), conn
            If Not rsNomes.EOF Then
                nomeCorretor = rsNomes("Nome")
                If IsNull(nomeCorretor) Then nomeCorretor = ""
            End If
            rsNomes.Close
            
            ' Busca nome do empreendimento
            nomeEmpreendimento = ""
            Dim rsEmp
            Set rsEmp = Server.CreateObject("ADODB.Recordset")
            rsEmp.Open "SELECT NomeEmpreendimento FROM Empreendimento WHERE Empreend_ID = " & empreend_id, conn
            If Not rsEmp.EOF Then
                nomeEmpreendimento = rsEmp("NomeEmpreendimento")
                If IsNull(nomeEmpreendimento) Then nomeEmpreendimento = ""
            End If
            rsEmp.Close
            Set rsEmp = Nothing
            Set rsNomes = Nothing
            
            ' Preparar SQL para inserção
            Dim sql
            sql = "INSERT INTO COMISSOES_A_PAGAR (ID_Venda, Empreend_ID, Empreendimento, Unidade, DataVenda, " & _
                  "UserIdDiretoria, UserIdGerencia, UserIdCorretor, PercDiretoria, ValorDiretoria, " & _
                  "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, " & _
                  "NomeDiretor, NomeGerente, NomeCorretor, " & _
                  "PremioDiretoria, PremioGerencia, PremioCorretor, StatusPagamento) " & _
                  "VALUES (" & CInt(vendaId) & ", " & CInt(empreend_id) & ", '" & Replace(nomeEmpreendimento, "'", "''") & "', '" & Replace(unidade, "'", "''") & "', '" & Replace(dataVenda, "'", "''") & "', " & _
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
                  Replace(CStr(premioCorretor), ",", ".") & ", 'PENDENTE')"
            
            ' Executar inserção
            On Error Resume Next
            connSales.Execute(sql)
            If Err.Number = 0 Then
                comissoesInseridas = comissoesInseridas + 1
                detalhesInseridas = detalhesInseridas & "<li>Venda " & vendaId & " - " & nomeEmpreendimento & " - " & unidade & " - R$ " & FormatNumber(valorComissaoGeral, 2) & "</li>"
            Else
                comissoesComErro = comissoesComErro + 1
                detalhesErros = detalhesErros & "<li>Venda " & vendaId & ": " & Replace(Err.Description, "'", "") & "</li>"
            End If
            On Error GoTo 0
            
            ' Move para próxima venda
            rsVendasPendentes.MoveNext
        End If
    Loop
Else
    ' Nenhuma venda pendente encontrada
    comissoesInseridas = 0
    comissoesComErro = 0
End If

rsVendasPendentes.Close
Set rsVendasPendentes = Nothing

' Fecha conexões
If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If
If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerar Todas as Comissões Pendentes</title>
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #f8f9fa;
            padding: 20px;
        }
        .container {
            background-color: #fff;
            border-radius: 10px;
            box-shadow: 0 0 15px rgba(0,0,0,0.1);
            padding: 30px;
            margin-top: 20px;
        }
        .header-title {
            color: #800000;
            border-bottom: 2px solid #800000;
            padding-bottom: 15px;
            margin-bottom: 25px;
        }
        .success-box {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .error-box {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .info-box {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .comissao-item {
            border-left: 4px solid #28a745;
            padding: 10px 15px;
            margin-bottom: 10px;
            background-color: #f8f9fa;
        }
        .erro-item {
            border-left: 4px solid #dc3545;
            padding: 10px 15px;
            margin-bottom: 10px;
            background-color: #f8f9fa;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row">
            <div class="col-12">
                <h1 class="text-center header-title">
                    <i class="fas fa-cogs me-2"></i>Gerar Todas as Comissões Pendentes
                </h1>
            </div>
        </div>

        <div class="row mb-4">
            <div class="col-md-6">
                <a href="gestao_vendas_list2x.asp" class="btn btn-primary">
                    <i class="fas fa-arrow-left me-2"></i>Voltar para Lista de Vendas
                </a>
            </div>
            <div class="col-md-6 text-end">
                <a href="gestao_vendas_gerenc_comissoes.asp" class="btn btn-info">
                    <i class="fas fa-coins me-2"></i>Ver Comissões
                </a>
            </div>
        </div>

        <!-- Resumo Geral -->
        <div class="row mb-4">
            <div class="col-md-4">
                <div class="card text-white bg-primary">
                    <div class="card-body text-center">
                        <h4><i class="fas fa-list"></i></h4>
                        <h5>Total Processado</h5>
                        <h3><%= comissoesInseridas + comissoesComErro %></h3>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card text-white bg-success">
                    <div class="card-body text-center">
                        <h4><i class="fas fa-check-circle"></i></h4>
                        <h5>Comissões Inseridas</h5>
                        <h3><%= comissoesInseridas %></h3>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card text-white bg-danger">
                    <div class="card-body text-center">
                        <h4><i class="fas fa-exclamation-triangle"></i></h4>
                        <h5>Comissões com Erro</h5>
                        <h3><%= comissoesComErro %></h3>
                    </div>
                </div>
            </div>
        </div>

        <!-- Comissões Inseridas com Sucesso -->
        <% If comissoesInseridas > 0 Then %>
        <div class="success-box">
            <h4><i class="fas fa-check-circle text-success me-2"></i>Comissões Inseridas com Sucesso</h4>
            <p><strong><%= comissoesInseridas %></strong> comissões foram geradas com sucesso:</p>
            <ul class="list-unstyled">
                <%= detalhesInseridas %>
            </ul>
        </div>
        <% End If %>

        <!-- Comissões com Erro -->
        <% If comissoesComErro > 0 Then %>
        <div class="error-box">
            <h4><i class="fas fa-exclamation-triangle text-danger me-2"></i>Comissões com Erro</h4>
            <p><strong><%= comissoesComErro %></strong> comissões não puderam ser geradas:</p>
            <ul class="list-unstyled">
                <%= detalhesErros %>
            </ul>
        </div>
        <% End If %>

        <!-- Nenhuma Comissão Pendente -->
        <% If comissoesInseridas = 0 And comissoesComErro = 0 Then %>
        <div class="info-box">
            <h4><i class="fas fa-info-circle text-info me-2"></i>Nenhuma Comissão Pendente</h4>
            <p>Não foram encontradas vendas sem comissões geradas. Todas as vendas já têm suas comissões registradas.</p>
        </div>
        <% End If %>

        <!-- Ações -->
        <div class="row mt-4">
            <div class="col-12 text-center">
                <a href="gestao_vendas_gerar_todas_comissoes.asp" class="btn btn-warning me-2">
                    <i class="fas fa-sync-alt me-2"></i>Executar Novamente
                </a>
                <a href="gestao_vendas_gerenc_comissoes.asp" class="btn btn-success">
                    <i class="fas fa-coins me-2"></i>Ver Todas as Comissões
                </a>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>