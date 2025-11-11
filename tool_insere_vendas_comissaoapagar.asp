<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

' Função auxiliar para formatar valores
Function FormatarValor(valor)
    If IsNull(valor) Or valor = "" Then
        FormatarValor = "0"
    Else
        valor = Replace(valor, ".", ",")
        valor = Replace(valor, ",", ".")
        FormatarValor = valor
    End If
End Function

' Função para converter valores monetários corretamente
Function ParseCurrency(value)
    On Error Resume Next
    If IsNumeric(value) Then
        ParseCurrency = CDbl(value)
        Exit Function
    End If
    If IsNull(value) Or value = "" Then
        ParseCurrency = 0
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

' Variáveis para estatísticas
Dim totalVendas, vendasProcessadas, comissoesInseridas, erros
totalVendas = 0
vendasProcessadas = 0
comissoesInseridas = 0
erros = 0

' Array para armazenar os status inseridos
Dim statusInseridos()
ReDim statusInseridos(0)

' Busca TODAS as vendas que ainda não têm comissão gerada
Dim rsVendas
Set rsVendas = Server.CreateObject("ADODB.Recordset")
Dim sqlVendas

sqlVendas = "SELECT v.* FROM Vendas v " & _
            "LEFT JOIN COMISSOES_A_PAGAR cp ON v.ID = cp.ID_Venda " & _
            "WHERE cp.ID_Venda IS NULL AND v.Excluido = 0 " & _
            "ORDER BY v.DataVenda, v.ID"

rsVendas.Open sqlVendas, connSales

If rsVendas.EOF Then
    ' COMENTADO: Redirecionamento automático
    ' Response.Write "<script>alert('Não há vendas pendentes para processar.');window.location.href='gestao_vendas_list3x.asp';</script>"
    ' rsVendas.Close
    ' Set rsVendas = Nothing
    ' Response.End
%>
    <!DOCTYPE html>
    <html>
    <head>
        <title>Processamento de Comissões</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
        <div class="container mt-4">
            <div class="alert alert-warning">
                <h4>Não há vendas pendentes para processar.</h4>
            </div>
            <a href="gestao_vendas_list3x.asp" class="btn btn-primary">Voltar para Lista de Vendas</a>
        </div>
    </body>
    </html>
<%
    rsVendas.Close
    Set rsVendas = Nothing
    Response.End
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Processamento de Comissões - Detalhes</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .status-pendente { background-color: #fff3cd; }
        .status-paga { background-color: #d1ecf1; }
        .status-cancelada { background-color: #f8d7da; }
        .table-hover tbody tr:hover { background-color: rgba(0,0,0,.075); }
    </style>
</head>
<body>
    <div class="container mt-4">
        <div class="card">
            <div class="card-header bg-primary text-white">
                <h4><i class="fas fa-cogs"></i> Processamento de Comissões em Lote</h4>
            </div>
            <div class="card-body">
<%

' Processa cada venda
Do While Not rsVendas.EOF
    totalVendas = totalVendas + 1
    Dim vendaId
    vendaId = rsVendas("ID")
    
    ' Verifica se a comissão já existe para esta venda
    Dim rsCheck
    Set rsCheck = Server.CreateObject("ADODB.Recordset")
    rsCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(vendaId), connSales
    
    If rsCheck.EOF Then
        ' Comissão não existe, pode inserir
        On Error Resume Next
        
        ' Obtém os dados da venda atual
        Dim empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
        Dim dataVenda, obs, m2, diretoriaId, gerenciaId, trimestre
        Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor
        Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
        Dim nomeDiretor, nomeGerente, nomeCorretor, nomeEmpreendimento
        Dim premioDiretoria, premioGerencia, premioCorretor

        empreend_id = rsVendas("Empreend_ID")
        unidade = Server.HTMLEncode(rsVendas("Unidade"))
        corretorId = rsVendas("CorretorId")
        diretoriaId = rsVendas("DiretoriaId")
        gerenciaId = rsVendas("GerenciaId")
        trimestre = rsVendas("Trimestre")
        dataVenda = rsVendas("DataVenda")
        obs = Server.HTMLEncode(rsVendas("Obs"))
        valorUnidade = ParseCurrency(rsVendas("ValorUnidade"))
        m2 = ParseCurrency(rsVendas("UnidadeM2"))

        comissaoPercentual = ParseCurrency(rsVendas("ComissaoPercentual"))
        comissaoDiretoria = ParseCurrency(rsVendas("ComissaoDiretoria"))
        comissaoGerencia = ParseCurrency(rsVendas("ComissaoGerencia"))
        comissaoCorretor = ParseCurrency(rsVendas("ComissaoCorretor"))

        ' Obtém os valores de premiação
        premioDiretoria = ParseCurrency(rsVendas("PremioDiretoria"))
        premioGerencia = ParseCurrency(rsVendas("PremioGerencia"))
        premioCorretor = ParseCurrency(rsVendas("PremioCorretor"))

        ' Cálculo das comissões
        valorComissaoGeral = valorUnidade * (comissaoPercentual / 100)
        valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / 100)
        valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / 100)
        valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / 100)

        ' Validações
        If IsEmpty(diretoriaId) Or IsNull(diretoriaId) Or diretoriaId = "" Then
            diretoriaId = 0
        End If
        If IsEmpty(gerenciaId) Or IsNull(gerenciaId) Or gerenciaId = "" Then
            gerenciaId = 0
        End If

        ' Arredondar valores decimais
        comissaoDiretoria = FormatarValor(comissaoDiretoria)
        comissaoGerencia = FormatarValor(comissaoGerencia)
        comissaoCorretor = FormatarValor(comissaoCorretor)
        valorComissaoDiretoria = FormatarValor(valorComissaoDiretoria)
        valorComissaoGerencia = FormatarValor(valorComissaoGerencia)
        valorComissaoCorretor = FormatarValor(valorComissaoCorretor)
        valorComissaoGeral = FormatarValor(valorComissaoGeral)

        ' Formatar valores de premiação
        premioDiretoria = FormatarValor(premioDiretoria)
        premioGerencia = FormatarValor(premioGerencia)
        premioCorretor = FormatarValor(premioCorretor)

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
        End If
        rsEmp.Close
        Set rsEmp = Nothing
        Set rsNomes = Nothing

        ' DEFINE O STATUS PAGAMENTO EXPLICITAMENTE
        Dim statusPagamento
        statusPagamento = "Pendente" ' Status fixo para novas inserções

        ' Insere na tabela COMISSOES_A_PAGAR COM STATUS
        Dim sql
        sql = "INSERT INTO COMISSOES_A_PAGAR (ID_Venda, Empreend_ID, Empreendimento, Unidade, DataVenda, " & _
              "UserIdDiretoria, UserIdGerencia, UserIdCorretor, PercDiretoria, ValorDiretoria, " & _
              "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, " & _
              "NomeDiretor, NomeGerente, NomeCorretor, StatusPagamento, " & _  
              "PremioDiretoria, PremioGerencia, PremioCorretor) " & _
              "VALUES (" & CInt(vendaId) & ", " & CInt(empreend_id) & ", '" & Replace(nomeEmpreendimento, "'", "''") & "', '" & Replace(unidade, "'", "''") & "', '" & Replace(dataVenda, "'", "''") & "', " & _
              CInt(diretoriaId) & ", " & CInt(gerenciaId) & ", " & CInt(corretorId) & ", " & _
              Replace(CStr(comissaoDiretoria), ",", ".") & ", " & Replace(CStr(valorComissaoDiretoria), ",", ".") & ", " & _
              Replace(CStr(comissaoGerencia), ",", ".") & ", " & Replace(CStr(valorComissaoGerencia), ",", ".") & ", " & _
              Replace(CStr(comissaoCorretor), ",", ".") & ", " & Replace(CStr(valorComissaoCorretor), ",", ".") & ", " & _
              Replace(CStr(valorComissaoGeral), ",", ".") & ", " & _
              "'" & Replace(nomeDiretor, "'", "''") & "', " & _
              "'" & Replace(nomeGerente, "'", "''") & "', " & _
              "'" & Replace(nomeCorretor, "'", "''") & "', " & _
              "'" & statusPagamento & "', " & _  
              Replace(CStr(premioDiretoria), ",", ".") & ", " & _
              Replace(CStr(premioGerencia), ",", ".") & ", " & _
              Replace(CStr(premioCorretor), ",", ".") & ")"

        ' Executa a inserção
        connSales.Execute(sql)
        
        If Err.Number = 0 Then
            comissoesInseridas = comissoesInseridas + 1
            
            ' Armazena o status inserido para exibição
            ReDim Preserve statusInseridos(comissoesInseridas)
            statusInseridos(comissoesInseridas) = statusPagamento
            
            ' Exibe detalhes da venda processada
            Response.Write "<div class='alert alert-success'>" & _
                          "<strong>✓ Venda " & vendaId & " processada:</strong> " & _
                          nomeEmpreendimento & " - " & unidade & " - " & _
                          "Status: <span class='badge bg-warning'>" & statusPagamento & "</span>" & _
                          "</div>"
        Else
            erros = erros + 1
            Response.Write "<div class='alert alert-danger'>" & _
                          "<strong>✗ Erro na venda " & vendaId & ":</strong> " & Err.Description & _
                          "</div>"
            Err.Clear
        End If
        On Error GoTo 0
        
    Else
        ' Comissão já existe, apenas conta como processada
        vendasProcessadas = vendasProcessadas + 1
        Response.Write "<div class='alert alert-info'>" & _
                      "<strong>ℹ Venda " & vendaId & ":</strong> Comissão já existente" & _
                      "</div>"
    End If
    
    rsCheck.Close
    Set rsCheck = Nothing
    
    vendasProcessadas = vendasProcessadas + 1
    rsVendas.MoveNext
Loop

' Fecha recordsets
rsVendas.Close
Set rsVendas = Nothing

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
            </div>
        </div>

        <!-- RESUMO DO PROCESSAMENTO -->
        <div class="card mt-4">
            <div class="card-header bg-info text-white">
                <h5><i class="fas fa-chart-bar"></i> Resumo do Processamento</h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h3 class="text-primary"><%= totalVendas %></h3>
                                <p class="text-muted">Vendas Encontradas</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h3 class="text-success"><%= comissoesInseridas %></h3>
                                <p class="text-muted">Comissões Inseridas</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h3 class="text-warning"><%= vendasProcessadas %></h3>
                                <p class="text-muted">Vendas Processadas</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card text-center">
                            <div class="card-body">
                                <h3 class="text-danger"><%= erros %></h3>
                                <p class="text-muted">Erros</p>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- DETALHES DOS STATUS -->
                <div class="mt-4">
                    <h5>Status de Pagamento Inseridos:</h5>
                    <table class="table table-bordered">
                        <thead>
                            <tr>
                                <th>Status</th>
                                <th>Quantidade</th>
                                <th>Observação</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr class="status-pendente">
                                <td><strong>Pendente</strong></td>
                                <td><%= comissoesInseridas %></td>
                                <td>Status padrão para novas comissões</td>
                            </tr>
                            <%
                            ' Verifica se há outros status no banco (para referência)
                            Set connSales = Server.CreateObject("ADODB.Connection")
                            connSales.Open StrConnSales
                            Dim rsStatus
                            Set rsStatus = connSales.Execute("SELECT StatusPagamento, COUNT(*) as Total FROM COMISSOES_A_PAGAR GROUP BY StatusPagamento")
                            
                            Do While Not rsStatus.EOF
                                If rsStatus("StatusPagamento") <> "Pendente" Then
                            %>
                            <tr class="status-<%= LCase(rsStatus("StatusPagamento")) %>">
                                <td><strong><%= rsStatus("StatusPagamento") %></strong></td>
                                <td><%= rsStatus("Total") %></td>
                                <td>Status existente no banco (não modificado)</td>
                            </tr>
                            <%
                                End If
                                rsStatus.MoveNext
                            Loop
                            rsStatus.Close
                            Set rsStatus = Nothing
                            connSales.Close
                            Set connSales = Nothing
                            %>
                        </tbody>
                    </table>
                </div>

                <!-- BOTÕES DE AÇÃO -->
                <div class="mt-4">
                    <a href="gestao_vendas_list3x.asp" class="btn btn-primary">
                        <i class="fas fa-arrow-left"></i> Voltar para Lista de Vendas
                    </a>
                    <a href="processar_comissoes_lote.asp" class="btn btn-success">
                        <i class="fas fa-redo"></i> Executar Novamente
                    </a>
                    <a href="ver_comissoes.asp" class="btn btn-info">
                        <i class="fas fa-eye"></i> Ver Comissões
                    </a>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <!-- Font Awesome -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/js/all.min.js"></script>
</body>
</html>