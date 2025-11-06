<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% 
Response.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!--#include file="conSunSales.asp"-->
<!--#include file="registra_log.asp"-->

<%
' Função para criar diretório se não existir
Function CriarDiretorio(caminho)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(caminho) Then
        fso.CreateFolder(caminho)
    End If
    
    Set fso = Nothing
End Function

' Função para sanitizar strings para JSON
Function SanitizeJSON(str)
    If IsNull(str) Then
        SanitizeJSON = ""
        Exit Function
    End If
    
    str = CStr(str)
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbTab, "\t")
    SanitizeJSON = str
End Function

' Função para formatar data para JSON
Function FormatDateForJSON(dt)
    If IsNull(dt) Or dt = "" Then
        FormatDateForJSON = ""
        Exit Function
    End If
    
    FormatDateForJSON = Year(dt) & "-" & _
                       Right("0" & Month(dt), 2) & "-" & _
                       Right("0" & Day(dt), 2) & "T" & _
                       Right("0" & Hour(dt), 2) & ":" & _
                       Right("0" & Minute(dt), 2) & ":" & _
                       Right("0" & Second(dt), 2) & "Z"
End Function

' Função para gerar JSON de uma venda
Function GerarJSONVenda(rs)
    Dim json
    json = "{"
    
    ' Dados básicos da venda
    json = json & """id"": " & rs("ID") & ","
    json = json & """empreendimento"": {"
    json = json & """id"": " & rs("Empreend_ID") & ","
    json = json & """nome"": """ & SanitizeJSON(rs("NomeEmpreendimento")) & """"
    json = json & "},"
    
    ' Informações da unidade
    json = json & """unidade"": {"
    json = json & """nome"": """ & SanitizeJSON(rs("Unidade")) & ""","
    json = json & """metragem"": " & rs("UnidadeM2") & ","
    json = json & """valor"": " & rs("ValorUnidade")
    json = json & "},"
    
    ' Data da venda
    json = json & """data_venda"": """ & FormatDateForJSON(rs("DataVenda")) & ""","
    json = json & """trimestre"": " & rs("Trimestre") & ","
    json = json & """ano"": " & rs("AnoVenda") & ","
    
    ' Informações de comissão
    json = json & """comissao"": {"
    json = json & """percentual"": " & rs("ComissaoPercentual") & ","
    json = json & """valor_total"": " & rs("ValorComissaoGeral") & ","
    json = json & """distribuicao"": {"
    json = json & """diretoria"": {"
    json = json & """percentual"": " & rs("ComissaoDiretoria") & ","
    json = json & """valor"": " & rs("ValorDiretoria")
    json = json & "},"
    json = json & """gerencia"": {"
    json = json & """percentual"": " & rs("ComissaoGerencia") & ","
    json = json & """valor"": " & rs("ValorGerencia")
    json = json & "},"
    json = json & """corretor"": {"
    json = json & """percentual"": " & rs("ComissaoCorretor") & ","
    json = json & """valor"": " & rs("ValorCorretor")
    json = json & "}"
    json = json & "}"
    json = json & "},"
    
    ' Premiações
    json = json & """premiacao"": {"
    json = json & """diretoria"": " & rs("PremioDiretoria") & ","
    json = json & """gerencia"": " & rs("PremioGerencia") & ","
    json = json & """corretor"": " & rs("PremioCorretor")
    json = json & "},"
    
    ' Equipe de vendas
    json = json & """equipe"": {"
    json = json & """diretoria"": {"
    json = json & """id"": " & rs("DiretoriaId") & ","
    json = json & """nome"": """ & SanitizeJSON(rs("Diretoria")) & """"
    json = json & "},"
    json = json & """gerencia"": {"
    json = json & """id"": " & rs("GerenciaId") & ","
    json = json & """nome"": """ & SanitizeJSON(rs("Gerencia")) & """"
    json = json & "},"
    json = json & """corretor"": {"
    json = json & """id"": " & rs("CorretorId") & ","
    json = json & """nome"": """ & SanitizeJSON(rs("Corretor")) & """"
    json = json & "}"
    json = json & "},"
    
    ' Informações adicionais
    json = json & """observacoes"": """ & SanitizeJSON(rs("Obs")) & ""","
    json = json & """usuario_registro"": """ & SanitizeJSON(rs("Usuario")) & ""","
    json = json & """data_registro"": """ & FormatDateForJSON(Now()) & """"
    
    json = json & "}"
    
    GerarJSONVenda = json
End Function

' Função principal para gerar todos os JSONs
Function GerarJSONsVendas()
    On Error Resume Next
    
    Dim conn, rs, fso, arquivo
    Dim pastaJSON, caminhoArquivo, jsonContent
    Dim contador, vendaId
    
    contador = 0
    
    ' Criar conexão
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open StrConnSales
    
    If Err.Number <> 0 Then
        GerarJSONsVendas = "Erro na conexão: " & Err.Description
        Exit Function
    End If
    
    ' Criar objeto FileSystemObject
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    ' Definir pasta JSON
    pastaJSON = Server.MapPath("json")
    
    ' Criar pasta se não existir
    Call CriarDiretorio(pastaJSON)
    
    ' Consultar todas as vendas
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM Vendas ORDER BY ID", conn, 1, 1
    
    If Err.Number <> 0 Then
        GerarJSONsVendas = "Erro ao consultar vendas: " & Err.Description
        Exit Function
    End If
    
    Do While Not rs.EOF
        vendaId = rs("ID")
        
        ' Gerar conteúdo JSON
        jsonContent = GerarJSONVenda(rs)
        
        ' Definir caminho do arquivo
        caminhoArquivo = pastaJSON & "\venda-" & vendaId & ".json"
        
        ' Criar e escrever no arquivo
        Set arquivo = fso.CreateTextFile(caminhoArquivo, True)
        arquivo.Write jsonContent
        arquivo.Close
        
        contador = contador + 1
        
        rs.MoveNext
    Loop
    
    ' Fechar conexões
    rs.Close
    conn.Close
    
    ' Limpar objetos
    Set arquivo = Nothing
    Set fso = Nothing
    Set rs = Nothing
    Set conn = Nothing
    
    GerarJSONsVendas = "JSONs gerados com sucesso! Total: " & contador & " arquivos criados na pasta 'json'"
    
    On Error GoTo 0
End Function

' Função para gerar JSON de uma venda específica
Function GerarJSONVendaEspecifica(vendaId)
    On Error Resume Next
    
    Dim conn, rs, fso, arquivo
    Dim pastaJSON, caminhoArquivo, jsonContent
    
    ' Criar conexão
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open StrConnSales
    
    If Err.Number <> 0 Then
        GerarJSONVendaEspecifica = "Erro na conexão: " & Err.Description
        Exit Function
    End If
    
    ' Consultar venda específica
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM Vendas WHERE ID = " & vendaId, conn, 1, 1
    
    If rs.EOF Then
        GerarJSONVendaEspecifica = "Venda não encontrada: ID " & vendaId
        rs.Close
        conn.Close
        Exit Function
    End If
    
    ' Criar objeto FileSystemObject
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    ' Definir pasta JSON
    pastaJSON = Server.MapPath("json")
    
    ' Criar pasta se não existir
    Call CriarDiretorio(pastaJSON)
    
    ' Gerar conteúdo JSON
    jsonContent = GerarJSONVenda(rs)
    
    ' Definir caminho do arquivo
    caminhoArquivo = pastaJSON & "\venda-" & vendaId & ".json"
    
    ' Criar e escrever no arquivo
    Set arquivo = fso.CreateTextFile(caminhoArquivo, True)
    arquivo.Write jsonContent
    arquivo.Close
    
    ' Fechar conexões
    rs.Close
    conn.Close
    
    ' Limpar objetos
    Set arquivo = Nothing
    Set fso = Nothing
    Set rs = Nothing
    Set conn = Nothing
    
    GerarJSONVendaEspecifica = "JSON da venda " & vendaId & " gerado com sucesso!"
    
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' EXECUÇÃO PRINCIPAL
' -----------------------------------------------------------------------------------
Dim acao, venda_id, resultado, mostrarResultado

acao = Request.QueryString("acao")
venda_id = Request.QueryString("venda_id")
mostrarResultado = False

' Processar ações
If acao = "gerar_todos" Then
    resultado = GerarJSONsVendas()
    Call InserirLog("JSON", "GERAR", "Gerados JSONs para todas as vendas")
    mostrarResultado = True
    
ElseIf acao = "gerar_especifico" And venda_id <> "" Then
    resultado = GerarJSONVendaEspecifica(venda_id)
    Call InserirLog("JSON", "GERAR", "Gerado JSON para venda ID: " & venda_id)
    mostrarResultado = True
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerador de JSON - Vendas</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        .container {
            max-width: 800px;
            margin: 0 auto;
        }
        
        .card {
            border: none;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
        }
        
        .card-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 15px 15px 0 0 !important;
            padding: 1.5rem;
            text-align: center;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            padding: 12px 30px;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .resultado {
            background: #f8f9fa;
            border-radius: 10px;
            padding: 15px;
            margin-top: 20px;
            border-left: 4px solid #667eea;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        
        .alert-success {
            border-left: 4px solid #28a745;
        }
        
        .alert-danger {
            border-left: 4px solid #dc3545;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="card-header">
                <h1 class="h3 mb-0">
                    <i class="fas fa-file-code"></i> Gerador de JSON - Vendas
                </h1>
                <p class="mb-0 mt-2">Gera arquivos JSON para todas as vendas do sistema</p>
            </div>
            
            <div class="card-body">
                <!-- Exibir resultado se houver processamento -->
                <% If mostrarResultado Then %>
                    <div class="alert alert-success resultado" role="alert">
                        <h5 class="alert-heading"><i class="fas fa-check-circle"></i> Processamento Concluído</h5>
                        <p class="mb-0"><%= resultado %></p>
                        <hr>
                        <a href="tool_venda_criar_json.asp" class="btn btn-sm btn-primary">
                            <i class="fas fa-arrow-left"></i> Voltar
                        </a>
                    </div>
                <% Else %>
                
                <div class="row">
                    <div class="col-md-6">
                        <div class="d-grid gap-2">
                            <a href="tool_venda_criar_json.asp?acao=gerar_todos" class="btn btn-primary btn-lg">
                                <i class="fas fa-sync-alt"></i> Gerar Todos os JSONs
                            </a>
                            <small class="text-muted text-center">Gera JSON para todas as vendas</small>
                        </div>
                    </div>
                    
                    <div class="col-md-6">
                        <form method="get" action="tool_venda_criar_json.asp">
                            <input type="hidden" name="acao" value="gerar_especifico">
                            <div class="input-group mb-3">
                                <input type="number" class="form-control" id="venda_id" name="venda_id" 
                                       placeholder="ID da Venda" required min="1">
                                <button type="submit" class="btn btn-outline-primary">
                                    <i class="fas fa-search"></i> Gerar
                                </button>
                            </div>
                            <small class="text-muted">Gera JSON para uma venda específica</small>
                        </form>
                    </div>
                </div>
                
                <div class="mt-4">
                    <h5><i class="fas fa-info-circle"></i> Informações:</h5>
                    <ul class="list-unstyled">
                        <li><i class="fas fa-folder text-primary"></i> Pasta de destino: <code>/json/</code></li>
                        <li><i class="fas fa-file text-success"></i> Formato do arquivo: <code>venda-&lt;id&gt;.json</code></li>
                        <li><i class="fas fa-database text-warning"></i> Estrutura completa dos dados da venda</li>
                        <li><i class="fas fa-sync text-info"></i> Atualização manual sob demanda</li>
                        <li><i class="fas fa-code text-secondary"></i> Codificação: UTF-8</li>
                    </ul>
                </div>
                
                <div class="alert alert-info mt-4">
                    <h6><i class="fas fa-lightbulb"></i> Dica:</h6>
                    <p class="mb-0">Os arquivos JSON serão gerados na pasta <code>json/</code> com todos os dados formatados e codificação UTF-8 para suporte a acentuação.</p>
                </div>
                
                <% End If %>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
</body>
</html>