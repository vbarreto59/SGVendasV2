<%@ Language=VBScript Codepage=65001 %>

<% 
' Configuração UTF-8 para todo o documento
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
' =========================================================================
' VARIÁVEIS ADO E GLOBAIS
' =========================================================================
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim ano, gerencia, whereClause, isFiltered
Dim connUsuarios

isFiltered = False

' =========================================================================
' 1. PROCESSAMENTO DE FORMULÁRIO E CONSTRUÇÃO DA WHERE CLAUSE
' =========================================================================

' Leitura dos filtros do formulário
ano = Request.Form("ano")
gerencia = Request.Form("gerencia")

' Cláusula WHERE inicial
whereClause = " WHERE [Vendas].[Excluido] = FALSE" 

' =========================================================================
' Construção da Cláusula WHERE de Vendas
' =========================================================================
If Not IsEmpty(ano) And ano <> "" And IsNumeric(ano) Then
    whereClause = whereClause & " AND CLng([Vendas].[AnoVenda]) = " & CLng(ano)
    isFiltered = True
Else
    ' Ano padrão caso não seja informado
    ano = Year(Date())
    whereClause = whereClause & " AND CLng([Vendas].[AnoVenda]) = " & CLng(ano)
    isFiltered = True
End If

' =========================================================================
' 2. ABERTURA DAS CONEXÕES
' =========================================================================
If isFiltered Then
    On Error Resume Next
    conn.Open StrConnSales
    If Err.Number <> 0 Then
        Response.Write "<!DOCTYPE html><html><body><div class='alert alert-danger'>Erro crítico ao conectar ao banco de dados de Vendas (StrConnSales): " & Err.Description & "</div></body></html>"
        Response.End
    End If
    On Error GoTo 0 
End If

' Conexão para tabela usuarios
Set connUsuarios = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connUsuarios.Open StrConn
If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>❌ Erro ao conectar ao banco de dados de Usuários (StrConn): " & Err.Description & "</div>"
End If
On Error GoTo 0 
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>Relatório de Unidades Vendidas por Corretor</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
        }
        .container {
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            padding: 30px;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 30px;
        }
        .card-header {
            background-color: #3498db;
            color: white;
            font-weight: bold;
        }
        .btn-primary {
            background-color: #3498db;
            border-color: #3498db;
        }
        .table th {
            background-color: #2c3e50;
            color: white;
            text-align: center;
            vertical-align: middle;
        }
        .table td {
            text-align: center;
            vertical-align: middle;
        }
        .gerencia-header {
            background-color: #e9ecef !important;
            font-weight: bold;
            font-size: 1.1em;
            color: #495057;
        }
        .mes-header {
            background-color: #17a2b8 !important;
            color: white;
            font-weight: bold;
        }
        .sem-venda {
            background-color: #ffebee !important;
            color: #c62828;
            font-weight: bold;
        }
        .com-venda {
            background-color: #e8f5e8 !important;
            color: #2e7d32;
            font-weight: bold;
        }
        .total-mes {
            background-color: #fff3cd !important;
            font-weight: bold;
            text-align: center;
        }
        .total-corretor {
            background-color: #d1ecf1 !important;
            font-weight: bold;
            text-align: center;
        }
        .table-responsive {
            max-height: 180vh;
            overflow: auto;
        }
        .sticky-column {
            position: sticky;
            left: 0;
            background-color: white;
            z-index: 1;
        }
        .sticky-header {
            position: sticky;
            top: 0;
            z-index: 2;
        }
    </style>
</head>
<body>

<div class="container mt-5">
    <h1>📊 Relatório de Unidades Vendidas</h1>

    <div class="card mb-4">
        <div class="card-header">Filtros do Relatório</div>
        <form method="POST" action="">
            <div class="card-body row">
                
<div class="form-group col-md-4">
    <label for="ano">Ano:</label>
    
    
    <select id="ano" name="ano" class="form-control">
        <option value="">Selecione o Ano</option>
        <%
        ' Define os anos disponíveis
        Dim anosDisponiveis(1)
        anosDisponiveis(0) = 2025
        anosDisponiveis(1) = 2026
        
        ' Loop para gerar as opções
        Dim k, anoOption
        
        For k = 0 To UBound(anosDisponiveis)
            anoOption = anosDisponiveis(k)
            
            Response.Write "<option value='" & anoOption & "'"
            ' Verifica se o ano submetido (variável 'ano') corresponde à opção atual
            If CStr(ano) = CStr(anoOption) Then 
                Response.Write " selected"
            End If
            Response.Write ">" & anoOption & "</option>"
        Next
        %>
    </select>
    
</div>
                
                <div class="form-group col-md-4">
                    <label for="gerencia">Gerência:</label>
                    <select id="gerencia" name="gerencia" class="form-control">
                        <option value="">Todas as Gerências</option>
                        <%
                        ' Consulta para obter todas as gerencias únicas
                        Dim sqlGerencias, rsGerencias
                        Set rsGerencias = Server.CreateObject("ADODB.Recordset")
                        
                        sqlGerencias = "SELECT DISTINCT Gerencia FROM usuarios " & _
                                      "WHERE permissao = 5 AND idEmp = 2 AND Gerencia IS NOT NULL AND Gerencia <> '' " & _
                                      "ORDER BY Gerencia ASC"
                        
                        rsGerencias.Open sqlGerencias, connUsuarios
                        
                        If Not rsGerencias.EOF Then
                            Do While Not rsGerencias.EOF
                                Dim gerenciaAtual
                                gerenciaAtual = rsGerencias("Gerencia")
                                Response.Write "<option value=""" & Server.HTMLEncode(gerenciaAtual) & """"
                                If gerencia = gerenciaAtual Then Response.Write " selected"
                                Response.Write ">" & Server.HTMLEncode(gerenciaAtual) & "</option>"
                                rsGerencias.MoveNext
                            Loop
                        End If
                        rsGerencias.Close
                        Set rsGerencias = Nothing
                        %>
                    </select>
                </div>
                
                <div class="col-12 mt-3">
                    <button type="submit" class="btn btn-primary">🔍 Gerar Relatório</button>
                    <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>" class="btn btn-secondary">🔄 Limpar Filtros</a>
                </div>

            </div>
        </form>
    </div>
    
    <% If Not isFiltered Then %>
        <div class="alert alert-info text-center">
            <h5>👋 Bem-vindo ao Relatório de Vendas por Corretor e Mês</h5>
            <p>Selecione o ano e/ou gerência desejados e clique em <strong>'Gerar Relatório'</strong> para visualizar as vendas de cada corretor por mês.</p>
        </div>
    <% Else %>
        <%
        ' =========================================================================
        ' 3. PROCESSAMENTO DOS DADOS
        ' =========================================================================
        
        ' Obter lista de todos os corretores ativos
        Dim sqlCorretores
        sqlCorretores = "SELECT UserId, nome, Gerencia FROM usuarios " & _
                        "WHERE permissao = 5 AND idEmp = 2 "
        
        ' Adicionar filtro de gerência se selecionado
        If gerencia <> "" Then
            sqlCorretores = sqlCorretores & " AND Gerencia = '" & Replace(gerencia, "'", "''") & "' "
        End If
        
        sqlCorretores = sqlCorretores & "ORDER BY Gerencia, nome ASC"
        
        Set rsCorretores = Server.CreateObject("ADODB.Recordset")
        rsCorretores.Open sqlCorretores, connUsuarios
        
        If rsCorretores.EOF Then
            Response.Write "<div class='alert alert-warning'>Nenhum corretor ativo encontrado para os filtros selecionados.</div>"
        Else
            ' Criar arrays para armazenar os dados
            Dim arrCorretores(), arrNomes(), arrGerencias(), arrVendas()
            Dim corretorCount
            corretorCount = 0
            
            ' Primeiro passada: contar corretores
            rsCorretores.MoveFirst
            Do While Not rsCorretores.EOF
                corretorCount = corretorCount + 1
                rsCorretores.MoveNext
            Loop
            
            ' Redimensionar arrays
            ReDim arrCorretores(corretorCount)
            ReDim arrNomes(corretorCount)
            ReDim arrGerencias(corretorCount)
            ReDim arrVendas(corretorCount, 12)
            
            ' Segunda passada: preencher arrays
            rsCorretores.MoveFirst
            Dim index
            index = 0
            Do While Not rsCorretores.EOF
                arrCorretores(index) = CStr(rsCorretores("UserId"))
                arrNomes(index) = rsCorretores("nome")
                arrGerencias(index) = rsCorretores("Gerencia")
                If IsNull(arrGerencias(index)) Or arrGerencias(index) = "" Then
                    arrGerencias(index) = "Sem Gerência Definida"
                End If
                
                ' Inicializar vendas com zero
                For i = 1 To 12
                    arrVendas(index, i) = 0
                Next
                
                index = index + 1
                rsCorretores.MoveNext
            Loop
            rsCorretores.Close
            
            ' Consultar vendas reais
            Dim sqlVendas
            sqlVendas = "SELECT corretorid, MesVenda, COUNT(*) as TotalVendas " & _
                        "FROM Vendas " & _
                        "WHERE Excluido = FALSE AND AnoVenda = " & ano & " "
            
            ' Adicionar filtro de corretores específicos se gerência foi filtrada
            If gerencia <> "" Then
                sqlVendas = sqlVendas & " AND corretorid IN ("
                For i = 0 To corretorCount - 1
                    If i > 0 Then sqlVendas = sqlVendas & ","
                    sqlVendas = sqlVendas & arrCorretores(i)
                Next
                sqlVendas = sqlVendas & ")"
            End If
            
            sqlVendas = sqlVendas & " GROUP BY corretorid, MesVenda " & _
                        "ORDER BY corretorid, MesVenda"
            
            Set rsVendas = Server.CreateObject("ADODB.Recordset")
            rsVendas.Open sqlVendas, conn
            
            If Not rsVendas.EOF Then
                Do While Not rsVendas.EOF
                    Dim corretorId, mesVenda, totalVendas
                    corretorId = CStr(rsVendas("corretorid"))
                    mesVenda = CInt(rsVendas("MesVenda"))
                    totalVendas = CInt(rsVendas("TotalVendas"))
                    
                    ' Encontrar o índice do corretor
                    For i = 0 To corretorCount - 1
                        If arrCorretores(i) = corretorId Then
                            If mesVenda >= 1 And mesVenda <= 12 Then
                                arrVendas(i, mesVenda) = totalVendas
                            End If
                            Exit For
                        End If
                    Next
                    rsVendas.MoveNext
                Loop
            End If
            rsVendas.Close
            
            ' Calcular totais
            Dim totalMes(12), totalGeral
            For i = 1 To 12
                totalMes(i) = 0
            Next
            totalGeral = 0
            
            For i = 0 To corretorCount - 1
                Dim totalCorretor
                totalCorretor = 0
                For j = 1 To 12
                    totalCorretor = totalCorretor + arrVendas(i, j)
                    totalMes(j) = totalMes(j) + arrVendas(i, j)
                Next
                totalGeral = totalGeral + totalCorretor
            Next
            
            ' Agrupar por gerência
            Set dictGerencia = Server.CreateObject("Scripting.Dictionary")
            
            For i = 0 To corretorCount - 1
                gerencia = arrGerencias(i)
                
                If Not dictGerencia.Exists(gerencia) Then
                    dictGerencia.Add gerencia, ""
                End If
            Next
            %>
            
            <div class="card mb-4 mt-4">
                <div class="card-body">
                    <h2 class="card-title">📈 Vendas por Corretor e Mês - Ano <%=ano%></h2>
                    
                    <% ' Exibir filtros ativos %>
                    <div class="alert alert-warning mb-4">
                        <h6>🔍 Filtros Aplicados:</h6>
                        <%
                        Dim filtrosAtivos
                        filtrosAtivos = ""
                        If ano <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-primary'>Ano: " & ano & "</span>"
                        If gerencia <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-info'>Gerência: " & gerencia & "</span>"
                        
                        If filtrosAtivos = "" Then
                            Response.Write "<em>Nenhum filtro específico aplicado</em>"
                        Else
                            Response.Write filtrosAtivos
                        End If
                        %>
                    </div>
                    
                    <div class="table-responsive">
                        <table class="table table-bordered table-hover">
                            <thead class="sticky-header">
                                <tr>
                                    <th class="sticky-column">Gerência / Corretor</th>
                                    <%
                                    ' Cabeçalho dos meses
                                    For i = 1 To 12
                                        Response.Write "<th class='mes-header'>" & MonthName(i, True) & "</th>"
                                    Next
                                    %>
                                    <th class="bg-warning">Total</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                ' Ordenar gerencias
                                Dim arrGerenciasUnicas
                                arrGerenciasUnicas = dictGerencia.Keys
                                
                                ' Exibir por gerência
                                For Each gerencia In arrGerenciasUnicas
                                    Dim totalGerencia
                                    totalGerencia = 0
                                    %>
                                    <tr class="gerencia-header">
                                        <td class="sticky-column"><strong>🏢 <%=gerencia%></strong></td>
                                        <%
                                        For i = 1 To 12
                                            Response.Write "<td></td>"
                                        Next
                                        %>
                                        <td></td>
                                    </tr>
                                    <%
                                    ' Exibir corretores desta gerência
                                    For i = 0 To corretorCount - 1
                                        If arrGerencias(i) = gerencia Then
                                           '' Dim totalCorretor
                                            totalCorretor = 0
                                            For j = 1 To 12
                                                totalCorretor = totalCorretor + arrVendas(i, j)
                                            Next
                                            totalGerencia = totalGerencia + totalCorretor
                                            %>
                                            <tr>
                                                <td class="sticky-column"><%=arrNomes(i)%></td>
                                                <%
                                                For j = 1 To 12
                                                    If arrVendas(i, j) > 0 Then
                                                        Response.Write "<td class='com-venda'>" & arrVendas(i, j) & "</td>"
                                                    Else
                                                        Response.Write "<td class='sem-venda'>0</td>"
                                                    End If
                                                Next
                                                %>
                                                <td class="total-corretor"><strong><%=totalCorretor%></strong></td>
                                            </tr>
                                            <%
                                        End If
                                    Next
                                    
                                    ' Linha de total da gerência
                                    %>
                                    <tr class="total-mes">
                                        <td class="sticky-column text-right"><strong>Total <%=gerencia%>:</strong></td>
                                        <%
                                        For j = 1 To 12
                                            Dim totalMesGerencia
                                            totalMesGerencia = 0
                                            For i = 0 To corretorCount - 1
                                                If arrGerencias(i) = gerencia Then
                                                    totalMesGerencia = totalMesGerencia + arrVendas(i, j)
                                                End If
                                            Next
                                            Response.Write "<td><strong>" & totalMesGerencia & "</strong></td>"
                                        Next
                                        %>
                                        <td class="total-corretor"><strong><%=totalGerencia%></strong></td>
                                    </tr>
                                    <%
                                Next
                                
                                ' Linha de totais gerais por mês
                                %>
                                <tr class="table-primary">
                                    <td class="sticky-column text-right"><strong>📊 TOTAL GERAL:</strong></td>
                                    <%
                                    For i = 1 To 12
                                        Response.Write "<td><strong>" & totalMes(i) & "</strong></td>"
                                    Next
                                    %>
                                    <td class="total-corretor"><strong><%=totalGeral%></strong></td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                    
                    <div class="mt-4">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="alert alert-info">
                                    <h6>🎨 Legenda:</h6>
                                    <p><span class="com-venda px-2">Verde</span> = Com vendas no mês</p>
                                    <p><span class="sem-venda px-2">Vermelho</span> = Sem vendas no mês</p>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="alert alert-light">
                                    <h6>📈 Estatísticas:</h6>
                                    <p><strong>Total de Corretores:</strong> <%=corretorCount%></p>
                                    <p><strong>Total de Gerencias:</strong> <%=dictGerencia.Count%></p>
                                    <p><strong>Total de Vendas no Ano:</strong> <%=totalGeral%></p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <%
        End If
    End If
    %>
    
</div>

<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

</body>
</html>

<%
' =========================================================================
' FECHAMENTO DE CONEXÕES E LIMPEZA DE OBJETOS
' =========================================================================

If IsObject(rsCorretores) Then
    If rsCorretores.State = 1 Then rsCorretores.Close
    Set rsCorretores = Nothing
End If

If IsObject(rsVendas) Then
    If rsVendas.State = 1 Then rsVendas.Close
    Set rsVendas = Nothing
End If

If IsObject(conn) Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If

If IsObject(connUsuarios) Then
    If connUsuarios.State = 1 Then connUsuarios.Close 
    Set connUsuarios = Nothing 
End If
%>