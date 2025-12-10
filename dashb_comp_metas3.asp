<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>

<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' ===================================================================
' 0. AUTENTICAÇÃO E FUNÇÕES DE DADOS (INÍCIO DO NOVO BLOCO VBSCRIPT)
' ===================================================================

if Session("Usuario") = "" then
    Response.redirect "gestao_login.asp"
end if 

' FUNÇÃO PARA OBTER ANOS DISPONÍVEIS
Function ObterAnosDisponiveis(conexao)
    Dim dicionarioAnos, recordsetAnos, consultaSQL
    Set dicionarioAnos = Server.CreateObject("Scripting.Dictionary")
    Set recordsetAnos = Server.CreateObject("ADODB.Recordset")
    
    On Error Resume Next
    
    ' Tentativa 1: Coluna AnoVenda (Prioridade conforme solicitado)
    consultaSQL = "SELECT DISTINCT AnoVenda FROM Vendas WHERE Excluido = 0 AND AnoVenda Is Not Null ORDER BY AnoVenda DESC"
    recordsetAnos.Open consultaSQL, conexao, 1, 1
    
    If Err.Number <> 0 Then
        Err.Clear
        ' Tentativa 2: Coluna [Ano Venda] (Fallback para nomes de colunas alternativos)
        consultaSQL = "SELECT DISTINCT [Ano Venda] AS AnoVenda FROM Vendas WHERE Excluido = 0 AND [Ano Venda] Is Not Null ORDER BY [Ano Venda] DESC"
        recordsetAnos.Close
        recordsetAnos.Open consultaSQL, conexao, 1, 1
    End If
    
    ' A tentativa de usar Year(DataVenda) foi removida conforme sua solicitação.
    
    On Error Goto 0
    
    If Not recordsetAnos.EOF Then
        Do While Not recordsetAnos.EOF
            If Not IsNull(recordsetAnos.Fields(0).Value) Then
                dicionarioAnos(CStr(recordsetAnos.Fields(0).Value)) = 1
            End If
            recordsetAnos.MoveNext
        Loop
    End If
    
    recordsetAnos.Close
    Set recordsetAnos = Nothing
    ObterAnosDisponiveis = dicionarioAnos.Keys
End Function

' FUNÇÃO AUXILIAR PARA EXTRAIR DADOS DO DICIONÁRIO (Não usada na função abaixo, mas mantida)
Function ExtrairDadosDicionario(dicionario, chave, indice)
    If dicionario.Exists(chave) Then
        Dim partesDados
        partesDados = Split(dicionario(chave), "|")
        If UBound(partesDados) >= indice Then
            ExtrairDadosDicionario = CDbl(partesDados(indice))
        Else
            ExtrairDadosDicionario = 0
        End If
    Else
        ExtrairDadosDicionario = 0
    End If
End Function

' FUNÇÃO PARA CALCULAR VGV E % META POR PERÍODO
Function CalcularDadosPeriodo(conexao, anoRef, tipoPeriodoRef, numPeriodoRef, filtroBase)
    Dim dicionarioDirPeriodo, dicionarioGerPeriodo
    Set dicionarioDirPeriodo = Server.CreateObject("Scripting.Dictionary")
    Set dicionarioGerPeriodo = Server.CreateObject("Scripting.Dictionary")
    
    Dim filtroSQLPeriodo, consultaSQLPeriodo, valorMetaPeriodo, recordsetMetaPeriodo, recordsetVendasPeriodo
    Dim nomeDirPeriodo, vgvDirPeriodo, percentDirPeriodo, nomeGerPeriodo, vgvGerPeriodo, percentGerPeriodo
    
    ' FILTRO OBRIGATÓRIO POR ANOVENDA
    filtroSQLPeriodo = filtroBase & " AND AnoVenda = " & anoRef
    
    Select Case tipoPeriodoRef
        Case "semestre"
            ' Usa a coluna Semestre da tabela Vendas
            filtroSQLPeriodo = filtroSQLPeriodo & " AND Semestre = " & numPeriodoRef
        Case "trimestre"
            ' Usa a coluna Trimestre da tabela Vendas
            filtroSQLPeriodo = filtroSQLPeriodo & " AND Trimestre = " & numPeriodoRef
        Case "mes"
            ' Usa a coluna MesVenda da tabela Vendas
            filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda = " & numPeriodoRef
    End Select
    
    ' --- CÁLCULO DA META TOTAL DO PERÍODO (Baseado em Ano e Mes da MetaEmpresa) ---
    valorMetaPeriodo = 0
    
    If tipoPeriodoRef = "ano" Then
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef
    ElseIf tipoPeriodoRef = "semestre" Then
        ' Mantém o cálculo por Mes, pois a MetaEmpresa é granular por Mês
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef & _
                        " AND Mes BETWEEN " & ((numPeriodoRef-1)*6+1) & " AND " & (numPeriodoRef*6)
    ElseIf tipoPeriodoRef = "trimestre" Then
        ' Mantém o cálculo por Mes, pois a MetaEmpresa é granular por Mês
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef & _
                        " AND Mes BETWEEN " & ((numPeriodoRef-1)*3+1) & " AND " & (numPeriodoRef*3)
    ElseIf tipoPeriodoRef = "mes" Then
        consultaSQLPeriodo = "SELECT Meta FROM MetaEmpresa WHERE Ano = " & anoRef & " AND Mes = " & numPeriodoRef
    End If
    
    Set recordsetMetaPeriodo = Server.CreateObject("ADODB.Recordset")
    recordsetMetaPeriodo.Open consultaSQLPeriodo, conexao

    On Error Resume Next
    
    ' Lógica para extrair a Meta (ajustada para tratar diferentes nomes de campo/tipos de retorno)
    If Not recordsetMetaPeriodo.EOF Then
        If recordsetMetaPeriodo.Fields.Count > 0 Then
            Dim campoNome : campoNome = "MetaTotal"
            If tipoPeriodoRef = "mes" Then campoNome = "Meta"

            ' Tentativa de pegar pelo nome da coluna (MetaTotal ou Meta)
            If Not IsNull(recordsetMetaPeriodo(campoNome).Value) Then
                valorMetaPeriodo = CDbl(recordsetMetaPeriodo(campoNome).Value)
            ElseIf Not IsNull(recordsetMetaPeriodo(0).Value) Then
                ' Tentativa de pegar pelo índice 0
                valorMetaPeriodo = CDbl(recordsetMetaPeriodo(0).Value)
            End If
        End If
    End If
    
    On Error Goto 0
    
    recordsetMetaPeriodo.Close
    Set recordsetMetaPeriodo = Nothing
    
    ' --- CÁLCULO DE VGV POR DIRETORIA --- (Usando filtroSQLPeriodo atualizado)
    consultaSQLPeriodo = "SELECT Diretoria, SUM(ValorUnidade) AS VGV FROM Vendas " & filtroSQLPeriodo & _
                " AND Diretoria IS NOT NULL AND Diretoria <> '' GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
    
    Set recordsetVendasPeriodo = Server.CreateObject("ADODB.Recordset")
    recordsetVendasPeriodo.Open consultaSQLPeriodo, conexao
    
    Do While Not recordsetVendasPeriodo.EOF
        nomeDirPeriodo = Trim(recordsetVendasPeriodo("Diretoria"))
        If Not IsNull(recordsetVendasPeriodo("VGV")) Then
            vgvDirPeriodo = CDbl(recordsetVendasPeriodo("VGV"))
        Else
            vgvDirPeriodo = 0
        End If
        
        If nomeDirPeriodo <> "" Then
            If valorMetaPeriodo > 0 And vgvDirPeriodo > 0 Then
                percentDirPeriodo = Round((vgvDirPeriodo / valorMetaPeriodo) * 100, 1)
            Else
                percentDirPeriodo = 0
            End If
            
            ' Armazena VGV | %Meta
            dicionarioDirPeriodo.Add nomeDirPeriodo, vgvDirPeriodo & "|" & percentDirPeriodo
        End If
        recordsetVendasPeriodo.MoveNext
    Loop
    recordsetVendasPeriodo.Close
    
    ' --- CÁLCULO DE VGV POR GERENCIA (Usando filtroSQLPeriodo atualizado) ---
    consultaSQLPeriodo = "SELECT Gerencia, SUM(ValorUnidade) AS VGV FROM Vendas " & filtroSQLPeriodo & _
                " AND Gerencia IS NOT NULL AND Gerencia <> '' GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
    
    recordsetVendasPeriodo.Open consultaSQLPeriodo, conexao
    
    Do While Not recordsetVendasPeriodo.EOF
        nomeGerPeriodo = Trim(recordsetVendasPeriodo("Gerencia"))
        If Not IsNull(recordsetVendasPeriodo("VGV")) Then
            vgvGerPeriodo = CDbl(recordsetVendasPeriodo("VGV"))
        Else
            vgvGerPeriodo = 0
        End If
        
        If nomeGerPeriodo <> "" Then
            ' Para a gerência, calcula a % meta em relação à Meta TOTAL do período (pode ser ajustado)
            If valorMetaPeriodo > 0 And vgvGerPeriodo > 0 Then
                percentGerPeriodo = Round((vgvGerPeriodo / valorMetaPeriodo) * 100, 1)
            Else
                percentGerPeriodo = 0
            End If
            
            dicionarioGerPeriodo.Add nomeGerPeriodo, vgvGerPeriodo & "|" & percentGerPeriodo
        End If
        recordsetVendasPeriodo.MoveNext
    Loop
    recordsetVendasPeriodo.Close
    Set recordsetVendasPeriodo = Nothing
    
    ' Retorna Array: [Dicionário de Diretoria, Dicionário de Gerência, Meta Total do Período]
    CalcularDadosPeriodo = Array(dicionarioDirPeriodo, dicionarioGerPeriodo, valorMetaPeriodo)
End Function


' SUB-ROTINA AUXILIAR PARA RENDERIZAR UMA TABELA DE NÍVEL (Diretoria ou Gerência)
Sub GerarTabela(dicionario, tituloNivel, valorMetaPeriodo, cor)
    
    If IsObject(dicionario) And dicionario.Count > 0 Then
        ' Título da Seção
        Response.Write "<h6 class=""mt-3 mb-2 pt-2 fw-bold text-" & cor & """><i class=""fas fa-sitemap""></i> " & tituloNivel & "</h6>"
        
        Response.Write "      <div class=""table-responsive table-scroll-mini"">" 
        Response.Write "          <table class=""table table-sm table-striped table-hover"">" 
        Response.Write "              <thead><tr><th>#</th><th>" & tituloNivel & "</th><th class=""text-end"">VGV</th><th class=""text-center"">% Meta</th></tr></thead>" 
        Response.Write "              <tbody>"

        Dim vgv_total, itemKey, itemData, vgv_val, perc_meta, badge_class, i
        vgv_total = 0
        i = 1
        
        For Each itemKey In dicionario.Keys
            itemData = Split(dicionario(itemKey), "|")
            vgv_val = CDbl(itemData(0))
            perc_meta = CDbl(itemData(1))
            
            vgv_total = vgv_total + vgv_val
            
            ' Determinação da classe do badge
            If perc_meta >= 100 Then
                badge_class = "percent-excelente"
            ElseIf perc_meta >= 75 Then
                badge_class = "percent-bom"
            Else
                badge_class = "percent-critico"
            End If 
            
            Response.Write "<tr>"
            Response.Write "  <td>" & i & "</td>"
            Response.Write "  <td><small class=""fw-bold"">" & itemKey & "</small></td>" 
            Response.Write "  <td class=""text-end""><small>R$ " & FormatNumber(vgv_val, 0) & "</small></td>"
            Response.Write "  <td class=""text-center""><span class=""percent-badge " & badge_class & """>" & perc_meta & "%</span></td>"
            Response.Write "</tr>"
            i = i + 1
        Next
        
        Response.Write "              </tbody>"
        Response.Write "          </table>"
        Response.Write "      </div>"
    End If
End Sub

' Função principal para inserir a tabela de resultados com DADOS REAIS (AGORA SUPORTANDO DIRETORIA E GERÊNCIA)
' Recebe o Array de 3 elementos retornado por CalcularDadosPeriodo
Sub InserirTabelaDeResultados(dadosResultado, cor, periodo)
    
    Dim dicionarioDir, dicionarioGer, valorMetaPeriodo, vgv_total_geral, itemKey, itemData
    
    ' Extrai os dicionários de Diretoria (índice 0) e Gerência (índice 1), e a Meta Total (índice 2)
    If IsArray(dadosResultado) And UBound(dadosResultado) = 2 Then
        Set dicionarioDir = dadosResultado(0) 
        Set dicionarioGer = dadosResultado(1) 
        valorMetaPeriodo = CDbl(dadosResultado(2))
    Else
        Set dicionarioDir = Nothing
        Set dicionarioGer = Nothing
        valorMetaPeriodo = 0
    End If
    
    ' Calcula o VGV total geral (baseado na Diretoria, pois a soma deve ser a mesma que Gerência)
    vgv_total_geral = 0
    If IsObject(dicionarioDir) Then
        For Each itemKey In dicionarioDir.Keys 
            itemData = Split(dicionarioDir(itemKey), "|")
            vgv_total_geral = vgv_total_geral + CDbl(itemData(0))
        Next
    End If

    Response.Write "<div class=""card h-100 shadow border-0"">"
    Response.Write "  <div class=""card-header bg-gradient-" & cor & " text-white border-0"">"
    Response.Write "      <h6 class=""mb-0 fw-bold""><i class=""fas fa-chart-line""></i> Resultados por " & periodo & "</h6>"
    Response.Write "  </div>"
    
    Response.Write "  <div class=""card-body p-3"">" 
    
    ' Flag para verificar se algum dado foi exibido
    Dim hasData : hasData = False
    
    ' 1. EXIBIÇÃO DA DIRETORIA
    If IsObject(dicionarioDir) And dicionarioDir.Count > 0 Then
        Call GerarTabela(dicionarioDir, "Diretoria", valorMetaPeriodo, "dark") ' Cor Dark para Diretoria
        hasData = True
    End If

    ' 2. SEPARADOR E EXIBIÇÃO DE GERÊNCIA
    If IsObject(dicionarioGer) And dicionarioGer.Count > 0 Then
        If hasData Then
            Response.Write "<hr class=""my-3 text-secondary"">"
        End If
        Call GerarTabela(dicionarioGer, "Gerência", valorMetaPeriodo, "primary") ' Cor Primary para Gerência
        hasData = True
    End If

    ' 3. TOTAIS GERAIS
    If hasData Then
        Response.Write "      <div class=""mt-3 pt-2 border-top text-end""><small class=""text-muted"">VGV Total: <strong class=""text-dark"">R$ " & FormatNumber(vgv_total_geral, 0) & "</strong></small>"
        Response.Write "      <br><small class=""text-muted"">Meta do Período: <strong class=""text-primary"">R$ " & FormatNumber(valorMetaPeriodo, 0) & "</strong></small></div>"
    Else 
        ' Mensagem de dados indisponíveis
        Response.Write "<div class=""empty-message text-center text-muted""><i class=""fas fa-box-open fa-3x mb-3""></i><p>Dados indisponíveis para este período.</p></div>"
    End If
    
    Response.Write "  </div>"
    Response.Write "</div>"
End Sub

' =======================================================
' INÍCIO DO PROCESSAMENTO PRINCIPAL E CÁLCULO DE DADOS
' =======================================================

Dim conexaoPrincipal
Set conexaoPrincipal = Server.CreateObject("ADODB.Connection")
conexaoPrincipal.Open strConnSales

Dim anosDisponiveisArray
anosDisponiveisArray = ObterAnosDisponiveis(conexaoPrincipal)

Dim anoAtualSelecionado
anoAtualSelecionado = Request.QueryString("ano")
If anoAtualSelecionado = "" And UBound(anosDisponiveisArray) >= 0 Then
    ' Seleciona o ano mais recente disponível se nenhum for selecionado
    anoAtualSelecionado = anosDisponiveisArray(0)
ElseIf anoAtualSelecionado = "" Then
    ' Padrão para o ano atual, se não houver dados no BD
    anoAtualSelecionado = Year(Now)
End If

Dim filtroGeral
filtroGeral = " WHERE Excluido = 0 AND Excluido IS NOT NULL"

Dim nomesMeses(12)
nomesMeses(1) = "Janeiro" : nomesMeses(7) = "Julho"
nomesMeses(2) = "Fevereiro" : nomesMeses(8) = "Agosto"
nomesMeses(3) = "Março" : nomesMeses(9) = "Setembro"
nomesMeses(4) = "Abril" : nomesMeses(10) = "Outubro"
nomesMeses(5) = "Maio" : nomesMeses(11) = "Novembro"
nomesMeses(6) = "Junho" : nomesMeses(12) = "Dezembro"

' --- CÁLCULO DE TODOS OS PERÍODOS ---
Dim resultadoAnoCompleto, resultadoSemestre1, resultadoSemestre2
Dim resultadoTrimestre1, resultadoTrimestre2, resultadoTrimestre3, resultadoTrimestre4
Dim resultadosMensais(12)

resultadoAnoCompleto = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "ano", 0, filtroGeral)
resultadoSemestre1 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "semestre", 1, filtroGeral)
resultadoSemestre2 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "semestre", 2, filtroGeral)
resultadoTrimestre1 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 1, filtroGeral)
resultadoTrimestre2 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 2, filtroGeral)
resultadoTrimestre3 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 3, filtroGeral)
resultadoTrimestre4 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 4, filtroGeral)

Dim contadorMes
For contadorMes = 1 to 12
    resultadosMensais(contadorMes) = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "mes", contadorMes, filtroGeral)
Next

conexaoPrincipal.Close
Set conexaoPrincipal = Nothing
%>

<!DOCTYPE html>
<html>
<head>
    <title>Dashboard de Metas</title>
    <meta charset="UTF-8"> <!-- ADICIONADO PARA GARANTIR ACENTUAÇÃO UTF-8 NO NAVEGADOR -->
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 1. INCLUSÃO DO BOOTSTRAP 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- 2. INCLUSÃO DO FONT AWESOME (ÍCONES) -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
    
    <style>
        /* Estilos de Cores Personalizadas para a Meta */
        .percent-excelente { 
            background-color: #d4edda; /* Verde claro */
            color: #198754; /* Verde forte */
            font-weight: bold;
            padding: 4px 8px; 
            border-radius: 6px; 
            min-width: 60px;
            display: inline-block;
            text-align: center;
            font-size: 0.8rem;
        }
        .percent-bom { 
            background-color: #fff3cd; /* Amarelo claro */
            color: #ffc107; /* Amarelo escuro */
            font-weight: bold;
            padding: 4px 8px; 
            border-radius: 6px; 
            min-width: 60px;
            display: inline-block;
            text-align: center;
            font-size: 0.8rem;
        }
        .percent-critico { 
            background-color: #f8d7da; /* Vermelho claro */
            color: #dc3545; /* Vermelho forte */
            font-weight: bold;
            padding: 4px 8px; 
            border-radius: 6px; 
            min-width: 60px;
            display: inline-block;
            text-align: center;
            font-size: 0.8rem;
        }
        
        /* Estilos do Dashboard */
body {
    background-color: #f4f7f9; /* Fundo cinza claro */
    font-family: 'Inter', sans-serif;
    
    /* 1. Aplica a redução de escala para 80% */
    transform: scale(0.8);
    
    /* 2. Define o ponto de origem da transformação no canto superior esquerdo */
    transform-origin: 0 0;
    
    /* 3. Ajusta a largura para compensar o zoom (100% / 0.8 = 125%) e evitar barra de rolagem horizontal */
    width: calc(100% / 0.8);
    
    /* 4. Propriedades opcionais para garantir que o body se comporte como um bloco de visualização */
    position: absolute;
    min-height: 100vh;
}
        .container-fluid {
            padding-top: 20px;
            padding-bottom: 50px;
        }
        h1 {
            color: #343a40;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 10px;
        }
        .period-section { 
            margin-bottom: 40px; 
        }
        .period-section h2 {
            font-size: 1.5rem;
            font-weight: 600;
            padding-bottom: 5px;
        }
        
        /* Ajuste do Card Header para gradiente */
        .bg-gradient-primary { background: linear-gradient(90deg, #0d6efd, #00b4d8) !important; }
        .bg-gradient-info { background: linear-gradient(90deg, #0dcaf0, #4895ef) !important; }
        .bg-gradient-success { background: linear-gradient(90deg, #198754, #44b967) !important; }
        .bg-gradient-warning { background: linear-gradient(90deg, #ffc107, #f79a00) !important; }

        /* Ajuste para rolagem da tabela dentro do card */
        .table-scroll {
            max-height: 350px; 
            overflow-y: auto;
        }
        /* Ajuste para tabelas menores (Gerência/Diretoria) */
        .table-scroll-mini {
            max-height: 250px; 
            overflow-y: auto;
        }

        /* Estilo da tabela */
        .table-hover tbody tr:hover {
            background-color: #e9f7ff;
        }

        /* Títulos de período menores */
        .period-section h3 {
            font-size: 1.1rem;
            font-weight: 500;
            color: #555;
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="d-flex justify-content-between align-items-center mb-5">
            <h1 class="mb-0 flex-grow-1"><i class="fas fa-tachometer-alt"></i> Dashboard de Metas</h1>
            
            <!-- SELETOR DE ANO -->
            <form method="GET" action="dashb_comp_metas2.asp" class="d-flex align-items-center">
                <label for="anoSelect" class="form-label mb-0 me-2 text-dark fw-bold">Ano de Análise:</label>
                <select name="ano" id="anoSelect" class="form-select form-select-lg" onchange="this.form.submit()">
                    <%
                    If UBound(anosDisponiveisArray) >= 0 Then
                        For Each ano in anosDisponiveisArray
                            Dim selectedStatus : selectedStatus = ""
                            If CStr(ano) = CStr(anoAtualSelecionado) Then
                                selectedStatus = "selected"
                            End If
                            Response.Write "<option value=""" & ano & """ " & selectedStatus & ">" & ano & "</option>"
                        Next
                    Else
                        Response.Write "<option value=""" & anoAtualSelecionado & """>" & anoAtualSelecionado & " (Sem Dados)</option>"
                    End If
                    %>
                </select>
            </form>
        </div>

        
        <!-- SEÇÃO 1: ANO COMPLETO -->
        <div class="period-section">
            <h2 class="mb-3 text-primary"><i class="fas fa-calendar-alt"></i> Análise Anual (<%=anoAtualSelecionado%>)</h2>
            <div class="row">
                <div class="col-12">
                    <% 
                    ' Passa o array de resultados reais
                    Call InserirTabelaDeResultados(resultadoAnoCompleto, "primary", "Ano Completo") 
                    %>
                </div>
            </div>
        </div>

        

        <!-- SEÇÃO 2: SEMESTRES -->
        <div class="period-section">
            <h2 class="mb-3 text-info"><i class="fas fa-sync-alt"></i> Análise Semestral</h2>
            <div class="row">
                <div class="col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoSemestre1, "info", "1º Semestre") %>
                </div>
                <div class="col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoSemestre2, "info", "2º Semestre") %>
                </div>
            </div>
        </div>

        

        <!-- SEÇÃO 3: TRIMESTRES -->
        <div class="period-section">
            <h2 class="mb-3 text-success"><i class="fas fa-chart-bar"></i> Análise Trimestral</h2>
            <div class="row">
                <div class="col-lg-3 col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoTrimestre1, "success", "1º Trimestre") %>
                </div>
                <div class="col-lg-3 col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoTrimestre2, "success", "2º Trimestre") %>
                </div>
                <div class="col-lg-3 col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoTrimestre3, "success", "3º Trimestre") %>
                </div>
                <div class="col-lg-3 col-md-6 mb-4">
                    
                    <% Call InserirTabelaDeResultados(resultadoTrimestre4, "success", "4º Trimestre") %>
                </div>
            </div>
        </div>

        
        
        <!-- SEÇÃO 4: MESES -->
        <div class="period-section">
            <h2 class="mb-3 text-success"><i class="fas fa-calendar-check"></i> Análise Mensal</h2>
            <div class="row">
                <br>
                <% 
                For contadorMes = 1 to 12 
                %>
                <div class="col-lg-3 col-md-6 mb-4">

                    
                    <% Call InserirTabelaDeResultados(resultadosMensais(contadorMes), "success", nomesMeses(contadorMes)) %>
                </div>
                <% 
                Next
                %>
            </div>
        </div>
        <!--#include file="footer.inc"-->

    </div>
    
    <!-- 3. INCLUSÃO DO JS DO BOOTSTRAP (Opcional, mas boa prática) -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>