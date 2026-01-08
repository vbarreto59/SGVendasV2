<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: FWNUUKVOKD          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  
 <!--#include file="usr_acoes_v4GVendas.inc"-->

<%
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    On Error Resume Next 
    set objMail = server.createobject("CDONTS.NewMail")
    if Err.Number <> 0 then 
        set objMail = Nothing ' Garante que a variável seja liberada, mesmo que não criada
    else
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
        objMail.Subject = "SV-DIR-VENDAS" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
        objMail.MailFormat = 0 ' 0 = Texto Simples
        objMail.Body = "Página Vendas Diretoria. " & Ucase(Session("Usuario"))
        objMail.Send
        set objMail = Nothing
    end if 
    On Error GoTo 0 
end if%>



<%
' Configuração para evitar cache
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"

' Função para converter número brasileiro para JavaScript
Function ConverterParaJS(valor)
    
    
    ' Remover formatação brasileira
    resultado = Replace(valor, ".", "") ' Remove separadores de milhar
    resultado = Replace(valor, ",", ".") ' Converte vírgula decimal para ponto
    
    If IsNumeric(resultado) Then
        ConverterParaJS = resultado
    Else
        ConverterParaJS = 0
    End If
End Function

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

' Parâmetros do filtro
Dim ano, mes, trimestre, semestre, diretoriaID, corretor, gerencia, whereClause, isFiltered
isFiltered = False

' *** MODIFICAÇÃO AQUI: Verificar se já existe filtro via POST, senão definir ano atual automaticamente ***
ano = Request.Form("ano")
mes = Request.Form("mes")
trimestre = Request.Form("trimestre")
semestre = Request.Form("semestre")
corretor = Request.Form("corretor") ' FILTRO: CORRETOR
gerencia = Request.Form("gerencia") ' NOVO FILTRO: GERÊNCIA

' **NOVA LÓGICA: Se ano não foi enviado via POST, definir automaticamente como ano atual**
Dim anoAtual, anoSelecionadoAutomaticamente
anoAtual = Year(Date()) ' Ano atual do sistema

If ano = "" Then
    ' Não foi enviado filtro via POST, vamos definir automaticamente
    anoSelecionadoAutomaticamente = True
Else
    anoSelecionadoAutomaticamente = False
End If

' **NOVA LÓGICA DE FILTRO DE DIRETORIA**
diretoriaID = Session("Dir_DiretoriaID")

paginaRedirecionamento1 = "http://www.gabnetweb.com.br/SunnyImob/login_v66a.asp"
paginaRedirecionamento2 = "http://localhost/SunnyImob/login_v66a.asp"
if diretoriaID = "" then
   Response.Write "Erro de processamento! (100)" 
   '*** verificar se é localhost ou no site'

   if (request.ServerVariables("remote_addr") <> "127.0.0.1")  then
      Response.Redirect paginaRedirecionamento1
   else
      Response.Redirect paginaRedirecionamento2
   end if   
  '' Response.end 
end if   


' Construir cláusula WHERE
whereClause = " WHERE Vendas.Excluido = 0"

' 1. Aplicar filtro de Diretoria se o diretor estiver logado
If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
    whereClause = whereClause & " AND Vendas.DiretoriaId = " & CLng(diretoriaID)
    isFiltered = True 
End If

' Abrir conexão
conn.Open StrConnSales

' --- BUSCAR ANOS DISPONÍVEIS NO BANCO DE DADOS ---
Dim sqlAnos, anosDisponiveis()
ReDim anosDisponiveis(0)
Dim anosCount
anosCount = 0

' Buscar anos distintos da tabela Vendas
sqlAnos = "SELECT DISTINCT AnoVenda FROM Vendas WHERE Vendas.Excluido = 0"
If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
    sqlAnos = sqlAnos & " AND Vendas.DiretoriaId = " & CLng(diretoriaID)
End If
sqlAnos = sqlAnos & " ORDER BY AnoVenda DESC"

Set rsAnos = Server.CreateObject("ADODB.Recordset")
rsAnos.Open sqlAnos, conn

Do While Not rsAnos.EOF
    If Not IsNull(rsAnos("AnoVenda")) Then
        ReDim Preserve anosDisponiveis(anosCount)
        anosDisponiveis(anosCount) = CInt(rsAnos("AnoVenda"))
        anosCount = anosCount + 1
    End If
    rsAnos.MoveNext
Loop
rsAnos.Close
Set rsAnos = Nothing
' --- FIM DA BUSCA DE ANOS ---

' --- BUSCAR GERENCIAS DISPONÍVEIS NO BANCO DE DADOS ---
Dim sqlGerencias, gerenciasDisponiveis()
ReDim gerenciasDisponiveis(0)
Dim gerenciasCount
gerenciasCount = 0

' Buscar gerências distintas da tabela Vendas
sqlGerencias = "SELECT DISTINCT Gerencia FROM Vendas WHERE Vendas.Excluido = 0 AND Gerencia IS NOT NULL AND TRIM(Gerencia) <> ''"
If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
    sqlGerencias = sqlGerencias & " AND Vendas.DiretoriaId = " & CLng(diretoriaID)
End If
sqlGerencias = sqlGerencias & " ORDER BY Gerencia"

Set rsGerencias = Server.CreateObject("ADODB.Recordset")
rsGerencias.Open sqlGerencias, conn

Do While Not rsGerencias.EOF
    If Not IsNull(rsGerencias("Gerencia")) And Trim(rsGerencias("Gerencia")) <> "" Then
        ReDim Preserve gerenciasDisponiveis(gerenciasCount)
        gerenciasDisponiveis(gerenciasCount) = Trim(rsGerencias("Gerencia"))
        gerenciasCount = gerenciasCount + 1
    End If
    rsGerencias.MoveNext
Loop
rsGerencias.Close
Set rsGerencias = Nothing
' --- FIM DA BUSCA DE GERENCIAS ---

' --- BUSCAR CORRETORES DISPONÍVEIS NO BANCO DE DADOS ORDENADOS POR VGV ---
Dim sqlCorretores, corretoresDisponiveis(), corretoresVGV()
ReDim corretoresDisponiveis(0)
ReDim corretoresVGV(0)
Dim corretoresCount
corretoresCount = 0

' Buscar corretores com VGV total, ordenados por VGV descendente
' APLICANDO OS MESMOS FILTROS DO RELATÓRIO
sqlCorretores = "SELECT Corretor, SUM(ValorUnidade) as TotalVGV " & _
                "FROM Vendas WHERE Vendas.Excluido = 0"

' Aplicar filtro de Diretoria
If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then
    sqlCorretores = sqlCorretores & " AND Vendas.DiretoriaId = " & CLng(diretoriaID)
End If

' Aplicar filtro de Ano (se definido automaticamente ou via POST)
If anoParaUsar <> "" And IsNumeric(anoParaUsar) Then
    sqlCorretores = sqlCorretores & " AND Vendas.AnoVenda = " & anoParaUsar
End If

' Aplicar filtro de Mês (se houver)
If mes <> "" And IsNumeric(mes) Then
    sqlCorretores = sqlCorretores & " AND Vendas.MesVenda = " & mes
End If

' Aplicar filtro de Trimestre (se houver)
If trimestre <> "" And IsNumeric(trimestre) Then
    sqlCorretores = sqlCorretores & " AND Vendas.Trimestre = " & trimestre
End If

' Aplicar filtro de Semestre (se houver)
If semestre <> "" And IsNumeric(semestre) Then
    sqlCorretores = sqlCorretores & " AND Vendas.Semestre = " & semestre
End If

' Aplicar filtro de Gerência (se houver) - IMPORTANTE: CORREÇÃO AQUI
If gerencia <> "" Then
    sqlCorretores = sqlCorretores & " AND Vendas.Gerencia = '" & Replace(gerencia, "'", "''") & "'"
End If

' Aplicar filtro de Corretor (se houver) - para quando já está filtrado
If corretor <> "" Then
    sqlCorretores = sqlCorretores & " AND Vendas.Corretor = '" & Replace(corretor, "'", "''") & "'"
End If

sqlCorretores = sqlCorretores & " GROUP BY Corretor ORDER BY SUM(ValorUnidade) DESC"

Set rsCorretores = Server.CreateObject("ADODB.Recordset")
rsCorretores.Open sqlCorretores, conn

Do While Not rsCorretores.EOF
    If Not IsNull(rsCorretores("Corretor")) And Trim(rsCorretores("Corretor")) <> "" Then
        ReDim Preserve corretoresDisponiveis(corretoresCount)
        ReDim Preserve corretoresVGV(corretoresCount)
        corretoresDisponiveis(corretoresCount) = UCase(Trim(rsCorretores("Corretor")))
        corretoresVGV(corretoresCount) = rsCorretores("TotalVGV")
        corretoresCount = corretoresCount + 1
    End If
    rsCorretores.MoveNext
Loop
rsCorretores.Close
Set rsCorretores = Nothing
' --- FIM DA BUSCA DE CORRETORES ---

' *** MODIFICAÇÃO AQUI: Determinar o ano a ser usado automaticamente ***
Dim anoParaUsar, anoExisteNoBanco
anoParaUsar = ano
anoExisteNoBanco = False

If ano = "" And anosCount > 0 Then
    ' Verificar se o ano atual existe no banco de dados
    For i = 0 To anosCount - 1
        If anosDisponiveis(i) = anoAtual Then
            anoExisteNoBanco = True
            Exit For
        End If
    Next
    
    ' Se ano atual existe no banco, usar ele
    If anoExisteNoBanco Then
        anoParaUsar = CStr(anoAtual)
    Else
        ' Se não existir, usar o ano mais recente disponível
        If anosCount > 0 Then
            anoParaUsar = CStr(anosDisponiveis(0)) ' Primeiro elemento é o mais recente (ordenado DESC)
        End If
    End If
    
    ' Definir que estamos usando filtro automático
    ano = anoParaUsar
    isFiltered = True
End If

' *** MODIFICAÇÃO AQUI: Continuar construindo a cláusula WHERE com o ano definido ***
If anoParaUsar <> "" And IsNumeric(anoParaUsar) Then
    whereClause = whereClause & " AND Vendas.AnoVenda = " & anoParaUsar
    isFiltered = True
End If

' Continuar com outros filtros (que podem vir do POST)
If mes <> "" And IsNumeric(mes) Then
    whereClause = whereClause & " AND Vendas.MesVenda = " & mes
    isFiltered = True
End If

If trimestre <> "" And IsNumeric(trimestre) Then
    whereClause = whereClause & " AND Vendas.Trimestre = " & trimestre
    isFiltered = True
End If

If semestre <> "" And IsNumeric(semestre) Then
    whereClause = whereClause & " AND Vendas.Semestre = " & semestre
    isFiltered = True
End If

' *** FILTRO: Corretor ***
If corretor <> "" Then
    whereClause = whereClause & " AND Vendas.Corretor = '" & Replace(corretor, "'", "''") & "'"
    isFiltered = True
End If

' *** NOVO FILTRO: Gerência ***
If gerencia <> "" Then
    whereClause = whereClause & " AND Vendas.Gerencia = '" & Replace(gerencia, "'", "''") & "'"
    isFiltered = True
End If
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Vendas - Diretoria <%=Session("Dir_Nome")%></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        body { 
            background: #f8f9fa;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .card { 
            margin-bottom: 20px; 
            border: none;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }
        .card-header {
            border-radius: 10px 10px 0 0 !important;
            font-weight: 600;
            padding: 15px 20px;
        }
        .table th { 
            background-color: #f8f9fa; 
            font-weight: 600;
        }
        .table td { vertical-align: middle; }
        .total { 
            font-weight: bold; 
            background-color: #e8f4ff !important;
        }
        .badge-percent {
            font-size: 0.8rem;
            padding: 4px 8px;
        }
        .metric-card {
            padding: 20px;
            border-radius: 10px;
            color: white;
            text-align: center;
            margin-bottom: 15px;
        }
        .metric-value {
            font-size: 2rem;
            font-weight: bold;
            margin: 5px 0;
        }
        .metric-label {
            font-size: 0.9rem;
            opacity: 0.9;
        }
        .btn-filter {
            padding: 10px 20px;
            font-weight: 600;
        }
        .chart-container {
            position: relative;
            height: 400px;
            width: 100%;
        }
        .progress-thin {
            height: 5px;
            border-radius: 2px;
        }
        .card-body {
            overflow-x: auto;
            overflow-y: hidden;
        }
        #graficoVGV {
            min-width: 100%;
        }
        
        @media (max-width: 768px) {
            .chart-container {
                height: 500px;
            }
        }
        
        .card-mes-vazio {
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            color: #6c757d;
        }
        .card-mes-com-dados {
            background-color: #e8f4ff;
            border: 1px solid #3498db;
            color: #2c3e50;
        }
        .card-mes-filtrado {
            background-color: #ffe8e8;
            border: 2px solid #e74c3c;
            color: #c0392b;
        }
        .auto-filter-badge {
            font-size: 0.75rem;
            padding: 2px 6px;
            margin-left: 5px;
        }
        
        /* Estilo para corretor filtrado */
        .corretor-filtrado {
            background-color: #fff3cd !important;
            border-left: 4px solid #ffc107 !important;
        }
        /* Estilo para gerência filtrada */
        .gerencia-filtrada {
            background-color: #d4edda !important;
            border-left: 4px solid #28a745 !important;
        }
        .lista-corretores {
            max-height: 400px;
            overflow-y: auto;
        }
        .corretor-item {
            transition: all 0.3s ease;
        }
        .corretor-item:hover {
            background-color: #f8f9fa;
        }
        
        /* Novos estilos para cards de mês */
        .card-mes-completo {
            border: 1px solid #dee2e6;
            transition: all 0.3s ease;
        }
        .card-mes-completo:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .unidades-mes {
            font-size: 0.85rem;
            color: #6c757d;
        }
        .valor-mes {
            font-size: 1.1rem;
            font-weight: bold;
            color: #2c3e50;
        }
        .card-mes-sem-dados {
            opacity: 0.6;
        }
        
        /* Estilo para ranking de corretores */
        .ranking-badge {
            font-size: 0.8rem;
            width: 24px;
            height: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            margin-right: 8px;
        }
        .ranking-1 { background-color: #ffd700; color: #000; }
        .ranking-2 { background-color: #c0c0c0; color: #000; }
        .ranking-3 { background-color: #cd7f32; color: #000; }
        .ranking-outros { background-color: #6c757d; color: #fff; }
        
        /* Estilo para filtros em colunas responsivas */
        @media (max-width: 768px) {
            .filtro-row {
                flex-direction: column;
            }
            .filtro-col {
                width: 100%;
                margin-bottom: 10px;
            }
        }
    </style>

<style>
    body {
        transform: scale(0.9); 
        transform-origin: 0 0; 
        width: calc(100% / 0.9); 
    }
    
    @media (max-width: 768px) {
        body {
            transform: scale(1); 
            width: 100%; 
        }
    }
</style>    
</head>
<body>
    <div class="container mt-4">
        <!-- HEADER -->
        <div class="card mb-4 border-0" style="background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);">
            <div class="card-body text-white p-4">
                <h1 class="mb-2"><i class="fas fa-chart-bar me-2"></i>Relatório de Vendas</h1>
                <p class="mb-0">Diretoria: <strong><%=Session("Dir_Nome")%></strong></p>
                <% If anoSelecionadoAutomaticamente And anoParaUsar <> "" Then %>
                <p class="mb-0 mt-2">
                    <small>
                        <i class="fas fa-robot me-1"></i>Filtro automático aplicado: <strong>Ano <%=anoParaUsar%></strong>
                        <% If anoExisteNoBanco Then %>
                            <span class="badge bg-success auto-filter-badge">Ano atual</span>
                        <% Else %>
                            <span class="badge bg-warning text-dark auto-filter-badge">Ano disponível mais recente</span>
                        <% End If %>
                    </small>
                </p>
                <% End If %>
            </div>
        </div>
        
        <!-- FILTROS -->
        <div class="card mb-4">
            <div class="card-header text-white" style="background: #3498db;">
                <i class="fas fa-filter me-2"></i>Filtros do Relatório
            </div>
            <div class="card-body">
                <form method="post" class="row g-3">
                    <div class="col-md-3">
                        <label for="ano" class="form-label fw-bold">Ano:</label>
                        <select class="form-select" id="ano" name="ano">
                            <option value="">Selecione o ano</option>
                            <%
                            For i = 0 To anosCount - 1
                                Response.Write "<option value=""" & anosDisponiveis(i) & """"
                                If CStr(anosDisponiveis(i)) = ano Then
                                    Response.Write " selected"
                                End If
                                Response.Write ">" & anosDisponiveis(i) & "</option>"
                            Next
                            
                            If anosCount = 0 Then
                                Response.Write "<option value="""">Nenhum ano disponível</option>"
                            End If
                            %>
                        </select>
                        <% If anosCount = 0 Then %>
                        <small class="text-warning">Nenhum ano com vendas encontrado</small>
                        <% End If %>
                    </div>

                    <div class="col-md-3">
                        <label for="mes" class="form-label fw-bold">Mês:</label>
                        <select class="form-select" id="mes" name="mes">
                            <option value="">Todos os meses</option>
                            <% For i = 1 To 12 %>
                            <option value="<%=i%>" <%If CStr(i) = mes Then Response.Write "selected"%>><%=MonthName(i, False)%></option>
                            <% Next %>
                        </select>
                    </div>
                    
                    <div class="col-md-3">
                        <label for="trimestre" class="form-label fw-bold">Trimestre:</label>
                        <select class="form-select" id="trimestre" name="trimestre">
                            <option value="">Todos trimestres</option>
                            <option value="1" <%If trimestre = "1" Then Response.Write "selected"%>>1º Trimestre</option>
                            <option value="2" <%If trimestre = "2" Then Response.Write "selected"%>>2º Trimestre</option>
                            <option value="3" <%If trimestre = "3" Then Response.Write "selected"%>>3º Trimestre</option>
                            <option value="4" <%If trimestre = "4" Then Response.Write "selected"%>>4º Trimestre</option>
                        </select>
                    </div>
                    
                    <div class="col-md-3">
                        <label for="semestre" class="form-label fw-bold">Semestre:</label>
                        <select class="form-select" id="semestre" name="semestre">
                            <option value="">Todos semestres</option>
                            <option value="1" <%If semestre = "1" Then Response.Write "selected"%>>1º Semestre</option>
                            <option value="2" <%If semestre = "2" Then Response.Write "selected"%>>2º Semestre</option>
                        </select>
                    </div>
                    
                    <!-- FILTRO: GERÊNCIA -->
                    <div class="col-md-4">
                        <label for="gerencia" class="form-label fw-bold">Gerência:</label>
                        <select class="form-select" id="gerencia" name="gerencia">
                            <option value="">Todas as gerências</option>
                            <%
                            For i = 0 To gerenciasCount - 1
                                Response.Write "<option value=""" & gerenciasDisponiveis(i) & """"
                                If gerenciasDisponiveis(i) = gerencia Then
                                    Response.Write " selected"
                                End If
                                Response.Write ">" & gerenciasDisponiveis(i) & "</option>"
                            Next
                            
                            If gerenciasCount = 0 Then
                                Response.Write "<option value="""">Nenhuma gerência disponível</option>"
                            End If
                            %>
                        </select>
                        <% If gerenciasCount = 0 Then %>
                        <small class="text-warning">Nenhuma gerência encontrada</small>
                        <% End If %>
                    </div>
                    
                    <!-- FILTRO: CORRETOR -->
                    <div class="col-md-4">
                        <label for="corretor" class="form-label fw-bold">Corretor:</label>
<select class="form-select" id="corretor" name="corretor">
    <option value="">Todos os corretores</option>
    <%
    If corretoresCount > 0 Then
        ' 1. Criar um array de índices (0, 1, 2, ...)
        Dim indices()
        ReDim indices(corretoresCount - 1)
        For i = 0 To corretoresCount - 1
            indices(i) = i
        Next

        ' 2. Ordenar apenas os índices baseando-se nos valores de corretoresDisponiveis
        Dim tempIdx
        For i = 0 To corretoresCount - 2
            For j = i + 1 To corretoresCount - 1
                ' Comparamos os nomes nos arrays originais usando os índices
                If UCase(corretoresDisponiveis(indices(i))) > UCase(corretoresDisponiveis(indices(j))) Then
                    ' Troca apenas a posição no array de índices
                    tempIdx = indices(i)
                    indices(i) = indices(j)
                    indices(j) = tempIdx
                End If
            Next
        Next

        ' 3. Gerar o HTML seguindo a ordem do array de índices
        For k = 0 To corretoresCount - 1
            Dim idx
            idx = indices(k) ' Pega o índice ordenado
            
            Response.Write "<option value=""" & corretoresDisponiveis(idx) & """"
            
            If corretoresDisponiveis(idx) = corretor Then
                Response.Write " selected"
            End If
            
            Dim vgvFormatado
            If corretoresVGV(idx) > 0 Then
                vgvFormatado = " (R$ " & FormatNumber(corretoresVGV(idx), 0) & ")"
            Else
                vgvFormatado = " (R$ 0)"
            End If
            
            Response.Write ">" & corretoresDisponiveis(idx) & vgvFormatado & "</option>"
        Next
    End If

    If corretoresCount = 0 Then
        Response.Write "<option value="""">Nenhum corretor disponível</option>"
    End If
    %>
</select>
                        <% If corretoresCount = 0 Then %>
                        <small class="text-warning">Nenhum corretor encontrado</small>
                        <% End If %>
                        <% If gerencia <> "" Then %>
                        <small class="text-info">
                            <i class="fas fa-info-circle me-1"></i>Mostrando apenas corretores da gerência: <strong><%=gerencia%></strong>
                        </small>
                        <% End If %>
                    </div>
                    
                    <div class="col-md-4"></div>
                    
                    <div class="col-12 d-flex gap-3 mt-2">
                        <button type="submit" class="btn btn-primary btn-filter">
                            <i class="fas fa-search me-2"></i>Aplicar Filtros
                        </button>
                        <a href="gestao_vendas_diretoria3.asp" class="btn btn-secondary btn-filter">
                            <i class="fas fa-times me-2"></i>Limpar Filtro
                        </a>
                    </div>
                </form>
            </div>
        </div>

        <% if isFiltered then %>
        <!-- FILTROS ATIVOS -->
        <div class="card mb-4 border-warning">
            <div class="card-header text-white" style="background: #ffc107;">
                <i class="fas fa-check-circle me-2"></i>Filtros Ativos
                <% If anoSelecionadoAutomaticamente Then %>
                <span class="float-end badge bg-info">Filtro automático</span>
                <% End If %>
            </div>
            <div class="card-body">
                <div class="row g-2">
                    <% if Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then %>
                    <div class="col-auto">
                        <span class="badge bg-primary p-2"><i class="fas fa-building me-1"></i> Diretoria ID: <%=diretoriaID%></span>
                    </div>
                    <% end if %>
                    <% if ano <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-info p-2">
                            <i class="fas fa-calendar me-1"></i> Ano: <%=ano%>
                            <% If anoSelecionadoAutomaticamente Then %>
                            <span class="badge bg-light text-dark ms-1">Automático</span>
                            <% End If %>
                        </span>
                    </div>
                    <% end if %>
                    <% if mes <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-success p-2"><i class="fas fa-calendar-alt me-1"></i> Mês: <%=MonthName(mes, False)%></span>
                    </div>
                    <% end if %>
                    <% if trimestre <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-warning text-dark p-2"><i class="fas fa-chart-pie me-1"></i> 
                        <%Select Case trimestre%>
                            <%Case 1%>1º Trim<%Case 2%>2º Trim<%Case 3%>3º Trim<%Case 4%>4º Trim<%End Select%>
                        </span>
                    </div>
                    <% end if %>
                    <% if semestre <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-danger p-2"><i class="fas fa-chart-line me-1"></i> 
                        <%Select Case semestre%><%Case 1%>1º Sem<%Case 2%>2º Sem<%End Select%>
                        </span>
                    </div>
                    <% end if %>
                    <% if gerencia <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-success p-2"><i class="fas fa-user-tie me-1"></i> Gerência: <%=gerencia%></span>
                    </div>
                    <% end if %>
                    <% if corretor <> "" then %>
                    <div class="col-auto">
                        <span class="badge bg-dark p-2"><i class="fas fa-user me-1"></i> Corretor: <%=UCase(corretor)%></span>
                    </div>
                    <% end if %>
                </div>
            </div>
        </div>
        <% end if %>
        
        <div class="resultados">
            <%
            If isFiltered Then
                ' 1. Total de unidades vendidas
                sql = "SELECT COUNT(*) as TotalUnidades FROM Vendas" & whereClause
                rs.Open sql, conn
                totalUnidades = rs("TotalUnidades")
                rs.Close
                
                ' 2. Total VGV (Valor Geral de Vendas)
                sql = "SELECT SUM(ValorUnidade) as TotalVGV FROM Vendas" & whereClause
                rs.Open sql, conn

                totalVGV = 0
                If Not rs.EOF Then
                    totalVGV = (rs("TotalVGV"))
                End If
                rs.Close

                ' 3. Dados para os cards de VGV por mês - **MODIFICADO: AGORA SEMPRE BUSCAR DADOS PARA OS CARDS**
                Dim mesesVGV(12), mesesUnidades(12), mesesLabels(12), arrMesesNome(12)
                Dim mesesParaExibir(12)
                
                ' Inicializar array de nomes dos meses
                For i = 1 To 12
                    arrMesesNome(i) = MonthName(i, False)
                    mesesLabels(i) = Left(arrMesesNome(i), 3)
                    mesesVGV(i) = 0
                    mesesUnidades(i) = 0
                    mesesParaExibir(i) = True ' Inicialmente todos os meses serão exibidos
                Next
                
                ' **VERIFICAR SE HÁ FILTRO DE MÊS, TRIMESTRE OU SEMESTRE**
                Dim filtroMesEspecifico, filtroTrimestre, filtroSemestre
                filtroMesEspecifico = (mes <> "")
                filtroTrimestre = (trimestre <> "")
                filtroSemestre = (semestre <> "")
                
                ' **SE HOUVER FILTRO DE MÊS ESPECÍFICO, mostrar apenas aquele mês**
                If filtroMesEspecifico Then
                    For i = 1 To 12
                        mesesParaExibir(i) = (CInt(mes) = i)
                    Next
                ' **SE HOUVER FILTRO DE TRIMESTRE, mostrar apenas meses do trimestre**
                ElseIf filtroTrimestre Then
                    Select Case CInt(trimestre)
                        Case 1 ' Jan-Mar
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 1 And i <= 3)
                            Next
                        Case 2 ' Abr-Jun
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 4 And i <= 6)
                            Next
                        Case 3 ' Jul-Set
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 7 And i <= 9)
                            Next
                        Case 4 ' Out-Dez
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 10 And i <= 12)
                            Next
                    End Select
                ' **SE HOUVER FILTRO DE SEMESTRE, mostrar apenas meses do semestre**
                ElseIf filtroSemestre Then
                    Select Case CInt(semestre)
                        Case 1 ' Jan-Jun
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 1 And i <= 6)
                            Next
                        Case 2 ' Jul-Dez
                            For i = 1 To 12
                                mesesParaExibir(i) = (i >= 7 And i <= 12)
                            Next
                    End Select
                End If
                
                ' Variável para controlar se podemos mostrar o gráfico
                ' MODIFICAÇÃO: Não mostrar gráfico se filtro de mês, gerência ou corretor estiver ativo
                Dim podeMostrarGrafico
                podeMostrarGrafico = (ano <> "" And IsNumeric(ano) And mes = "" And gerencia = "" And corretor = "")
                
                ' **SEMPRE BUSCAR DADOS PARA OS CARDOS DOS MESES (VGV E UNIDADES)**
                ' Buscar dados de VGV e unidades por mês
                sql = "SELECT MesVenda, COUNT(*) as Unidades, SUM(ValorUnidade) as VGVMes FROM Vendas" & whereClause & " GROUP BY MesVenda ORDER BY MesVenda"
                rs.Open sql, conn
                
                Do While Not rs.EOF
                    If Not IsNull(rs("MesVenda")) Then
                        mesNum = CInt(rs("MesVenda"))
                        If mesNum >= 1 And mesNum <= 12 Then
                            mesesVGV(mesNum) = rs("VGVMes")
                            mesesUnidades(mesNum) = rs("Unidades")
                        End If
                    End If
                    rs.MoveNext
                Loop
                rs.Close
                
                ' **CALCULAR TÍTULO DINÂMICO BASEADO NOS FILTROS**
                Dim tituloPeriodo
                If filtroMesEspecifico Then
                    tituloPeriodo = MonthName(CInt(mes), False) & " de " & ano
                ElseIf filtroTrimestre Then
                    tituloPeriodo = trimestre & "º Trimestre de " & ano
                ElseIf filtroSemestre Then
                    tituloPeriodo = semestre & "º Semestre de " & ano
                Else
                    tituloPeriodo = ano
                End If
                
                ' Adicionar gerência ao título se filtrada
                If gerencia <> "" Then
                    tituloPeriodo = tituloPeriodo & " - Gerência: " & gerencia
                End If
                
                ' Adicionar corretor ao título se filtrado
                If corretor <> "" Then
                    tituloPeriodo = tituloPeriodo & " - Corretor: " & UCase(corretor)
                End If
                %>
                
                <!-- Título do período exibido -->
                <div class="text-center mb-3">
                    <h5 class="text-muted">
                        VGV por mês em <strong><%=tituloPeriodo%></strong>
                        <% If anoSelecionadoAutomaticamente Then %>
                            <span class="badge bg-info ms-2"><i class="fas fa-robot me-1"></i> Filtro automático</span>
                        <% End If %>
                        <% If mes <> "" And IsNumeric(mes) Then %>
                            <span class="badge bg-info ms-2"><%=arrMesesNome(CInt(mes))%> selecionado</span>
                        <% End If %>
                        <% If gerencia <> "" Then %>
                            <span class="badge bg-success ms-2"><i class="fas fa-user-tie me-1"></i> Gerência: <%=gerencia%></span>
                        <% End If %>
                        <% If corretor <> "" Then %>
                            <span class="badge bg-dark ms-2"><i class="fas fa-user me-1"></i> Corretor: <%=UCase(corretor)%></span>
                        <% End If %>
                    </h5>
                </div>
                
                <!-- 12 CARDS DE VGV POR MÊS - **SEMPRE MOSTRAR, MESMO COM FILTRO DE GERÊNCIA OU CORRETOR** -->
                <div class="row mb-4 g-3">
                    <%
                    Dim valorMesFormatado, unidadesMesFormatado, bgClass, isMesFiltrado, mesesExibidos, cardClass
                    mesesExibidos = 0
                    
                    For i = 1 To 12
                        ' Verificar se este mês deve ser exibido baseado nos filtros
                        If mesesParaExibir(i) Then
                            mesesExibidos = mesesExibidos + 1
                            
                            ' Determinar se é o mês filtrado
                            isMesFiltrado = False
                            If mes <> "" And IsNumeric(mes) Then
                                If CInt(mes) = i Then
                                    isMesFiltrado = True
                                End If
                            End If
                            
                            ' Determinar classe CSS baseada nos dados
                            If mesesVGV(i) > 0 Then
                                If isMesFiltrado Then
                                    bgClass = "card-mes-filtrado"
                                    cardClass = "card-mes-completo"
                                Else
                                    bgClass = "card-mes-com-dados"
                                    cardClass = "card-mes-completo"
                                End If
                                
                                ' Formatar valor para exibição
                                If mesesVGV(i) >= 1000000 Then
                                    valorMesFormatado = "R$ " & FormatNumber(mesesVGV(i)/1000000, 2) & " M"
                                ElseIf mesesVGV(i) >= 1000 Then
                                    valorMesFormatado = "R$ " & FormatNumber(mesesVGV(i)/1000, 0) & " mil"
                                Else
                                    valorMesFormatado = "R$ " & FormatNumber(mesesVGV(i), 0)
                                End If
                                
                                ' Formatar unidades
                                unidadesMesFormatado = FormatNumber(mesesUnidades(i), 0) & " unid."
                            Else
                                bgClass = "card-mes-vazio"
                                cardClass = "card-mes-completo card-mes-sem-dados"
                                valorMesFormatado = "R$ 0"
                                unidadesMesFormatado = "0 unid."
                            End If
                    %>
                    <div class="col-6 col-md-3 col-lg-2">
                        <div class="card h-100 shadow-sm <%=cardClass%> <%=bgClass%>">
                            <div class="card-body text-center p-3">
                                <h6 class="mb-2 fw-bold">
                                    <%=UCase(mesesLabels(i))%>
                                </h6>
                                <div class="valor-mes mb-2">
                                    <%=valorMesFormatado%>
                                </div>
                                <div class="unidades-mes">
                                    <%=unidadesMesFormatado%>
                                </div>
                                <% If isMesFiltrado Then %>
                                    <small class="badge bg-danger text-white mt-2">FILTRADO</small>
                                <% ElseIf mesesVGV(i) = 0 Then %>
                                    <small class="text-muted mt-2">Sem vendas</small>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                    <%
                        End If ' If mesesParaExibir(i)
                    Next
                    
                    ' Se nenhum mês for exibido (filtro muito restritivo)
                    If mesesExibidos = 0 Then
                        %>
                        <div class="col-12">
                            <div class="alert alert-info text-center">
                                <i class="fas fa-info-circle me-2"></i>Nenhum mês corresponde aos filtros aplicados.
                            </div>
                        </div>
                        <%
                    End If
                    %>
                </div>
                
                <%
                ' Exibir totais gerais
                %>
                <div class="row mb-4">
                    <div class="col-md-6">
                        <div class="metric-card" style="background: #3498db;">
                            <div class="metric-label"><i class="fas fa-cube me-2"></i> Total de Unidades Vendidas</div>
                            <div class="metric-value"><%=FormatNumber(totalUnidades, 0)%></div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="metric-card" style="background: #2ecc71;">
                            <div class="metric-label"><i class="fas fa-money-bill-wave me-2"></i> Total VGV</div>
                            <div class="metric-value">R$ <%=FormatNumber(totalVGV, 2)%></div>
                        </div>
                    </div>
                </div>
                
                <!-- GRÁFICO DE VGV POR MÊS - **NÃO EXIBIR QUANDO MÊS, GERÊNCIA OU CORRETOR ESTIVER FILTRADO** -->
                <% If podeMostrarGrafico Then %>
                <div class="card mb-4">
                    <div class="card-header text-white" style="background: #9b59b6;">
                        <i class="fas fa-chart-line me-2"></i>VGV por Mês
                        <% 
                        If filtroMesEspecifico Then
                            Response.Write "<span class='float-end badge bg-info'>Mês específico: " & MonthName(CInt(mes), False) & "</span>"
                        ElseIf filtroTrimestre Then
                            Response.Write "<span class='float-end badge bg-info'>Trimestre " & trimestre & "</span>"
                        ElseIf filtroSemestre Then
                            Response.Write "<span class='float-end badge bg-info'>Semestre " & semestre & "</span>"
                        End If
                        %>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoVGV"></canvas>
                        </div>
                    </div>
                </div>
                <% ElseIf mes <> "" Or gerencia <> "" Or corretor <> "" Then %>
                <div class="alert alert-info">
                    <i class="fas fa-info-circle me-2"></i>Gráfico não exibido quando filtro de <strong>mês</strong>, <strong>gerência</strong> ou <strong>corretor</strong> está ativo.
                </div>
                <% Else %>
                <div class="alert alert-warning">
                    <i class="fas fa-exclamation-triangle me-2"></i>Selecione um ano para visualizar o gráfico
                </div>
                <% End If %>
                
                <!-- SEÇÕES DE ANÁLISE -->
                <div class="row">
                    <!-- VENDAS POR DIRETORIA -->
                    <div class="col-md-6">
                        <div class="card h-100">
                            <div class="card-header text-white" style="background: #e74c3c;">
                                <i class="fas fa-building me-2"></i>Vendas por Diretoria
                            </div>
                            <div class="card-body">
                                <%
                                sql = "SELECT Vendas.Diretoria, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                                      "FROM Vendas" & whereClause & " GROUP BY Diretoria ORDER BY SUM(Vendas.ValorUnidade) DESC"
                                rs.Open sql, conn
                                
                                If Not rs.EOF Then
                                    totalUnidadesDiretoria = 0
                                    totalVGVDiretoria = 0
                                    %>
                                    <div class="table-responsive">
                                        <table class="table table-sm">
                                            <thead>
                                                <tr>
                                                    <th>Diretoria</th>
                                                    <th class="text-center">Unid.</th>
                                                    <th class="text-end">VGV</th>
                                                    <th class="text-end">%</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <%
                                                Do While Not rs.EOF
                                                    vgvDiretoria = (rs("VGV"))
                                                    If totalVGV <> 0 Then
                                                        percentual = (vgvDiretoria / totalVGV) * 100
                                                    Else
                                                        percentual = 0
                                                    End If
                                                    
                                                    totalUnidadesDiretoria = totalUnidadesDiretoria + rs("Unidades")
                                                    totalVGVDiretoria = totalVGVDiretoria + vgvDiretoria
                                                    %>
                                                    <tr>
                                                        <td><%=rs("Diretoria")%></td>
                                                        <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                        <td class="text-end">R$ <%=FormatNumber(vgvDiretoria, 0)%></td>
                                                        <td class="text-end"><span class="badge bg-primary badge-percent"><%=FormatNumber(percentual, 1)%>%</span></td>
                                                    </tr>
                                                    <%
                                                    rs.MoveNext
                                                Loop
                                                %>
                                                <tr class="total">
                                                    <td><strong>TOTAL</strong></td>
                                                    <td class="text-center"><strong><%=FormatNumber(totalUnidadesDiretoria, 0)%></strong></td>
                                                    <td class="text-end"><strong>R$ <%=FormatNumber(totalVGVDiretoria, 0)%></strong></td>
                                                    <td class="text-end"><strong>100%</strong></td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                    <%
                                Else
                                    %>
                                    <p class="text-center text-muted py-4">Nenhuma venda encontrada</p>
                                    <%
                                End If
                                rs.Close
                                %>
                            </div>
                        </div>
                    </div>
                    
                    <!-- VENDAS POR GERÊNCIA - MODIFICADO PARA DESTACAR GERÊNCIA FILTRADA -->
                    <div class="col-md-6">
                        <div class="card h-100">
                            <div class="card-header text-white" style="background: #3498db;">
                                <i class="fas fa-user-tie me-2"></i>Vendas por Gerência
                                <% If gerencia <> "" Then %>
                                <span class="float-end badge bg-success">Filtrada: <%=gerencia%></span>
                                <% End If %>
                            </div>
                            <div class="card-body">
                                <%
                                sql = "SELECT Vendas.Gerencia, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                                      "FROM Vendas" & whereClause & " GROUP BY Gerencia ORDER BY SUM(Vendas.ValorUnidade) DESC"
                                rs.Open sql, conn
                                
                                If Not rs.EOF Then
                                    totalUnidadesGerencia = 0
                                    totalVGVGerencia = 0
                                    %>
                                    <div class="table-responsive">
                                        <table class="table table-sm">
                                            <thead>
                                                <tr>
                                                    <th>Gerência</th>
                                                    <th class="text-center">Unid.</th>
                                                    <th class="text-end">VGV</th>
                                                    <th class="text-end">%</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <%
                                                Do While Not rs.EOF
                                                    vgvGerencia = (rs("VGV"))
                                                    If totalVGV <> 0 Then
                                                        percentual = (vgvGerencia / totalVGV) * 100
                                                    Else
                                                        percentual = 0
                                                    End If
                                                    
                                                    totalUnidadesGerencia = totalUnidadesGerencia + rs("Unidades")
                                                    totalVGVGerencia = totalVGVGerencia + vgvGerencia
                                                    
                                                    ' Verificar se é a gerência filtrada
                                                    Dim linhaClass
                                                    linhaClass = ""
                                                    If gerencia <> "" And rs("Gerencia") = gerencia Then
                                                        linhaClass = "gerencia-filtrada"
                                                    End If
                                                    %>
                                                    <tr class="<%=linhaClass%>">
                                                        <td>
                                                            <%=rs("Gerencia")%>
                                                            <% If gerencia <> "" And rs("Gerencia") = gerencia Then %>
                                                            <span class="badge bg-success ms-2">FILTRADA</span>
                                                            <% End If %>
                                                        </td>
                                                        <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                        <td class="text-end">R$ <%=FormatNumber(vgvGerencia, 0)%></td>
                                                        <td class="text-end"><span class="badge bg-success badge-percent"><%=FormatNumber(percentual, 1)%>%</span></td>
                                                    </tr>
                                                    <%
                                                    rs.MoveNext
                                                Loop
                                                %>
                                                <tr class="total">
                                                    <td><strong>TOTAL</strong></td>
                                                    <td class="text-center"><strong><%=FormatNumber(totalUnidadesGerencia, 0)%></strong></td>
                                                    <td class="text-end"><strong>R$ <%=FormatNumber(totalVGVGerencia, 0)%></strong></td>
                                                    <td class="text-end"><strong>100%</strong></td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                    <%
                                Else
                                    %>
                                    <p class="text-center text-muted py-4">Nenhuma venda encontrada</p>
                                    <%
                                End If
                                rs.Close
                                %>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- OUTRAS SEÇÕES -->
                <div class="row mt-4">
                    <!-- VENDAS POR EMPREENDIMENTO -->
                    <div class="col-md-4">
                        <div class="card h-100">
                            <div class="card-header text-white" style="background: #2ecc71;">
                                <i class="fas fa-building me-2"></i> Empreendimentos
                            </div>
                            <div class="card-body">
                                <%
                                sql = "SELECT NomeEmpreendimento, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                                      "FROM Vendas" & whereClause & " GROUP BY NomeEmpreendimento ORDER BY SUM(Vendas.ValorUnidade) DESC"
                                rs.Open sql, conn
                                
                                If Not rs.EOF Then
                                    %>
                                    <div class="list-group list-group-flush">
                                        <%
                                        cont = 0
                                        Do While Not rs.EOF
                                            cont = cont + 1
                                            vgvEmpreendimento = (rs("VGV"))
                                            If totalVGV <> 0 Then
                                                percentual = (vgvEmpreendimento / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            %>
                                            <div class="list-group-item border-0 py-2 px-0">
                                                <div class="d-flex justify-content-between align-items-center">
                                                    <div>
                                                        <strong><%=cont &"-"&rs("NomeEmpreendimento")%></strong><br>
                                                        <small class="text-muted"><%=rs("Unidades")%> unidades</small>
                                                    </div>
                                                    <div class="text-end">
                                                        <div class="fw-bold">R$ <%=FormatNumber(vgvEmpreendimento, 0)%></div>
                                                        <small class="text-muted"><%=FormatNumber(percentual, 1)%>%</small>
                                                    </div>
                                                </div>
                                            </div>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                    </div>
                                    <%
                                Else
                                    %>
                                    <p class="text-center text-muted py-4">Nenhum dado</p>
                                    <%
                                End If
                                rs.Close
                                %>
                            </div>
                        </div>
                    </div>
                    
                    <!-- VENDAS POR EMPRESA -->
                    <div class="col-md-4">
                        <div class="card h-100">
                            <div class="card-header text-white" style="background: #9b59b6;">
                                <i class="fas fa-industry me-2"></i>Vendas por Empresa
                            </div>
                            <div class="card-body">
                                <%
                                sql = "SELECT NomeEmpresa, COUNT(*) as Unidades, SUM(ValorUnidade) as VGV " & _
                                      "FROM Vendas" & whereClause & " GROUP BY Nomeempresa ORDER BY SUM(ValorUnidade) DESC"
                                rs.Open sql, conn
                                
                                If Not rs.EOF Then
                                    %>
                                    <div class="list-group list-group-flush">
                                        <%
                                        cont = 0
                                        Do While Not rs.EOF
                                            cont = cont + 1 
                                            vgvEmpresa = (rs("VGV"))
                                            If totalVGV <> 0 Then
                                                percentual = (vgvEmpresa / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            %>
                                            <div class="list-group-item border-0 py-2 px-0">
                                                <div class="d-flex justify-content-between align-items-center">
                                                    <div>
                                                        <strong><%=cont &"-"& rs("NomeEmpresa")%></strong><br>
                                                        <small class="text-muted"><%=rs("Unidades")%> unidades</small>
                                                    </div>
                                                    <div class="text-end">
                                                        <div class="fw-bold">R$ <%=FormatNumber(vgvEmpresa, 0)%></div>
                                                        <small class="text-muted"><%=FormatNumber(percentual, 1)%>%</small>
                                                    </div>
                                                </div>
                                            </div>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                    </div>
                                    <%
                                Else
                                    %>
                                    <p class="text-center text-muted py-4">Nenhum dado</p>
                                    <%
                                End If
                                rs.Close
                                %>
                            </div>
                        </div>
                    </div>
                    
                    <!-- LISTA COMPLETA DE CORRETORES ORDENADA POR VGV -->
                    <div class="col-md-4">
                        <div class="card h-100">
                            <div class="card-header text-white" style="background: #f39c12;">
                                <i class="fas fa-users me-2"></i> 
                                <% If corretor <> "" Then %>
                                    Corretor Selecionado
                                <% Else %>
                                    Lista de Corretores (Ordenados por VGV)
                                <% End If %>
                                <% If gerencia <> "" And corretor = "" Then %>
                                <span class="float-end badge bg-success">Filtrado por gerência: <%=gerencia%></span>
                                <% Else %>
                                <span class="float-end badge bg-light text-dark">
                                    <% If corretor <> "" Then %>
                                        1 corretor
                                    <% Else %>
                                        <%=corretoresCount%> corretores
                                    <% End If %>
                                </span>
                                <% End If %>
                            </div>
                            <div class="card-body">
                                <%
                                ' Verificar se há corretores
                                If corretoresCount > 0 Then
                                    ' Os corretores já estão ordenados por VGV descendente no array
                                    %>
                                    <div class="list-group list-group-flush">
                                        <%
                                        Dim corretoresExibidos, rankingClass
                                        corretoresExibidos = 0
                                        
                                        For i = 0 To corretoresCount - 1
                                            ' *** MODIFICAÇÃO: Se houver filtro de corretor, mostrar apenas o corretor filtrado ***
                                            If corretor = "" Or corretoresDisponiveis(i) = corretor Then
                                                corretoresExibidos = corretoresExibidos + 1
                                                
                                                ' Determinar classe de ranking
                                                If i = 0 Then
                                                    rankingClass = "ranking-1"
                                                ElseIf i = 1 Then
                                                    rankingClass = "ranking-2"
                                                ElseIf i = 2 Then
                                                    rankingClass = "ranking-3"
                                                Else
                                                    rankingClass = "ranking-outros"
                                                End If
                                                
                                                ' Usar VGV já armazenado no array
                                                unidadesCorretor = 0
                                                vgvCorretor = corretoresVGV(i)
                                                
                                                ' Se precisar buscar unidades, podemos fazer uma consulta adicional
                                                ' Mas para simplificar, vamos usar o VGV que já temos
                                                
                                                ' Determinar classe do item
                                                Dim listItemClass
                                                listItemClass = "list-group-item corretor-item"
                                                If corretoresDisponiveis(i) = corretor Then
                                                    listItemClass = listItemClass & " corretor-filtrado"
                                                End If
                                                %>
                                                <a href="javascript:void(0)" onclick="document.getElementById('corretor').value='<%=corretoresDisponiveis(i)%>'; document.forms[0].submit();" class="<%=listItemClass%> text-decoration-none">
                                                    <div class="d-flex justify-content-between align-items-center">
                                                        <div class="d-flex align-items-center">
                                                            <% If corretor = "" Then %>
                                                            <span class="ranking-badge <%=rankingClass%>"><%=i+1%></span>
                                                            <% End If %>
                                                            <div>
                                                                <strong><%=corretoresDisponiveis(i)%></strong>
                                                                <% If corretoresDisponiveis(i) = corretor Then %>
                                                                    <span class="badge bg-warning text-dark ms-2">Filtrado</span>
                                                                <% End If %>
                                                            </div>
                                                        </div>
                                                        <div class="text-end">
                                                            <div>
                                                                <span class="ms-2"><strong>R$ <%=FormatNumber(vgvCorretor, 0)%></strong></span>
                                                            </div>
                                                            <% If totalVGV <> 0 And vgvCorretor > 0 Then %>
                                                                <small class="text-muted"><%=FormatNumber((vgvCorretor/totalVGV)*100, 1)%>% do total</small>
                                                            <% End If %>
                                                        </div>
                                                    </div>
                                                </a>
                                                <%
                                            End If
                                        Next
                                        
                                        ' Se não houver corretores para exibir (filtro muito específico)
                                        If corretoresExibidos = 0 Then
                                            %>
                                            <div class="list-group-item text-center text-muted py-4">
                                                Nenhum corretor corresponde aos filtros aplicados.
                                            </div>
                                            <%
                                        End If
                                        %>
                                    </div>
                                    
                                    <!-- RESUMO DE CORRETORES -->
                                    <div class="mt-3">
                                        <div class="alert alert-light">
                                            <div class="row">
                                                <div class="col-12">
                                                    <strong>Total de Corretores:</strong> 
                                                    <% If corretor <> "" Then %>
                                                        <span class="badge bg-warning text-dark">1 filtrado</span> de <%=corretoresCount%>
                                                    <% Else %>
                                                        <%=corretoresCount%>
                                                        <% If gerencia <> "" Then %>
                                                            <span class="badge bg-success ms-2">da gerência: <%=gerencia%></span>
                                                        <% End If %>
                                                    <% End If %>
                                                </div>
                                                <% If corretor = "" Then %>
                                                <div class="col-12 mt-1">
                                                    <strong>Média VGV/Corretor:</strong> 
                                                    <% If corretoresCount > 0 Then %>
                                                        R$ <%=FormatNumber(totalVGV/corretoresCount, 0)%>
                                                    <% Else %>
                                                        R$ 0
                                                    <% End If %>
                                                </div>
                                                <% End If %>
                                            </div>
                                        </div>
                                        
                                        <% If corretor <> "" Then %>
                                        <div class="d-grid gap-2">
                                            <a href="javascript:void(0)" onclick="document.getElementById('corretor').value=''; document.forms[0].submit();" class="btn btn-sm btn-outline-warning">
                                                <i class="fas fa-times me-1"></i> Remover Filtro de Corretor
                                            </a>
                                        </div>
                                        <% ElseIf corretoresCount > 1 Then %>
                                        <div class="d-grid gap-2">
                                            <button type="button" class="btn btn-sm btn-outline-secondary" onclick="mostrarTodosCorretores()">
                                                <i class="fas fa-list me-1"></i> Ver todos os corretores
                                            </button>
                                        </div>
                                        <% End If %>
                                    </div>
                                    <%
                                Else
                                    %>
                                    <div class="text-center text-muted py-4">
                                        Nenhum corretor encontrado
                                        <% If gerencia <> "" Then %>
                                        <br><small>para a gerência: <strong><%=gerencia%></strong></small>
                                        <% End If %>
                                    </div>
                                    <%
                                End If
                                %>
                            </div>
                        </div>
                    </div>
                </div>
                
                <%
            Else ' Se não houver filtro
                %>
                <div class="card border-info">
                    <div class="card-header text-white" style="background: #17a2b8;">
                        <i class="fas fa-info-circle me-2"></i>Instruções
                    </div>
                    <div class="card-body text-center py-5">
                        <i class="fas fa-filter fa-3x text-muted mb-3"></i>
                        <h4 class="text-muted mb-3">Selecione os filtros acima para visualizar os dados de vendas</h4>
                        <% If Not IsNull(diretoriaID) And Trim(CStr(diretoriaID)) <> "" And IsNumeric(diretoriaID) Then %>
                            <div class="alert alert-info">
                                O filtro de <strong>Diretoria (ID: <%=diretoriaID%>)</strong> já está aplicado automaticamente para o seu perfil.
                            </div>
                        <% End If %>
                        <% If anosCount = 0 Then %>
                            <div class="alert alert-warning mt-3">
                                <i class="fas fa-exclamation-triangle me-2"></i>Nenhum ano com vendas encontrado para sua diretoria.
                            </div>
                        <% End If %>
                    </div>
                </div>
                <%
            End If
            
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
            %>
        </div>
    </div>
    
    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    
    <script>
    <%
    If isFiltered And podeMostrarGrafico Then
        ' Preparar dados JavaScript para o gráfico - **APENAS MESES RELEVANTES AO FILTRO**
        Response.Write "const mesesLabels = ["
        Dim labelCount
        labelCount = 0
        
        For i = 1 To 12
            If mesesParaExibir(i) Then
                If labelCount > 0 Then
                    Response.Write ", "
                End If
                Response.Write """" & UCase(Left(arrMesesNome(i), 3)) & """"
                labelCount = labelCount + 1
            End If
        Next
        Response.Write "];" & vbCrLf
                       
        Response.Write "const vgvData = ["
        Dim dataCount
        dataCount = 0
        
        For i = 1 To 12
            If mesesParaExibir(i) Then
                If dataCount > 0 Then
                    Response.Write ", "
                End If
                
                ' *** CORREÇÃO: Garantir formato numérico com ponto decimal ***
                Dim valorParaJS
                valorParaJS = mesesVGV(i)
                
                ' Se for string, converter para número com ponto decimal
                If VarType(valorParaJS) = vbString Then
                    ' Remover formatação brasileira
                    valorParaJS = Replace(valorParaJS, ".", "") ' Remove separadores de milhar
                    valorParaJS = Replace(valorParaJS, ",", ".") ' Converte vírgula para ponto
                    
                    ' Remover caracteres não numéricos
                    valorParaJS = Replace(valorParaJS, "R$", "")
                    valorParaJS = Replace(valorParaJS, " ", "")
                    valorParaJS = Trim(valorParaJS)
                    
                    If IsNumeric(valorParaJS) Then
                        valorParaJS = CDbl(valorParaJS)
                    Else
                        valorParaJS = 0
                    End If
                End If
                
                ' Escrever valor garantindo ponto decimal
                Response.Write Replace(CStr(valorParaJS), ",", ".")
                dataCount = dataCount + 1
            End If
        Next
        Response.Write "];" & vbCrLf
        
        ' *** ADICIONAR CONVERSÃO SEGURA NO LADO DO JAVASCRIPT ***
        Response.Write "// Converter strings para números se necessário" & vbCrLf
        Response.Write "if (vgvData.length > 0) {" & vbCrLf
        Response.Write "    for(let i = 0; i < vgvData.length; i++) {" & vbCrLf
        Response.Write "        if (typeof vgvData[i] === 'string') {" & vbCrLf
        Response.Write "            // Remover qualquer vírgula e converter para número" & vbCrLf
        Response.Write "            vgvData[i] = parseFloat(vgvData[i].replace(',', '.'));" & vbCrLf
        Response.Write "        }" & vbCrLf
        Response.Write "        // Garantir que seja número (NaN = 0)" & vbCrLf
        Response.Write "        if (isNaN(vgvData[i])) {" & vbCrLf
        Response.Write "            vgvData[i] = 0;" & vbCrLf
        Response.Write "        }" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "}" & vbCrLf
        
        Response.Write "console.log('Dados do gráfico (filtrados):', vgvData);" & vbCrLf
        Response.Write "console.log('Labels (filtrados):', mesesLabels);" & vbCrLf
        
        ' Encontrar qual índice corresponde ao mês filtrado (para destaque)
        Dim indiceFiltrado, contadorIndice
        indiceFiltrado = -1
        contadorIndice = 0
        
        If filtroMesEspecifico Then
            For i = 1 To 12
                If mesesParaExibir(i) Then
                    If CInt(mes) = i Then
                        indiceFiltrado = contadorIndice
                        Exit For
                    End If
                    contadorIndice = contadorIndice + 1
                End If
            Next
        End If
    %>
    
    // Aguardar o carregamento da página
    document.addEventListener('DOMContentLoaded', function() {
        // Configuração do gráfico
        const ctx = document.getElementById('graficoVGV').getContext('2d');
        
        // Verificar se já existe um gráfico e destruí-lo
        if (window.myChart) {
            window.myChart.destroy();
        }
        
        // Verificar se há dados
        console.log('Dados do gráfico:', vgvData);
        console.log('Tipo dos dados:', typeof vgvData[0]);
        console.log('Labels:', mesesLabels);
        
        // *** CORREÇÃO FINAL: Garantir que todos os dados sejam números ***
        const vgvDataNumerico = vgvData.map(item => {
            if (typeof item === 'string') {
                // Remover vírgula e converter para número
                return parseFloat(item.replace(',', '.'));
            }
            return Number(item);
        });
        
        console.log('Dados numéricos corrigidos:', vgvDataNumerico);
        
        // Registrar o plugin datalabels
        Chart.register(ChartDataLabels);
        
        // Criar novo gráfico
        window.myChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: mesesLabels,
                datasets: [{
                    label: 'VGV por Mês',
                    data: vgvDataNumerico,
                    backgroundColor: function(context) {
                        const index = context.dataIndex;
                        const value = context.dataset.data[index];
                        // Destacar a barra do mês filtrado (se houver)
                        <% 
                        If filtroMesEspecifico And indiceFiltrado >= 0 Then
                            Response.Write "if (index == " & indiceFiltrado & ") {"
                            Response.Write "return 'rgba(231, 76, 60, 0.7)';"
                            Response.Write "}"
                        End If
                        %>
                        return value > 0 ? 'rgba(255, 165, 0, 0.7)' : 'rgba(200, 200, 200, 0.3)';
                    },
                    borderColor: function(context) {
                        const index = context.dataIndex;
                        <% 
                        If filtroMesEspecifico And indiceFiltrado >= 0 Then
                            Response.Write "if (index == " & indiceFiltrado & ") {"
                            Response.Write "return 'rgb(231, 76, 60)';"
                            Response.Write "}"
                        End If
                        %>
                        return 'rgb(52, 152, 219)';
                    },
                    borderWidth: 1,
                    borderRadius: 4,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const value = context.parsed.y;
                                return 'VGV: R$ ' + value.toLocaleString('pt-BR', {
                                    minimumFractionDigits: 2, 
                                    maximumFractionDigits: 2
                                });
                            }
                        }
                    },
                    // CONFIGURAÇÃO DOS LABELS NAS BARRAS
                    datalabels: {
                        anchor: 'center',
                        align: 'center',
                        color: '#2c3e50',
                        font: {
                            weight: 'bold',
                            size: 11
                        },
                        rotation: -90, // Rotação de -90 graus (vertical)
                        formatter: function(value, context) {
                            // Só exibir se valor > 0
                            if (value <= 0) return '';
                            
                            // Formatar o valor
                            if (value >= 1000000) {
                                return 'R$' + (value / 1000000).toFixed(1).replace('.', ',') + 'M';
                            }
                            if (value >= 1000) {
                                return 'R$' + (value / 1000).toFixed(0) + 'K';
                            }
                            return 'R$' + value.toFixed(0);
                        }
                    }
                },
                scales: {
                    x: {
                        ticks: {
                            autoSkip: false,
                            maxRotation: 45,
                            minRotation: 45
                        },
                        grid: {
                            display: true
                        }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                if (value >= 1000000) {
                                    return 'R$' + (value / 1000000).toFixed(1).replace('.', ',') + 'M';
                                }
                                if (value >= 1000) {
                                    return 'R$' + (value / 1000).toFixed(0) + 'K';
                                }
                                return 'R$' + value.toFixed(0);
                            }
                        }
                    }
                }
            },
            plugins: [ChartDataLabels]
        });
        
        // Redimensionar gráfico quando a janela for redimensionada
        window.addEventListener('resize', function() {
            if (window.myChart) {
                window.myChart.resize();
            }
        });
    });
    
    // Função para mostrar todos os corretores
    function mostrarTodosCorretores() {
        // Esta função pode ser expandida para mostrar uma modal com todos os corretores
        alert('Para ver todos os corretores, remova o filtro de corretor selecionando "Todos os corretores" no campo de filtro.');
    }
    </script>
    <%
    End If
    %>
</body>
</html>