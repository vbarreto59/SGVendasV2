<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->           
<!--#include file="conSunSales.asp"-->       
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' =============================================
'      META – GERENCIAMENTO (BLOCO ANTES DO HTML)
' =============================================
Response.CodePage = 65001
Response.Charset = "utf-8"

Const ANO_MINIMO = 2026

' -----------------------------
'  ABRIR CONEXÕES
' -----------------------------
Dim connOrg, connSales
Set connOrg = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

connOrg.Open strConn
connSales.Open strConnSales


AtualizarMetaDiretoria()

' -----------------------------
'  RECEBER ANO E MÊS
' -----------------------------
Dim anoSelecionado, mesSelecionado
anoSelecionado = Request("ano")
mesSelecionado = Request("mes")

If anoSelecionado = "" Then anoSelecionado = Year(Now)
If mesSelecionado = "" Then mesSelecionado = Month(Now)

If IsNumeric(anoSelecionado) Then
    If CInt(anoSelecionado) < ANO_MINIMO Then anoSelecionado = ANO_MINIMO
Else
    anoSelecionado = ANO_MINIMO
End If

' -----------------------------
'  PREPARAR DICIONÁRIOS
' -----------------------------
Dim rs, sql

Dim dicGerencias
Set dicGerencias = Server.CreateObject("Scripting.Dictionary")

sql = "SELECT GerenciaID, NomeGerencia, DiretoriaID FROM Gerencias ORDER BY NomeGerencia"
Set rs = connOrg.Execute(sql)

Dim arrGerenciaInfo

Do While Not rs.EOF
    Dim nome, diretoria

    nome = CStr("" & rs("NomeGerencia"))
    diretoria = CLng("" & rs("DiretoriaID"))

    arrGerenciaInfo = Array(nome, diretoria)

    dicGerencias.Add CStr(rs("GerenciaID")), arrGerenciaInfo

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Dim dicMetas
Set dicMetas = Server.CreateObject("Scripting.Dictionary")

sql = "SELECT GerenciaID, ValorMeta FROM MetaGerencia WHERE Ano=" & anoSelecionado & " AND Mes=" & mesSelecionado
Set rs = connSales.Execute(sql)

Do While Not rs.EOF
    If Not dicMetas.Exists(CStr(rs("GerenciaID"))) Then
        dicMetas.Add CStr(rs("GerenciaID")), CDbl(rs("ValorMeta"))
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing


' -----------------------------
'  PROCESSAR POST (SALVAR METAS)
' -----------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    connSales.BeginTrans
    Dim totalInseridos
    totalInseridos = 0

    Dim key, valorRaw, valorConv, gerID, diretoriaID, gerenciaInfo

    For Each key In Request.Form
        If Left(key,8) = "gerencia" Then
            
            valorRaw = Trim(Request.Form(key))

            If valorRaw <> "" Then
                
                valorConv = Replace(valorRaw, ".", "")
                valorConv = Replace(valorConv, ",", ".")

                If IsNumeric(valorConv) Then
                    valorConv = CDbl(valorConv)

                    If valorConv > 0 Then

                        gerID = Mid(key, 9)
                        If Left(gerID,1) = "_" Then gerID = Mid(gerID,2)

                        If IsNumeric(gerID) Then
                            
                            gerID = CInt(gerID)

                            If dicGerencias.Exists(CStr(gerID)) Then
                                
                                gerenciaInfo = dicGerencias.Item(CStr(gerID))
                                diretoriaID = gerenciaInfo(1)

                                ' VALIDAR DIRETORIA
                                If Not IsNull(diretoriaID) And diretoriaID <> "" Then

                                    If Not IsNumeric(diretoriaID) Then diretoriaID = CLng(diretoriaID)

                                    ' Verificar se já existe meta
                                    Dim rsCheck, sqlCheck
                                    sqlCheck = "SELECT MetaGer_ID FROM MetaGerencia " & _
                                               "WHERE GerenciaID=" & gerID & " AND Ano=" & anoSelecionado & " AND Mes=" & mesSelecionado

                                    Set rsCheck = Server.CreateObject("ADODB.Recordset")
                                    rsCheck.Open sqlCheck, connSales, 1, 3

                                    

                                    valorConv = valorConv / 100
                                    valorConv = Replace(valorConv, ".","")
                                    valorConv = Replace(valorConv, ",",".")

                                    If rsCheck.EOF Then
                                        connSales.Execute "INSERT INTO MetaGerencia (GerenciaID, DiretoriaID, Ano, Mes, ValorMeta, Usuario) VALUES (" & _
                                            gerID & ", " & diretoriaID & ", " & anoSelecionado & ", " & mesSelecionado & ", " & Replace(valorConv, ",", ".") & ", '" & _
                                            Replace(Session("Usuario"), "'", "''") & "')"

                                    Else
                                        connSales.Execute "UPDATE MetaGerencia SET ValorMeta=" & valorConv & _
                                            ", Usuario='" & Replace(Session("Usuario"), "'", "''") & "' WHERE MetaGer_ID=" & rsCheck("MetaGer_ID")

                                    End If

                                    rsCheck.Close
                                    Set rsCheck = Nothing

                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

    If Err.Number = 0 Then
        connSales.CommitTrans
    Else
        connSales.RollbackTrans
        Err.Clear
    End If

    ' ===========================================
    '    CHAMAR A FUNÇÃO PARA ATUALIZAR A DIRETORIA
    ' ===========================================
    Call AtualizarMetaDiretoria()

    Response.Redirect "meta_gerenciamento.asp?ano=" & anoSelecionado & "&mes=" & mesSelecionado & "&saved=1"
End If


' -----------------------------
'  FUNÇÃO DE LOG
' -----------------------------
Sub RegistrarLog(msg)
    On Error Resume Next
    Dim fso, arq, caminho
    caminho = Server.MapPath("metas.log")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(caminho) Then
        Set arq = fso.OpenTextFile(caminho, 8, True)
    Else
        Set arq = fso.CreateTextFile(caminho, True)
    End If
    arq.WriteLine Now() & " - " & msg
    arq.Close
    Set arq = Nothing
    Set fso = Nothing
End Sub
%>
<%
' ===========================================
'     FUNÇÃO **NOVA** → ATUALIZAR METADIRETORIA
' ===========================================
Function AtualizarMetaDiretoria()

    Dim sql, rs, sqlInsert, usuarioLogado
    Dim DiretoriaId, Ano, Mes, TotalMetas
    
    usuarioLogado = Session("Usuario")
    If usuarioLogado = "" Then usuarioLogado = "Sistema"
    
    On Error Resume Next
    
    ' Limpar tabela
    connSales.Execute "DELETE FROM MetaDiretoria"
    If Err.Number <> 0 Then Exit Function
    
    ' Buscar soma das metas agrupadas por diretoria/ano/mês
    'sql = "SELECT g.DiretoriaID, mg.Ano, mg.Mes, SUM(mg.ValorMeta) AS Total " & _
       ''   "FROM MetaGerencia mg " & _
      ''    "INNER JOIN Gerencias g ON mg.GerenciaID = g.GerenciaID " & _
       ''   "WHERE mg.ValorMeta > 0 " & _
      ''    "GROUP BY g.DiretoriaID, mg.Ano, mg.Mes"

    sql = "SELECT MetaGerencia.DiretoriaId, MetaGerencia.Ano, MetaGerencia.Mes, Sum(MetaGerencia.ValorMeta) AS SomaDeValorMeta FROM MetaGerencia GROUP BY MetaGerencia.DiretoriaId, MetaGerencia.Ano, MetaGerencia.Mes HAVING (((Sum(MetaGerencia.ValorMeta))>0));"      
    
    Set rs = connSales.Execute(sql)
    
    Do Until rs.EOF
        DiretoriaId = rs("DiretoriaID")
        Ano = rs("Ano")
        Mes = rs("Mes")
        TotalMetas = rs("SomaDeValorMeta")
        
        
        sqlInsert = "INSERT INTO MetaDiretoria " & _
                        "(DiretoriaID, Ano, Mes, TotalMetas, Usuario) " & _
                        "VALUES (" & DiretoriaId & "," & Ano & "," & Mes & "," & _
                        TotalMetas & "," & "'" & Replace(usuarioLogado, "'", "''") & "')"
        'Response.write sqlInsert
        'Response.end                 
        
        connSales.Execute sqlInsert
    
        
        rs.MoveNext
    Loop
    
    ' Atualizar meta da empresa
    Call AtualizarMetaEmpresa()
    
    On Error GoTo 0
    
End Function

' ===============================================

Function AtualizarMetaEmpresa()

    Dim sql, rs, sqlInsert
    Dim Ano, Mes, TotalEmpresa
    
    On Error Resume Next
    
    ' Limpar tabela (apenas anos a partir de 2026)
    connSales.Execute "DELETE FROM MetaEmpresa WHERE Ano >= 2026"
    If Err.Number <> 0 Then Exit Function
    
    ' Buscar soma das metas da empresa agrupadas por ano/mês
    sql = "SELECT " & _
          "MetaDiretoria.Ano, " & _
          "MetaDiretoria.Mes, " & _
          "Sum(MetaDiretoria.TotalMetas) AS SomaDeTotalMetas " & _
          "FROM MetaDiretoria " & _
          "WHERE MetaDiretoria.Ano >= 2026 " & _
          "GROUP BY MetaDiretoria.Ano, MetaDiretoria.Mes " & _
          "HAVING Sum(MetaDiretoria.TotalMetas) > 0"
          'response.write sql 
          'response.end  
    
    Set rs = connSales.Execute(sql)
    
    Do Until rs.EOF
        Ano = rs("Ano")
        Mes = rs("Mes")
        TotalEmpresa = rs("SomaDeTotalMetas")
        
        ' Verificar se o valor não é nulo
        If Not IsNull(TotalEmpresa) And TotalEmpresa > 0 Then
            sqlInsert = "INSERT INTO MetaEmpresa " & _
                        "(Ano, Mes, Meta) " & _
                        "VALUES (" & Ano & "," & Mes & "," & _
                        TotalEmpresa & ")"
          'response.write sqlInsert 
          'response.end                         
            
            connSales.Execute sqlInsert
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    On Error GoTo 0
    
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerenciamento de Metas</title>
    <!-- Bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-color: #4361ee;
            --secondary-color: #3a0ca3;
            --success-color: #2ec4b6;
        }
        .card-header-custom {
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
        }
        .diretoria-card {
            border-left: 5px solid var(--primary-color);
        }
        .meta-input {
            font-weight: 600;
            text-align: right;
        }
        .loading-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 9999;
            justify-content: center;
            align-items: center;
        }
        .spinner {
            width: 50px;
            height: 50px;
            border: 5px solid #f3f3f3;
            border-top: 5px solid var(--primary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .campo-vazio {
            border-color: #ffc107 !important;
            background-color: #fff3cd !important;
        }
        .campo-valido {
            border-color: #198754 !important;
            background-color: #d1e7dd !important;
        }
    </style>
</head>
<body>
    <!-- Loading Overlay -->
    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="text-white mt-3">Processando, aguarde...</div>
    </div>

    <div class="container-fluid py-4">
        <!-- Cabeçalho -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-header card-header-custom">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <h4 class="mb-0"><i class="bi bi-bullseye me-2"></i>Sistema de Metas</h4>
                                <small>Gestão hierárquica de metas organizacionais</small>
                            </div>
                            <div>
                                <span class="badge bg-light text-dark">
                                    <i class="bi bi-calendar me-1"></i>
                                    <%
                                    Dim mesesArray
                                    mesesArray = Array("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
                                    
                                    If mesSelecionado <> "" Then
                                        Response.Write mesesArray(CInt(mesSelecionado)) & "/" & anoSelecionado
                                    Else
                                        Response.Write "Selecione um período"
                                    End If
                                    %>
                                </span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Mensagens -->
        <%
        If Session("msg_sucesso") <> "" Then
            Response.Write "<div class='alert alert-success alert-dismissible fade show mb-4'>" & _
                         "<i class='bi bi-check-circle me-2'></i>" & Session("msg_sucesso") & _
                         "<button type='button' class='btn-close' data-bs-dismiss='alert'></button></div>"
            Session("msg_sucesso") = ""
        End If
        
        If Session("msg_erro") <> "" Then
            Response.Write "<div class='alert alert-danger alert-dismissible fade show mb-4'>" & _
                         "<i class='bi bi-exclamation-triangle me-2'></i>" & Session("msg_erro") & _
                         "<button type='button' class='btn-close' data-bs-dismiss='alert'></button></div>"
            Session("msg_erro") = ""
        End If
        %>

        <!-- Filtros -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-body">
                        <h5 class="card-title mb-3">
                            <i class="bi bi-funnel me-2"></i>Filtrar Período
                        </h5>
                        <form method="GET" class="row g-3" id="filtroForm">
                            <div class="col-md-3">
                                <label class="form-label">Ano</label>
<select name="ano" class="form-select" id="anoSelect" required>
        <option value="">Selecione o ano</option>
        <%
        ' Gerar opções de ano a partir de 2026
        Dim anoInicial, anoFinal
        anoInicial = 2026
        anoFinal = Year(Now) + 1 ' Até 3 anos no futuro
        
        For i = anoInicial to anoFinal
            Response.Write "<option value='" & i & "'"
            If CStr(i) = CStr(anoSelecionado) Then
                Response.Write " selected"
            End If
            Response.Write ">" & i & "</option>"
        Next
        %>
    </select>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">Mês</label>
                                <select name="mes" class="form-select" id="mesSelect">
                                    <option value="">Selecione o mês</option>
                                    <%
                                    For i = 1 to 12
                                        Response.Write "<option value='" & i & "'"
                                        If CStr(i) = CStr(mesSelecionado) Then
                                            Response.Write " selected"
                                        End If
                                        Response.Write ">" & mesesArray(i) & "</option>"
                                    Next
                                    %>
                                </select>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">Diretoria</label>
                                <select name="diretoria" class="form-select">
                                    <option value="">Todas as Diretorias</option>
                                    <%
                                    'Dim sql
                                    sql = "SELECT * FROM Diretorias ORDER BY NomeDiretoria"
                                    Set rs = connOrg.Execute(sql)
                                    
                                    Do While Not rs.EOF
                                        Response.Write "<option value='" & rs("DiretoriaID") & "'"
                                        If CStr(rs("DiretoriaID")) = Request.QueryString("diretoria") Then
                                            Response.Write " selected"
                                        End If
                                        Response.Write ">" & rs("NomeDiretoria") & "</option>"
                                        rs.MoveNext
                                    Loop
                                    rs.Close
                                    %>
                                </select>
                            </div>
                            <div class="col-md-3 d-flex align-items-end">
                                <div class="d-grid gap-2 d-md-flex">
                                    <button type="submit" class="btn btn-primary">
                                        <i class="bi bi-search me-1"></i>Filtrar
                                    </button>
                                    <a href="meta_historico2.asp" class="btn btn-outline-secondary" target="_blank">
                                        <i class="bi bi-clock-history me-1"></i>Histórico
                                    </a>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Formulário de Metas -->
        <%
        If mesSelecionado <> "" Then
        %>
        <form method="POST" action="meta_gerenciamento.asp" id="formMetas">
            <!-- Campos hidden com valores atuais -->
            <input type="hidden" name="acao" value="salvar_metas">
            <input type="hidden" name="ano" id="anoHidden" value="<%=anoSelecionado%>">
            <input type="hidden" name="mes" id="mesHidden" value="<%=mesSelecionado%>">
            
            <!-- Instruções -->
            <div class="alert alert-info mb-4">
                <i class="bi bi-info-circle me-2"></i>
                <strong>Instruções:</strong> Preencha apenas os campos onde deseja definir metas. Campos vazios ou com valor zero serão ignorados.
            </div>
            
            <!-- Metas por Diretoria -->
            <%
            ' Buscar diretorias com suas gerencias
            Dim filtroDiretoriaSQL
            filtroDiretoriaSQL = ""
            If Request.QueryString("diretoria") <> "" Then
                filtroDiretoriaSQL = " WHERE d.DiretoriaID = " & Request.QueryString("diretoria")
            End If
            
            sql = "SELECT d.DiretoriaID, d.NomeDiretoria, " & _
                  "g.GerenciaID, g.NomeGerencia " & _
                  "FROM Diretorias d " & _
                  "INNER JOIN Gerencias g ON d.DiretoriaID = g.DiretoriaID " & _
                  filtroDiretoriaSQL & _
                  " ORDER BY d.NomeDiretoria, g.NomeGerencia"
            
            Set rsDiretorias = Server.CreateObject("ADODB.Recordset")
            rsDiretorias.Open sql, connOrg, 1, 3
            
            Dim currentDiretoriaID, diretoriaTotal
            currentDiretoriaID = 0
            diretoriaTotal = 0
            
            Do While Not rsDiretorias.EOF
                If currentDiretoriaID <> rsDiretorias("DiretoriaID") Then
                    If currentDiretoriaID > 0 Then
                        ' Exibir total da diretoria
                        %>
                        </div>
                        <div class="card-footer bg-light">
                            <div class="row align-items-center">
                                <div class="col-md-8">
                                    <strong>Total da Diretoria:</strong>
                                    <small class="text-muted ms-2">Soma automática (apenas valores > 0)</small>
                                </div>
                                <div class="col-md-4">
                                    <div class="input-group">
                                        <span class="input-group-text bg-white">
                                            <i class="bi bi-calculator"></i>
                                        </span>
                                        <input type="text" class="form-control bg-white meta-input" value="<%=FormatNumber(diretoriaTotal, 2, -1, -1, -1)%>" readonly>
                                        <span class="input-group-text bg-white">R$</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                        <%
                        diretoriaTotal = 0
                    End If
                    
                    currentDiretoriaID = rsDiretorias("DiretoriaID")
                    %>
                    <div class="diretoria-card card shadow mb-4">
                        <div class="card-header card-header-custom">
                            <h5 class="mb-0">
                                <i class="bi bi-building me-2"></i>
                                <%=rsDiretorias("NomeDiretoria")%>
                            </h5>
                        </div>
                        <div class="card-body">
                    <%
                End If
                
                ' Buscar meta atual
                Dim metaAtualGerencia
                metaAtualGerencia = 0
                
                sql = "SELECT ValorMeta FROM MetaGerencia WHERE GerenciaID = " & rsDiretorias("GerenciaID") & _
                      " AND Ano = " & anoSelecionado & " AND Mes = " & mesSelecionado
                
                Set rsMeta = Server.CreateObject("ADODB.Recordset")
                rsMeta.Open sql, connSales, 1, 3
                
                If Not rsMeta.EOF And Not IsNull(rsMeta("ValorMeta")) Then
                    metaAtualGerencia = CDbl(rsMeta("ValorMeta"))
                End If
                rsMeta.Close
                Set rsMeta = Nothing
                
                If metaAtualGerencia > 0 Then
                    diretoriaTotal = diretoriaTotal + metaAtualGerencia
                End If
                
                ' Determinar classe CSS baseada no valor
                Dim classeInput
                If metaAtualGerencia > 0 Then
                    classeInput = "campo-valido"
                ElseIf metaAtualGerencia = 0 Then
                    classeInput = "campo-vazio"
                Else
                    classeInput = ""
                End If
                %>
                <div class="row align-items-center mb-2">
                    <div class="col-md-5">
                        <label class="form-label mb-0">
                            <i class="bi bi-diagram-2 me-2 text-primary"></i>
                            <%=rsDiretorias("NomeGerencia")%>
                            <% If metaAtualGerencia > 0 Then %>
                            <span class="badge bg-success ms-2">✓ Definida</span>
                            <% End If %>
                        </label>
                    </div>
                    <div class="col-md-7">
                        <div class="input-group">
                            <span class="input-group-text">R$</span>
<input type="text" 
       name="gerencia<%=rsDiretorias("GerenciaID")%>"
       class="form-control meta-input <%=classeInput%>"
       value="<% 
                 If metaAtualGerencia > 0 Then 
                     Response.Write FormatNumber(metaAtualGerencia, 2, -1, -1, -1) 
                 Else 
                     Response.Write "" 
                 End If 
              %>"
       placeholder="0,00"
       onkeyup="formatarMoeda(this)"
       onblur="validarCampo(this)"
       data-gerenciaid="<%=rsDiretorias("GerenciaID")%>">
                            <button type="button" class="btn btn-outline-secondary" onclick="limparCampo(this)">
                                <i class="bi bi-x"></i>
                            </button>
                        </div>
                        <div class="form-text">
                            <% If metaAtualGerencia > 0 Then %>
                            <small class="text-success">Meta atual: R$ <%=FormatNumber(metaAtualGerencia, 2, -1, -1, -1)%></small>
                            <% Else %>
                            <small class="text-muted">Deixe vazio para manter sem meta</small>
                            <% End If %>
                        </div>
                    </div>
                </div>
                <%
                rsDiretorias.MoveNext
            Loop
            
            ' Fechar última diretoria
            If currentDiretoriaID > 0 Then
                %>
                        </div>
                        <div class="card-footer bg-light">
                            <div class="row align-items-center">
                                <div class="col-md-8">
                                    <strong>Total da Diretoria:</strong>
                                    <small class="text-muted ms-2">Soma automática (apenas valores > 0)</small>
                                </div>
                                <div class="col-md-4">
                                    <div class="input-group">
                                        <span class="input-group-text bg-white">
                                            <i class="bi bi-calculator"></i>
                                        </span>
                                        <input type="text" class="form-control bg-white meta-input" value="<%=FormatNumber(diretoriaTotal, 2, -1, -1, -1)%>" readonly>
                                        <span class="input-group-text bg-white">R$</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                <%
            End If
            
            rsDiretorias.Close
            Set rsDiretorias = Nothing
            
            If currentDiretoriaID = 0 Then
                Response.Write "<div class='alert alert-warning'>Nenhuma diretoria/gerencia encontrada.</div>"
            End If
            %>
            
            <!-- Botões -->
            <div class="row mt-4">
                <div class="col-12">
                    <div class="card shadow">
                        <div class="card-body">
                            <div class="d-flex justify-content-between align-items-center">
                                <div>
                                    <button type="button" class="btn btn-outline-warning" onclick="limparTodosCampos()">
                                        <i class="bi bi-eraser me-1"></i>Limpar Todos
                                    </button>
                                    <button type="button" class="btn btn-outline-info ms-2" onclick="validarFormulario()">
                                        <i class="bi bi-check2-circle me-1"></i>Validar Antes de Salvar
                                    </button>
                                </div>
                                <div>
                                    <button type="button" class="btn btn-secondary" onclick="window.location.href='meta_gerenciamento.asp'">
                                        <i class="bi bi-x-circle me-1"></i>Cancelar
                                    </button>
                                    <button type="submit" class="btn btn-success ms-2" id="btnSalvar">
                                        <i class="bi bi-check-circle me-1"></i>Salvar Metas Válidas
                                    </button>
                                </div>
                            </div>
                            <div class="mt-2">
                                <small class="text-muted">
                                    <i class="bi bi-lightbulb me-1"></i>
                                    Apenas campos preenchidos com valores maiores que zero serão salvos.
                                </small>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </form>
        <%
        Else
        %>
        <div class="row">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-body text-center py-5">
                        <i class="bi bi-bullseye text-muted" style="font-size: 4rem;"></i>
                        <h4 class="mt-3">Selecione um período</h4>
                        <p class="text-muted">Escolha um ano e mês para gerenciar as metas.</p>
                    </div>
                </div>
            </div>
        </div>
        <%
        End If
        %>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
// ===============================
// FORMATAR MOEDA
// ===============================
function formatarMoeda(input) {
    let valor = input.value.replace(/\D/g, '');
    if (valor === '') {
        input.value = '';
        validarCampo(input);
        return;
    }
    valor = (parseInt(valor) / 100).toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    input.value = valor;
    validarCampo(input);
}

// ===============================
// VALIDAR CAMPO INDIVIDUAL
// ===============================
function validarCampo(input) {
    let valor = input.value.replace(/\./g, '').replace(',', '.');

    input.classList.remove('campo-vazio', 'campo-valido', 'is-invalid');

    if (input.value.trim() === '') {
        input.classList.add('campo-vazio');
        return;
    }

    if (isNaN(parseFloat(valor))) {
        input.classList.add('is-invalid');
        return;
    }

    let num = parseFloat(valor);

    if (num > 0) {
        input.classList.add('campo-valido');
    } else {
        input.classList.add('campo-vazio');
    }
}

// ===============================
// LIMPAR CAMPO INDIVIDUAL
// ===============================
function limparCampo(botao) {
    const input = botao.closest('.input-group').querySelector('input');
    if (!input) return;
    input.value = '';
    validarCampo(input);
    input.focus();
}

// ===============================
// LIMPAR TODOS
// ===============================
function limparTodosCampos() {
    if (!confirm("Tem certeza que deseja limpar TODOS os campos?")) return;

    const inputs = document.querySelectorAll('input[name^="gerencia"]');
    inputs.forEach(input => {
        input.value = "";
        validarCampo(input);
    });
}

// ===============================
// VALIDAR FORMULÁRIO COMPLETO
// ===============================
function validarFormulario() {
    const inputs = document.querySelectorAll('input[name^="gerencia"]');

    let validos = 0;
    let vazios = 0;
    let invalidos = 0;

    inputs.forEach(input => {
        let val = input.value.replace(/\./g, '').replace(',', '.');

        if (input.value.trim() === '') {
            vazios++;
        }
        else if (isNaN(parseFloat(val))) {
            invalidos++;
        }
        else if (parseFloat(val) > 0) {
            validos++;
        }
        else {
            vazios++;
        }
    });

    alert(
        "Validação:\n\n" +
        "✓ Válidos (>0): " + validos + "\n" +
        "○ Vazios/Zero: " + vazios + "\n" +
        "✗ Inválidos: " + invalidos + "\n"
    );

    return validos > 0;
}

// ===============================
// MOSTRAR LOADING
// ===============================
function mostrarLoading() {
    const overlay = document.getElementById("loadingOverlay");
    const btn = document.getElementById("btnSalvar");

    if (overlay) overlay.style.display = "flex";
    if (btn) btn.disabled = true;
}

// ===============================
// PROCESSAR SUBMIT
// ===============================
function configurarEnvioFormulario() {
    const form = document.getElementById("formMetas");
    if (!form) return; // evita erro

    form.addEventListener("submit", function (e) {
        e.preventDefault();

        if (!validarFormulario()) {
            alert("Preencha ao menos um campo > 0.");
            return;
        }

        mostrarLoading();

        const inputs = form.querySelectorAll('input[name^="gerencia"]');
        let enviados = 0;

        inputs.forEach(input => {
            let valOriginal = input.value;
            let val = valOriginal.replace(/\./g, '').replace(',', '.');

            if (val === "" || isNaN(parseFloat(val)) || parseFloat(val) <= 0) {
                input.disabled = true; // não enviar
            } else {
                input.value = parseFloat(val).toFixed(2);
                enviados++;
            }
        });

        if (enviados === 0) {
            alert("Nenhum valor válido encontrado!");
            inputs.forEach(i => i.disabled = false);
            return;
        }

        setTimeout(() => form.submit(), 500);
    });
}

// ===============================
// SINCRONIZAR SELECTS COM HIDDEN
// ===============================
function configurarSelectAnoMes() {
    const anoSel = document.getElementById("anoSelect");
    const mesSel = document.getElementById("mesSelect");
    const anoHid = document.getElementById("anoHidden");
    const mesHid = document.getElementById("mesHidden");

    if (anoSel && anoHid) {
        anoHid.value = anoSel.value;
        anoSel.addEventListener("change", () => {
            anoHid.value = anoSel.value;
        });
    }

    if (mesSel && mesHid) {
        mesHid.value = mesSel.value;
        mesSel.addEventListener("change", () => {
            mesHid.value = mesSel.value;
        });
    }
}

// ===============================
// INICIALIZAR AO CARREGAR A PÁGINA
// ===============================
document.addEventListener("DOMContentLoaded", function () {
    configurarEnvioFormulario();
    configurarSelectAnoMes();

    const inputs = document.querySelectorAll('input[name^="gerencia"]');
    inputs.forEach(input => validarCampo(input));
});
</script>

</body>
</html>
<%
' Fechar conexões
If IsObject(connOrg) Then
    connOrg.Close
    Set connOrg = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>