<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 05/12/2024                       -->
<!-- PÁGINA: vendas_ultimos_dois_anos.asp   -->
<!-- OBS: Envia email com vendas dos últimos 2 anos -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="gestao_header.inc"-->

<%
' Verifica se usuário está logado
if Session("Usuario") = "" then
   Session("Usuario")  = "System"
''   Response.redirect "gestao_login.asp"
end if 

' ===============================================
' CONFIGURAÇÕES INICIAIS
' ===============================================
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

Dim connSales
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================
Function GetNomeMes(numeroMes)
    Select Case numeroMes
        Case 1: GetNomeMes = "Jan"
        Case 2: GetNomeMes = "Fev"
        Case 3: GetNomeMes = "Mar"
        Case 4: GetNomeMes = "Abr"
        Case 5: GetNomeMes = "Mai"
        Case 6: GetNomeMes = "Jun"
        Case 7: GetNomeMes = "Jul"
        Case 8: GetNomeMes = "Ago"
        Case 9: GetNomeMes = "Set"
        Case 10: GetNomeMes = "Out"
        Case 11: GetNomeMes = "Nov"
        Case 12: GetNomeMes = "Dez"
        Case Else: GetNomeMes = "Mês " & numeroMes
    End Select
End Function

Function FormatNumberBr(valor)
    If IsNull(valor) Or Not IsNumeric(valor) Or valor = 0 Then
        FormatNumberBr = ""
    Else
        FormatNumberBr = FormatNumber(valor, 2, -1, -2, -2)
    End If
End Function

' ===============================================
' DADOS DOS ÚLTIMOS DOIS ANOS
' ===============================================
Dim anoAtual, anoAnterior
anoAtual = Year(Date())
anoAnterior = anoAtual - 1

' Arrays para armazenar os totais por mês
Dim totaisAtual(12), totaisAnterior(12)
Dim mes, i

' Inicializa arrays
For i = 1 To 12
    totaisAtual(i) = 0
    totaisAnterior(i) = 0
Next

' ===============================================
' BUSCA VENDAS DO ANO ATUAL
' ===============================================
Dim sqlVendas, rsVendas
sqlVendas = "SELECT MesVenda, SUM(ValorUnidade) AS TotalMes FROM Vendas " & _
            "WHERE Excluido = 0 AND AnoVenda = " & anoAtual & " " & _
            "GROUP BY MesVenda ORDER BY MesVenda"

Set rsVendas = connSales.Execute(sqlVendas)

If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        mes = CInt(rsVendas("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            If Not IsNull(rsVendas("TotalMes")) Then
                totaisAtual(mes) = CDbl(rsVendas("TotalMes"))
            End If
        End If
        rsVendas.MoveNext
    Loop
End If
rsVendas.Close

' ===============================================
' BUSCA VENDAS DO ANO ANTERIOR
' ===============================================
sqlVendas = "SELECT MesVenda, SUM(ValorUnidade) AS TotalMes FROM Vendas " & _
            "WHERE Excluido = 0 AND AnoVenda = " & anoAnterior & " " & _
            "GROUP BY MesVenda ORDER BY MesVenda"

Set rsVendas = connSales.Execute(sqlVendas)

If Not rsVendas.EOF Then
    Do While Not rsVendas.EOF
        mes = CInt(rsVendas("MesVenda"))
        If mes >= 1 And mes <= 12 Then
            If Not IsNull(rsVendas("TotalMes")) Then
                totaisAnterior(mes) = CDbl(rsVendas("TotalMes"))
            End If
        End If
        rsVendas.MoveNext
    Loop
End If
rsVendas.Close

' ===============================================
' CALCULA TOTAIS GERAIS
' ===============================================
Dim totalGeralAtual, totalGeralAnterior
totalGeralAtual = 0
totalGeralAnterior = 0

For i = 1 To 12
    totalGeralAtual = totalGeralAtual + totaisAtual(i)
    totalGeralAnterior = totalGeralAnterior + totaisAnterior(i)
Next

' ===============================================
' ENVIO DE EMAIL
' ===============================================
If (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") Then
    On Error Resume Next
    Set objMail = Server.CreateObject("CDONTS.NewMail")
    
    If Err.Number <> 0 Then
        Response.Write "<div class='alert alert-danger'>Erro ao criar objeto de email: " & Err.Description & "</div>"
        Err.Clear
    Else
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
        objMail.Subject = "RELATÓRIO VENDAS " & anoAnterior & "-" & anoAtual & " - " & Ucase(Session("Usuario")) & " - " & Date
        
        objMail.MailFormat = 0 ' 0 = Texto Simples
        
        ' ===============================================
        ' CONSTRUIR CORPO DO EMAIL
        ' ===============================================
        Dim corpoEmail
        
        corpoEmail = "========================================================" & vbCrLf
        corpoEmail = corpoEmail & "RELATÓRIO DE VENDAS - ÚLTIMOS DOIS ANOS" & vbCrLf
        corpoEmail = corpoEmail & "========================================================" & vbCrLf & vbCrLf
        corpoEmail = corpoEmail & "Período: " & anoAnterior & " e " & anoAtual & vbCrLf
        corpoEmail = corpoEmail & "Data de geração: " & Now() & vbCrLf
        corpoEmail = corpoEmail & "Usuário: " & Ucase(Session("Usuario")) & vbCrLf & vbCrLf
        
        corpoEmail = corpoEmail & "--------------------------------------------------------" & vbCrLf
        corpoEmail = corpoEmail & "VENDAS POR MÊS " & vbCrLf
        corpoEmail = corpoEmail & "--------------------------------------------------------" & vbCrLf & vbCrLf
        
        ' Cabeçalho da tabela
        corpoEmail = corpoEmail & String(15, " ") & "| " & String(7, " ") & anoAnterior & String(7, " ") & " | " & String(10, " ") & anoAtual & String(10, " ") & " |" & vbCrLf
        corpoEmail = corpoEmail & String(50, "-") & vbCrLf
        
        ' Linhas com os meses
        For i = 1 To 12
            Dim nomeMes, valorAnterior, valorAtual
            
            nomeMes = GetNomeMes(i)
            valorAnterior = FormatNumberBr(totaisAnterior(i))
            valorAtual = FormatNumberBr(totaisAtual(i))
            
            ' Ajusta espaçamento para alinhar
            Dim espacosNome
            espacosNome = 10 - Len(nomeMes)
            If espacosNome < 0 Then espacosNome = 0
            
            corpoEmail = corpoEmail & nomeMes & String(espacosNome, " ") & " | "
            
            ' Valor ano anterior
            If valorAnterior = "" Then
                corpoEmail = corpoEmail & String(24, " ") & " | "
            Else
                Dim espacosAnterior
                espacosAnterior = 24 - Len(valorAnterior)
                If espacosAnterior < 0 Then espacosAnterior = 0
                corpoEmail = corpoEmail & String(espacosAnterior, " ") & valorAnterior & " | "
            End If
            
            ' Valor ano atual
            If valorAtual = "" Then
                corpoEmail = corpoEmail & String(24, " ")
            Else
                Dim espacosAtual
                espacosAtual = 24 - Len(valorAtual)
                If espacosAtual < 0 Then espacosAtual = 0
                corpoEmail = corpoEmail & String(espacosAtual, " ") & valorAtual
            End If
            
            corpoEmail = corpoEmail & vbCrLf
        Next
        
        corpoEmail = corpoEmail & String(50, "-") & vbCrLf
        
        ' Totais gerais
        corpoEmail = corpoEmail & "TOTAL GERAL" & String(4, " ") & " | "
        
        Dim totalAnteriorFormatado, totalAtualFormatado
        totalAnteriorFormatado = FormatNumber(totalGeralAnterior, 2)
        totalAtualFormatado = FormatNumber(totalGeralAtual, 2)
        
        Dim espacosTotalAnterior, espacosTotalAtual
        espacosTotalAnterior = 24 - Len(totalAnteriorFormatado)
        If espacosTotalAnterior < 0 Then espacosTotalAnterior = 0
        
        espacosTotalAtual = 24 - Len(totalAtualFormatado)
        If espacosTotalAtual < 0 Then espacosTotalAtual = 0
        
        corpoEmail = corpoEmail & String(espacosTotalAnterior, " ") & totalAnteriorFormatado & " | "
        corpoEmail = corpoEmail & String(espacosTotalAtual, " ") & totalAtualFormatado & vbCrLf
        
        corpoEmail = corpoEmail & String(50, "=") & vbCrLf & vbCrLf
        
        ' Crescimento/Variação
        If totalGeralAnterior > 0 Then
            Dim crescimento
            crescimento = ((totalGeralAtual - totalGeralAnterior) / totalGeralAnterior) * 100
            
            corpoEmail = corpoEmail & "CRESCIMENTO: " & FormatNumber(crescimento, 2) & "%" & vbCrLf
            corpoEmail = corpoEmail & "Valor: R$ " & FormatNumber(totalGeralAtual - totalGeralAnterior, 2) & vbCrLf & vbCrLf
        End If
        
        ' Resumo por ano
        corpoEmail = corpoEmail & "RESUMO POR ANO:" & vbCrLf
        corpoEmail = corpoEmail & "----------------" & vbCrLf
        corpoEmail = corpoEmail & anoAnterior & ": R$ " & FormatNumber(totalGeralAnterior, 2) & vbCrLf
        corpoEmail = corpoEmail & anoAtual & ": R$ " & FormatNumber(totalGeralAtual, 2) & vbCrLf & vbCrLf
        
        corpoEmail = corpoEmail & "========================================================" & vbCrLf
        corpoEmail = corpoEmail & "IP: " & request.serverVariables("REMOTE_ADDR") & vbCrLf
        corpoEmail = corpoEmail & "========================================================"
        
        objMail.Body = corpoEmail
        objMail.Send
        Set objMail = Nothing
        
        If Err.Number <> 0 Then
            Response.Write "<div class='alert alert-danger'>Erro ao enviar email: " & Err.Description & "</div>"
        Else
            Response.Write "<div class='alert alert-success'>Email enviado com sucesso!</div>"
        End If
    End If
    On Error GoTo 0
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Relatório Vendas - Últimos 2 Anos</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
    body { background: #f5f7fb; padding: 20px; }
    .card { margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .table-responsive { margin-top: 20px; }
    .ano-header { background-color: #e9ecef; font-weight: bold; }
    .mes-row:hover { background-color: #f8f9fa; }
    .total-row { font-weight: bold; background-color: #e9ecef; }
    .crescimento-positivo { color: #28a745; font-weight: bold; }
    .crescimento-negativo { color: #dc3545; font-weight: bold; }
</style>
</head>
<body>

<div class="container">
    <div class="row">
        <div class="col-12">
            <div class="card">
                <div class="card-header bg-primary text-white">
                    <h4 class="mb-0">
                        <i class="fas fa-chart-line me-2"></i>
                        Relatório de Vendas - Últimos 2 Anos
                    </h4>
                </div>
                <div class="card-body">
                    <div class="row mb-4">
                        <div class="col-md-6">
                            <div class="card">
                                <div class="card-body text-center">
                                    <h5 class="text-muted mb-2"><%= anoAnterior %></h5>
                                    <h3 class="text-primary">R$ <%= FormatNumber(totalGeralAnterior, 2) %></h3>
                                    <p class="text-muted mb-0">Total do ano anterior</p>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="card">
                                <div class="card-body text-center">
                                    <h5 class="text-muted mb-2"><%= anoAtual %></h5>
                                    <h3 class="text-success">R$ <%= FormatNumber(totalGeralAtual, 2) %></h3>
                                    <p class="text-muted mb-0">Total do ano atual</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="table-responsive">
                        <table class="table table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th width="20%">Mês</th>
                                    <th width="40%" class="text-center"><%= anoAnterior %></th>
                                    <th width="40%" class="text-center"><%= anoAtual %></th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                For i = 1 To 12
                                    Dim nomeMesDisplay, valorAnteriorDisplay, valorAtualDisplay
                                    Dim crescimentoMes, crescimentoSinal, crescimentoClasse
                                    
                                    nomeMesDisplay = GetNomeMes(i)
                                    valorAnteriorDisplay = totaisAnterior(i)
                                    valorAtualDisplay = totaisAtual(i)
                                    
                                    ' Calcula crescimento para destacar
                                    If valorAnteriorDisplay > 0 And valorAtualDisplay > 0 Then
                                        crescimentoMes = ((valorAtualDisplay - valorAnteriorDisplay) / valorAnteriorDisplay) * 100
                                        
                                        ' Determina sinal e classe sem IIF
                                        If crescimentoMes > 0 Then
                                            crescimentoSinal = "▲ +"
                                            crescimentoClasse = "crescimento-positivo"
                                        ElseIf crescimentoMes < 0 Then
                                            crescimentoSinal = "▼ "
                                            crescimentoClasse = "crescimento-negativo"
                                        Else
                                            crescimentoSinal = "● "
                                            crescimentoClasse = "text-muted"
                                        End If
                                    End If
                                %>
                                <tr class="mes-row">
                                    <td><strong><%= nomeMesDisplay %></strong></td>
                                    <td class="text-end">
                                        <% If valorAnteriorDisplay > 0 Then %>
                                            R$ <%= FormatNumber(valorAnteriorDisplay, 2) %>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </td>
                                    <td class="text-end">
                                        <% If valorAtualDisplay > 0 Then %>
                                            R$ <%= FormatNumber(valorAtualDisplay, 2) %>
                                            <% If valorAnteriorDisplay > 0 Then %>
                                                <br>
                                                <small class="<%= crescimentoClasse %>">
                                                    <%= crescimentoSinal %><%= FormatNumber(Abs(crescimentoMes), 1) %>% 
                                                </small>
                                            <% End If %>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </td>
                                </tr>
                                <% Next %>
                                
                                <tr class="total-row">
                                    <td><strong>TOTAL GERAL</strong></td>
                                    <td class="text-end">
                                        <strong>R$ <%= FormatNumber(totalGeralAnterior, 2) %></strong>
                                    </td>
                                    <td class="text-end">
                                        <strong>R$ <%= FormatNumber(totalGeralAtual, 2) %></strong>
                                        <% If totalGeralAnterior > 0 Then %>
                                            <br>
                                            <%
                                            Dim crescimentoTotal, crescimentoTotalSinal, crescimentoTotalClasse
                                            crescimentoTotal = ((totalGeralAtual - totalGeralAnterior) / totalGeralAnterior) * 100
                                            
                                            ' Determina sinal e classe sem IIF
                                            If crescimentoTotal > 0 Then
                                                crescimentoTotalSinal = "▲ +"
                                                crescimentoTotalClasse = "crescimento-positivo"
                                            ElseIf crescimentoTotal < 0 Then
                                                crescimentoTotalSinal = "▼ "
                                                crescimentoTotalClasse = "crescimento-negativo"
                                            Else
                                                crescimentoTotalSinal = "● "
                                                crescimentoTotalClasse = "text-muted"
                                            End If
                                            %>
                                            <small class="<%= crescimentoTotalClasse %>">
                                                <%= crescimentoTotalSinal %><%= FormatNumber(Abs(crescimentoTotal), 1) %>% 
                                                (R$ <%= FormatNumber(totalGeralAtual - totalGeralAnterior, 2) %>)
                                            </small>
                                        <% End If %>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                    
                    <div class="row mt-4">
                        <div class="col-12">
                            <div class="card">
                                <div class="card-header bg-info text-white">
                                    <h5 class="mb-0">Resumo Comparativo</h5>
                                </div>
                                <div class="card-body">
                                    <div class="row text-center">
                                        <div class="col-md-4">
                                            <h6 class="text-muted">Média Mensal <%= anoAnterior %></h6>
                                            <h4 class="text-primary">
                                                R$ <%= FormatNumber(totalGeralAnterior / 12, 2) %>
                                            </h4>
                                        </div>
                                        <div class="col-md-4">
                                            <h6 class="text-muted">Média Mensal <%= anoAtual %></h6>
                                            <h4 class="text-success">
                                                R$ <%= FormatNumber(totalGeralAtual / 12, 2) %>
                                            </h4>
                                        </div>
                                        <div class="col-md-4">
                                            <h6 class="text-muted">Crescimento</h6>
                                            <%
                                            Dim diferencaTotal, crescimentoFormatado, crescimentoClasseFormatado
                                            diferencaTotal = totalGeralAtual - totalGeralAnterior
                                            
                                            If diferencaTotal > 0 Then
                                                crescimentoClasseFormatado = "crescimento-positivo"
                                            ElseIf diferencaTotal < 0 Then
                                                crescimentoClasseFormatado = "crescimento-negativo"
                                            Else
                                                crescimentoClasseFormatado = "text-muted"
                                            End If
                                            
                                            If totalGeralAnterior > 0 Then
                                                crescimentoFormatado = FormatNumber((diferencaTotal / totalGeralAnterior) * 100, 1)
                                            Else
                                                crescimentoFormatado = "0,0"
                                            End If
                                            %>
                                            <h4 class="<%= crescimentoClasseFormatado %>">
                                                <%= crescimentoFormatado %>%
                                            </h4>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="card-footer">
                    <div class="d-flex justify-content-between">
                        <div>
                            <i class="fas fa-calendar me-1"></i>
                            Gerado em: <%= Now() %>
                        </div>
                        <div>
                            <i class="fas fa-user me-1"></i>
                            Usuário: <%= Session("Usuario") %>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/js/all.min.js"></script>

</body>
</html>

<%
' Fecha conexão
If Not connSales Is Nothing Then
    If connSales.State = 1 Then connSales.Close
    Set connSales = Nothing
End If
%>