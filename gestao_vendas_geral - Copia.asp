<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
' Configuração para evitar cache
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

' Parâmetros do filtro
Dim ano, mes, trimestre, semestre, whereClause, isFiltered
isFiltered = False

ano = Request.Form("ano")
mes = Request.Form("mes")
trimestre = Request.Form("trimestre")
semestre = Request.Form("semestre")

' Construir cláusula WHERE
whereClause = " WHERE Vendas.Excluido = 0"

If ano <> "" And IsNumeric(ano) Then
    whereClause = whereClause & " AND Vendas.AnoVenda = " & ano
    isFiltered = True
End If

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

' Abrir conexão
conn.Open StrConnSales
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Vendas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" xintegrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">
    <style>
        body { background-color: #f8f9fa; }
        .card { margin-bottom: 20px; }
        .table th, .table td { vertical-align: middle; }
        .total { font-weight: bold; background-color: #e8f4ff; }
        .text-center { text-align: center !important; }
        .text-left { text-align: left !important; }
        .text-right { text-align: right !important; }
    </style>
</head>
<body>
    <div class="container mt-4">
        <h1 class="mb-4">SGVendas - Relatório de Vendas</h1>
        
        <div class="card">
            <div class="card-body">
                <form method="post" class="row g-3">
                    
                <div class="col-md-3">
                    <label for="ano" class="form-label">Ano:</label>
                    <select class="form-select" id="ano" name="ano">
                        <option value="2025">2025</option>
                        <option value="2026">2026</option>
                    </select>
                </div>


                    <div class="col-md-3">
                        <label for="mes" class="form-label">Mês:</label>
                        <select class="form-select" id="mes" name="mes">
                            <option value="">Todos</option>
                            <% For i = 1 To 12 %>
                            <option value="<%=i%>" <%If CStr(i) = mes Then Response.Write "selected"%>><%=MonthName(i, True)%></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="trimestre" class="form-label">Trimestre:</label>
                        <select class="form-select" id="trimestre" name="trimestre">
                            <option value="">Todos</option>
                            <option value="1" <%If trimestre = "1" Then Response.Write "selected"%>>1º</option>
                            <option value="2" <%If trimestre = "2" Then Response.Write "selected"%>>2º</option>
                            <option value="3" <%If trimestre = "3" Then Response.Write "selected"%>>3º</option>
                            <option value="4" <%If trimestre = "4" Then Response.Write "selected"%>>4º</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="semestre" class="form-label">Semestre:</label>
                        <select class="form-select" id="semestre" name="semestre">
                            <option value="">Todos</option>
                            <option value="1" <%If semestre = "1" Then Response.Write "selected"%>>1º</option>
                            <option value="2" <%If semestre = "2" Then Response.Write "selected"%>>2º</option>
                            <option value="3" <%If semestre = "3" Then Response.Write "selected"%>>3º</option>
                            <option value="4" <%If semestre = "4" Then Response.Write "selected"%>>4º</option>
                        </select>
                    </div>
                    <div class="col-12 d-flex gap-2">
                        <button type="submit" class="btn btn-primary">Filtrar</button>
                        <a href="gestao_vendas_geral.asp" class="btn btn-secondary">Limpar Filtro</a>
                    </div>
                </form>
            </div>
        </div>

        <% if isFiltered then %>
        <div class="card mb-4">
            <div class="card-body">
                <h5 class="card-title">Filtros Ativos</h5>
                <ul class="list-group list-group-flush">
                    <% if ano <> "" then %>
                    <li class="list-group-item"><strong>Ano:</strong> <%=ano%></li>
                    <% end if %>
                    <% if mes <> "" then %>
                    <li class="list-group-item"><strong>Mês:</strong> <%=MonthName(mes, False)%></li>
                    <% end if %>
                    <% if trimestre <> "" then %>
                    <li class="list-group-item"><strong>Trimestre:</strong>
                        <%
                        Select Case trimestre
                            Case 1
                                Response.Write "1º"
                            Case 2
                                Response.Write "2º"
                            Case 3
                                Response.Write "3º"
                            Case 4
                                Response.Write "4º"
                        End Select
                        %>
                    </li>
                    <% end if %>
                    <% if semestre <> "" then %>
                    <li class="list-group-item"><strong>Semestre:</strong>
                        <%
                        Select Case semestre
                            Case 1
                                Response.Write "1º"
                            Case 2
                                Response.Write "2º"
                        End Select
                        %>
                    </li>
                    <% end if %>
                </ul>
            </div>
        </div>
        <% end if %>
        <div class="resultados">
            <%
            If whereClause <> " WHERE Vendas.Excluido = 0" Then
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
                    totalVGV = rs("TotalVGV")
                    If IsNull(totalVGV) Then totalVGV = 0
                End If
                rs.Close
                
                ' Exibir totais gerais
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Totais Gerais</h2>
                        <p class="card-text"><strong>Total de Unidades Vendidas:</strong> <%=FormatNumber(totalUnidades, 0)%></p>
                        <p class="card-text"><strong>Total VGV:</strong> R$ <%=FormatNumber(totalVGV, 2)%></p>
                    </div>
                </div>
                
                <%
                ' 3. Total vendas por Diretoria
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Vendas por Diretoria</h2>
                        <%
                        sql = "SELECT Vendas.Diretoria, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                              "FROM Vendas" & whereClause & " GROUP BY Diretoria ORDER BY SUM(Vendas.ValorUnidade) DESC"
                              'Response.Write sql
                              'Response.end 
                        rs.Open sql, conn
                        
                        If Not rs.EOF Then
                            Dim totalUnidadesDiretoria, totalVGVDiretoria
                            totalUnidadesDiretoria = 0
                            totalVGVDiretoria = 0
                            %>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered">
                                    <thead>
                                        <tr>
                                            <th>Diretoria</th>
                                            <th class="text-center">Unidades</th>
                                            <th class="text-right">VGV</th>
                                            <th class="text-right">% do Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            If totalVGV <> 0 Then
                                                percentual = (rs("VGV") / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            
                                            totalUnidadesDiretoria = totalUnidadesDiretoria + rs("Unidades")
                                            totalVGVDiretoria = totalVGVDiretoria + rs("VGV")
                                            %>
                                            <tr>
                                                <td><%=rs("Diretoria")%></td>
                                                <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                <td class="text-right">R$ <%=FormatNumber(rs("VGV"), 2)%></td>
                                                <td class="text-right"><%=FormatNumber(percentual, 2)%>%</td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                        <tr class="total">
                                            <td><strong>TOTAL</strong></td>
                                            <td class="text-center"><strong><%=FormatNumber(totalUnidadesDiretoria, 0)%></strong></td>
                                            <td class="text-right"><strong>R$ <%=FormatNumber(totalVGVDiretoria, 2)%></strong></td>
                                            <td class="text-right"><strong>100%</strong></td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        Else
                            %>
                            <p class="card-text">Nenhuma venda encontrada para os filtros selecionados.</p>
                            <%
                        End If
                        rs.Close
                        %>
                    </div>
                </div>

                <%
                ' 4. Total vendas por Gerência (nova tabela)
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Vendas por Gerência</h2>
                        <%
                        sql = "SELECT Vendas.Gerencia, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                              "FROM Vendas" & whereClause & " GROUP BY Gerencia ORDER BY SUM(Vendas.ValorUnidade) DESC"
                        rs.Open sql, conn
                        
                        If Not rs.EOF Then
                            Dim totalUnidadesGerencia, totalVGVGerencia
                            totalUnidadesGerencia = 0
                            totalVGVGerencia = 0
                            %>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered">
                                    <thead>
                                        <tr>
                                            <th>Gerência</th>
                                            <th class="text-center">Unidades</th>
                                            <th class="text-right">VGV</th>
                                            <th class="text-right">% do Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            If totalVGV <> 0 Then
                                                percentual = (rs("VGV") / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            
                                            totalUnidadesGerencia = totalUnidadesGerencia + rs("Unidades")
                                            totalVGVGerencia = totalVGVGerencia + rs("VGV")
                                            %>
                                            <tr>
                                                <td><%=rs("Gerencia")%></td>
                                                <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                <td class="text-right">R$ <%=FormatNumber(rs("VGV"), 2)%></td>
                                                <td class="text-right"><%=FormatNumber(percentual, 2)%>%</td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                        <tr class="total">
                                            <td><strong>TOTAL</strong></td>
                                            <td class="text-center"><strong><%=FormatNumber(totalUnidadesGerencia, 0)%></strong></td>
                                            <td class="text-right"><strong>R$ <%=FormatNumber(totalVGVGerencia, 2)%></strong></td>
                                            <td class="text-right"><strong>100%</strong></td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        Else
                            %>
                            <p class="card-text">Nenhuma venda encontrada para os filtros selecionados.</p>
                            <%
                        End If
                        rs.Close
                        %>
                    </div>
                </div>

                <%
                ' 5. Total vendas por Localidade
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Vendas por Localidade</h2>
                        <%
                        sql = "SELECT Vendas.Localidade, COUNT(*) as Unidades, SUM(Vendas.ValorUnidade) as VGV " & _
                              "FROM Vendas" & whereClause & " GROUP BY Localidade ORDER BY SUM(Vendas.ValorUnidade) DESC"
                        rs.Open sql, conn
                        
                        If Not rs.EOF Then
                            Dim totalUnidadesLocalidade, totalVGVLocalidade
                            totalUnidadesLocalidade = 0
                            totalVGVLocalidade = 0
                            %>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered">
                                    <thead>
                                        <tr>
                                            <th>Localidade</th>
                                            <th class="text-center">Unidades</th>
                                            <th class="text-right">VGV</th>
                                            <th class="text-right">% do Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            If totalVGV <> 0 Then
                                                percentual = (rs("VGV") / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            
                                            totalUnidadesLocalidade = totalUnidadesLocalidade + rs("Unidades")
                                            totalVGVLocalidade = totalVGVLocalidade + rs("VGV")
                                            %>
                                            <tr>
                                                <td><%=rs("Localidade")%></td>
                                                <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                <td class="text-right">R$ <%=FormatNumber(rs("VGV"), 2)%></td>
                                                <td class="text-right"><%=FormatNumber(percentual, 2)%>%</td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                        <tr class="total">
                                            <td><strong>TOTAL</strong></td>
                                            <td class="text-center"><strong><%=FormatNumber(totalUnidadesLocalidade, 0)%></strong></td>
                                            <td class="text-right"><strong>R$ <%=FormatNumber(totalVGVLocalidade, 2)%></strong></td>
                                            <td class="text-right"><strong>100%</strong></td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        Else
                            %>
                            <p class="card-text">Nenhuma venda encontrada para os filtros selecionados.</p>
                            <%
                        End If
                        rs.Close
                        %>
                    </div>
                </div>
                
                <%
                ' 6. Total vendas por Empresa
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Vendas por Empresa</h2>
                        <%
                        sql = "SELECT NomeEmpresa, COUNT(*) as Unidades, SUM(ValorUnidade) as VGV " & _
                              "FROM Vendas" & whereClause & " GROUP BY Nomeempresa ORDER BY SUM(ValorUnidade) DESC"
                        rs.Open sql, conn
                        
                        If Not rs.EOF Then
                            Dim totalUnidadesEmpresa, totalVGVEmpresa
                            totalUnidadesEmpresa = 0
                            totalVGVEmpresa = 0
                            %>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered">
                                    <thead>
                                        <tr>
                                            <th>Empresa</th>
                                            <th class="text-center">Unidades</th>
                                            <th class="text-right">VGV</th>
                                            <th class="text-right">% do Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            If totalVGV <> 0 Then
                                                percentual = (rs("VGV") / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            
                                            totalUnidadesEmpresa = totalUnidadesEmpresa + rs("Unidades")
                                            totalVGVEmpresa = totalVGVEmpresa + rs("VGV")
                                            %>
                                            <tr>
                                                <td><%=rs("NomeEmpresa")%></td>
                                                <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                <td class="text-right">R$ <%=FormatNumber(rs("VGV"), 2)%></td>
                                                <td class="text-right"><%=FormatNumber(percentual, 2)%>%</td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                        <tr class="total">
                                            <td><strong>TOTAL</strong></td>
                                            <td class="text-center"><strong><%=FormatNumber(totalUnidadesEmpresa, 0)%></strong></td>
                                            <td class="text-right"><strong>R$ <%=FormatNumber(totalVGVEmpresa, 2)%></strong></td>
                                            <td class="text-right"><strong>100%</strong></td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        End If
                        rs.Close
                        %>
                    </div>
                </div>
                
                <%
                ' 7. Total vendas por Empreendimento
                %>
                <div class="card mb-4">
                    <div class="card-body">
                        <h2 class="card-title">Vendas por Empreendimento</h2>
                        <%
                        sql = "SELECT NomeEmpreendimento, COUNT(*) as Unidades, SUM(ValorUnidade) as VGV " & _
                              "FROM Vendas" & whereClause & " GROUP BY NomeEmpreendimento ORDER BY SUM(Vendas.ValorUnidade) DESC"
                        rs.Open sql, conn
                        
                        If Not rs.EOF Then
                            Dim totalUnidadesEmpreendimento, totalVGVEmpreendimento
                            totalUnidadesEmpreendimento = 0
                            totalVGVEmpreendimento = 0
                            %>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered">
                                    <thead>
                                        <tr>
                                            <th>Empreendimento</th>
                                            <th class="text-center">Unidades</th>
                                            <th class="text-right">VGV</th>
                                            <th class="text-right">% do Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            If totalVGV <> 0 Then
                                                percentual = (rs("VGV") / totalVGV) * 100
                                            Else
                                                percentual = 0
                                            End If
                                            
                                            totalUnidadesEmpreendimento = totalUnidadesEmpreendimento + rs("Unidades")
                                            totalVGVEmpreendimento = totalVGVEmpreendimento + rs("VGV")
                                            %>
                                            <tr>
                                                <td><%=rs("NomeEmpreendimento")%></td>
                                                <td class="text-center"><%=FormatNumber(rs("Unidades"), 0)%></td>
                                                <td class="text-right">R$ <%=FormatNumber(rs("VGV"), 2)%></td>
                                                <td class="text-right"><%=FormatNumber(percentual, 2)%>%</td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        %>
                                        <tr class="total">
                                            <td><strong>TOTAL</strong></td>
                                            <td class="text-center"><strong><%=FormatNumber(totalUnidadesEmpreendimento, 0)%></strong></td>
                                            <td class="text-right"><strong>R$ <%=FormatNumber(totalVGVEmpreendimento, 2)%></strong></td>
                                            <td class="text-right"><strong>100%</strong></td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        End If
                        rs.Close
                        %>
                    </div>
                </div>
                <%
            Else
                %>
                <div class="alert array alert-info" role="alert">
                    Selecione os filtros desejados e clique em 'Filtrar' para gerar o relatório.
                </div>
                <%
            End If
            %>
        </div>
    </div>
             <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" xintegrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
</body>
</html>

<%
' Fechar conexões
If rs.State = 1 Then rs.Close
If conn.State = 1 Then conn.Close
Set rs = Nothing
Set conn = Nothing
%>