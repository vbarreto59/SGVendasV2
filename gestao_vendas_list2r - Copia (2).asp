<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% If Len(StrConn) = 0 Then %>
    <!--#include file="conexao.asp"-->
<% End If %>

<% If Len(StrConnSales) = 0 Then %>
    <!--#include file="conSunSales.asp"-->
<%End If%>

<!--#include file="AtualizarVendas.asp"-->
<!--#include file="gestao_atu_localizacao.asp"-->
<%
    'Response.Write strConnSales
    'Response.end
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"    

%>

<%
'============================= LOG ============================================'
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    set objMail = server.createobject("CDONTS.NewMail")
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    objMail.Subject = "SGVendas-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    objMail.MailFormat = 0
    objMail.Body = "Página de Vendas (Gestão Vendas)"
    objMail.Send
    set objMail = Nothing
end if 
'============================= ATUALIZANDO O BANCO DE DADOS ==================='
%>


<% 
'Modificação para separar banco de dados em 08 08 2025'
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)
%>

<%
'============================= ATUALIZANDO O BANCO DE DADOS ============================================================================'
Response.Buffer = True
Response.Expires = -1
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute(sqlUpdate1)

sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId];"
connSales.Execute(sqlUpdate2)

sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute(sqlUpdateCorretor)

sql = "UPDATE Vendas SET Semestre = SWITCH(Trimestre IN (1, 2), 1, Trimestre IN (3, 4), 2) WHERE Trimestre IS NOT NULL;"
On Error Resume Next
connSales.Execute sql
If Err.Number <> 0 Then
    Response.Write "Ocorreu um erro ao atualizar o campo Semestre: " & Err.Description
End If
On Error GoTo 0
%>

<%
' Função para remover números e asteriscos de uma string
Function RemoverNumeros(texto)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "[0-9*-]"
    regex.Global = True
    RemoverNumerosEAsteriscos = regex.Replace(texto, "")
    RemoverNumeros = Trim(Replace(RemoverNumerosEAsteriscos, "  ", " "))
End Function

Dim mensagem
mensagem = Request.QueryString("mensagem")

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales
Set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT Vendas.*, Usuarios.Nome AS UsuarioNome FROM Vendas LEFT JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId WHERE (Vendas.Excluido <> -1 OR Vendas.Excluido IS NULL) ORDER BY Vendas.ID DESC;"
rs.CursorLocation = 3
rs.CursorType = 1
rs.LockType = 1
rs.Open sql, connSales

Response.Write "<script>"
Response.Write "var salesData = [];"
totalValorScript = 0
totalComissaoScript = 0

If Not rs.EOF Then
    Do While Not rs.EOF
        comissaoPercentual = FormatNumber(rs("ComissaoPercentual"), 2)
        valorComissaoGeral = FormatNumber(rs("ValorComissaoGeral"), 2)
        totalValorScript = totalValorScript + CDbl(rs("ValorUnidade"))
        totalComissaoScript = totalComissaoScript + CDbl(rs("ValorComissaoGeral"))
        vAno = Right(rs("AnoVenda"), 2)
        Response.Write "salesData.push({"
        Response.Write "id: '" & rs("ID") & "',"
        Response.Write "anoMes: '" & rs("AnoVenda") & "-" & Right("0"&rs("MesVenda"),2) & "',"
        Response.Write "empreendimento: '" & Replace(RemoverNumeros(rs("NomeEmpreendimento")), "'", "\'") & "',"
        Response.Write "unidade: '" & Replace(rs("Unidade"), "'", "\'") & "',"
        Response.Write "unidadeM2: '" & Replace(rs("UnidadeM2"), "'", "\'") & "',"
        Response.Write "diretoria: '" & Replace(rs("Diretoria"), "'", "\'") & "',"
        Response.Write "gerencia: '" & Replace(rs("Gerencia"), "'", "\'") & "',"
        Response.Write "trimestre: '" & vAno&Replace("T"&rs("Trimestre"), "'", "\'") & "',"
        Response.Write "corretor: '" & Replace(rs("Corretor"), "'", "\'") & "',"
        Response.Write "dataVenda: '" & FormatDateTime(rs("DataVenda"), 2) & "',"
        Response.Write "valorUnidade: '" & FormatNumber(rs("ValorUnidade"), 2) & "',"
        Response.Write "valorUnidadeRaw: " & CDbl(rs("ValorUnidade")) & ","
        Response.Write "comissaoPercentual: '" & comissaoPercentual & "',"
        Response.Write "valorComissaoGeral: '" & valorComissaoGeral & "',"
        Response.Write "valorComissaoGeralRaw: " & CDbl(rs("ValorComissaoGeral")) & ","
        Response.Write "dataRegistro: '" & FormatDateTime(rs("DataRegistro"),2) & "',"
        Response.Write "usuarioRegistro: '" & Replace(rs("Usuario"), "'", "\'") & "',"
        Response.Write "});"
        rs.MoveNext
    Loop
    rs.MoveFirst
End If
Response.Write "</script>"
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Vendas</title>
    <meta http-equiv="refresh" content="600">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background-color: #E4F0F6; color: #fff; padding: 20px; }
        .table { background-color: #fff; color: #000; }
        th { background-color: #800000; color: #fff; }
        .card-body { background-color: #D37676; }
        .btn-maroon { background-color: #800000; color: white; }
        .btn-maroon:hover { background-color: #a00; color: white; }
        .dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter,
        .dataTables_wrapper .dataTables_info, .dataTables_wrapper .dataTables_paginate { color: white !important; }
        .dataTables_filter input { background-color: white; }
        .badge-comissao { background-color: #17a2b8; color: white; padding: 0.3em 0.6em; font-size: 0.85em; }
        table.dataTable thead th { background-color: #800000 !important; color: #fff !important; }
        table.dataTable tbody td { color: #000; }
        tfoot th { background-color: #f8f9fa; color: #000; font-weight: bold; }
        .total-row { background-color: #e9ecef !important; }
        .mobile-search-container { display: none; margin-bottom: 15px; background-color: #fff; padding: 10px; border-radius: 5px; }
        #tabelaVendas { display: table; }
        .dataTables_wrapper { display: block; }
        #mobileCardsContainer { display: none; }
        @media (max-width: 767.98px) {
            body { padding: 10px; font-size: 0.9em; }
            .container-fluid { padding-top: 10px; padding-left: 5px; padding-right: 5px; }
            h2 { font-size: 1.3rem; margin-top: 1rem !important; margin-bottom: 1rem !important; }
            .mb-4 { margin-bottom: 0.8rem !important; }
            .btn { width: 100%; margin-bottom: 8px; font-size: 0.9em; }
            .btn-sm { font-size: 0.8em; padding: 0.2rem 0.5rem; }
            .mobile-search-container { display: block; }
            .desktop-search { display: none; }
            #tabelaVendas, .dataTables_wrapper { display: none !important; }
            #mobileCardsContainer { display: block !important; margin-top: 15px; }
            .sale-card { background-color: #F7F3F3; color: #333; border-radius: 8px; padding: 15px; margin-bottom: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
            .card-title { font-size: 1.1em; font-weight: bold; color: #800000; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
            .card-item { display: flex; justify-content: space-between; align-items: center; padding: 4px 0; border-bottom: 1px dotted #f0f0f0; }
            .card-item:last-child { border-bottom: none; }
            .card-label { font-weight: bold; color: #555; flex-basis: 40%; }
            .card-value { text-align: right; flex-basis: 60%; }
            .card-actions { margin-top: 10px; text-align: right; }
            .card-actions .btn { width: auto; margin-left: 5px; }
            .card-registro { font-size: 0.75em; color: #777; text-align: right; margin-top: 5px; }
            .card-comissao .badge-comissao { font-size: 0.9em; }
            .card-total { background-color: #e9ecef; color: #000; font-weight: bold; padding: 10px 15px; margin-top: 15px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; }
            .card-total-label { font-size: 1em; }
            .card-total-value { font-size: 1.1em; }
            .no-results { text-align: center; padding: 20px; color: #666; font-size: 1.1em; }
        }
    </style>
</head>

<body>
    <div class="container-fluid">
        <div class="card">
            <div class="card-body">
                <h2 class="card-title mt-4 mb-4">
                    <i class="fas fa-handshake"></i> Gestão de Vendas
                </h2>
            </div>
        </div>
        <br>
        <% If mensagem <> "" Then %>
            <div class="alert alert-success alert-dismissible fade show">
                <%= mensagem %>
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            </div>
        <% End If %>
        
        <div class="mb-4 d-grid gap-2 d-md-block">
            <button type="button" onclick="window.close();" class="btn btn-success">
                <i class="fas fa-times me-2"></i>Fechar
            </button>
            <a href="gestao_vendas_create2.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Nova Venda</a>
            <a href="gestao_vendas_gerenc_comissoes.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Comissões</a>

            <a href="gestao_vendas_list_excluidos.asp" class="btn btn-success" target="_blank"><i class="fas fa-plus"></i> Excluídos</a>
            <a href="inserirVendasTeste2.asp" class="btn btn-primary" target="_blank"><i class="fas fa-plus"></i> Inserir Testes</a>
            <a href="excluir_testes.asp" class="btn btn-warning" target="_blank"><i class="fas fa-plus"></i> Excluir Testes</a>

        </div>
        
        <div class="mobile-search-container">
            <div class="input-group">
                <input type="text" id="mobileSearchInput" class="form-control" placeholder="Pesquisar vendas...">
                <button class="btn btn-maroon" id="mobileSearchBtn">
                    <i class="fas fa-search"></i>
                </button>
            </div>
        </div>
        
        <div class="card">
            <div class="card-body">
                <div class="table-responsive">
                    <table id="tabelaVendas" class="table table-striped table-bordered" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Ano-Mês</th>
                                <th>Data Venda</th>
                                <th>Empreendimento</th>
                                <th>Unidade</th>
                                <th>Diretoria</th>
                                <th>Gerência</th>
                                <th>Corretor</th>
                                <th>Valor (R$)</th>
                                <th>Comissão Unidade</th>
                                <th>Registro</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            Dim totalValorHtml, totalComissaoHtml
                            totalValorHtml = 0
                            totalComissaoHtml = 0

                            If Not rs.EOF Then
                                rs.MoveFirst
                                Do While Not rs.EOF
                                    Dim sqlPagamentos, rsPagamentos
                                    Dim totalPagoDiretoria, totalPagoGerencia, totalPagoCorretor
                                    Dim dataPagamentoDiretoria, dataPagamentoGerencia, dataPagamentoCorretor
                                    Dim tooltipDiretoria, tooltipGerencia, tooltipCorretor
                                    Dim pagoDiretoria, pagoGerencia, pagoCorretor
                                    totalPagoDiretoria = 0
                                    totalPagoGerencia = 0
                                    totalPagoCorretor = 0
                                    dataPagamentoDiretoria = ""
                                    dataPagamentoGerencia = ""
                                    dataPagamentoCorretor = ""
                                    tooltipDiretoria = ""
                                    tooltipGerencia = ""
                                    tooltipCorretor = ""
                                    pagoDiretoria = False
                                    pagoGerencia = False
                                    pagoCorretor = False

                                    sqlPagamentos = "SELECT * FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & rs("ID") & " ORDER BY DataPagamento ASC;"
                                    Set rsPagamentos = connSales.Execute(sqlPagamentos)

                                    If Not rsPagamentos.EOF Then
                                        Do While Not rsPagamentos.EOF
                                            Dim detalhePagamento
                                            detalhePagamento = "Data: " & FormatDateTime(rsPagamentos("DataPagamento"), 2) & " | Valor: R$ " & FormatNumber(rsPagamentos("ValorPago"), 2) & " | Status: " & rsPagamentos("Status")
                                            Select Case LCase(rsPagamentos("TipoRecebedor"))
                                                Case "diretoria"
                                                    If tooltipDiretoria <> "" Then tooltipDiretoria = tooltipDiretoria & Chr(13)
                                                    tooltipDiretoria = tooltipDiretoria & detalhePagamento
                                                    totalPagoDiretoria = totalPagoDiretoria + CDbl(rsPagamentos("ValorPago"))
                                                    dataPagamentoDiretoria = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                                Case "gerencia"
                                                    If tooltipGerencia <> "" Then tooltipGerencia = tooltipGerencia & Chr(13)
                                                    tooltipGerencia = tooltipGerencia & detalhePagamento
                                                    totalPagoGerencia = totalPagoGerencia + CDbl(rsPagamentos("ValorPago"))
                                                    dataPagamentoGerencia = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                                Case "corretor"
                                                    If tooltipCorretor <> "" Then tooltipCorretor = tooltipCorretor & Chr(13)
                                                    tooltipCorretor = tooltipCorretor & detalhePagamento
                                                    totalPagoCorretor = totalPagoCorretor + CDbl(rsPagamentos("ValorPago"))
                                                    dataPagamentoCorretor = FormatDateTime(rsPagamentos("DataPagamento"), 2)
                                            End Select
                                            rsPagamentos.MoveNext
                                        Loop
                                    End If
                                    rsPagamentos.Close
                                    Set rsPagamentos = Nothing

                                    If rs("ValorDiretoria") > 0 And totalPagoDiretoria >= CDbl(rs("ValorDiretoria")) Then pagoDiretoria = True
                                    If rs("ValorDiretoria") = 0 Then pagoDiretoria = True
                                    If rs("ValorGerencia") > 0 And totalPagoGerencia >= CDbl(rs("ValorGerencia")) Then pagoGerencia = True
                                    If rs("ValorGerencia") = 0 Then pagoGerencia = True
                                    If rs("ValorCorretor") > 0 And totalPagoCorretor >= CDbl(rs("ValorCorretor")) Then pagoCorretor = True
                                    If rs("ValorCorretor") = 0 Then pagoCorretor = True

                                    Dim comissaoText
                                    comissaoText = FormatNumber(rs("ComissaoPercentual"), 2) & "%"
                                    If rs("ValorComissaoGeral") > 0 Then
                                        comissaoText = comissaoText & " (R$ " & FormatNumber(rs("ValorComissaoGeral"), 2) & ")"
                                    End If

                                    totalValorHtml = totalValorHtml + CDbl(rs("ValorUnidade"))
                                    totalComissaoHtml = totalComissaoHtml + CDbl(rs("ValorComissaoGeral"))
                                    vAno = Right(rs("AnoVenda"), 2)

                                    ' Verifica se a venda já possui comissão cadastrada
                                    Dim rsComissaoCheck, comissaoExiste
                                    Set rsComissaoCheck = Server.CreateObject("ADODB.Recordset")
                                    rsComissaoCheck.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Venda = " & CInt(rs("ID")), connSales
                                    comissaoExiste = Not rsComissaoCheck.EOF
                                    rsComissaoCheck.Close
                                    Set rsComissaoCheck = Nothing
                            %>
                            <tr>
                                <td><%= rs("ID") %></td>
                                <td><%= rs("AnoVenda") & "-" & Right("0"&rs("MesVenda"),2) %><br><%= vAno & "T" & rs("Trimestre") %></td>
                                <td><%= FormatDateTime(rs("DataVenda"), 2) %></td>
                                <td><b><%= rs("Empreend_ID") %>-<%= RemoverNumeros(rs("NomeEmpreendimento")) %><br><%= RemoverNumeros(rs("Localidade")) %></b></td>
                                <td><%= rs("Unidade") %></td>
                                
                                <!-- --------------------------------------------------------------------------- -->
<td>
                                    <small style="color: red;"><b><%= rs("ComissaoDiretoria") %>%</b>-<b><%= FormatNumber(rs("ValorDiretoria"), 2) %></b></small>
                                    <% If Not IsEmpty(dataPagamentoDiretoria) Then %>
                                        <br><b><small style="color: blue;"><b><%= dataPagamentoDiretoria %></b></small></b>
                                    <% End If %>
                                    <% If pagoDiretoria Then %>
                                        <b><span class="badge bg-success" title="<%= Server.HTMLEncode(tooltipDiretoria) %>">PAGO</span></b>
                                    <% End If %>
                                    <b><%= rs("Diretoria") %></b><b><br><b><%= rs("UserIdDiretoria") & "-" & rs("NomeDiretor") %></b></b>
                                </td>
                                <td>
                                    <small style="color: red;"><b><%= rs("ComissaoGerencia") %>%</b>-<b><%= FormatNumber(rs("ValorGerencia"), 2) %></b></small>
                                    
                                    <% If Not IsEmpty(dataPagamentoGerencia) Then %>
                                        <br><b><small style="color: blue;"><b><%= dataPagamentoGerencia %></b></small></b>
                                    <% End If %>
                                    <% If pagoGerencia Then %>
                                        <b><span class="badge bg-success" title="<%= Server.HTMLEncode(tooltipGerencia) %>">PAGO</span></b>
                                    <% End If %>
                                    <b><b><%= rs("Gerencia") %></b><br><b><%= rs("UserIdGerencia") & "-" & rs("NomeGerente") %></b></b>
                                </td>
                                <td>
                                    <small style="color: red;"><b><%= rs("ComissaoCorretor") %>%</b>-<b><%= FormatNumber(rs("ValorCorretor"), 2) %></b></small>
                                    
                                    <% If Not IsEmpty(dataPagamentoCorretor) Then %>
                                        <br><b><small style="color: blue;"><b><%= dataPagamentoCorretor %></b></small></b>
                                    <% End If %>
                                    <% If pagoCorretor Then %>
                                        <b><span class="badge bg-success" title="<%= Server.HTMLEncode(tooltipCorretor) %>">PAGO</span></b>
                                    <% End If %>
                                    <br>
                                   <b><b><%= rs("CorretorId") & "-" & rs("Corretor") %></b></b> 
                                </td>
                                <!-- ------------------------------------------------------------------------- -->

                                <td data-order="<%= rs("ValorUnidade") %>" style="text-align: right;"><%= FormatNumber(rs("ValorUnidade"), 2) %></td>
                                <td data-order="<%= rs("ValorComissaoGeral") %>" style="text-align: right;"><span class="badge badge-comissao"><%= comissaoText %></span></td>
                                <td>
                                    <small>
                                        <%= FormatDateTime(rs("DataRegistro"),2) %><br>
                                        por <%= rs("Usuario") %>
                                    </small>
                                </td>
                                <td>
                                    <a href="gestao_vendas_update2.asp?id=<%= rs("id") %>" class="btn btn-warning btn-sm"><i class="fas fa-edit"></i> Editar</a>
                                    <% If Not comissaoExiste Then %>
                                        <a href="gestao_vendas_inserir_comissao1.asp?id=<%= rs("id") %>" class="btn btn-primary btn-sm"><i class="fas fa-edit"></i> Inserir comissão</a>
                                    <% End If %>
                                    <% If UCase(Session("Usuario")) = "BARRETO" Then %>
                                        <a href="gestao_vendas_delete.asp?id=<%= rs("id") %>" class="btn btn-danger btn-sm" onclick="return confirm('Confirma exclusão desta venda?');"><i class="fas fa-trash"></i> Excluir</a>
                                    <% End If %>
                                </td>
                            </tr>
                            <%
                                    rs.MoveNext
                                Loop
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="total-row">
                                <th colspan="8" style="text-align: right;">Totais:</th>
                                <th id="totalValor2" style="text-align: right;"><%= FormatNumber(totalValorHtml, 2) %></th>
                                <th id="totalComissao2" style="text-align: right;"><%= FormatNumber(totalComissaoHtml, 2) %></th>
                                <th colspan="2"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
                <div id="mobileCardsContainer"></div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
    $(document).ready(function () {
        var table;
        var originalSalesData = [];

        function initDataTable() {
            if (!$.fn.DataTable.isDataTable('#tabelaVendas')) {
                table = $('#tabelaVendas').DataTable({
                    language: { url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json" },
                    pageLength: 100,
                    order: [[0, "desc"]],
                    columnDefs: [
                        { orderable: false, targets: [11] },
                        { type: "date-eu", targets: [7] },
                        { type: "num-fmt", targets: [8, 9] }
                    ],
                    footerCallback: function (row, data, start, end, display) {
                        var api = this.api();
                        var intVal = function (i) {
                            if (typeof i === 'string') {
                                return parseFloat(i.replace(/\./g, '').replace(',', '.')) || 0;
                            }
                            return typeof i === 'number' ? i : 0;
                        };
                        var totalValor = api.column(8, { page: 'current' }).data().reduce(function (a, b) {
                            return intVal(a) + intVal(b);
                        }, 0);
                        var totalComissao = 0;
                        api.column(9, { page: 'current' }).nodes().each(function (node) {
                            totalComissao += intVal($(node).attr('data-order'));
                        });
                        $('#totalValor2').html(totalValor.toLocaleString('pt-BR', { minimumFractionDigits: 2 }));
                        $('#totalComissao2').html(totalComissao.toLocaleString('pt-BR', { minimumFractionDigits: 2 }));
                    }
                });
                table.on('draw', function () {
                    var footerCallback = table.init().footerCallback;
                    if (footerCallback) footerCallback.call(table);
                });
            }
        }

        function destroyDataTable() {
            if ($.fn.DataTable.isDataTable('#tabelaVendas')) {
                $('#tabelaVendas').DataTable().destroy();
            }
        }

        function searchMobileCards(searchTerm) {
            if (!searchTerm || searchTerm.trim() === '') {
                renderMobileCards(originalSalesData);
                return;
            }
            searchTerm = searchTerm.toLowerCase();
            const filteredData = originalSalesData.filter(function(sale) {
                return (
                    (sale.id && sale.id.toString().toLowerCase().includes(searchTerm)) ||
                    (sale.empreendimento && sale.empreendimento.toLowerCase().includes(searchTerm)) ||
                    (sale.unidade && sale.unidade.toLowerCase().includes(searchTerm)) ||
                    (sale.diretoria && sale.diretoria.toLowerCase().includes(searchTerm)) ||
                    (sale.gerencia && sale.gerencia.toLowerCase().includes(searchTerm)) ||
                    (sale.corretor && sale.corretor.toLowerCase().includes(searchTerm)) ||
                    (sale.dataVenda && sale.dataVenda.toLowerCase().includes(searchTerm)) ||
                    (sale.valorUnidade && sale.valorUnidade.toLowerCase().includes(searchTerm)) ||
                    (sale.comissaoPercentual && sale.comissaoPercentual.toLowerCase().includes(searchTerm)) ||
                    (sale.trimestre && sale.trimestre.toLowerCase().includes(searchTerm))
                );
            });
            renderMobileCards(filteredData);
        }

        function renderMobileCards(data) {
            const container = $('#mobileCardsContainer');
            container.empty();

            let totalValorCards = 0;
            let totalComissaoCards = 0;

            if (data.length === 0) {
                container.append('<div class="no-results">Nenhuma venda encontrada com os critérios de pesquisa.</div>');
                container.append('<div class="card-total mt-3"><span class="card-total-label">Total Valor:</span><span class="card-total-value">R$ 0,00</span></div>');
                container.append('<div class="card-total mt-2"><span class="card-total-label">Total Comissão:</span><span class="card-total-value">R$ 0,00</span></div>');
                return;
            }

            data.forEach(function (sale) {
                totalValorCards += sale.valorUnidadeRaw;
                totalComissaoCards += sale.valorComissaoGeralRaw;

                let comissaoHtml = sale.comissaoPercentual + '%';
                if (parseFloat(sale.valorComissaoGeralRaw) > 0) {
                    comissaoHtml += ' (R$ ' + sale.valorComissaoGeral + ')';
                }

                const cardHtml = `
<div class="sale-card">
    <div class="card-title">Venda ID: ${sale.id} - ${sale.empreendimento}</div>
    <div class="card-item combined-fields">
        <div class="field-group">
            <span class="card-label">Unidade:</span> <span class="card-value">${sale.unidade}</span>
        </div>
        <div class="field-group">
            <span class="card-label">M2:</span> <span class="card-value">${sale.unidadeM2}</span>
        </div>
    </div>
    <div class="card-item combined-fields">
        <div class="field-group">
            <span class="card-label">Mês:</span> <span class="card-value">${sale.anoMes}</span>
        </div>
        <div class="field-group">
            <span class="card-label">Trimestre:</span> <span class="card-value">${sale.trimestre}</span>
        </div>
    </div>
    <div class="card-item"><span class="card-label">Diretoria:</span><span class="card-value">${sale.diretoria}</span></div>
    <div class="card-item"><span class="card-label">Gerência:</span><span class="card-value">${sale.gerencia}</span></div>
    <div class="card-item"><span class="card-label">Corretor:</span><span class="card-value">${sale.corretor}</span></div>
    <div class="card-item"><span class="card-label">Data Venda:</span><span class="card-value">${sale.dataVenda}</span></div>
    <div class="card-item"><span class="card-label">Valor (R$):</span><span class="card-value">${sale.valorUnidade}</span></div>
    <div class="card-item card-comissao"><span class="card-label">Comissão:</span><span class="card-value"><span class="badge badge-comissao">${comissaoHtml}</span></span></div>
    <div class="card-registro">${sale.dataRegistro}<br>por ${sale.usuarioRegistro || ''}</div>
    <div class="card-actions">
        <a href="gestao_vendas_update2.asp?id=${sale.id}" class="btn btn-warning btn-sm"><i class="fas fa-edit"></i></a>
        <a href="gestao_vendas_excluir.asp?id=${sale.id}" class="btn btn-danger btn-sm" onclick="return confirm('Confirma exclusão desta venda?');"><i class="fas fa-trash"></i></a>
    </div>
</div>
                `;
                container.append(cardHtml);
            });

            container.append(`
                <div class="card-total mt-3">
                    <span class="card-total-label">Total Valor:</span>
                    <span class="card-total-value">R$ ${totalValorCards.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                </div>
                <div class="card-total mt-2">
                    <span class="card-total-label">Total Comissão:</span>
                    <span class="card-total-value">R$ ${totalComissaoCards.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                </div>
            `);
        }

        function checkScreenSize() {
            if (window.matchMedia('(max-width: 767.98px)').matches) {
                destroyDataTable();
                $('#tabelaVendas, .dataTables_wrapper').hide();
                $('#mobileCardsContainer').show();
                originalSalesData = salesData;
                renderMobileCards(originalSalesData);
                $('#mobileSearchBtn').click(function() {
                    searchMobileCards($('#mobileSearchInput').val());
                });
                $('#mobileSearchInput').on('keyup', function(e) {
                    if (e.key === 'Enter') {
                        searchMobileCards($('#mobileSearchInput').val());
                    }
                });
            } else {
                $('#mobileCardsContainer').hide();
                $('#tabelaVendas, .dataTables_wrapper').show();
                initDataTable();
                if (table) {
                    table.columns.adjust().draw();
                }
            }
        }

        checkScreenSize();
        $(window).on('resize', checkScreenSize);
    });
    </script>
</body>
</html>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
connSales.Close
Set connSales = Nothing
' --------------- gestao_vendas_list2r.asp --------------- 21 08 2025 19:50'
%>