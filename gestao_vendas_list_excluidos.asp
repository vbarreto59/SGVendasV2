<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
'Option Explicit 

' =======================================================
' === DEFINIÇÃO DE CONSTANTES ADO ===
' =======================================================
Const adOpenKeyset = 1  
Const adLockReadOnly = 1 

' Function to safely get string field (returns "" if NULL)
Function GetStringField(rs, fieldName)
    If IsNull(rs(fieldName)) Then
        GetStringField = ""
    Else
        GetStringField = Trim(rs(fieldName))
    End If
End Function

' Function to remove numbers and asterisks from a string
Function RemoverNumeros(texto)
    If IsNull(texto) Or texto = "" Then
        RemoverNumeros = ""
        Exit Function
    End If
    
    Dim regex
    Set regex = New RegExp
    
    regex.Pattern = "[\-0-9*]" 
    regex.Global = True
    
    RemoverNumeros = regex.Replace(texto, "")
    RemoverNumeros = Trim(Replace(RemoverNumeros, "  ", " "))
    
    Set regex = Nothing
End Function
%>

<%
' Variáveis
Dim mensagem, conn, rs, sql

mensagem = Request.QueryString("mensagem")

' Create and open ADODB connection
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next 
conn.Open StrConnSales 
If Err.Number <> 0 Then
    Response.Write "<h2>ERRO DE CONEXÃO COM O BANCO DE DADOS!</h2>"
    Response.Write "<p>Detalhes: " & Err.Description & "</p>"
    Response.End
End If
On Error Goto 0 

' Create and open ADODB Recordset
Set rs = Server.CreateObject("ADODB.Recordset")
        
sql = "SELECT Vendas.* FROM Vendas WHERE Vendas.Excluido = -1 ORDER BY Vendas.ID DESC;"

' Abertura do Recordset
On Error Resume Next
rs.Open sql, conn, adOpenKeyset, adLockReadOnly

If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>Erro ao acessar os dados: " & Err.Description & "</div>"
End If
On Error Goto 0
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vendas Excluídas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <style>
        body {
            background-color: #f8f9fa;
            padding: 20px;
        }
        .table {
            background-color: #fff;
        }
        th {
            background-color: #800000;
            color: #fff;
        }
        .btn-maroon {
            background-color: #800000;
            color: white;
        }
        .btn-maroon:hover {
            background-color: #a00;
            color: white;
        }
        .badge-comissao {
            background-color: #17a2b8;
            color: white;
            padding: 0.3em 0.6em;    
            font-size: 0.85em;    
        }
        table.dataTable thead th {
            background-color: #800000 !important;
            color: #fff !important;
        }
        .total-row {
            background-color: #e9ecef !important;
            font-weight: bold;
        }
tfoot th {
    background-color: #800000 !important;
    color: #fff !important;
}        
    </style>
<style>
    body {
        /* Define a escala de 0.8 (80%) */
        transform: scale(0.8); 
        
        /* Define o ponto de origem para o canto superior esquerdo */
        transform-origin: 0 0; 
        
        /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
        /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
        width: calc(100% / 0.8); 
    }
</style>    
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4"><i class="fas fa-trash"></i> Vendas Excluídas</h2>
        
        <% If mensagem <> "" Then %>
            <div class="alert alert-success alert-dismissible fade show">
                <%= mensagem %>
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            </div>
        <% End If %>
        
        <div class="mb-4">
            <button type="button" onclick="window.close();" class="btn btn-success">
                <i class="fas fa-times me-2"></i>Fechar
            </button>
        </div>
        
        <div class="card">
            <div class="card-body">
                <%
                If rs.State = 1 And Not rs.EOF Then
                %>
                <div class="table-responsive">
                    <table id="tabelaVendas" class="table table-striped table-bordered" style="width:100%">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Ano-Mês</th>
                                <th>Empreendimento</th>
                                <th>Unidade</th>
                                <th>Diretoria</th>
                                <th>Gerência</th>
                                <th>Corretor</th>
                                <th>Data Venda</th>
                                <th>Valor (R$)</th>
                                <th>Comissão</th>
                                <th>Registro</th>
                                <th>Exclusão</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%    
                            Dim totalValorHtml, totalComissaoHtml, comissaoText, valorAcumulado
                            Dim valorCorretor, valorDiretoria, valorGerencia, valorUnidadeHtml, valorComissaoHtml
                            Dim vAno
                            
                            totalValorHtml = 0
                            totalComissaoHtml = 0
                            
                            Do While Not rs.EOF
                                ' === TRATAMENTO DE NULLS PARA HTML ===
                                On Error Resume Next
                                
                                valorCorretor = 0
                                If Not IsNull(rs("ValorCorretor")) Then
                                    valorCorretor = CDbl(rs("ValorCorretor"))
                                End If
                                
                                valorDiretoria = 0
                                If Not IsNull(rs("ValorDiretoria")) Then
                                    valorDiretoria = CDbl(rs("ValorDiretoria"))
                                End If
                                
                                valorGerencia = 0
                                If Not IsNull(rs("ValorGerencia")) Then
                                    valorGerencia = CDbl(rs("ValorGerencia"))
                                End If
                                
                                comissaoText = "0,00%"
                                If Not IsNull(rs("ComissaoPercentual")) Then
                                    comissaoText = FormatNumber(rs("ComissaoPercentual"), 2) & "%"
                                End If
                                
                                If Not IsNull(rs("ValorComissaoGeral")) And CDbl(rs("ValorComissaoGeral")) > 0 Then
                                    comissaoText = comissaoText & " (R$ " & FormatNumber(rs("ValorComissaoGeral"), 2) & ")"
                                End If
                                
                                valorUnidadeHtml = 0
                                If Not IsNull(rs("ValorUnidade")) Then
                                    valorUnidadeHtml = CDbl(rs("ValorUnidade"))
                                End If
                                
                                valorComissaoHtml = 0
                                If Not IsNull(rs("ValorComissaoGeral")) Then
                                    valorComissaoHtml = CDbl(rs("ValorComissaoGeral"))
                                End If
                                
                                totalValorHtml = totalValorHtml + valorUnidadeHtml
                                totalComissaoHtml = totalComissaoHtml + valorComissaoHtml
                                vAno = Right(GetStringField(rs, "AnoVenda"), 2)
                                On Error Goto 0
                            %>
                                <tr>
                                    <td><%= GetStringField(rs, "ID") %></td>
                                    <td>
                                        <%= GetStringField(rs, "AnoVenda") & "-" & Right("0" & GetStringField(rs, "MesVenda"), 2) %>
                                        <br>
                                        <small><%= vAno & "T" & GetStringField(rs, "Trimestre") %></small>
                                    </td>
                                    <td><%= RemoverNumeros(GetStringField(rs, "NomeEmpreendimento")) %></td>
                                    <td><%= GetStringField(rs, "Unidade") %></td>
                                    <td>
                                        <%= GetStringField(rs, "Diretoria") %>
                                        <% If GetStringField(rs, "ComissaoDiretoria") <> "" Then %>
                                        <br><small style="color: red;"><%= GetStringField(rs, "ComissaoDiretoria") %>% - R$ <%= FormatNumber(valorDiretoria, 2) %></small>
                                        <% End If %>
                                    </td>
                                    <td>
                                        <%= GetStringField(rs, "Gerencia") %>
                                        <% If GetStringField(rs, "ComissaoGerencia") <> "" Then %>
                                        <br><small style="color: red;"><%= GetStringField(rs, "ComissaoGerencia") %>% - R$ <%= FormatNumber(valorGerencia, 2) %></small>
                                        <% End If %>
                                    </td>
                                    <td>
                                        <%= GetStringField(rs, "Corretor") %>
                                        <% If GetStringField(rs, "ComissaoCorretor") <> "" Then %>
                                        <br><small style="color: red;"><%= GetStringField(rs, "ComissaoCorretor") %>% - R$ <%= FormatNumber(valorCorretor, 2) %></small>
                                        <% End If %>
                                    </td>
                                    <td><%= FormatDateTime(GetStringField(rs, "DataVenda"), 2) %></td>
                                    <td style="text-align: right;"><%= FormatNumber(valorUnidadeHtml, 2) %></td>
                                    <td style="text-align: right;"><span class="badge badge-comissao"><%= comissaoText %></span></td>
                                    <td>
                                        <small>
                                            <%= FormatDateTime(GetStringField(rs, "DataRegistro"), 2) %><br>
                                            por <%= GetStringField(rs, "Usuario") %>
                                        </small>
                                    </td>
                                    <td>

                                    </td>
                                    <td>
                                        <a href="gestao_vendas_restaurar.asp?id=<%= GetStringField(rs, "ID") %>" class="btn btn-info btn-sm" onclick="return confirm('Confirma restauração desta venda?');">
                                            <i class="fas fa-undo"></i> Restaurar
                                        </a>
                                    </td>
                                </tr>
                                <%
                                rs.MoveNext
                            Loop
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="total-row">
                                <th colspan="8" style="text-align: right;">Totais:</th>
                                <th style="text-align: right;"><%= FormatNumber(totalValorHtml, 2) %></th>
                                <th style="text-align: right;"><%= FormatNumber(totalComissaoHtml, 2) %></th>
                                <th colspan="3"></th>
                            </tr>
                        </tfoot>
                    </table>
                </div>
                <%
                Else
                %>
                <div class="alert alert-warning text-center">
                    <i class="fas fa-exclamation-triangle fa-2x mb-3"></i><br>
                    <h4>Não há vendas excluídas para exibir</h4>
                    <p class="mb-0">Todos os registros marcados como excluídos aparecerão aqui.</p>
                </div>
                <%
                End If
                %>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
    $(document).ready(function () {
        // Inicializa o DataTable apenas se a tabela existir e tiver dados
        if ($('#tabelaVendas').length) {
            $('#tabelaVendas').DataTable({
                language: {
                    url: "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
                },
                pageLength: 50,
                order: [[0, "desc"]],
                responsive: true,
                columnDefs: [
                    { orderable: false, targets: [10, 11, 12] },
                    { type: "date-eu", targets: [7] },
                    { 
                        type: "num-fmt", 
                        targets: [8, 9],
                        render: function (data, type, row) {
                            if (type === 'sort') {
                                return data.replace(/\./g, '').replace(',', '.');
                            }
                            return data;
                        }
                    }
                ]
            });
        }
    });
    </script>
</body>
</html>
<%
' Close ADODB connections
If rs.State = 1 Then rs.Close
Set rs = Nothing
If conn.State = 1 Then conn.Close
Set conn = Nothing
%>