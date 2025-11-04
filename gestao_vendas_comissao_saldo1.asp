<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' =======================================================
' Configuração da conexão com o banco de dados e consulta
' =======================================================

' Declaração das variáveis para a tabela
Dim conn, rs, sql

' Cria o objeto de conexão ADODB
Set conn = Server.CreateObject("ADODB.Connection")

' Abre a conexão usando a string de conexão do arquivo 'conSunSales.asp'
conn.Open StrConnSales

' Constrói a consulta SQL para totalizar as colunas por Nome, incluindo o Saldo
sql = "SELECT Nome, " & _
      "Sum(TotalVenda) AS SomaDeTotalVenda, " & _
      "Sum(TotalComissao) AS SomaDeTotalComissao, " & _
      "Sum(TotalComissaoPago) AS SomaDeTotalComissaoPago, " & _
      "Sum(TotalComissao) - Sum(TotalComissaoPago) AS Saldo " & _
      "FROM ComissaoSaldo GROUP BY Nome ORDER BY Nome;"

' Cria o Recordset e executa a consulta
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

' =======================================================
' Fim das consultas
' =======================================================
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Saldo Comissões Dinâmicas</title>
    
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
</head>
<body>

<div class="container mt-5">
    <h1 class="mb-4">Saldo Comissões a Pagar (Totais Dinâmicos)</h1>

    <table id="myTable" class="table table-striped table-bordered table-sm" style="width:100%">
        <thead>
            <tr class="bg-primary text-white">
                <th>Nome</th>
                <th class="text-end">Total Venda</th>
                <th class="text-end">Total Comissao</th>
                <th class="text-end">Total Pago</th>
                <th class="text-end">Saldo</th>
            </tr>
        </thead>
        <tbody>
            <% 
            If Not rs.EOF Then
                Do While Not rs.EOF
                    ' =======================================================
                    ' PRE-CÁLCULO DOS VALORES PARA DISPLAY E DATA-ATTRIBUTES
                    ' =======================================================
                    Dim vTotalVenda, vTotalComissao, vTotalComissaoPago, vSaldo
                    
                    vTotalVenda = 0
                    If Not IsNull(rs("SomaDeTotalVenda")) Then vTotalVenda = rs("SomaDeTotalVenda")
                    
                    vTotalComissao = 0
                    If Not IsNull(rs("SomaDeTotalComissao")) Then vTotalComissao = rs("SomaDeTotalComissao")
                    
                    vTotalComissaoPago = 0
                    If Not IsNull(rs("SomaDeTotalComissaoPago")) Then vTotalComissaoPago = rs("SomaDeTotalComissaoPago")
                    
                    vSaldo = vTotalComissao - vTotalComissaoPago
                %>
            <tr>
                <td><%= rs("Nome") %></td>
                <td class="text-end total-venda-col" data-numeric-value="<%= vTotalVenda %>">
                    <% Response.Write "R$ " & FormatNumber(vTotalVenda, 2) %>
                </td>
                <td class="text-end total-comissao-col" data-numeric-value="<%= vTotalComissao %>">
                    <% Response.Write "R$ " & FormatNumber(vTotalComissao, 2) %>
                </td>
                <td class="text-end total-pago-col" data-numeric-value="<%= vTotalComissaoPago %>">
                    <% Response.Write "R$ " & FormatNumber(vTotalComissaoPago, 2) %>
                </td>
                <td class="text-end saldo-col" data-numeric-value="<%= vSaldo %>">
                    <%
                    ' Adiciona classes CSS baseadas no Saldo para melhor visualização
                    Dim saldoClass
                    If vSaldo < 0 Then
                        saldoClass = "text-success fw-bold" ' Pago a mais/Ajuste
                    ElseIf vSaldo > 0 Then
                        saldoClass = "text-danger fw-bold" ' A Pagar
                    Else
                        saldoClass = "text-secondary" ' Saldo Zero
                    End If
                    %>
                    <span class="<%= saldoClass %>">
                        <%Response.Write "R$ " & FormatNumber(vSaldo, 2)%>
                    </span>
                </td>
            </tr>
            <% 
                    rs.MoveNext
                Loop
            End If
            
            ' Fecha o Recordset
            If Not rs Is Nothing Then
                If rs.State = 1 Then rs.Close
            End If
            Set rs = Nothing
            
            ' Fecha a conexão com o banco de dados
            If Not conn Is Nothing Then
                If conn.State = 1 Then conn.Close
            End If
            Set conn = Nothing
            %>
        </tbody>
        <!-- Seção TFOOT adicionada para exibir os totais -->
        <tfoot>
            <tr class="bg-dark text-white fw-bold">
                <th>Total Geral (Filtro)</th>
                <!-- IDs para as células de total, que serão atualizadas pelo JavaScript -->
                <th id="total-venda" class="text-end">R$ 0,00</th>
                <th id="total-comissao" class="text-end">R$ 0,00</th>
                <th id="total-pago" class="text-end">R$ 0,00</th>
                <th id="total-saldo" class="text-end">R$ 0,00</th>
            </tr>
        </tfoot>
    </table>
</div>

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>

<script>
    /**
     * Recalcula e exibe os totais das colunas visíveis/filtradas lendo o atributo data-numeric-value.
     * @param {object} api Instância da API do DataTables.
     */
    function updateTotals(api) {
        // Mapeamento: Coluna Index -> ID do elemento no rodapé
        // Lembre-se: Colunas são 0 (Nome), 1 (Venda), 2 (Comissão), 3 (Pago), 4 (Saldo)
        const columnsMap = {
            1: '#total-venda',    // Total Venda
            2: '#total-comissao', // Total Comissao
            3: '#total-pago',     // Total Pago
            4: '#total-saldo'     // Saldo
        };

        Object.keys(columnsMap).forEach(colIndex => {
            const totalId = columnsMap[colIndex];
            const columnIndex = parseInt(colIndex);

            // 1. Usa o seletor DataTables para obter os elementos <td> APENAS das linhas VISÍVEIS (filtradas)
            const total = api
                .column(columnIndex, { search: 'applied' }) // Garante que apenas as linhas filtradas sejam consideradas
                .nodes() // Obtém os elementos <td>
                .reduce(function(a, b) {
                    // 2. LÊ O VALOR NUMÉRICO BRUTO DO ATRIBUTO data-numeric-value
                    // O .data('numeric-value') do jQuery lê o atributo 'data-numeric-value'
                    const val = $(b).data('numeric-value') || 0;
                    
                    // Garante que o valor seja um número
                    const numericValue = parseFloat(val);
                    return a + (isNaN(numericValue) ? 0 : numericValue);
                }, 0); // 0 é o valor inicial do acumulador

            // 3. Formata o total de volta para o formato de moeda brasileiro (R$ X.XXX,XX)
            const formattedTotal = 'R$ ' + total.toLocaleString('pt-BR', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });

            // 4. Atualiza a célula no rodapé
            $(totalId).text(formattedTotal);
        });
    }

    $(document).ready(function() {
        const table = $('#myTable').DataTable({
            "order": [
                [0, "asc"]
            ],
            "pageLength": 100,
            "language": {
                "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
            },
            // Função DataTables que é chamada após a inicialização
            "initComplete": function(settings, json) {
                const api = this.api();
                updateTotals(api); // Calcula o total inicial (total geral)
            }
        });

        // Evento que dispara após qualquer operação de redesenho (filtro, busca, paginação, etc.)
        // Isso garante que os totais sejam sempre atualizados com base nos dados visíveis.
        table.on('draw', function() {
            const api = table.api();
            updateTotals(api); 
        });

        // Também atualiza os totais quando há mudanças na busca/filtro
        table.on('search.dt', function() {
            const api = table.api();
            updateTotals(api);
        });
    });
</script>

</body>
</html>
