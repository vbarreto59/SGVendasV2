<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
' ===========================================================
' PAINEL BOLSA DE VALORES - VGV DOS ÚLTIMOS 2 ANOS
' ===========================================================
Function GetBolsaVGV()
    Dim connSales, rsAnos, rsVendas, sql
    Dim anoAntigo, anoRecente, totalAntigo, totalRecente
    
    ' Verificar se a conexão existe
    If Len(StrConnSales) = 0 Then
        Server.Execute("conSunSales.asp")
    End If
    
    Set connSales = Server.CreateObject("ADODB.Connection")
    connSales.Open StrConnSales
    
    ' Buscar anos distintos com vendas
    sql = "SELECT DISTINCT TOP 2 AnoVenda FROM Vendas WHERE (Excluido <> -1 OR Excluido IS NULL) AND AnoVenda IS NOT NULL ORDER BY AnoVenda DESC"
    Set rsAnos = Server.CreateObject("ADODB.Recordset")
    rsAnos.Open sql, connSales
    
    If Not rsAnos.EOF Then
        rsAnos.MoveFirst
        anoRecente = rsAnos("AnoVenda")
        rsAnos.MoveNext
        
        If Not rsAnos.EOF Then
            anoAntigo = rsAnos("AnoVenda")
        Else
            ' Se só tem um ano, usa o anterior
            anoAntigo = anoRecente - 1
        End If
    Else
        ' Se não tem nenhum registro
        anoRecente = Year(Date())
        anoAntigo = anoRecente - 1
    End If
    
    rsAnos.Close
    Set rsAnos = Nothing
    
    ' Buscar total do ano antigo
    sql = "SELECT SUM(ValorUnidade) as Total FROM Vendas WHERE (AnoVenda = " & anoAntigo & " OR YEAR(DataVenda) = " & anoAntigo & ") AND (Excluido <> -1 OR Excluido IS NULL)"
    Set rsVendas = Server.CreateObject("ADODB.Recordset")
    rsVendas.Open sql, connSales
    
    If Not rsVendas.EOF Then
        If Not IsNull(rsVendas("Total")) Then
            totalAntigo = CDbl(rsVendas("Total"))
        Else
            totalAntigo = 0
        End If
    Else
        totalAntigo = 0
    End If
    
    rsVendas.Close
    
    ' Buscar total do ano recente
    sql = "SELECT SUM(ValorUnidade) as Total FROM Vendas WHERE (AnoVenda = " & anoRecente & " OR YEAR(DataVenda) = " & anoRecente & ") AND (Excluido <> -1 OR Excluido IS NULL)"
    rsVendas.Open sql, connSales
    
    If Not rsVendas.EOF Then
        If Not IsNull(rsVendas("Total")) Then
            totalRecente = CDbl(rsVendas("Total"))
        Else
            totalRecente = 0
        End If
    Else
        totalRecente = 0
    End If
    
    rsVendas.Close
    Set rsVendas = Nothing
    connSales.Close
    Set connSales = Nothing
    
    ' Calcular variação
    Dim variacao, variacaoValor, classeCor, icone
    variacaoValor = totalRecente - totalAntigo
    
    If totalAntigo > 0 Then
        variacao = ((totalRecente - totalAntigo) / totalAntigo) * 100
    Else
        If totalRecente > 0 Then
            variacao = 100
        Else
            variacao = 0
        End If
    End If
    
    ' Determinar cor
    If variacao > 0 Then
        classeCor = "text-success"
        icone = "fa-arrow-up"
    ElseIf variacao < 0 Then
        classeCor = "text-danger"
        icone = "fa-arrow-down"
    Else
        classeCor = "text-secondary"
        icone = "fa-minus"
    End If
    
    GetBolsaVGV = Array(anoAntigo, anoRecente, totalAntigo, totalRecente, variacao, classeCor, icone, variacaoValor)
End Function

' Obter dados
Dim bolsaVGV
bolsaVGV = GetBolsaVGV()

Dim bAnoAntigo, bAnoRecente, bTotalAntigo, bTotalRecente, bVariacao, bClasseCor, bIcone, bDiff
bAnoAntigo = bolsaVGV(0)
bAnoRecente = bolsaVGV(1)
bTotalAntigo = bolsaVGV(2)
bTotalRecente = bolsaVGV(3)
bVariacao = bolsaVGV(4)
bClasseCor = bolsaVGV(5)
bIcone = bolsaVGV(6)
bDiff = bolsaVGV(7)
%>

<style>
/* ================================================== */
/* PAINEL BOLSA DE VALORES - VGV */
/* ================================================== */
.bolsa-vgv-container {
    background: #000000;
    border-bottom: 1px solid #333;
    padding: 3px 0;
    font-family: 'Courier New', monospace;
    font-size: 14px;
    color: #fff;
    overflow: hidden;
    white-space: nowrap;
}

.bolsa-vgv-ticker {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0 10px;
}

.bolsa-vgv-item {
    display: flex;
    align-items: center;
    margin: 0 8px;
}

.bolsa-vgv-label {
    color: #999;
    margin-right: 4px;
    font-size: 14px;
}

.bolsa-vgv-value {
    font-weight: bold;
    font-size: 14px;
}

.bolsa-vgv-year {
    background: #222;
    padding: 1px 4px;
    border-radius: 2px;
    margin: 0 2px;
    font-size: 14px;
}

.bolsa-vgv-variacao {
    padding: 1px 4px;
    border-radius: 2px;
    font-size: 9px;
    margin-left: 4px;
}

.bolsa-vgv-separator {
    color: #333;
    margin: 0 6px;
}

/* Animações estilo bolsa */
@keyframes ticker {
    0% { transform: translateX(100%); }
    100% { transform: translateX(-100%); }
}

.bolsa-vgv-scrolling {
    animation: ticker 30s linear infinite;
}

/* Efeitos de cor */
.text-bolsa-up { color: #00ff00 !important; }
.text-bolsa-down { color: #ff0000 !important; }
.text-bolsa-neutral { color: #cccccc !important; }

/* Tooltip para valores completos */
[data-bs-toggle="tooltip"] {
    cursor: help;
}

/* Versão compacta para telas pequenas */
@media (max-width: 768px) {
    .bolsa-vgv-container {
        font-size: 9px;
        padding: 2px 0;
    }
    
    .bolsa-vgv-ticker {
        padding: 0 5px;
    }
    
    .bolsa-vgv-item {
        margin: 0 4px;
    }
    
    .bolsa-vgv-label {
        display: none;
    }
}
</style>

<!-- PAINEL BOLSA DE VALORES -->
<div class="bolsa-vgv-container">
    <div class="bolsa-vgv-ticker">
        <!-- Título -->
        <div class="bolsa-vgv-item">
            <span class="bolsa-vgv-label">VGV</span>
            <span class="bolsa-vgv-value" style="color: #ffff00;">SGVENDAS</span>
        </div>
        
        <div class="bolsa-vgv-separator">|</div>
        
        <!-- Ano Antigo -->
        <div class="bolsa-vgv-item">
            <span class="bolsa-vgv-year"><%= bAnoAntigo %></span>
            <span class="bolsa-vgv-value" data-bs-toggle="tooltip" title="R$ <%= FormatNumber(bTotalAntigo, 2) %>">
                R$<%= FormatNumber(bTotalAntigo/1000000, 1) %>M
            </span>
        </div>
        
        <div class="bolsa-vgv-separator">|</div>
        
        <!-- Ano Recente -->
        <div class="bolsa-vgv-item">
            <span class="bolsa-vgv-year"><%= bAnoRecente %></span>
            <span class="bolsa-vgv-value" data-bs-toggle="tooltip" title="R$ <%= FormatNumber(bTotalRecente, 2) %>">
                R$<%= FormatNumber(bTotalRecente/1000000, 1) %>M
            </span>
        </div>
        
        <div class="bolsa-vgv-separator">|</div>
        
        <!-- Variação -->
        <div class="bolsa-vgv-item">
            <span class="bolsa-vgv-label">VAR</span>
            <span class="bolsa-vgv-value <%= bClasseCor %>">
                <i class="fas <%= bIcone %>" style="font-size: 8px;"></i>
                <%= FormatNumber(bVariacao, 1) %>%
            </span>
            <span class="bolsa-vgv-variacao <%= Replace(bClasseCor, "text-", "bg-") %>" 
                  data-bs-toggle="tooltip" 
                  title="Diferença: R$ <%= FormatNumber(bDiff, 0) %>">
                <%= iif(bDiff > 0, "+", "") %><%= FormatNumber(bDiff/1000, 0) %>K
            </span>
        </div>
        
        <div class="bolsa-vgv-separator">|</div>
        
        <!-- Info adicional -->
        <div class="bolsa-vgv-item">
            <span class="bolsa-vgv-label">ATUAL</span>
            <span class="bolsa-vgv-value" style="color: #00ffff;">
                <%= Hour(Now) %>:<%= Right("0" & Minute(Now), 2) %>
            </span>
        </div>
    </div>
</div>

<script>
// Atualizar a hora a cada minuto
function atualizarHoraBolsa() {
    var agora = new Date();
    var horas = agora.getHours();
    var minutos = agora.getMinutes();
    var strHora = horas + ':' + (minutos < 10 ? '0' : '') + minutos;
    
    // Atualiza todos os elementos com a hora
    document.querySelectorAll('.bolsa-vgv-value[style*="00ffff"]').forEach(function(el) {
        el.textContent = strHora;
    });
}

// Inicializar tooltips do Bootstrap
document.addEventListener('DOMContentLoaded', function() {
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
    
    // Atualizar hora inicialmente
    atualizarHoraBolsa();
    
    // Atualizar a cada minuto
    setInterval(atualizarHoraBolsa, 60000);
});
</script>

<% 
' Função IIF
Function iif(condicao, verdadeiro, falso)
    If condicao Then
        iif = verdadeiro
    Else
        iif = falso
    End If
End Function
%>