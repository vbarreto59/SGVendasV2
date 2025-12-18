<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 04/12/2025                       -->
<!-- CODIGO_ARQUIVO: FICHA_CORRETOR         -->
<!-- OBS: Ficha Detalhada do Corretor       -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' ===============================================
' CONFIGURAÇÃO UTF-8
' ===============================================
Response.CodePage = 65001 ' UTF-8
Response.CharSet = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"
%>

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>Erro ao conectar ao banco de dados: " & Err.Description & "</div>"
    Response.End
End If
On Error GoTo 0

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno, filtroCorretor, modoRelatorio
filtroAno = Request.QueryString("ano")
filtroCorretor = Request.QueryString("corretor")
modoRelatorio = Request.QueryString("modo") ' "completo" ou "resumido"

If filtroAno = "" Then 
    filtroAno = Year(Date())
End If

If filtroCorretor = "" Then 
    filtroCorretor = "Todos"
End If

If modoRelatorio = "" Then 
    modoRelatorio = "completo"
End If

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

Function GetUniqueValues(tableName, columnName, whereClause)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    sqlQuery = "SELECT DISTINCT " & columnName & " FROM " & tableName & " "
    If whereClause <> "" Then
        sqlQuery = sqlQuery & " " & whereClause
    End If
    sqlQuery = sqlQuery & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = connSales.Execute(sqlQuery)
    If Err.Number <> 0 Then
        GetUniqueValues = Array()
        Exit Function
    End If
    On Error GoTo 0
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                dict.Add CStr(rs(0)), 1
            End If
            rs.MoveNext
        Loop
    End If
    
    If Not rs Is Nothing Then 
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Março"
arrMesesNome(4) = "Abril"
arrMesesNome(5) = "Maio"
arrMesesNome(6) = "Junho"
arrMesesNome(7) = "Julho"
arrMesesNome(8) = "Agosto"
arrMesesNome(9) = "Setembro"
arrMesesNome(10) = "Outubro"
arrMesesNome(11) = "Novembro"
arrMesesNome(12) = "Dezembro"

' ===============================================
' FUNÇÃO PARA CONVERTER VALORES NUMÉRICOS
' ===============================================
Function ConverterValor(valorString)
    Dim valorConvertido
    valorConvertido = 0
    
    If Not IsNull(valorString) And Trim(valorString) <> "" Then
        ' Remove caracteres não numéricos, exceto ponto e vírgula
        Dim valorTemp
        valorTemp = Trim(valorString)
        
        ' Verifica se já é um número
        If IsNumeric(valorTemp) Then
            valorConvertido = CDbl(valorTemp)
        Else
            ' Remove o símbolo de moeda se existir
            valorTemp = Replace(valorTemp, "R$", "")
            valorTemp = Replace(valorTemp, "$", "")
            valorTemp = Trim(valorTemp)
            
            valorTemp = Replace(valorTemp, ".","")
            valorTemp = Replace(valorTemp, ",",".")
            

            
            ' Remove quaisquer outros caracteres não numéricos
            Dim i, char, valorLimpo
            valorLimpo = ""
            For i = 1 To Len(valorTemp)
                char = Mid(valorTemp, i, 1)
                If IsNumeric(char) Or char = "." Or char = "-" Then
                    valorLimpo = valorLimpo & char
                End If
            Next
            
            If IsNumeric(valorLimpo) Then
                valorConvertido = CDbl(valorLimpo)
            End If
        End If
    End If
    
    ConverterValor = valorConvertido
End Function

' ===============================================
' OBTER LISTA DE CORRETORES
' ===============================================

Dim uniqueCorretores, uniqueAnos
uniqueCorretores = GetUniqueValues("Vendas", "Corretor", "WHERE Corretor IS NOT NULL AND Corretor <> '' AND Corretor <> ' '")
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")

' ===============================================
' DADOS PRINCIPAIS - APENAS SE ANO ESTIVER PREENCHIDO
' ===============================================

Dim dadosCorretor, totalGeralVendas, totalGeralVGV, totalGeralComissao
Dim empreendimentosDict, localidadesDict, mesesComVendas, mesesSemVendas
Dim vendasPorMesDetalhado, vendasPorMesQuantidade, empreendimentosPorLocalidade
Dim vendasPorLocalidade, vendasPorEmpreendimento ' NOVO: Dicionário para armazenar vendas por empreendimento
Set dadosCorretor = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    
    ' Construir WHERE clause baseado no filtro
    Dim whereClause, sqlSafeCorretor
    whereClause = "WHERE Excluido = 0 AND AnoVenda = " & filtroAno
    
    If filtroCorretor <> "Todos" Then
        sqlSafeCorretor = Replace(filtroCorretor, "'", "''")
        whereClause = whereClause & " AND Corretor = '" & sqlSafeCorretor & "'"
    End If
    
    ' ===============================================
    ' 1. DADOS MENSAIS DETALHADOS
    ' ===============================================
    
    Dim sqlDadosMensais, rsDadosMensais
    sqlDadosMensais = "SELECT " & _
                     "Corretor, " & _
                     "MesVenda, " & _
                     "COUNT(*) as QtdVendas, " & _
                     "SUM(ValorUnidade) as TotalVGV, " & _
                     "SUM(ValorCorretor) as TotalComissao " & _
                     "FROM Vendas " & _
                     whereClause & _
                     " GROUP BY Corretor, MesVenda " & _
                     "ORDER BY Corretor, MesVenda"

    Set rsDadosMensais = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsDadosMensais.Open sqlDadosMensais, connSales
    
    If Err.Number <> 0 Then
        Response.Write "<div class='alert alert-danger'>Erro na consulta: " & Err.Description & "</div>"
    Else
        ' Inicializar dicionários
        Set empreendimentosDict = Server.CreateObject("Scripting.Dictionary")
        Set localidadesDict = Server.CreateObject("Scripting.Dictionary")
        Set mesesComVendas = Server.CreateObject("Scripting.Dictionary")
        Set mesesSemVendas = Server.CreateObject("Scripting.Dictionary")
        Set vendasPorMesDetalhado = Server.CreateObject("Scripting.Dictionary")
        Set vendasPorMesQuantidade = Server.CreateObject("Scripting.Dictionary")
        Set vendasPorLocalidade = Server.CreateObject("Scripting.Dictionary")
        Set vendasPorEmpreendimento = Server.CreateObject("Scripting.Dictionary") ' NOVO: Inicializar dicionário
        
        totalGeralVendas = 0
        totalGeralVGV = 0
        totalGeralComissao = 0
        
        If Not rsDadosMensais.EOF Then
            
            Do While Not rsDadosMensais.EOF
                Dim corretorNome, mes, qtdVendas, totalVGV, totalComissao
                corretorNome = CStr(rsDadosMensais("Corretor"))
                mes = CStr(rsDadosMensais("MesVenda"))
                qtdVendas = CLng(rsDadosMensais("QtdVendas"))
                totalVGV = 0
                totalComissao = 0
                
                If Not IsNull(rsDadosMensais("TotalVGV")) Then
                    totalVGV = ConverterValor(rsDadosMensais("TotalVGV"))
                    totalVGV = Replace(totalVGV, ".","")
                    totalGeralVGV = totalGeralVGV + totalVGV
                    totalVGV = Replace(totalVGV, ",",".")                    
                End If
                
                If Not IsNull(rsDadosMensais("TotalComissao")) Then
                    totalComissao = ConverterValor(rsDadosMensais("TotalComissao"))
                End If
                
                ' Adicionar corretor ao dicionário principal se não existir
                If Not dadosCorretor.Exists(corretorNome) Then
                    Dim infoCorretor
                    Set infoCorretor = Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "TotalVendas", 0
                    infoCorretor.Add "TotalVGV", 0
                    infoCorretor.Add "TotalComissao", 0
                    infoCorretor.Add "Empreendimentos", Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "Localidades", Server.CreateObject("Scripting.Dictionary")
                    dadosCorretor.Add corretorNome, infoCorretor
                End If
                
                Set infoCorretor = dadosCorretor(corretorNome)
                
                ' Adicionar dados do mês
                infoCorretor("Meses").Add mes, Array(qtdVendas, totalVGV, totalComissao)
                
                ' Atualizar totais do corretor
                infoCorretor("TotalVendas") = infoCorretor("TotalVendas") + qtdVendas

                    totalVGV = ConverterValor(rsDadosMensais("TotalVGV"))
                    totalVGV = Replace(totalVGV, ".","")
                
                infoCorretor("TotalVGV") = infoCorretor("TotalVGV") + totalVGV
                infoCorretor("TotalComissao") = infoCorretor("TotalComissao") + totalComissao
                
                ' Adicionar aos dicionários das novas tabelas
                ' 1. Vendas por mês detalhado
                If Not vendasPorMesDetalhado.Exists(mes) Then
                    vendasPorMesDetalhado.Add mes, Array(totalVGV, qtdVendas, totalComissao)
                Else
                    Dim dadosMesExistente
                    dadosMesExistente = vendasPorMesDetalhado(mes)
                    dadosMesExistente(0) = dadosMesExistente(0) + totalVGV
                    dadosMesExistente(1) = dadosMesExistente(1) + qtdVendas
                    dadosMesExistente(2) = dadosMesExistente(2) + totalComissao
                    vendasPorMesDetalhado(mes) = dadosMesExistente
                End If
                
                ' 2. Quantidade por mês
                If Not vendasPorMesQuantidade.Exists(mes) Then
                    vendasPorMesQuantidade.Add mes, qtdVendas
                Else
                    vendasPorMesQuantidade(mes) = vendasPorMesQuantidade(mes) + qtdVendas
                End If
                
                ' Marcar mês como com vendas
                mesesComVendas.Add mes, 1
                
                ' Atualizar totais gerais
                totalGeralVendas = totalGeralVendas + qtdVendas
               
                totalGeralComissao = totalGeralComissao + totalComissao
                
                rsDadosMensais.MoveNext
            Loop
            
            ' Identificar meses sem vendas
            For i = 1 To 12
                If Not mesesComVendas.Exists(CStr(i)) Then
                    mesesSemVendas.Add CStr(i), arrMesesNome(i)
                End If
            Next
            
        Else
            Response.Write "<div class='alert alert-warning'>Nenhum registro encontrado para os filtros selecionados.</div>"
        End If
        
        rsDadosMensais.Close
    End If
    Set rsDadosMensais = Nothing
    
    ' ===============================================
    ' 2. OBTER EMPREENDIMENTOS E VENDAS POR EMPREENDIMENTO
    ' ===============================================
    
    If dadosCorretor.Count > 0 Then
        Dim sqlEmpreendVendas, rsEmpreendVendas
        
        ' NOVA CONSULTA: Obter vendas por empreendimento
        sqlEmpreendVendas = "SELECT " & _
                           "NomeEmpreendimento, " & _
                           "COUNT(*) as QtdVendas, " & _
                           "SUM(ValorUnidade) as TotalVGV, " & _
                           "MIN(ValorUnidade) as MenorVGV, " & _
                           "MAX(ValorUnidade) as MaiorVGV, " & _
                           "AVG(ValorUnidade) as MediaVGV " & _
                           "FROM Vendas " & _
                           whereClause & _
                           " AND NomeEmpreendimento IS NOT NULL " & _
                           " AND NomeEmpreendimento <> '' " & _
                           "GROUP BY NomeEmpreendimento " & _
                           "ORDER BY NomeEmpreendimento"
        
        Set rsEmpreendVendas = Server.CreateObject("ADODB.Recordset")
        On Error Resume Next
        rsEmpreendVendas.Open sqlEmpreendVendas, connSales
        
        If Err.Number <> 0 Then
            Response.Write "<div class='alert alert-danger'>Erro na consulta de empreendimentos: " & Err.Description & "</div>"
        Else
            If Not rsEmpreendVendas.EOF Then
                Do While Not rsEmpreendVendas.EOF
                    Dim empreendimento, qtdVendasEmp, totalVGVEmp, menorVGVEmp, maiorVGVEmp, mediaVGVEmp
                    empreendimento = CStr(rsEmpreendVendas("NomeEmpreendimento"))
                    qtdVendasEmp = CLng(rsEmpreendVendas("QtdVendas"))
                    
                    totalVGVEmp = 0
                    menorVGVEmp = 0
                    maiorVGVEmp = 0
                    mediaVGVEmp = 0
                    
                    If Not IsNull(rsEmpreendVendas("TotalVGV")) Then
                        totalVGVEmp = ConverterValor(rsEmpreendVendas("TotalVGV"))
                    End If
                    
                    If Not IsNull(rsEmpreendVendas("MenorVGV")) Then
                        menorVGVEmp = ConverterValor(rsEmpreendVendas("MenorVGV"))
                    End If
                    
                    If Not IsNull(rsEmpreendVendas("MaiorVGV")) Then
                        maiorVGVEmp = ConverterValor(rsEmpreendVendas("MaiorVGV"))
                    End If
                    
                    If Not IsNull(rsEmpreendVendas("MediaVGV")) Then
                        mediaVGVEmp = ConverterValor(rsEmpreendVendas("MediaVGV"))
                    End If
                    
                    ' Adicionar ao dicionário de empreendimentos
                    If Not empreendimentosDict.Exists(empreendimento) Then
                        empreendimentosDict.Add empreendimento, 1
                    End If
                    
                    ' Adicionar ao dicionário de vendas por empreendimento
                    Dim infoEmpreendVendas
                    Set infoEmpreendVendas = Server.CreateObject("Scripting.Dictionary")
                    infoEmpreendVendas.Add "QtdVendas", qtdVendasEmp
                    infoEmpreendVendas.Add "TotalVGV", totalVGVEmp
                    infoEmpreendVendas.Add "MenorVGV", menorVGVEmp
                    infoEmpreendVendas.Add "MaiorVGV", maiorVGVEmp
                    infoEmpreendVendas.Add "MediaVGV", mediaVGVEmp
                    
                    vendasPorEmpreendimento.Add empreendimento, infoEmpreendVendas
                    
                    rsEmpreendVendas.MoveNext
                Loop
            End If
            rsEmpreendVendas.Close
        End If
        Set rsEmpreendVendas = Nothing
        
        ' ===============================================
        ' 3. OBTER LOCALIDADES PARA OUTRAS ABAS (CORREÇÃO)
        ' ===============================================
        ' Usar o campo correto: Vendas.Localidade
        Dim campoLocalidade
        campoLocalidade = "Localidade" ' Campo correto conforme informação
        
        ' Consulta para obter localidades
        Dim sqlLocalidades, rsLocalidades
        sqlLocalidades = "SELECT " & _
                        campoLocalidade & ", " & _
                        "COUNT(*) as QtdVendas, " & _
                        "SUM(ValorUnidade) as TotalVGV " & _
                        "FROM Vendas " & _
                        whereClause & _
                        " AND " & campoLocalidade & " IS NOT NULL " & _
                        " AND " & campoLocalidade & " <> '' " & _
                        "GROUP BY " & campoLocalidade & " " & _
                        "ORDER BY " & campoLocalidade
        
        Set rsLocalidades = Server.CreateObject("ADODB.Recordset")
        On Error Resume Next
        rsLocalidades.Open sqlLocalidades, connSales
        
        If Err.Number <> 0 Then
            Response.Write "<div class='alert alert-warning'>Não foi possível obter dados de localidades. Campo de localização: " & campoLocalidade & "</div>"
        Else
            If Not rsLocalidades.EOF Then
                Do While Not rsLocalidades.EOF
                    Dim localidadeGeral, qtdVendasLocal, totalVGGLocal
                    
                    ' Verificar se o campo Localidade não é nulo
                    If Not IsNull(rsLocalidades(campoLocalidade)) Then
                        localidadeGeral = CStr(rsLocalidades(campoLocalidade))
                        
                        ' Verificar se a localidade não está vazia
                        If Trim(localidadeGeral) <> "" Then
                            qtdVendasLocal = CLng(rsLocalidades("QtdVendas"))
                            totalVGGLocal = 0
                            
                            If Not IsNull(rsLocalidades("TotalVGV")) Then
                                totalVGGLocal = ConverterValor(rsLocalidades("TotalVGV"))
                            End If
                            
                            ' Adicionar ao dicionário de localidades
                            If Not localidadesDict.Exists(localidadeGeral) Then
                                localidadesDict.Add localidadeGeral, 1
                            End If
                            
                            ' Adicionar ao dicionário de vendas por localidade
                            If Not vendasPorLocalidade.Exists(localidadeGeral) Then
                                Set infoVendasLocal = Server.CreateObject("Scripting.Dictionary")
                                infoVendasLocal.Add "QtdVendas", qtdVendasLocal
                                infoVendasLocal.Add "TotalVGV", totalVGGLocal
                                infoVendasLocal.Add "Empreendimentos", Server.CreateObject("Scripting.Dictionary")
                                vendasPorLocalidade.Add localidadeGeral, infoVendasLocal
                            End If
                        End If
                    End If
                    
                    rsLocalidades.MoveNext
                Loop
            End If
        End If
        
        If Not rsLocalidades Is Nothing Then
            If rsLocalidades.State = 1 Then rsLocalidades.Close
            Set rsLocalidades = Nothing
        End If
        On Error GoTo 0
        
        ' ===============================================
        ' 3.1 CONSULTA ALTERNATIVA: Se não encontrar localidades, tentar consulta mais simples
        ' ===============================================
        If localidadesDict.Count = 0 Then
            Dim sqlLocalidadesAlternativa, rsLocalidadesAlternativa
            sqlLocalidadesAlternativa = "SELECT DISTINCT Localidade FROM Vendas " & _
                                      whereClause & _
                                      " AND Localidade IS NOT NULL " & _
                                      " AND Localidade <> '' " & _
                                      "ORDER BY Localidade"
            
            Set rsLocalidadesAlternativa = Server.CreateObject("ADODB.Recordset")
            On Error Resume Next
            rsLocalidadesAlternativa.Open sqlLocalidadesAlternativa, connSales
            
            If Err.Number = 0 Then
                If Not rsLocalidadesAlternativa.EOF Then
                    Do While Not rsLocalidadesAlternativa.EOF
                        If Not IsNull(rsLocalidadesAlternativa("Localidade")) Then
                            Dim localidadeAlt
                            localidadeAlt = CStr(rsLocalidadesAlternativa("Localidade"))
                            
                            If Trim(localidadeAlt) <> "" And Not localidadesDict.Exists(localidadeAlt) Then
                                localidadesDict.Add localidadeAlt, 1
                                
                                ' Criar entrada no dicionário de vendas por localidade (sem dados de VGV)
                                If Not vendasPorLocalidade.Exists(localidadeAlt) Then
                                    Set infoVendasLocal = Server.CreateObject("Scripting.Dictionary")
                                    infoVendasLocal.Add "QtdVendas", 0
                                    infoVendasLocal.Add "TotalVGV", 0
                                    infoVendasLocal.Add "Empreendimentos", Server.CreateObject("Scripting.Dictionary")
                                    vendasPorLocalidade.Add localidadeAlt, infoVendasLocal
                                End If
                            End If
                        End If
                        rsLocalidadesAlternativa.MoveNext
                    Loop
                End If
                rsLocalidadesAlternativa.Close
            End If
            Set rsLocalidadesAlternativa = Nothing
            On Error GoTo 0
        End If
        
        ' ===============================================
        ' 3.2 OBTER EMPREENDIMENTOS POR LOCALIDADE
        ' ===============================================
        If localidadesDict.Count > 0 Then
            For Each localidadeKey In localidadesDict.Keys
                Dim sqlEmpreendLocal, rsEmpreendLocal
                sqlEmpreendLocal = "SELECT DISTINCT NomeEmpreendimento FROM Vendas " & _
                                  whereClause & _
                                  " AND Localidade = '" & Replace(localidadeKey, "'", "''") & "' " & _
                                  " AND NomeEmpreendimento IS NOT NULL " & _
                                  " AND NomeEmpreendimento <> ''"
                
                Set rsEmpreendLocal = Server.CreateObject("ADODB.Recordset")
                On Error Resume Next
                rsEmpreendLocal.Open sqlEmpreendLocal, connSales
                
                If Err.Number = 0 Then
                    If Not rsEmpreendLocal.EOF Then
                        Do While Not rsEmpreendLocal.EOF
                            If Not IsNull(rsEmpreendLocal("NomeEmpreendimento")) Then
                                Dim empreendLocal
                                empreendLocal = CStr(rsEmpreendLocal("NomeEmpreendimento"))
                                
                                If Trim(empreendLocal) <> "" Then
                                    ' Adicionar ao dicionário de empreendimentos da localidade
                                    If vendasPorLocalidade.Exists(localidadeKey) Then
                                        Set infoVendasLocal = vendasPorLocalidade(localidadeKey)
                                        If Not infoVendasLocal("Empreendimentos").Exists(empreendLocal) Then
                                            infoVendasLocal("Empreendimentos").Add empreendLocal, 1
                                        End If
                                    End If
                                End If
                            End If
                            rsEmpreendLocal.MoveNext
                        Loop
                    End If
                    rsEmpreendLocal.Close
                End If
                Set rsEmpreendLocal = Nothing
                On Error GoTo 0
            Next
        End If        
        ' ===============================================
        ' 4. CONSULTA ALTERNATIVA: Se não encontrar empreendimentos, tentar buscar diretamente
        ' ===============================================
        If empreendimentosDict.Count = 0 Then
            Dim sqlEmpreendimentosDirect, rsEmpreendimentosDirect
            sqlEmpreendimentosDirect = "SELECT DISTINCT NomeEmpreendimento FROM Vendas " & _
                                      whereClause & _
                                      " AND NomeEmpreendimento IS NOT NULL " & _
                                      " AND NomeEmpreendimento <> '' " & _
                                      "ORDER BY NomeEmpreendimento"
            
            Set rsEmpreendimentosDirect = Server.CreateObject("ADODB.Recordset")
            rsEmpreendimentosDirect.Open sqlEmpreendimentosDirect, connSales
            
            If Not rsEmpreendimentosDirect.EOF Then
                Do While Not rsEmpreendimentosDirect.EOF
                    If Not IsNull(rsEmpreendimentosDirect("NomeEmpreendimento")) Then
                        Dim empNome
                        empNome = CStr(rsEmpreendimentosDirect("NomeEmpreendimento"))
                        If empNome <> "" And Not empreendimentosDict.Exists(empNome) Then
                            empreendimentosDict.Add empNome, 1
                        End If
                    End If
                    rsEmpreendimentosDirect.MoveNext
                Loop
            End If
            rsEmpreendimentosDirect.Close
            Set rsEmpreendimentosDirect = Nothing
        End If
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Ficha do Corretor</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <!-- Chart.js para gráficos -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            background-color: #f8f9fa;
            padding: 20px;
            color: #333;
        }
        .container-fluid {
            max-width: 1800px;
            margin: 0 auto;
        }
        .filter-container {
            background-color: #FFF;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .card-dashboard {
            background-color: #FFF;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
            min-height: 100px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .kpi-card h5 {
            font-size: 0.9rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            font-size: 1.2rem;
            font-weight: bold;
        }
        .bg-primary-kpi { background-color: #007bff; }
        .bg-success-kpi { background-color: #28a745; }
        .bg-info-kpi { background-color: #17a2b8; }
        .bg-warning-kpi { background-color: #ffc107; color: #000; }
        .bg-danger-kpi { background-color: #dc3545; }
        .bg-purple-kpi { background-color: #6f42c1; }
        .bg-pink-kpi { background-color: #e83e8c; }
        .bg-teal-kpi { background-color: #20c997; }
        
        .table th {
            background-color: #800000;
            color: white;
        }
        .mes-com-venda { background-color: #d4edda !important; }
        .mes-sem-venda { background-color: #f8d7da !important; }
        
        .tab-content {
            background-color: #FFF;
            padding: 20px;
            border-radius: 0 0 8px 8px;
            border: 1px solid #dee2e6;
            border-top: none;
        }
        
        .nav-tabs .nav-link.active {
            background-color: #800000;
            color: white;
            border-color: #800000;
        }
        
        .nav-tabs .nav-link {
            color: #800000;
            font-weight: bold;
        }
        
        .chart-container {
            position: relative;
            height: 300px;
            width: 100%;
            margin-bottom: 30px;
        }
        
        .progress-bar-custom {
            background-color: #28a745;
            height: 20px;
            border-radius: 5px;
        }
        
        .badge-corretor {
            background-color: #6c757d;
            color: white;
            margin: 2px;
        }
        
        .table-sm th, .table-sm td {
            padding: 0.3rem;
            font-size: 0.9rem;
        }
        
        .localidade-total {
            font-size: 0.9rem;
            color: #666;
            font-weight: bold;
        }
        
        .vgv-detalhes {
            font-size: 0.85rem;
            color: #666;
        }
        
        .progress-thin {
            height: 8px;
            margin-top: 5px;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-user-tie"></i> SGVendas - Ficha do Corretor
        </h2>
        
        <!-- Filtros -->
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-3">
                        <div class="mb-3">
                            <label class="form-label">Ano</label>
                            <select class="form-select" name="ano" id="anoFilter" required>
                                <option value="">Selecione o ano</option>
                                <%
                                If IsArray(uniqueAnos) Then
                                    For Each ano In uniqueAnos
                                        Response.Write "<option value=""" & ano & """"
                                        If CStr(filtroAno) = CStr(ano) Then Response.Write " selected"
                                        Response.Write ">" & ano & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label class="form-label">Corretor</label>
                            <select class="form-select" name="corretor" id="corretorFilter">
                                <option value="Todos">Todos os Corretores</option>
                                <%
                                If IsArray(uniqueCorretores) Then
                                    For Each corretor In uniqueCorretores
                                        Response.Write "<option value=""" & Server.HTMLEncode(corretor) & """"
                                        If CStr(filtroCorretor) = CStr(corretor) Then Response.Write " selected"
                                        Response.Write ">" & corretor & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-3">
                        <div class="mb-3">
                            
                            <select class="form-select" name="modo" hidden>
                                <option value="completo" <% If modoRelatorio = "completo" Then Response.Write "selected" %>>Relatório Completo</option>
                                <option value="resumido" <% If modoRelatorio = "resumido" Then Response.Write "selected" %>>Relatório Resumido</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-2 d-flex align-items-end">
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Relatório
                        </button>
                    </div>
                </div>
            </form>
        </div>
        
        <% If filtroAno = "" Then %>
            <div class="alert alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar a ficha do corretor.
            </div>
        <% ElseIf dadosCorretor.Count = 0 Then %>
            <div class="alert alert-info text-center">
                <i class="fas fa-info-circle"></i> Nenhum dado encontrado para os filtros selecionados.
            </div>
        <% Else %>
        
        <!-- KPIs Gerais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <h5>Total de Corretores</h5>
                    <p><%= dadosCorretor.Count %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <h5>Total de Vendas</h5>
                    <p><%= totalGeralVendas %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <h5>VGV Total (R$)</h5>
                    <p><%= FormatNumber(totalGeralVGV, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <h5>Comissão Total (R$)</h5>
                    <p><%= FormatNumber(totalGeralComissao, 2) %></p>
                </div>
            </div>
        </div>
        
        <!-- Tabs de Navegação -->
        <ul class="nav nav-tabs mt-4" id="myTab" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="resumo-vgv-tab" data-bs-toggle="tab" data-bs-target="#resumo-vgv" type="button" role="tab">Resumo por VGV</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="resumo-tab" data-bs-toggle="tab" data-bs-target="#resumo" type="button" role="tab">Resumo Geral QTD</button>
            </li>
             <% If filtroCorretor <> "Todos" Then %>
                <li class="nav-item" role="presentation">
                    <button class="nav-link" id="vendas-mes-tab" data-bs-toggle="tab" data-bs-target="#vendas-mes" type="button" role="tab">Vendas por Mês</button>
                </li>
            <%end if%>    
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="quantidade-mes-tab" data-bs-toggle="tab" data-bs-target="#quantidade-mes" type="button" role="tab">Quantidade por Mês</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="empreendimentos-tab" data-bs-toggle="tab" data-bs-target="#empreendimentos" type="button" role="tab">Empreendimentos</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="localidades-tab" data-bs-toggle="tab" data-bs-target="#localidades" type="button" role="tab">Localidades</button>
            </li>
        </ul>
        
        <div class="tab-content" id="myTabContent">

            <!-- Tab 1: Resumo por VGV (NOVA ABA NA PRIMEIRA POSIÇÃO) -->
            <div class="tab-pane fade show active" id="resumo-vgv" role="tabpanel">
                <div class="row">
                    <div class="col-md-12">
                        <h4>Resumo por VGV - Ano <%= filtroAno %></h4>
                        <p class="text-muted">Corretores ordenados por Valor Geral de Vendas (VGV)</p>
                        
                        <%
                        ' Ordenar corretores por VGV (decrescente)
                        Dim arrCorretoresVGV
                        If dadosCorretor.Count > 0 Then
                            arrCorretoresVGV = dadosCorretor.Keys
                            
                            ' Bubble sort por VGV
                            For i = 0 To UBound(arrCorretoresVGV)
                                For j = i + 1 To UBound(arrCorretoresVGV)
                                    If dadosCorretor(arrCorretoresVGV(j))("TotalVGV") > dadosCorretor(arrCorretoresVGV(i))("TotalVGV") Then
                                        Dim tempCorretorVGV
                                        tempCorretorVGV = arrCorretoresVGV(i)
                                        arrCorretoresVGV(i) = arrCorretoresVGV(j)
                                        arrCorretoresVGV(j) = tempCorretorVGV
                                    End If
                                Next
                            Next
                            
                            ' Calcular totais para percentuais
                            Dim totalVGVGeral, totalComissaoGeral
                            totalVGVGeral = 0
                            totalComissaoGeral = 0
                            
                            For Each corretorKey In arrCorretoresVGV
                                Set infoCorretor = dadosCorretor(corretorKey)
                                totalVGVGeral = totalVGVGeral + infoCorretor("TotalVGV")
                                totalComissaoGeral = totalComissaoGeral + infoCorretor("TotalComissao")
                            Next
                        %>
                        
                        <div class="row mb-4">
                            <div class="col-md-3">
                                <div class="kpi-card bg-primary-kpi">
                                    <h5>Corretores Listados</h5>
                                    <p><%= dadosCorretor.Count %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-success-kpi">
                                    <h5>VGV Total (R$)</h5>
                                    <p><%= FormatNumber(totalVGVGeral, 2) %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-info-kpi">
                                    <h5>Comissão Total (R$)</h5>
                                    <p><%= FormatNumber(totalComissaoGeral, 2) %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-warning-kpi">
                                    <h5>VGV Médio por Corretor</h5>
                                    <p>
                                        <%
                                        If dadosCorretor.Count > 0 Then
                                            Response.Write FormatNumber(totalVGVGeral / dadosCorretor.Count, 2)
                                        Else
                                            Response.Write "0,00"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="table-responsive">
                            <table class="table table-striped table-hover">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Corretor</th>
                                        <th class="text-center">VGV Total (R$)</th>
                                        <th class="text-center">% do Total</th>
                                        <th class="text-center">Quantidade de Vendas</th>
                                        <th class="text-center">Comissão (R$)</th>
                                        <th class="text-center">VGV Médio por Venda</th>
                                        <th class="text-center">Meses com Vendas</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim contadorVGV
                                    contadorVGV = 0
                                    
                                    For Each corretorKey In arrCorretoresVGV
                                        contadorVGV = contadorVGV + 1
                                        Set infoCorretor = dadosCorretor(corretorKey)
                                        
                                        Dim percentualVGV, mediaVendaVGV
                                        
                                        ' Calcular percentual do VGV total
                                        If totalVGVGeral > 0 Then
                                            percentualVGV = (infoCorretor("TotalVGV") / totalVGVGeral) * 100
                                        Else
                                            percentualVGV = 0
                                        End If
                                        
                                        ' Calcular VGV médio por venda
                                        If infoCorretor("TotalVendas") > 0 Then
                                            mediaVendaVGV = infoCorretor("TotalVGV") / infoCorretor("TotalVendas")
                                        Else
                                            mediaVendaVGV = 0
                                        End If
                                        
                                        ' Contar meses com vendas para este corretor
                                        Dim mesesComVendasCorretor
                                        mesesComVendasCorretor = 0
                                        If Not infoCorretor("Meses") Is Nothing Then
                                            mesesComVendasCorretor = infoCorretor("Meses").Count
                                        End If
                                    %>
                                    <tr>
                                        <td><strong><%= contadorVGV %></strong></td>
                                        <td><strong><%= corretorKey %></strong></td>
                                        <td class="text-end">
                                            <span class="badge bg-success" style="font-size: 1.1em;">
                                                <%= FormatNumber(infoCorretor("TotalVGV"), 2) %>
                                            </span>
                                        </td>
                                        <td>
                                            <div class="progress" style="height: 20px;">
                                                <div class="progress-bar bg-info" role="progressbar" 
                                                     style="width: <%= percentualVGV %>%"
                                                     aria-valuenow="<%= percentualVGV %>" 
                                                     aria-valuemin="0" 
                                                     aria-valuemax="100">
                                                    <%= FormatNumber(percentualVGV, 1) %>%
                                                </div>
                                            </div>
                                        </td>
                                        <td class="text-center">
                                            <span class="badge bg-primary"><%= infoCorretor("TotalVendas") %></span>
                                        </td>
                                        <td class="text-end">
                                            <span class="text-muted"><%= FormatNumber(infoCorretor("TotalComissao"), 2) %></span>
                                        </td>
                                        <td class="text-end">
                                            <small><%= FormatNumber(mediaVendaVGV, 2) %></small>
                                        </td>
                                        <td class="text-center">
                                            <span class="badge bg-secondary"><%= mesesComVendasCorretor %></span>
                                        </td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td colspan="2"><strong>TOTAIS / MÉDIAS</strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalVGVGeral, 2) %></strong></td>
                                        <td><strong>100%</strong></td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalComissaoGeral, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalGeralVendas > 0 Then
                                                Response.Write FormatNumber(totalVGVGeral / totalGeralVendas, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not mesesComVendas Is Nothing Then
                                                Response.Write mesesComVendas.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                        
                        <div class="row mt-4">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Top 5 Corretores por VGV</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="list-group">
                                            <%
                                            Dim contadorTopVGV
                                            contadorTopVGV = 0
                                            
                                            For Each corretorKey In arrCorretoresVGV
                                                If contadorTopVGV < 5 Then
                                                    Set infoCorretor = dadosCorretor(corretorKey)
                                                    Dim percentualTopVGV
                                                    
                                                    If totalVGVGeral > 0 Then
                                                        percentualTopVGV = (infoCorretor("TotalVGV") / totalVGVGeral) * 100
                                                    Else
                                                        percentualTopVGV = 0
                                                    End If
                                            %>
                                            <div class="list-group-item list-group-item-action">
                                                <div class="d-flex w-100 justify-content-between">
                                                    <h6 class="mb-1"><%= contadorTopVGV + 1 %>. <%= corretorKey %></h6>
                                                    <small class="badge bg-success rounded-pill"><%= FormatNumber(infoCorretor("TotalVGV"), 0) %></small>
                                                </div>
                                                <p class="mb-1">VGV: <strong>R$ <%= FormatNumber(infoCorretor("TotalVGV"), 2) %></strong></p>
                                                <div class="d-flex justify-content-between">
                                                    <small>Vendas: <%= infoCorretor("TotalVendas") %></small>
                                                    <small><%= FormatNumber(percentualTopVGV, 1) %>% do total</small>
                                                </div>
                                                <div class="progress progress-thin mt-1">
                                                    <div class="progress-bar bg-success" role="progressbar" 
                                                         style="width: <%= percentualTopVGV %>%"
                                                         aria-valuenow="<%= percentualTopVGV %>" 
                                                         aria-valuemin="0" 
                                                         aria-valuemax="100">
                                                    </div>
                                                </div>
                                            </div>
                                            <%
                                                    contadorTopVGV = contadorTopVGV + 1
                                                End If
                                            Next
                                            %>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Distribuição de VGV por Corretor</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="graficoVGVCorretores"></canvas>
                                        </div>
                                        <div class="mt-3">
                                            <p class="mb-1"><i class="fas fa-trophy text-warning"></i> 
                                                <strong>Líder em VGV:</strong>
                                                <%
                                                If dadosCorretor.Count > 0 Then
                                                    Set infoCorretor = dadosCorretor(arrCorretoresVGV(0))
                                                    Response.Write "<br>" & arrCorretoresVGV(0) & " - R$ " & FormatNumber(infoCorretor("TotalVGV"), 2)
                                                End If
                                                %>
                                            </p>
                                            <p class="mb-1"><i class="fas fa-chart-line text-success"></i> 
                                                <strong>VGV Médio:</strong>
                                                <%
                                                If dadosCorretor.Count > 0 Then
                                                    Response.Write "R$ " & FormatNumber(totalVGVGeral / dadosCorretor.Count, 2) & " por corretor"
                                                End If
                                                %>
                                            </p>
                                            <p class="mb-1"><i class="fas fa-percentage text-info"></i> 
                                                <strong>Top 3 concentram:</strong>
                                                <%
                                                If dadosCorretor.Count >= 3 Then
                                                    Dim vgvTop3, percentualTop3
                                                    vgvTop3 = 0
                                                    For i = 0 To 2
                                                        If i <= UBound(arrCorretoresVGV) Then
                                                            Set infoCorretor = dadosCorretor(arrCorretoresVGV(i))
                                                            vgvTop3 = vgvTop3 + infoCorretor("TotalVGV")
                                                        End If
                                                    Next
                                                    
                                                    If totalVGVGeral > 0 Then
                                                        percentualTop3 = (vgvTop3 / totalVGVGeral) * 100
                                                        Response.Write FormatNumber(percentualTop3, 1) & "% do VGV total"
                                                    End If
                                                End If
                                                %>
                                            </p>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-warning text-center">
                            <i class="fas fa-exclamation-triangle"></i> Nenhum dado disponível para exibir o resumo por VGV.
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            
            <!-- Tab 2: Resumo Geral QTD (AGORA SEGUNDA POSIÇÃO) -->
            <div class="tab-pane fade" id="resumo" role="tabpanel">
                <% 
                Dim arrCorretoresResumo
                arrCorretoresResumo = dadosCorretor.Keys
                
                If IsArray(arrCorretoresResumo) Then
                    For i = 0 To UBound(arrCorretoresResumo)
                        For j = i + 1 To UBound(arrCorretoresResumo)
                            If dadosCorretor(arrCorretoresResumo(j))("TotalVendas") > dadosCorretor(arrCorretoresResumo(i))("TotalVendas") Then
                                Dim tempCorretor
                                tempCorretor = arrCorretoresResumo(i)
                                arrCorretoresResumo(i) = arrCorretoresResumo(j)
                                arrCorretoresResumo(j) = tempCorretor
                            End If
                        Next
                    Next
                End If
                %>
                
                <div class="row">
                    <div class="col-md-8">
                        <h4>Resumo por Corretor - Ano <%= filtroAno %></h4>
                        <div class="table-responsive">
                            <table class="table table-striped table-hover">
                                <thead>
                                    <tr>
                                        <th>Corretor</th>
                                        <th class="text-center">Vendas</th>
                                        <th class="text-end">VGV (R$)</th>
                                        <th class="text-end">Comissão (R$)</th>
                                        <th class="text-center">Média/Venda</th>
                                        <th class="text-center">Empreend.</th>
                                        <th class="text-center">Localidades</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    If IsArray(arrCorretoresResumo) Then
                                        vCont = 0
                                        For Each corretorKey In arrCorretoresResumo
                                            Set infoCorretor = dadosCorretor(corretorKey)
                                            Dim mediaVenda
                                            If infoCorretor("TotalVendas") > 0 Then
                                                mediaVenda = infoCorretor("TotalVGV") / infoCorretor("TotalVendas")
                                            Else
                                                mediaVenda = 0
                                            End If
                                            vCont = vCont + 1
                                    %>
                                        <tr>
                                            <td><strong><%=vCont%>-<%= corretorKey %></strong></td>
                                            <td class="text-center"><%= infoCorretor("TotalVendas") %></td>
                                            <td class="text-end"><%= FormatNumber(infoCorretor("TotalVGV"), 2) %></td>
                                            <td class="text-end"><%= FormatNumber(infoCorretor("TotalComissao"), 2) %></td>
                                            <td class="text-end"><%= FormatNumber(mediaVenda, 2) %></td>
                                            <td class="text-center">
                                                <%
                                                If Not infoCorretor("Empreendimentos") Is Nothing Then
                                                    'Response.Write infoCorretor("Empreendimentos").Count
                                                Else
                                                    'Response.Write "0"
                                                End If
                                                %>
                                            </td>
                                            <td class="text-center">
                                                <%
                                                If Not infoCorretor("Localidades") Is Nothing Then
                                                    'Response.Write infoCorretor("Localidades").Count
                                                Else
                                                    'Response.Write "0"
                                                End If
                                                %>
                                            </td>
                                        </tr>
                                        <%
                                        Next
                                    End If
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td><strong>TOTAIS</strong></td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralVGV, 2) %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralComissao, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalGeralVendas > 0 Then
                                                Response.Write FormatNumber(totalGeralVGV / totalGeralVendas, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not empreendimentosDict Is Nothing Then
                                                Response.Write empreendimentosDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not localidadesDict Is Nothing Then
                                                Response.Write localidadesDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <h4>Estatísticas do Período</h4>
                        
                        <div class="row mb-4">
                            <div class="col-6">
                                <div class="kpi-card bg-teal-kpi">
                                    <h5>Meses com Vendas</h5>
                                    <p>
                                        <%
                                        If Not mesesComVendas Is Nothing Then
                                            Response.Write mesesComVendas.Count
                                        Else
                                            Response.Write "0"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                            <div class="col-6">
                                <div class="kpi-card bg-pink-kpi">
                                    <h5>Meses sem Vendas</h5>
                                    <p>
                                        <%
                                        If Not mesesSemVendas Is Nothing Then
                                            Response.Write mesesSemVendas.Count
                                        Else
                                            Response.Write "12"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                        </div>
                        
                        <% If Not mesesSemVendas Is Nothing And mesesSemVendas.Count > 0 Then %>
                        <div class="card mb-3">
                            <div class="card-header">
                                <h5 class="mb-0">Meses sem Vendas</h5>
                            </div>
                            <div class="card-body">
                                <%
                                Dim arrMesesSemVenda
                                arrMesesSemVenda = mesesSemVendas.Keys
                                
                                For Each mesKey In arrMesesSemVenda
                                    Response.Write "<span class='badge bg-danger me-1 mb-1'>" & mesesSemVendas(mesKey) & "</span>"
                                Next
                                %>
                            </div>
                        </div>
                        <% End If %>
                        
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0">Top Corretores</h5>
                            </div>
                            <div class="card-body">
                                <div class="list-group">
                                    <%
                                    If IsArray(arrCorretoresResumo) Then
                                        Dim contadorTop
                                        contadorTop = 0
                                        For Each corretorKey In arrCorretoresResumo
                                            If contadorTop < 10 Then
                                                Set infoCorretor = dadosCorretor(corretorKey)
                                    %>
                                    <div class="list-group-item list-group-item-success">
                                        <div class="d-flex w-100 justify-content-between">
                                            <h6 class="mb-1"><%= contadorTop + 1 %>. <%= corretorKey %></h6>
                                            <small><%= infoCorretor("TotalVendas") %> vendas</small>
                                        </div>
                                        <p class="mb-1">VGV: R$ <%= FormatNumber(infoCorretor("TotalVGV"), 2) %></p>
                                        <small>Comissão: R$ <%= FormatNumber(infoCorretor("TotalComissao"), 2) %></small>
                                    </div>
                                    <%
                                            contadorTop = contadorTop + 1
                                            End If
                                        Next
                                    End If
                                    %>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Tab 3: Vendas por Mês (valores e quantidades) -->
        <% If filtroCorretor <> "Todos" Then %>    
            <div class="tab-pane fade" id="vendas-mes" role="tabpanel">
                <h4>Vendas por Mês - Ano <%= filtroAno %></h4>
                <p class="text-muted">Valor das vendas e quantidades por mês</p>
                
                <div class="row">
                    <div class="col-md-8">
                        <div class="table-responsive">
                            <table class="table table-striped table-bordered">
                                <thead>
                                    <tr>
                                        <th class="text-center">Mês</th>
                                        <th class="text-center">Quantidade de Vendas</th>
                                        <th class="text-center">VGV Total (R$)</th>
                                        <th class="text-center">Comissão Total (R$)</th>
                                        <th class="text-center">Média por Venda (R$)</th>
                                        <th class="text-center">% do Total</th>
                                    </tr>
                                </thead>
                                <!-- =========================================== -->
<tbody>
    <%
    If Not vendasPorMesDetalhado Is Nothing And vendasPorMesDetalhado.Count > 0 Then
        Dim arrMesesDetalhados
        arrMesesDetalhados = vendasPorMesDetalhado.Keys
        arrMesesDetalhados = SortArrayNumeric(arrMesesDetalhados)
        
        For Each mesNum In arrMesesDetalhados
            Dim dadosMesDetalhado
            dadosMesDetalhado = vendasPorMesDetalhado(mesNum)
            
            ' VALORES INICIAIS
            Dim vgvMes, qtdVendasMes, comissaoMes
            Dim mediaVendaMes, percentualMes
            
            vgvMes = 0
            qtdVendasMes = 0
            comissaoMes = 0
            mediaVendaMes = 0
            percentualMes = 0
            
            ' EXTRAIR VALORES SIMPLES
            If IsArray(dadosMesDetalhado) Then
                ' Valor VGV - índice 0
                If Not IsNull(dadosMesDetalhado(0)) Then
                    vgvMes = dadosMesDetalhado(0)
                End If
                
                ' Quantidade - índice 1
                If Not IsNull(dadosMesDetalhado(1)) Then
                    qtdVendasMes = dadosMesDetalhado(1)
                End If
                
                ' Comissão - índice 2
                If Not IsNull(dadosMesDetalhado(2)) Then
                    comissaoMes = dadosMesDetalhado(2)
                End If
            End If
            
            ' VERIFICAR SE OS VALORES SÃO NÚMEROS
            If Not IsNumeric(vgvMes) Then
                vgvMes = 0
            End If
            
            If Not IsNumeric(qtdVendasMes) Then
                qtdVendasMes = 0
            End If
            
            If Not IsNumeric(comissaoMes) Then
                comissaoMes = 0
            End If
            
            ' CONVERTER PARA NÚMEROS
            vgvMes = CDbl(vgvMes)
            qtdVendasMes = CLng(qtdVendasMes)
            comissaoMes = CDbl(comissaoMes)
            
            ' CÁLCULOS
            If qtdVendasMes > 0 Then
                mediaVendaMes = vgvMes / qtdVendasMes
            End If
            
            If totalGeralVGV > 0 Then
                percentualMes = (vgvMes / totalGeralVGV) * 100
            End If
    %>
    <tr>
        <td class="text-center"><strong><%= arrMesesNome(CInt(mesNum)) %></strong></td>
        <td class="text-center"><%= qtdVendasMes %></td>
        <td class="text-end"><strong><%= FormatNumber(vgvMes, 2) %></strong></td>
        <td class="text-end"><%= FormatNumber(comissaoMes, 2) %></td>
        <td class="text-end"><%= FormatNumber(mediaVendaMes, 2) %></td>
        <td>
            <div class="progress" style="height: 20px;">
                <div class="progress-bar bg-success" role="progressbar" 
                     style="width: <%= percentualMes %>%"
                     aria-valuenow="<%= percentualMes %>" 
                     aria-valuemin="0" 
                     aria-valuemax="100">
                    <%= FormatNumber(percentualMes, 1) %>%
                </div>
            </div>
        </td>
    </tr>
    <%
        Next
    Else
    %>
    <tr>
        <td colspan="6" class="text-center">Nenhum dado disponível</td>
    </tr>
    <%
    End If
    %>
</tbody>
                                <!-- ================== -->
                                <tfoot>
                                    <tr class="table-dark">
                                        <td class="text-center"><strong>TOTAL ANUAL</strong></td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralVGV, 2) %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralComissao, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalGeralVendas > 0 Then
                                                Response.Write FormatNumber(totalGeralVGV / totalGeralVendas, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td><strong>100%</strong></td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <h5>Distribuição Mensal</h5>
                        <div class="chart-container">
                            <canvas id="graficoVendasMes"></canvas>
                        </div>
                        
                        <!-- ==================================================== -->
<div class="card mt-3">
    <div class="card-header">
        <h6 class="mb-0">Resumo</h6>
    </div>
    <div class="card-body">
        <p><i class="fas fa-chart-line text-primary"></i> <strong>Melhor Mês:</strong>
        <%
        ' FUNÇÃO SEGURA PARA FORMATAR NÚMEROS
        Function SafeFormatNumber(valor, casasDecimais)
            SafeFormatNumber = "0,00"
            If Not IsNull(valor) Then
                If IsNumeric(valor) Then
                    On Error Resume Next
                    SafeFormatNumber = FormatNumber(CDbl(valor), casasDecimais)
                    If Err.Number <> 0 Then
                        SafeFormatNumber = "0,00"
                    End If
                    On Error GoTo 0
                End If
            End If
        End Function
        
        If Not vendasPorMesDetalhado Is Nothing And vendasPorMesDetalhado.Count > 0 Then
            Dim melhorMes, melhorValor, melhorMesNome
            melhorValor = 0
            
            For Each mesNum In vendasPorMesDetalhado.Keys
                Dim dadosMesTemp
                dadosMesTemp = vendasPorMesDetalhado(mesNum)
                
                ' Extrair valor VGV de forma segura
                Dim valorTemp
                valorTemp = 0
                
                If IsArray(dadosMesTemp) Then
                    If Not IsNull(dadosMesTemp(0)) Then
                        ' Converter para número de forma segura
                        If IsNumeric(dadosMesTemp(0)) Then
                            valorTemp = CDbl(dadosMesTemp(0))
                        Else
                            ' Tentar converter string para número
                            Dim strValor
                            strValor = CStr(dadosMesTemp(0))
                            strValor = Replace(strValor, ".", "")
                            strValor = Replace(strValor, ",", ".")
                            
                            If IsNumeric(strValor) Then
                                valorTemp = CDbl(strValor)
                            End If
                        End If
                    End If
                End If
                
                If valorTemp > melhorValor Then
                    melhorValor = valorTemp
                    melhorMes = mesNum
                End If
            Next
            
            If melhorMes <> "" Then
                Response.Write arrMesesNome(CInt(melhorMes)) & " - R$ " & SafeFormatNumber(melhorValor, 2)
            Else
                Response.Write "Nenhum"
            End If
        Else
            Response.Write "Nenhum"
        End If
        %>
        </p>
        
        <p><i class="fas fa-calendar-alt text-success"></i> <strong>Meses com Maior Quantidade:</strong>
        <%
        If Not vendasPorMesQuantidade Is Nothing And vendasPorMesQuantidade.Count > 0 Then
            Dim mesesMaisVendas, contadorMeses
            Set mesesMaisVendas = Server.CreateObject("Scripting.Dictionary")
            
            For Each mesNum In vendasPorMesQuantidade.Keys
                mesesMaisVendas.Add mesNum, vendasPorMesQuantidade(mesNum)
            Next
            
            Dim arrMesesQtd
            arrMesesQtd = mesesMaisVendas.Keys
            arrMesesQtd = SortArrayByValue(mesesMaisVendas, arrMesesQtd)
            
            contadorMeses = 0
            For Each mesNum In arrMesesQtd
                If contadorMeses < 3 Then
                    If contadorMeses > 0 Then Response.Write ", "
                    Response.Write arrMesesNome(CInt(mesNum)) & " (" & mesesMaisVendas(mesNum) & ")"
                    contadorMeses = contadorMeses + 1
                End If
            Next
        Else
            Response.Write "Nenhum"
        End If
        %>
        </p>
    </div>
</div>
<!-- ================================================= -->

                    </div>
                </div>
            </div>
<%end if%>            
            
            <!-- Tab 4: Quantidade por Mês -->
            <div class="tab-pane fade" id="quantidade-mes" role="tabpanel">
                <h4>Quantidade Vendida por Mês - Ano <%= filtroAno %></h4>
                
                <div class="row">
                    <div class="col-md-8">
                        <div class="table-responsive">
                            <table class="table table-striped table-bordered table-sm">
                                <thead>
                                    <tr>
                                        <th class="text-center">Mês</th>
                                        <%
                                        ' Cabeçalho dos corretores
                                        If IsArray(arrCorretoresResumo) Then
                                            For Each corretorKey In arrCorretoresResumo
                                                Response.Write "<th class='text-center small'>" & Left(corretorKey, 10) & "</th>"
                                            Next
                                        End If
                                        %>
                                        <th class="text-center table-dark">Total</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    For i = 1 To 12
                                        Dim totalMesQuantidade
                                        totalMesQuantidade = 0
                                    %>
                                    <tr>
                                        <td class="text-center"><strong><%= arrMesesNome(i) %></strong></td>
                                        <%
                                        If IsArray(arrCorretoresResumo) Then
                                            For Each corretorKey In arrCorretoresResumo
                                                Set infoCorretor = dadosCorretor(corretorKey)
                                                Set mesesCorretor = infoCorretor("Meses")
                                                
                                                If mesesCorretor.Exists(CStr(i)) Then
                                                    Dim dadosMesCorretor
                                                    dadosMesCorretor = mesesCorretor(CStr(i))
                                                    totalMesQuantidade = totalMesQuantidade + dadosMesCorretor(0)
                                                    Response.Write "<td class='text-center'>" & dadosMesCorretor(0) & "</td>"
                                                Else
                                                    Response.Write "<td class='text-center'>-</td>"
                                                End If
                                            Next
                                        End If
                                        %>
                                        <td class="text-center table-dark"><strong>
                                            <%
                                            If Not vendasPorMesQuantidade Is Nothing Then
                                                If vendasPorMesQuantidade.Exists(CStr(i)) Then
                                                    Response.Write vendasPorMesQuantidade(CStr(i))
                                                Else
                                                    Response.Write "0"
                                                End If
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                        </strong></td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td class="text-center"><strong>TOTAIS</strong></td>
                                        <%
                                        If IsArray(arrCorretoresResumo) Then
                                            For Each corretorKey In arrCorretoresResumo
                                                Set infoCorretor = dadosCorretor(corretorKey)
                                                Response.Write "<td class='text-center'><strong>" & infoCorretor("TotalVendas") & "</strong></td>"
                                            Next
                                        End If
                                        %>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <h5>Gráfico de Quantidade por Mês</h5>
                        <div class="chart-container">
                            <canvas id="graficoQuantidadeMes"></canvas>
                        </div>
                        
                        <div class="card mt-3">
                            <div class="card-header">
                                <h6 class="mb-0">Análise de Performance</h6>
                            </div>
                            <div class="card-body">
                                <p><i class="fas fa-trophy text-warning"></i> <strong>Mês com Mais Vendas:</strong>
                                <%
                                If Not vendasPorMesQuantidade Is Nothing And vendasPorMesQuantidade.Count > 0 Then
                                    Dim mesMaisVendas, qtdMaisVendas
                                    qtdMaisVendas = 0
                                    
                                    For Each mesNum In vendasPorMesQuantidade.Keys
                                        If vendasPorMesQuantidade(mesNum) > qtdMaisVendas Then
                                            qtdMaisVendas = vendasPorMesQuantidade(mesNum)
                                            mesMaisVendas = mesNum
                                        End If
                                    Next
                                    
                                    If mesMaisVendas <> "" Then
                                        Response.Write arrMesesNome(CInt(mesMaisVendas)) & " (" & qtdMaisVendas & " vendas)"
                                    End If
                                End If
                                %>
                                </p>
                                
                                <p><i class="fas fa-chart-bar text-info"></i> <strong>Média Mensal:</strong>
                                <%
                                If mesesComVendas.Count > 0 Then
                                    Response.Write FormatNumber(totalGeralVendas / mesesComVendas.Count, 1) & " vendas/mês"
                                Else
                                    Response.Write "0 vendas/mês"
                                End If
                                %>
                                </p>
                                
                                <p><i class="fas fa-percentage text-success"></i> <strong>Variação:</strong>
                                <%
                                If mesesComVendas.Count > 1 Then
                                    ' Calcular variação simples entre primeiro e último mês com vendas
                                    Dim mesesArray, primeiroMes, ultimoMes
                                    mesesArray = vendasPorMesQuantidade.Keys
                                    mesesArray = SortArrayNumeric(mesesArray)
                                    
                                    If UBound(mesesArray) >= 1 Then
                                        primeiroMes = mesesArray(0)
                                        ultimoMes = mesesArray(UBound(mesesArray))
                                        
                                        Dim crescimento
                                        crescimento = ((vendasPorMesQuantidade(ultimoMes) - vendasPorMesQuantidade(primeiroMes)) / vendasPorMesQuantidade(primeiroMes)) * 100
                                        
                                        If crescimento > 0 Then
                                            Response.Write "<span class='text-success'>+" & FormatNumber(crescimento, 1) & "%</span>"
                                        ElseIf crescimento < 0 Then
                                            Response.Write "<span class='text-danger'>" & FormatNumber(crescimento, 1) & "%</span>"
                                        Else
                                            Response.Write "<span class='text-muted'>" & FormatNumber(crescimento, 1) & "%</span>"
                                        End If
                                    End If
                                End If
                                %>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Tab 5: Empreendimentos -->
            <div class="tab-pane fade" id="empreendimentos" role="tabpanel">
                <div class="row">
                    <div class="col-md-12">
                        <h4>Empreendimentos Vendidos - Ano <%= filtroAno %></h4>
                        <p class="text-muted mb-3">Detalhamento dos valores de VGV por empreendimento</p>
                        
                        <%
                        If Not vendasPorEmpreendimento Is Nothing And vendasPorEmpreendimento.Count > 0 Then
                            Dim arrEmpreendimentosVendas
                            arrEmpreendimentosVendas = vendasPorEmpreendimento.Keys
                            
                            ' Ordenar empreendimentos por VGV total (decrescente)
                            arrEmpreendimentosVendas = SortArrayByVGV(vendasPorEmpreendimento, arrEmpreendimentosVendas)
                            
                            ' Calcular totais
                            Dim totalVendasEmpreendimentos, totalVGVEmpreendimentos
                            totalVendasEmpreendimentos = 0
                            totalVGVEmpreendimentos = 0
                            
                            For Each empreendKey In arrEmpreendimentosVendas
                                Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                totalVendasEmpreendimentos = totalVendasEmpreendimentos + infoEmpreend("QtdVendas")
                                totalVGVEmpreendimentos = totalVGVEmpreendimentos + infoEmpreend("TotalVGV")
                            Next
                        %>
                        
                        <div class="row mb-4">
                            <div class="col-md-3">
                                <div class="kpi-card bg-primary-kpi">
                                    <h5>Total Empreendimentos</h5>
                                    <p><%= vendasPorEmpreendimento.Count %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-success-kpi">
                                    <h5>Vendas nos Empreendimentos</h5>
                                    <p><%= totalVendasEmpreendimentos %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-info-kpi">
                                    <h5>VGV Total (R$)</h5>
                                    <p><%= FormatNumber(totalVGVEmpreendimentos, 2) %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-warning-kpi">
                                    <h5>Média por Empreendimento</h5>
                                    <p>
                                        <%
                                        If vendasPorEmpreendimento.Count > 0 Then
                                            Response.Write FormatNumber(totalVGVEmpreendimentos / vendasPorEmpreendimento.Count, 2)
                                        Else
                                            Response.Write "0,00"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Nome do Empreendimento</th>
                                        <th class="text-center">Quantidade de Vendas</th>
                                        <th class="text-center">% do Total</th>
                                        <th class="text-center">VGV Total (R$)</th>
                                        <th class="text-center">VGV Mínimo (R$)</th>
                                        <th class="text-center">VGV Máximo (R$)</th>
                                        <th class="text-center">VGV Médio (R$)</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim contadorEmp
                                    contadorEmp = 0
                                    
                                    For Each empreendKey In arrEmpreendimentosVendas
                                        contadorEmp = contadorEmp + 1
                                        Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                        
                                        Dim percentualEmp
                                        If totalVendasEmpreendimentos > 0 Then
                                            percentualEmp = (infoEmpreend("QtdVendas") / totalVendasEmpreendimentos) * 100
                                        Else
                                            percentualEmp = 0
                                        End If
                                    %>
                                    <tr>
                                        <td><%= contadorEmp %></td>
                                        <td><strong><%= empreendKey %></strong></td>
                                        <td class="text-center">
                                            <span class="badge bg-success" style="font-size: 1.1em;">
                                                <%= infoEmpreend("QtdVendas") %>
                                            </span>
                                        </td>
                                        <td>
                                            <div class="progress progress-thin">
                                                <div class="progress-bar bg-info" role="progressbar" 
                                                     style="width: <%= percentualEmp %>%"
                                                     aria-valuenow="<%= percentualEmp %>" 
                                                     aria-valuemin="0" 
                                                     aria-valuemax="100">
                                                </div>
                                            </div>
                                            <div class="vgv-detalhes text-center">
                                                <%= FormatNumber(percentualEmp, 1) %>%
                                            </div>
                                        </td>
                                        <td class="text-end">
                                            <strong><%= FormatNumber(infoEmpreend("TotalVGV"), 2) %></strong>
                                            <div class="vgv-detalhes">
                                                <%
                                                If totalVGVEmpreendimentos > 0 Then
                                                    'Dim percentualVGV
                                                    percentualVGV = (infoEmpreend("TotalVGV") / totalVGVEmpreendimentos) * 100
                                                    Response.Write FormatNumber(percentualVGV, 1) & "% do total"
                                                End If
                                                %>
                                            </div>
                                        </td>
                                        <td class="text-end">
                                            <span class="vgv-detalhes"><%= FormatNumber(infoEmpreend("MenorVGV"), 2) %></span>
                                        </td>
                                        <td class="text-end">
                                            <span class="vgv-detalhes"><%= FormatNumber(infoEmpreend("MaiorVGV"), 2) %></span>
                                        </td>
                                        <td class="text-end">
                                            <strong><%= FormatNumber(infoEmpreend("MediaVGV"), 2) %></strong>
                                        </td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td colspan="2"><strong>TOTAIS / MÉDIAS</strong></td>
                                        <td class="text-center"><strong><%= totalVendasEmpreendimentos %></strong></td>
                                        <td><strong>100%</strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalVGVEmpreendimentos, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            ' Calcular menor VGV entre todos os empreendimentos
                                            Dim menorVGVGeral
                                            menorVGVGeral = 0
                                            If vendasPorEmpreendimento.Count > 0 Then
                                                For Each empreendKey In arrEmpreendimentosVendas
                                                    Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                                    If menorVGVGeral = 0 Or infoEmpreend("MenorVGV") < menorVGVGeral Then
                                                        menorVGVGeral = infoEmpreend("MenorVGV")
                                                    End If
                                                Next
                                                Response.Write FormatNumber(menorVGVGeral, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            ' Calcular maior VGV entre todos os empreendimentos
                                            Dim maiorVGVGeral
                                            maiorVGVGeral = 0
                                            If vendasPorEmpreendimento.Count > 0 Then
                                                For Each empreendKey In arrEmpreendimentosVendas
                                                    Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                                    If infoEmpreend("MaiorVGV") > maiorVGVGeral Then
                                                        maiorVGVGeral = infoEmpreend("MaiorVGV")
                                                    End If
                                                Next
                                                Response.Write FormatNumber(maiorVGVGeral, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalVendasEmpreendimentos > 0 Then
                                                Response.Write FormatNumber(totalVGVEmpreendimentos / totalVendasEmpreendimentos, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                        
                        <div class="row mt-4">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Top 5 Empreendimentos por VGV</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="list-group">
                                            <%
                                            Dim contadorTopEmp
                                            contadorTopEmp = 0
                                            
                                            For Each empreendKey In arrEmpreendimentosVendas
                                                If contadorTopEmp < 5 Then
                                                    Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                            %>
                                            <div class="list-group-item">
                                                <div class="d-flex w-100 justify-content-between">
                                                    <h6 class="mb-1"><%= contadorTopEmp + 1 %>. <%= empreendKey %></h6>
                                                    <small class="badge bg-success rounded-pill"><%= infoEmpreend("QtdVendas") %></small>
                                                </div>
                                                <p class="mb-1">VGV Total: <strong>R$ <%= FormatNumber(infoEmpreend("TotalVGV"), 2) %></strong></p>
                                                <small>
                                                    Mínimo: R$ <%= FormatNumber(infoEmpreend("MenorVGV"), 2) %> | 
                                                    Máximo: R$ <%= FormatNumber(infoEmpreend("MaiorVGV"), 2) %> | 
                                                    Médio: R$ <%= FormatNumber(infoEmpreend("MediaVGV"), 2) %>
                                                </small>
                                            </div>
                                            <%
                                                    contadorTopEmp = contadorTopEmp + 1
                                                End If
                                            Next
                                            %>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Distribuição de VGV por Empreendimento</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="graficoEmpreendimentos"></canvas>
                                        </div>
                                        <div class="mt-3">
                                            <p class="mb-1"><i class="fas fa-chart-pie text-primary"></i> 
                                                <strong>Empreendimento com Maior VGV:</strong>
                                                <%
                                                If vendasPorEmpreendimento.Count > 0 Then
                                                    Dim empreendMaiorVGV, maiorVGVValor
                                                    maiorVGVValor = 0
                                                    
                                                    For Each empreendKey In arrEmpreendimentosVendas
                                                        Set infoEmpreend = vendasPorEmpreendimento(empreendKey)
                                                        If infoEmpreend("TotalVGV") > maiorVGVValor Then
                                                            maiorVGVValor = infoEmpreend("TotalVGV")
                                                            empreendMaiorVGV = empreendKey
                                                        End If
                                                    Next
                                                    
                                                    If empreendMaiorVGV <> "" Then
                                                        Response.Write "<br>" & empreendMaiorVGV & " - R$ " & FormatNumber(maiorVGVValor, 2)
                                                    End If
                                                End If
                                                %>
                                            </p>
                                            <p class="mb-1"><i class="fas fa-balance-scale text-success"></i> 
                                                <strong>Média de Vendas por Empreendimento:</strong>
                                                <%
                                                If vendasPorEmpreendimento.Count > 0 Then
                                                    Response.Write FormatNumber(totalVendasEmpreendimentos / vendasPorEmpreendimento.Count, 1) & " vendas"
                                                End If
                                                %>
                                            </p>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-warning">
                            <i class="fas fa-exclamation-triangle"></i> Nenhum dado de empreendimento disponível para os filtros selecionados.
                            <br><small>Verifique se existem vendas com o campo 'NomeEmpreendimento' preenchido no ano selecionado.</small>
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            
            <!-- Tab 6: Localidades -->
            <div class="tab-pane fade" id="localidades" role="tabpanel">
                <div class="row">
                    <div class="col-md-12">
                        <h4>Localidades - Ano <%= filtroAno %></h4>
                        <div class="alert alert-info mb-3">
                            <i class="fas fa-info-circle"></i> 
                            Esta seção mostra o quantitativo de vendas por localidade conforme os filtros aplicados.
                            <% If filtroCorretor <> "Todos" Then %>
                                <br><strong>Filtro ativo:</strong> Corretor: <%= filtroCorretor %>
                            <% End If %>
                        </div>
                        
                        <%
                        If Not vendasPorLocalidade Is Nothing And vendasPorLocalidade.Count > 0 Then
                            Dim arrLocalidadesVendas
                            arrLocalidadesVendas = vendasPorLocalidade.Keys
                            
                            ' Ordenar localidades alfabeticamente
                            arrLocalidadesVendas = SortArrayAlphabetical(arrLocalidadesVendas)
                            
                            ' Calcular totais
                            Dim totalVendasLocalidades, totalVGGLocalidades
                            totalVendasLocalidades = 0
                            totalVGGLocalidades = 0
                            
                            For Each localidadeKey In arrLocalidadesVendas
                                Set infoVendasLocal = vendasPorLocalidade(localidadeKey)
                                totalVendasLocalidades = totalVendasLocalidades + infoVendasLocal("QtdVendas")
                                totalVGGLocalidades = totalVGGLocalidades + infoVendasLocal("TotalVGV")
                            Next
                        %>
                        
                        <div class="row mb-4">
                            <div class="col-md-3">
                                <div class="kpi-card bg-primary-kpi">
                                    <h5>Total Localidades</h5>
                                    <p><%= vendasPorLocalidade.Count %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-success-kpi">
                                    <h5>Vendas nas Localidades</h5>
                                    <p><%= totalVendasLocalidades %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-info-kpi">
                                    <h5>VGV nas Localidades (R$)</h5>
                                    <p><%= FormatNumber(totalVGGLocalidades, 2) %></p>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="kpi-card bg-warning-kpi">
                                    <h5>Média por Localidade</h5>
                                    <p>
                                        <%
                                        If vendasPorLocalidade.Count > 0 Then
                                            Response.Write FormatNumber(totalVendasLocalidades / vendasPorLocalidade.Count, 1) & " vendas"
                                        Else
                                            Response.Write "0"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Localidade</th>
                                        <th class="text-center">Quantidade de Vendas</th>
                                        <th class="text-center">% do Total</th>
                                        <th class="text-center">VGV Total (R$)</th>
                                        <th class="text-center">Média por Venda (R$)</th>
                                        <th class="text-center">Empreendimentos</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    For Each localidadeKey In arrLocalidadesVendas
                                        Set infoVendasLocal = vendasPorLocalidade(localidadeKey)
                                        
                                        Dim percentualLocal, mediaVendaLocal
                                        
                                        If totalVendasLocalidades > 0 Then
                                            percentualLocal = (infoVendasLocal("QtdVendas") / totalVendasLocalidades) * 100
                                        Else
                                            percentualLocal = 0
                                        End If
                                        
                                        If infoVendasLocal("QtdVendas") > 0 Then
                                            mediaVendaLocal = infoVendasLocal("TotalVGV") / infoVendasLocal("QtdVendas")
                                        Else
                                            mediaVendaLocal = 0
                                        End If
                                    %>
                                    <tr>
                                        <td><strong><%= localidadeKey %></strong></td>
                                        <td class="text-center">
                                            <span class="badge bg-success" style="font-size: 1.1em;">
                                                <%= infoVendasLocal("QtdVendas") %>
                                            </span>
                                        </td>
                                        <td>
                                            <div class="progress" style="height: 20px;">
                                                <div class="progress-bar bg-info" role="progressbar" 
                                                     style="width: <%= percentualLocal %>%"
                                                     aria-valuenow="<%= percentualLocal %>" 
                                                     aria-valuemin="0" 
                                                     aria-valuemax="100">
                                                    <%= FormatNumber(percentualLocal, 1) %>%
                                                </div>
                                            </div>
                                        </td>
                                        <td class="text-end">
                                            <strong><%= FormatNumber(infoVendasLocal("TotalVGV"), 2) %></strong>
                                        </td>
                                        <td class="text-end">
                                            <span class="localidade-total"><%= FormatNumber(mediaVendaLocal, 2) %></span>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            ' Obter empreendimentos para esta localidade
                                            Dim sqlEmpreendPorLocal, rsEmpreendPorLocal
                                            sqlEmpreendPorLocal = "SELECT DISTINCT NomeEmpreendimento FROM Vendas " & _
                                                                whereClause & _
                                                                " AND (Cidade = '" & Replace(localidadeKey, "'", "''") & "' OR Municipio = '" & Replace(localidadeKey, "'", "''") & "') " & _
                                                                " AND NomeEmpreendimento IS NOT NULL " & _
                                                                " AND NomeEmpreendimento <> ''"
                                            
                                            Set rsEmpreendPorLocal = Server.CreateObject("ADODB.Recordset")
                                            On Error Resume Next
                                            rsEmpreendPorLocal.Open sqlEmpreendPorLocal, connSales
                                            
                                            Dim contadorEmpreendLocal
                                            contadorEmpreendLocal = 0
                                            Dim listaEmpreendLocal
                                            listaEmpreendLocal = ""
                                            
                                            If Err.Number = 0 Then
                                                If Not rsEmpreendPorLocal.EOF Then
                                                    Do While Not rsEmpreendPorLocal.EOF
                                                        If Not IsNull(rsEmpreendPorLocal("NomeEmpreendimento")) Then
                                                            contadorEmpreendLocal = contadorEmpreendLocal + 1
                                                            If listaEmpreendLocal <> "" Then listaEmpreendLocal = listaEmpreendLocal & ", "
                                                            listaEmpreendLocal = listaEmpreendLocal & rsEmpreendPorLocal("NomeEmpreendimento")
                                                        End If
                                                        rsEmpreendPorLocal.MoveNext
                                                    Loop
                                                End If
                                            End If
                                            
                                            If Not rsEmpreendPorLocal Is Nothing Then
                                                If rsEmpreendPorLocal.State = 1 Then rsEmpreendPorLocal.Close
                                                Set rsEmpreendPorLocal = Nothing
                                            End If
                                            On Error GoTo 0
                                            
                                            If contadorEmpreendLocal > 0 Then
                                                Response.Write "<span class='badge bg-primary' data-bs-toggle='tooltip' title='" & listaEmpreendLocal & "'>" & contadorEmpreendLocal & "</span>"
                                            Else
                                                Response.Write "<span class='badge bg-secondary'>0</span>"
                                            End If
                                            %>
                                        </td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td><strong>TOTAIS</strong></td>
                                        <td class="text-center"><strong><%= totalVendasLocalidades %></strong></td>
                                        <td><strong>100%</strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalVGGLocalidades, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalVendasLocalidades > 0 Then
                                                Response.Write FormatNumber(totalVGGLocalidades / totalVendasLocalidades, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not empreendimentosDict Is Nothing Then
                                                Response.Write empreendimentosDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                        
                        <div class="row mt-4">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Top 5 Localidades por Vendas</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="list-group">
                                            <%
                                            ' Ordenar localidades por quantidade de vendas (decrescente)
                                            Dim arrLocalidadesOrdenadas
                                            arrLocalidadesOrdenadas = vendasPorLocalidade.Keys
                                            
                            ' Bubble sort simples para ordenar por quantidade de vendas
                            If IsArray(arrLocalidadesOrdenadas) Then
                                For i = 0 To UBound(arrLocalidadesOrdenadas)
                                    For j = i + 1 To UBound(arrLocalidadesOrdenadas)
                                        Set infoLocalI = vendasPorLocalidade(arrLocalidadesOrdenadas(i))
                                        Set infoLocalJ = vendasPorLocalidade(arrLocalidadesOrdenadas(j))
                                        
                                        If infoLocalJ("QtdVendas") > infoLocalI("QtdVendas") Then
                                            Dim tempLocal
                                            tempLocal = arrLocalidadesOrdenadas(i)
                                            arrLocalidadesOrdenadas(i) = arrLocalidadesOrdenadas(j)
                                            arrLocalidadesOrdenadas(j) = tempLocal
                                        End If
                                    Next
                                Next
                            End If
                            
                            Dim contadorTopLocal
                            contadorTopLocal = 0
                            
                            If IsArray(arrLocalidadesOrdenadas) Then
                                For Each localidadeTop In arrLocalidadesOrdenadas
                                    If contadorTopLocal < 5 Then
                                        Set infoLocalTop = vendasPorLocalidade(localidadeTop)
                            %>
                            <div class="list-group-item">
                                <div class="d-flex w-100 justify-content-between">
                                    <h6 class="mb-1"><%= contadorTopLocal + 1 %>. <%= localidadeTop %></h6>
                                    <small class="badge bg-success rounded-pill"><%= infoLocalTop("QtdVendas") %></small>
                                </div>
                                <p class="mb-1">VGV: R$ <%= FormatNumber(infoLocalTop("TotalVGV"), 2) %></p>
                                <%
                                ' Obter empreendimentos para esta localidade
                                Dim sqlEmpTopLocal, rsEmpTopLocal
                                sqlEmpTopLocal = "SELECT COUNT(DISTINCT NomeEmpreendimento) as TotalEmpreend FROM Vendas " & _
                                                whereClause & _
                                                " AND (Cidade = '" & Replace(localidadeTop, "'", "''") & "' OR Municipio = '" & Replace(localidadeTop, "'", "''") & "') " & _
                                                " AND NomeEmpreendimento IS NOT NULL"
                                
                                Set rsEmpTopLocal = Server.CreateObject("ADODB.Recordset")
                                On Error Resume Next
                                rsEmpTopLocal.Open sqlEmpTopLocal, connSales
                                
                                Dim totalEmpreendTop
                                totalEmpreendTop = 0
                                
                                If Err.Number = 0 Then
                                    If Not rsEmpTopLocal.EOF Then
                                        If Not IsNull(rsEmpTopLocal("TotalEmpreend")) Then
                                            totalEmpreendTop = rsEmpTopLocal("TotalEmpreend")
                                        End If
                                    End If
                                End If
                                
                                If Not rsEmpTopLocal Is Nothing Then
                                    If rsEmpTopLocal.State = 1 Then rsEmpTopLocal.Close
                                    Set rsEmpTopLocal = Nothing
                                End If
                                On Error GoTo 0
                                %>
                                <small>Empreendimentos: <%= totalEmpreendTop %></small>
                            </div>
                            <%
                                        contadorTopLocal = contadorTopLocal + 1
                                    End If
                                Next
                            End If
                            %>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <h5 class="mb-0">Distribuição por Localidade</h5>
                                    </div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="graficoLocalidades"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-warning">
                            <i class="fas fa-exclamation-triangle"></i> Nenhum dado de localidade disponível para os filtros selecionados.
                            <br><small>Verifique se existem vendas com informações de localização (Cidade/Município) no ano selecionado.</small>
                            <br><small>Total de registros encontrados: <%= totalGeralVendas %></small>
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            
        </div>
        
        <!-- Botões de Ação -->
        <div class="row mt-4">
            <div class="col-12">

            </div>
        </div>
        
        <% End If %>
    </div>

    <!-- Scripts -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    
    <script>
    $(document).ready(function() {
        // Inicializar tabs do Bootstrap
        var triggerTabList = [].slice.call(document.querySelectorAll('#myTab button'))
        triggerTabList.forEach(function (triggerEl) {
            var tabTrigger = new bootstrap.Tab(triggerEl)
            triggerEl.addEventListener('click', function (event) {
                event.preventDefault()
                tabTrigger.show()
            })
        });
        
        // Inicializar tooltips
        var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
        var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl)
        });
        
        // Renderizar gráficos
        setTimeout(function() {
            renderizarGraficos();
        }, 500);
    });
    
    function renderizarGraficos() {
        // Gráfico de Vendas por Mês (valores)
        if (document.getElementById('graficoVendasMes')) {
            const ctxVendas = document.getElementById('graficoVendasMes').getContext('2d');
            
            // Dados para o gráfico
            const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
            const valores = [
                <%= GetDataGraficoVendas() %>
            ];
            
            new Chart(ctxVendas, {
                type: 'bar',
                data: {
                    labels: meses,
                    datasets: [{
                        label: 'VGV (R$)',
                        data: valores,
                        backgroundColor: 'rgba(40, 167, 69, 0.7)',
                        borderColor: 'rgba(40, 167, 69, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: {
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR');
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return 'VGV: R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                                }
                            }
                        }
                    }
                }
            });
        }
        
        // Gráfico de Quantidade por Mês
        if (document.getElementById('graficoQuantidadeMes')) {
            const ctxQuantidade = document.getElementById('graficoQuantidadeMes').getContext('2d');
            
            // Dados para o gráfico
            const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
            const quantidades = [
                <%= GetDataGraficoQuantidade() %>
            ];
            
            new Chart(ctxQuantidade, {
                type: 'line',
                data: {
                    labels: meses,
                    datasets: [{
                        label: 'Quantidade de Vendas',
                        data: quantidades,
                        backgroundColor: 'rgba(23, 162, 184, 0.2)',
                        borderColor: 'rgba(23, 162, 184, 1)',
                        borderWidth: 3,
                        tension: 0.1,
                        fill: true
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: {
                                stepSize: 1
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        }
                    }
                }
            });
        }
        
        // Gráfico de Localidades (somente se houver dados)
        if (document.getElementById('graficoLocalidades')) {
            const ctxLocalidades = document.getElementById('graficoLocalidades').getContext('2d');
            
            // Dados para o gráfico (top 10 localidades por vendas)
            <%
            If Not vendasPorLocalidade Is Nothing And vendasPorLocalidade.Count > 0 Then
                Dim arrLocalidadesGrafico, arrVendasGrafico, arrVGVGrafico
                arrLocalidadesGrafico = vendasPorLocalidade.Keys
                
                ' Ordenar por quantidade de vendas (decrescente)
                If IsArray(arrLocalidadesGrafico) Then
                    For i = 0 To UBound(arrLocalidadesGrafico)
                        For j = i + 1 To UBound(arrLocalidadesGrafico)
                            Set infoLocalI = vendasPorLocalidade(arrLocalidadesGrafico(i))
                            Set infoLocalJ = vendasPorLocalidade(arrLocalidadesGrafico(j))
                            
                            If infoLocalJ("QtdVendas") > infoLocalI("QtdVendas") Then
                                Dim tempLocalGrafico
                                tempLocalGrafico = arrLocalidadesGrafico(i)
                                arrLocalidadesGrafico(i) = arrLocalidadesGrafico(j)
                                arrLocalidadesGrafico(j) = tempLocalGrafico
                            End If
                        Next
                    Next
                End If
                
                ' Limitar a 10 localidades para o gráfico
                Dim limiteGrafico
                If UBound(arrLocalidadesGrafico) < 9 Then
                    limiteGrafico = UBound(arrLocalidadesGrafico)
                Else
                    limiteGrafico = 9
                End If
                
                If limiteGrafico >= 0 Then
            %>
            
            const localidadesLabels = [
                <%
                For i = 0 To limiteGrafico
                    If i <= UBound(arrLocalidadesGrafico) Then
                        Response.Write "'" & arrLocalidadesGrafico(i) & "', "
                    End If
                Next
                %>
            ];
            
            const localidadesVendas = [
                <%
                For i = 0 To limiteGrafico
                    If i <= UBound(arrLocalidadesGrafico) Then
                        Set infoLocalGrafico = vendasPorLocalidade(arrLocalidadesGrafico(i))
                        Response.Write infoLocalGrafico("QtdVendas") & ", "
                    End If
                Next
                %>
            ];
            
            const localidadesVGV = [
                <%
                For i = 0 To limiteGrafico
                    If i <= UBound(arrLocalidadesGrafico) Then
                        Set infoLocalGrafico = vendasPorLocalidade(arrLocalidadesGrafico(i))
                        Response.Write infoLocalGrafico("TotalVGV") & ", "
                    End If
                Next
                %>
            ];
            
            new Chart(ctxLocalidades, {
                type: 'bar',
                data: {
                    labels: localidadesLabels,
                    datasets: [{
                        label: 'Quantidade de Vendas',
                        data: localidadesVendas,
                        backgroundColor: 'rgba(0, 123, 255, 0.7)',
                        borderColor: 'rgba(0, 123, 255, 1)',
                        borderWidth: 1,
                        yAxisID: 'y'
                    }, {
                        label: 'VGV (R$)',
                        data: localidadesVGV,
                        backgroundColor: 'rgba(220, 53, 69, 0.7)',
                        borderColor: 'rgba(220, 53, 69, 1)',
                        borderWidth: 1,
                        yAxisID: 'y1'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: 'Quantidade de Vendas'
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: 'VGV (R$)'
                            },
                            grid: {
                                drawOnChartArea: false
                            },
                            ticks: {
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR');
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    if (context.datasetIndex === 0) {
                                        return 'Vendas: ' + context.parsed.y;
                                    } else {
                                        return 'VGV: R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                                    }
                                }
                            }
                        }
                    }
                }
            });
            <%
                Else
                Response.Write "<!-- Não há dados suficientes para o gráfico de localidades -->"
                End If
            End If
            %>
        }
        
        // Gráfico de Empreendimentos
        if (document.getElementById('graficoEmpreendimentos')) {
            const ctxEmpreendimentos = document.getElementById('graficoEmpreendimentos').getContext('2d');
            
            // Dados para o gráfico (top 10 empreendimentos por VGV)
            <%
            If Not vendasPorEmpreendimento Is Nothing And vendasPorEmpreendimento.Count > 0 Then
                Dim arrEmpreendimentosGrafico, arrVGVEmpreendimentos
                arrEmpreendimentosGrafico = vendasPorEmpreendimento.Keys
                
                ' Ordenar por VGV total (decrescente)
                arrEmpreendimentosGrafico = SortArrayByVGV(vendasPorEmpreendimento, arrEmpreendimentosGrafico)
                
                ' Limitar a 10 empreendimentos para o gráfico
                Dim limiteGraficoEmp
                If UBound(arrEmpreendimentosGrafico) < 9 Then
                    limiteGraficoEmp = UBound(arrEmpreendimentosGrafico)
                Else
                    limiteGraficoEmp = 9
                End If
                
                If limiteGraficoEmp >= 0 Then
            %>
            
            const empreendimentosLabels = [
                <%
                For i = 0 To limiteGraficoEmp
                    If i <= UBound(arrEmpreendimentosGrafico) Then
                        Response.Write "'" & Left(arrEmpreendimentosGrafico(i), 15) & "', "
                    End If
                Next
                %>
            ];
            
            const empreendimentosVGV = [
                <%
                For i = 0 To limiteGraficoEmp
                    If i <= UBound(arrEmpreendimentosGrafico) Then
                        Set infoEmpreendGrafico = vendasPorEmpreendimento(arrEmpreendimentosGrafico(i))
                        Response.Write infoEmpreendGrafico("TotalVGV") & ", "
                    End If
                Next
                %>
            ];
            
            const empreendimentosVendas = [
                <%
                For i = 0 To limiteGraficoEmp
                    If i <= UBound(arrEmpreendimentosGrafico) Then
                        Set infoEmpreendGrafico = vendasPorEmpreendimento(arrEmpreendimentosGrafico(i))
                        Response.Write infoEmpreendGrafico("QtdVendas") & ", "
                    End If
                Next
                %>
            ];
            
            new Chart(ctxEmpreendimentos, {
                type: 'bar',
                data: {
                    labels: empreendimentosLabels,
                    datasets: [{
                        label: 'VGV Total (R$)',
                        data: empreendimentosVGV,
                        backgroundColor: 'rgba(255, 193, 7, 0.7)',
                        borderColor: 'rgba(255, 193, 7, 1)',
                        borderWidth: 1,
                        yAxisID: 'y'
                    }, {
                        label: 'Quantidade de Vendas',
                        data: empreendimentosVendas,
                        backgroundColor: 'rgba(40, 167, 69, 0.7)',
                        borderColor: 'rgba(40, 167, 69, 1)',
                        borderWidth: 1,
                        yAxisID: 'y1'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: 'VGV (R$)'
                            },
                            ticks: {
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR');
                                }
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: 'Quantidade de Vendas'
                            },
                            grid: {
                                drawOnChartArea: false
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    if (context.datasetIndex === 0) {
                                        return 'VGV: R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                                    } else {
                                        return 'Vendas: ' + context.parsed.y;
                                    }
                                }
                            }
                        }
                    }
                }
            });
            <%
                Else
                Response.Write "<!-- Não há dados suficientes para o gráfico de empreendimentos -->"
                End If
            End If
            %>
        }
        
        // NOVO: Gráfico de VGV por Corretor
        if (document.getElementById('graficoVGVCorretores')) {
            const ctxVGVCorretores = document.getElementById('graficoVGVCorretores').getContext('2d');
            
            // Dados para o gráfico (top 10 corretores por VGV)
            <%
            If dadosCorretor.Count > 0 Then
                ' Reutilizar array já ordenado por VGV
                Dim arrCorretoresGraficoVGV
                arrCorretoresGraficoVGV = arrCorretoresVGV ' Usar o array já ordenado acima
                
                ' Limitar a 10 corretores para o gráfico
                Dim limiteGraficoVGV
                If UBound(arrCorretoresGraficoVGV) < 9 Then
                    limiteGraficoVGV = UBound(arrCorretoresGraficoVGV)
                Else
                    limiteGraficoVGV = 9
                End If
                
                If limiteGraficoVGV >= 0 Then
            %>
            
            const corretoresLabelsVGV = [
                <%
                For i = 0 To limiteGraficoVGV
                    If i <= UBound(arrCorretoresGraficoVGV) Then
                        Response.Write "'" & Left(arrCorretoresGraficoVGV(i), 12) & "', "
                    End If
                Next
                %>
            ];
            
            const corretoresValoresVGV = [
                <%
                For i = 0 To limiteGraficoVGV
                    If i <= UBound(arrCorretoresGraficoVGV) Then
                        Set infoCorretorGrafico = dadosCorretor(arrCorretoresGraficoVGV(i))
                        Response.Write infoCorretorGrafico("TotalVGV") & ", "
                    End If
                Next
                %>
            ];
            
            const corretoresVendasVGV = [
                <%
                For i = 0 To limiteGraficoVGV
                    If i <= UBound(arrCorretoresGraficoVGV) Then
                        Set infoCorretorGrafico = dadosCorretor(arrCorretoresGraficoVGV(i))
                        Response.Write infoCorretorGrafico("TotalVendas") & ", "
                    End If
                Next
                %>
            ];
            
            new Chart(ctxVGVCorretores, {
                type: 'bar',
                data: {
                    labels: corretoresLabelsVGV,
                    datasets: [{
                        label: 'VGV (R$)',
                        data: corretoresValoresVGV,
                        backgroundColor: 'rgba(220, 53, 69, 0.7)',
                        borderColor: 'rgba(220, 53, 69, 1)',
                        borderWidth: 1,
                        yAxisID: 'y'
                    }, {
                        label: 'Quantidade de Vendas',
                        data: corretoresVendasVGV,
                        backgroundColor: 'rgba(0, 123, 255, 0.7)',
                        borderColor: 'rgba(0, 123, 255, 1)',
                        borderWidth: 1,
                        yAxisID: 'y1'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: 'VGV (R$)'
                            },
                            ticks: {
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR');
                                }
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: 'Quantidade de Vendas'
                            },
                            grid: {
                                drawOnChartArea: false
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    if (context.datasetIndex === 0) {
                                        return 'VGV: R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                                    } else {
                                        return 'Vendas: ' + context.parsed.y;
                                    }
                                }
                            }
                        }
                    }
                }
            });
            <%
                Else
                Response.Write "<!-- Não há dados suficientes para o gráfico de VGV por corretor -->"
                End If
            End If
            %>
        }
    }
    
    function exportToExcel() {
        // Função básica de exportação
        alert('Funcionalidade de exportação para Excel será implementada em breve!');
    }
    </script>
</body>
</html>

<%
' ===============================================
' FUNÇÕES AUXILIARES
' ===============================================

Function SortArrayNumeric(arr)
    ' Ordena array numericamente
    If IsArray(arr) Then
        Dim i, j, temp
        For i = 0 To UBound(arr)
            For j = i + 1 To UBound(arr)
                If CInt(arr(j)) < CInt(arr(i)) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next
        Next
    End If
    SortArrayNumeric = arr
End Function

Function SortArrayAlphabetical(arr)
    ' Ordena array alfabeticamente
    If IsArray(arr) Then
        Dim i, j, temp
        For i = 0 To UBound(arr)
            For j = i + 1 To UBound(arr)
                If arr(j) < arr(i) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next
        Next
    End If
    SortArrayAlphabetical = arr
End Function

Function SortArrayByValue(dict, arr)
    ' Ordena array com base nos valores do dicionário (decrescente)
    If IsArray(arr) Then
        Dim i, j, temp
        For i = 0 To UBound(arr)
            For j = i + 1 To UBound(arr)
                If dict(arr(j)) > dict(arr(i)) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next
        Next
    End If
    SortArrayByValue = arr
End Function

Function SortArrayByVGV(dict, arr)
    ' Ordena array com base no VGV total (decrescente)
    If IsArray(arr) Then
        Dim i, j, temp
        For i = 0 To UBound(arr)
            For j = i + 1 To UBound(arr)
                Set infoI = dict(arr(i))
                Set infoJ = dict(arr(j))
                
                If infoJ("TotalVGV") > infoI("TotalVGV") Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next
        Next
    End If
    SortArrayByVGV = arr
End Function

Function GetDataGraficoVendas()
    ' Retorna dados formatados para o gráfico de vendas
    Dim dados, i
    dados = ""
    
    If Not vendasPorMesDetalhado Is Nothing Then
        For i = 1 To 12
            If vendasPorMesDetalhado.Exists(CStr(i)) Then
                dados = dados & vendasPorMesDetalhado(CStr(i))(0) & ", "
            Else
                dados = dados & "0, "
            End If
        Next
        ' Remove a última vírgula e espaço
        If Len(dados) > 0 Then
            dados = Left(dados, Len(dados) - 2)
        End If
    Else
        dados = "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
    End If
    
    GetDataGraficoVendas = dados
End Function

Function GetDataGraficoQuantidade()
    ' Retorna dados formatados para o gráfico de quantidade
    Dim dados, i
    dados = ""
    
    If Not vendasPorMesQuantidade Is Nothing Then
        For i = 1 To 12
            If vendasPorMesQuantidade.Exists(CStr(i)) Then
                dados = dados & vendasPorMesQuantidade(CStr(i)) & ", "
            Else
                dados = dados & "0, "
            End If
        Next
        ' Remove a última vírgula e espaço
        If Len(dados) > 0 Then
            dados = Left(dados, Len(dados) - 2)
        End If
    Else
        dados = "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
    End If
    
    GetDataGraficoQuantidade = dados
End Function

' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>