<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Response.Buffer = True
Response.ContentType = "text/html"
Response.Charset = "UTF-8"
Response.Write "<html><head><meta charset='UTF-8'></head><body>"
Randomize Timer

If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "<p style='color:red'>Erro: Conexões não configuradas (StrConn / StrConnSales).</p>"
    Response.End
End If

Dim connMain, connSales
Set connMain  = Server.CreateObject("ADODB.Connection")
connMain.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Function SqlQuote(s)
    If IsNull(s) Then
        SqlQuote = "NULL"
    Else
        SqlQuote = "'" & Replace(CStr(s), "'", "''") & "'"
    End If
End Function

Function SqlDate(d)
    SqlDate = "'" & Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "'"
End Function

Function SqlNum(n)
    If IsNull(n) Then
        SqlNum = "0"
    Else
        ' Converte para número, formata com duas casas decimais sem separadores de milhares, e substitui vírgula por ponto
        SqlNum = Replace(FormatNumber(CDbl(n), 2, -1, -1, 0), ",", ".")
    End If
End Function

Function GetRowsOrNull(sql, oConn)
    Dim rs, arr
    On Error Resume Next
    Set rs = oConn.Execute(sql)
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red'>Erro executando SQL: " & Server.HTMLEncode(sql) & "<br>Detalhe: " & Err.Description & "</div>"
        Err.Clear
        GetRowsOrNull = Null
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close : Set rs = Nothing
        GetRowsOrNull = Null
        Exit Function
    End If

    arr = rs.GetRows()
    rs.Close : Set rs = Nothing
    GetRowsOrNull = arr
End Function

Function RandIndex(maxVal)
    RandIndex = Int((maxVal + 1) * Rnd)
End Function

Function RandomDate(y, m)
    Dim lastDay
    lastDay = Day(DateSerial(y, m + 1, 0))
    RandomDate = DateSerial(y, m, Int(lastDay * Rnd) + 1)
End Function

Function PickUnidadeCodigo()
    Dim letras, a, b, num
    letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    a = Mid(letras, Int(26 * Rnd) + 1, 1)
    b = Mid(letras, Int(26 * Rnd) + 1, 1)
    num = Int(900 * Rnd) + 100
    PickUnidadeCodigo = "AP-" & a & b & "-" & num
End Function

Function GerarPremio()
    ' Gera prêmios entre R$ 500 e R$ 1000, sem decimais, com 70% de chance de ser R$ 0
    If Rnd() > 0.3 Then
        GerarPremio = 0
    Else
        ' Gera múltiplos de 50 entre 500 e 1000 (500, 550, 600, ..., 1000)
        Dim baseVal
        baseVal = (Int((1000 - 500) / 50) + 1)*100
        GerarPremio = baseVal
    End If
End Function

Sub FlushLine(msg)
    Response.Write msg & "<br>" & vbCrLf
    Response.Flush
End Sub

Dim empreendimentos, diretorias, corretores
empreendimentos = GetRowsOrNull("SELECT empreend_id, NomeEmpreendimento, ComissaoVenda FROM empreendimento ORDER BY NomeEmpreendimento", connMain)
diretorias      = GetRowsOrNull("SELECT DiretoriaID, NomeDiretoria FROM Diretorias WHERE NomeDiretoria IN ('VILARIM','VERAS') ORDER BY NomeDiretoria", connMain)
corretores      = GetRowsOrNull("SELECT UserId, Nome FROM Usuarios WHERE Funcao='Corretor' AND Nome IS NOT NULL AND Nome <> '' ORDER BY Nome", connMain)

If IsNull(empreendimentos) Then
    FlushLine "<span style='color:red'>Nenhum Empreendimento encontrado.</span>"
    Response.End
End If
If IsNull(diretorias) Then
    FlushLine "<span style='color:red'>Nenhuma Diretoria encontrada.</span>"
    Response.End
End If
If IsNull(corretores) Then
    FlushLine "<span style='color:red'>Nenhum Corretor encontrado.</span>"
    Response.End
End If

Dim maxEmp, maxDir, maxCor
maxEmp = UBound(empreendimentos, 2)
maxDir = UBound(diretorias, 2)
maxCor = UBound(corretores, 2)

FlushLine "<strong>Iniciando geração de massa para 2025…</strong>"
FlushLine "Empreendimentos: " & (maxEmp + 1) & " | Diretorias: " & (maxDir + 1) & " | Corretores: " & (maxCor + 1)

Dim anoAlvo: anoAlvo = 2025
Dim qtdPorMes: qtdPorMes = 1

Dim percDirDefault: percDirDefault = 5
Dim percGerDefault: percGerDefault = 10
Dim percCorDefault: percCorDefault = 35

Dim usuarioAtual
usuarioAtual = Session("Usuario")
If Trim(usuarioAtual & "") = "" Then usuarioAtual = "SISTEMA"

Dim mes, k
For mes = 1 To 12
    FlushLine "<hr><strong>Mês " & Right("0" & mes, 2) & "/" & anoAlvo & "</strong>"

    For k = 1 To qtdPorMes

        Dim idxEmp, idEmp, nomeEmp, comissaoVendaEmp
        idxEmp = RandIndex(maxEmp)
        idEmp = empreendimentos(0, idxEmp)
        nomeEmp = empreendimentos(1, idxEmp)
        comissaoVendaEmp = empreendimentos(2, idxEmp)
        If IsNull(comissaoVendaEmp) Or comissaoVendaEmp = "" Then comissaoVendaEmp = 3

        Dim idxDir, idDir, nomeDir, userIdDir, nomeDirCompleto
        idxDir = RandIndex(maxDir)
        idDir = diretorias(0, idxDir)
        nomeDir = diretorias(1, idxDir)
        ' Buscar UserId e Nome da Diretoria
        Dim dirArr
        dirArr = GetRowsOrNull("SELECT DiretoriaID, NomeDiretoria, UserId, Nome FROM Diretorias WHERE DiretoriaID=" & idDir, connMain)
        If IsNull(dirArr) Then
            userIdDir = 0
            nomeDirCompleto = "Não aplicável"
        Else
            userIdDir = dirArr(2, 0)
            If IsNull(userIdDir) Then userIdDir = 0
            nomeDirCompleto = dirArr(3, 0)
            If IsNull(nomeDirCompleto) Then nomeDirCompleto = "Não aplicável"
        End If

        Dim gerArr, idGer, nomeGer, userIdGer, nomeGerCompleto
        gerArr = GetRowsOrNull("SELECT GerenciaID, NomeGerencia, UserId, Nome FROM Gerencias WHERE DiretoriaID=" & idDir & " ORDER BY NomeGerencia", connMain)
        If IsNull(gerArr) Then
            idGer = 0
            nomeGer = "Não aplicável"
            userIdGer = 0
            nomeGerCompleto = "Não aplicável"
        Else
            Dim maxGer, idxGer
            maxGer = UBound(gerArr, 2)
            idxGer = RandIndex(maxGer)
            idGer = gerArr(0, idxGer)
            nomeGer = gerArr(1, idxGer)
            userIdGer = gerArr(2, idxGer)
            If IsNull(userIdGer) Then userIdGer = 0
            nomeGerCompleto = gerArr(3, idxGer)
            If IsNull(nomeGerCompleto) Then nomeGerCompleto = "Não aplicável"
        End If

        Dim idxCor, idCor, nomeCor
        idxCor = RandIndex(maxCor)
        idCor = corretores(0, idxCor)
        nomeCor = corretores(1, idxCor)

        Dim valorUnidade, m2, dataVenda, unidadeCod
        valorUnidade = Int((500000 - 200000 + 1) * Rnd + 200000)
        m2 = Int((50 - 20 + 1) * Rnd + 20)
        dataVenda = RandomDate(anoAlvo, mes)
        unidadeCod = PickUnidadeCodigo()

        Dim diaVenda, mesVenda, anoVenda, trimestre
        diaVenda = Day(dataVenda)
        mesVenda = Month(dataVenda)
        anoVenda = Year(dataVenda)
        trimestre = Int((mesVenda - 1) / 3) + 1

        Dim percTotal, comissaoTotal, valorDir, valorGer, valorCor
        percTotal = CDbl(comissaoVendaEmp)
        comissaoTotal = CDbl(valorUnidade) * (percTotal / 100)

        valorDir = comissaoTotal * (percDirDefault / 100)
        valorGer = comissaoTotal * (percGerDefault / 100)
        valorCor = comissaoTotal * (percCorDefault / 100)

        ' Gerar prêmios entre 500 e 1000, sem decimais
        Dim premioDiretoria, premioGerencia, premioCorretor
        premioDiretoria = 1000
        premioGerencia = 1100
        premioCorretor = 1200


        ' Log de depuração para verificar valores antes da inserção
        FlushLine "Depuração: valorUnidade = " & valorUnidade & ", percTotal = " & percTotal & ", comissaoTotal = " & comissaoTotal & _
                  ", valorDir = " & valorDir & ", valorGer = " & valorGer & ", valorCor = " & valorCor & _
                  ", Prêmios: Dir=" & premioDiretoria & " Ger=" & premioGerencia & " Cor=" & premioCorretor

        Dim sqlVendas, vendaId
        sqlVendas = "INSERT INTO Vendas (" & _
                    "Empreend_ID, NomeEmpreendimento, Unidade, UnidadeM2, " & _
                    "Corretor, CorretorId, ValorUnidade, ComissaoPercentual, ValorComissaoGeral, " & _
                    "DataVenda, DiaVenda, MesVenda, AnoVenda, Trimestre, " & _
                    "Obs, Usuario, DiretoriaId, Diretoria, UserIdDiretoria, NomeDiretor, " & _
                    "GerenciaId, Gerencia, UserIdGerencia, NomeGerente, " & _
                    "ComissaoDiretoria, ValorDiretoria, ComissaoGerencia, ValorGerencia, ComissaoCorretor, ValorCorretor, " & _
                    "PremioDiretoria, PremioGerencia, PremioCorretor) VALUES (" & _
                    idEmp & ", " & SqlQuote(nomeEmp) & ", " & SqlQuote(unidadeCod) & ", " & SqlNum(m2) & ", " & _
                    SqlQuote(nomeCor) & ", " & idCor & ", " & SqlNum(valorUnidade) & ", " & SqlNum(percTotal) & ", " & SqlNum(comissaoTotal) & ", " & _
                    SqlDate(dataVenda) & ", " & diaVenda & ", " & mesVenda & ", " & anoVenda & ", " & trimestre & ", " & _
                    SqlQuote("Massa 2025 - auto") & ", " & SqlQuote(usuarioAtual) & ", " & idDir & ", " & SqlQuote(nomeDir) & ", " & _
                    userIdDir & ", " & SqlQuote(nomeDirCompleto) & ", " & idGer & ", " & SqlQuote(nomeGer) & ", " & _
                    userIdGer & ", " & SqlQuote(nomeGerCompleto) & ", " & _
                    SqlNum(percDirDefault) & ", " & SqlNum(valorDir) & ", " & SqlNum(percGerDefault) & ", " & SqlNum(valorGer) & ", " & SqlNum(percCorDefault) & ", " & SqlNum(valorCor) & ", " & _
                    SqlNum(premioDiretoria) & ", " & SqlNum(premioGerencia) & ", " & SqlNum(premioCorretor) & ")"

        On Error Resume Next
        connSales.Execute sqlVendas
        
        If Err.Number = 0 Then
            ' Obter o ID da venda recém-inserida
            Set rsLastID = connSales.Execute("SELECT MAX(ID) AS NewID FROM Vendas")
            If Not rsLastID.EOF Then 
                vendaId = rsLastID("NewID")
            Else
                vendaId = 0
            End If
            rsLastID.Close
            Set rsLastID = Nothing
            
            FlushLine "[" & Right("0" & diaVenda, 2) & "/" & Right("0" & mesVenda, 2) & "/" & anoVenda & "] " & _
                      "Empreendimento: <strong>" & nomeEmp & "</strong> | Diretoria: " & nomeDir & _
                      " | Gerência: " & nomeGer & _
                      " | Corretor: " & nomeCor & _
                      " | Valor: R$ " & FormatNumber(valorUnidade, 2) & " | m²: " & m2 & " → <span style='color:green'>OK (ID: " & vendaId & ")</span>"

            ' ============ INSERIR NA TABELA COMISSOES_A_PAGAR ============
            If vendaId > 0 Then
                Dim sqlComissaoPagar
                sqlComissaoPagar = "INSERT INTO COMISSOES_A_PAGAR (" & _
                                   "ID_Venda, Empreendimento, Unidade, DataVenda, " & _
                                   "UserIdDiretoria, NomeDiretor, UserIdGerencia, NomeGerente, UserIdCorretor, NomeCorretor, " & _
                                   "PercDiretoria, ValorDiretoria, PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, " & _
                                   "TotalComissao, StatusPagamento, Usuario, " & _
                                   "PremioDiretoria, PremioGerencia, PremioCorretor) VALUES (" & _
                                   vendaId & ", " & SqlQuote(nomeEmp) & ", " & SqlQuote(unidadeCod) & ", " & SqlDate(dataVenda) & ", " & _
                                   userIdDir & ", " & SqlQuote(nomeDirCompleto) & ", " & userIdGer & ", " & SqlQuote(nomeGerCompleto) & ", " & idCor & ", " & SqlQuote(nomeCor) & ", " & _
                                   SqlNum(percDirDefault) & ", " & SqlNum(valorDir) & ", " & SqlNum(percGerDefault) & ", " & SqlNum(valorGer) & ", " & SqlNum(percCorDefault) & ", " & SqlNum(valorCor) & ", " & _
                                   SqlNum(comissaoTotal) & ", 'Pendente', " & SqlQuote(usuarioAtual) & ", " & _
                                   SqlNum(premioDiretoria) & ", " & SqlNum(premioGerencia) & ", " & SqlNum(premioCorretor) & ")"
                
                connSales.Execute sqlComissaoPagar
                
                If Err.Number = 0 Then
                    FlushLine "<span style='color:green; margin-left: 20px;'>✓ Inserido em COMISSOES_A_PAGAR com sucesso!</span>"
                Else
                    FlushLine "<span style='color:red; margin-left: 20px;'>✗ Erro inserindo em COMISSOES_A_PAGAR: " & Err.Description & "</span>"
                    Err.Clear
                End If
            End If
            ' ============ FIM DA INSERÇÃO EM COMISSOES_A_PAGAR ============

            ' Verificação dos dados inseridos na tabela Vendas
            Dim rsVerify, sqlVerify
            sqlVerify = "SELECT TOP 1 NomeEmpreendimento, UserIdDiretoria, NomeDiretor, UserIdGerencia, NomeGerente, " & _
                        "ValorUnidade, ComissaoPercentual, ValorComissaoGeral, ValorDiretoria, ValorGerencia, ValorCorretor, " & _
                        "PremioDiretoria, PremioGerencia, PremioCorretor " & _
                        "FROM Vendas WHERE ID = " & vendaId
            Set rsVerify = connSales.Execute(sqlVerify)
            If Not rsVerify.EOF Then
                FlushLine "<span style='color:blue; margin-left: 20px;'>Verificação Vendas: " & _
                          "ValorUnidade = " & rsVerify("ValorUnidade") & _
                          ", ComissaoTotal = " & rsVerify("ValorComissaoGeral") & _
                          ", Prêmios = [" & rsVerify("PremioDiretoria") & "," & rsVerify("PremioGerencia") & "," & rsVerify("PremioCorretor") & "]</span>"
            Else
                FlushLine "<span style='color:red; margin-left: 20px;'>Erro: Não foi possível verificar os dados inseridos.</span>"
            End If
            rsVerify.Close
            Set rsVerify = Nothing

        Else
            FlushLine "<span style='color:red'>Erro inserindo Vendas: " & Err.Description & "</span><br><code>" & Server.HTMLEncode(sqlVendas) & "</code>"
            Err.Clear
        End If
        On Error GoTo 0

    Next
Next

FlushLine "<hr><strong>Inserção de massa concluída!</strong>"

If Not connMain Is Nothing Then If connMain.State = 1 Then connMain.Close : Set connMain = Nothing
If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close : Set connSales = Nothing

' ======================= ATUALIZAÇÃO FINAL DO BANCO DE DADOS =======================

Set connMain  = Server.CreateObject("ADODB.Connection")
connMain.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

On Error Resume Next

' --- UPDATE Diretoria ---
sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) " & _
             "SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute sqlUpdate1
If Err.Number = 0 Then
    FlushLine "<span style='color:green'>Atualização de Diretoria concluída com sucesso!</span>"
    Dim rsVerifyDir, sqlVerifyDir
    sqlVerifyDir = "SELECT TOP 1 UserIdDiretoria, NomeDiretor FROM Vendas WHERE AnoVenda = " & anoAlvo
    Set rsVerifyDir = connSales.Execute(sqlVerifyDir)
    If Not rsVerifyDir.EOF Then
        FlushLine "Verificação Diretoria: UserIdDiretoria = " & rsVerifyDir("UserIdDiretoria") & ", NomeDiretor = " & rsVerifyDir("NomeDiretor")
    Else
        FlushLine "<span style='color:red'>Erro: Não foi possível verificar UserIdDiretoria/NomeDiretor.</span>"
    End If
    rsVerifyDir.Close
    Set rsVerifyDir = Nothing
Else
    FlushLine "<span style='color:red'>Erro na atualização de Diretoria: " & Err.Description & "</span>"
    Err.Clear
End If

' --- UPDATE Gerência ---
sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) " & _
             "SET Vendas.NomeGerente = [Gerencias].[Nome], Vendas.UserIdGerencia = [Gerencias].[UserId];"
connSales.Execute sqlUpdate2
If Err.Number = 0 Then
    FlushLine "<span style='color:green'>Atualização de Gerência concluída com sucesso!</span>"
    Dim rsVerifyGer, sqlVerifyGer
    sqlVerifyGer = "SELECT TOP 1 UserIdGerencia, NomeGerente FROM Vendas WHERE AnoVenda = " & anoAlvo
    Set rsVerifyGer = connSales.Execute(sqlVerifyGer)
    If Not rsVerifyGer.EOF Then
        FlushLine "Verificação Gerência: UserIdGerencia = " & rsVerifyGer("UserIdGerencia") & ", NomeGerente = " & rsVerifyGer("NomeGerente")
    Else
        FlushLine "<span style='color:red'>Erro: Não foi possível verificar UserIdGerencia/NomeGerente.</span>"
    End If
    rsVerifyGer.Close
    Set rsVerifyGer = Nothing
Else
    FlushLine "<span style='color:red'>Erro na atualização de Gerência: " & Err.Description & "</span>"
    Err.Clear
End If

' --- UPDATE Corretor ---
sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) " & _
                    "SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute sqlUpdateCorretor

' --- UPDATE Semestre ---
sqlUpdateSemestre = "UPDATE Vendas " & _
                    "SET Semestre = SWITCH(" & _
                    "    Trimestre IN (1, 2), 1, " & _
                    "    Trimestre IN (3, 4), 2" & _
                    ") " & _
                    "WHERE Trimestre IS NOT NULL;"
connSales.Execute sqlUpdateSemestre

' --- Atualizar COMISSOES_A_PAGAR com dados das Diretorias/Gerencias ---
sqlUpdateComissaoDir = "UPDATE (COMISSOES_A_PAGAR INNER JOIN [;DATABASE=" & dbSunnyPath & "].Diretorias ON COMISSOES_A_PAGAR.UserIdDiretoria = Diretorias.UserId) " & _
                       "SET COMISSOES_A_PAGAR.NomeDiretor = Diretorias.Nome;"
connSales.Execute sqlUpdateComissaoDir

sqlUpdateComissaoGer = "UPDATE (COMISSOES_A_PAGAR INNER JOIN [;DATABASE=" & dbSunnyPath & "].Gerencias ON COMISSOES_A_PAGAR.UserIdGerencia = Gerencias.UserId) " & _
                       "SET COMISSOES_A_PAGAR.NomeGerente = Gerencias.Nome;"
connSales.Execute sqlUpdateComissaoGer

' --- Verificação de erros geral ---
If Err.Number <> 0 Then
    Response.Write "<span style='color:red'>Ocorreu um erro ao atualizar o banco de dados: " & Err.Description & "</span><br>"
Else
    FlushLine "<span style='color:green'>Atualização final do banco de dados concluída com sucesso!</span>"
End If
On Error GoTo 0

' ======================= FIM DA ATUALIZAÇÃO =======================

If Not connMain Is Nothing Then If connMain.State = 1 Then connMain.Close : Set connMain = Nothing
If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close : Set connSales = Nothing

Response.Write "</body></html>"
%>