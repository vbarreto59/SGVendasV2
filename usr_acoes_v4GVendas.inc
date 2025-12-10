<%
Function UserFuncao()
  Dim strAcoesPermitidas
  strAcoesPermitidas = "" ' Inicializa a variável

  ' Verifica se o UserId na sessão existe e é um número válido
  If Session("UserId") <> "" And IsNumeric(Session("UserId")) Then
    Dim objConn, objRS, strSQL, intUserId

    ' Recupera o UserId da sessão
    intUserId = Session("UserId")

    ' Cria e abre a conexão com o banco de dados
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open StrConn ' Certifique-se que 'StrConn' está definida com a sua string de conexão

    ' Define a instrução SQL para buscar as ações permitidas do usuário
    strSQL = "SELECT * FROM Usuarios WHERE UserId = " & intUserId

    ' Cria e executa o Recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn

    ' Verifica se algum registro foi encontrado
    If Not objRS.EOF Then
      strAcoesPermitidas = objRS("Funcao")
    End If

    ' Fecha e libera os objetos ADO
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
  Else
    ' Caso o UserId na sessão seja inválido ou não exista, retorna uma string vazia ou um valor indicativo de erro
    strAcoesPermitidas = "" ' Ou poderia ser algo como "ERRO: Usuário não logado"
  End If

  ' Retorna a string de ações permitidas
  UserFuncao = strAcoesPermitidas
End Function


Function AcoesPermitidas()
  Dim strAcoesPermitidas
  strAcoesPermitidas = "" ' Inicializa a variável

  ' Verifica se o UserId na sessão existe e é um número válido
  If Session("UserId") <> "" And IsNumeric(Session("UserId")) Then
    Dim objConn, objRS, strSQL, intUserId

    ' Recupera o UserId da sessão
    intUserId = Session("UserId")

    ' Cria e abre a conexão com o banco de dados
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open StrConn ' Certifique-se que 'StrConn' está definida com a sua string de conexão

    ' Define a instrução SQL para buscar as ações permitidas do usuário
    strSQL = "SELECT * FROM Usuarios WHERE UserId = " & intUserId

    ' Cria e executa o Recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn

    ' Verifica se algum registro foi encontrado
    If Not objRS.EOF Then
      ' Atribui o valor do campo 'ACOES' à variável de retorno
      strAcoesPermitidas = objRS("Acoes")
    End If

    ' Fecha e libera os objetos ADO
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
  Else
    ' Caso o UserId na sessão seja inválido ou não exista, retorna uma string vazia ou um valor indicativo de erro
    strAcoesPermitidas = "" ' Ou poderia ser algo como "ERRO: Usuário não logado"
  End If

  ' Retorna a string de ações permitidas
  AcoesPermitidas = strAcoesPermitidas
  Session("Funcao")
End Function
%>

<%
Function UsuarioPodeEditar()
  Dim podeEditar
  podeEditar = False
  If Session("Acoes") <> "" Then
          If InStr(1, Session("Acoes"), "/EDITAR/", vbTextCompare) > 0 Then
              podeEditar = True
          End If
  Else
      response.redirect("gestao_login.asp")
  End If
  UsuarioPodeEditar = podeEditar
End Function
%>

<%
Function UsuarioGestor()
  Dim podeEditar
  podeEditar = False
  If Session("Acoes") <> "" Then
      If InStr(1, Session("Acoes"), "/GESTOR/", vbTextCompare) > 0 Then
          podeEditar = True
      End If
  Else
      response.redirect("gestao_login.asp")
  End If
  UsuarioGestor = podeEditar
End Function
%>

<%
Function UsuarioAdmin()
  Dim podeEditar
  podeEditar = False
  If Session("Acoes") <> "" Then
          If InStr(1, Session("Acoes"), "/ADMIN/", vbTextCompare) > 0 Then
              podeEditar = True
          End If
          If InStr(1, Session("Acoes"), "/SISTEMA/", vbTextCompare) > 0 Then
              podeEditar = True
          End If
  Else
      response.redirect("gestao_login.asp")
  End If
  UsuarioAdmin = podeEditar
End Function
%>

<%
' =================================================================================
' FUNÇÃO PARA VERIFICAR BLOQUEIO DE EMAIL
' Retorna True se o arquivo 'NoEmail.txt' existir (bloqueado), False caso contrário.
' =================================================================================
Function BloqueioEmail()
    Dim objFSO, strCaminhoArquivo, blnBloqueado
    
    blnBloqueado = False
    
    ' Caminho absoluto para o arquivo de bloqueio
    strCaminhoArquivo = Server.MapPath("NoEmail.txt") 
    
    ' Checa a existência do arquivo NoEmail.txt usando FileSystemObject
    On Error Resume Next ' Inicia tratamento de erro para FSO
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    
    If Err.Number = 0 Then
        If objFSO.FileExists(strCaminhoArquivo) Then
            blnBloqueado = True ' Arquivo de bloqueio encontrado
            Call InserirLog ("Sistema", "EMAIL", "Envio de e-mail bloqueado: NoEmail.txt encontrado.")
        End If
        Set objFSO = Nothing
    End If
    Err.Clear
    On Error GoTo 0 ' Restaura tratamento de erro
    
    If blnBloqueado then
       Session("EnviaEmail") = "NoEmail"
    Else
       Session("EnviaEmail") = "**"
    End If   
    'response.write blnBloqueado
    'response.end 
    BloqueioEmail = blnBloqueado
End Function
%>
