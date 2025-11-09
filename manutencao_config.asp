<%
' Função para verificar se está em manutenção
Function EstaEmManutencao()
    Dim fs, arquivoManutencao
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    arquivoManutencao = Server.MapPath("manut.txt")
    
    EstaEmManutencao = fs.FileExists(arquivoManutencao)
    
    Set fs = Nothing
End Function

' Função para verificar se é o usuário admin
Function EAdmin()
    EAdmin = (LCase(Session("Usuario")) = "barreto")
End Function

' Redirecionar para manutenção se necessário
If EstaEmManutencao() And Not EAdmin() Then
    If Request.ServerVariables("SCRIPT_NAME") <> "manutencao.asp" Then
        Response.Redirect "manutencao.asp"
        Response.End
    End If
End If
%>