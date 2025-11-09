<%@ Language="VBSCRIPT" %>
<!--#include file="manutencao_config.asp"-->
<%
' Verificar se é admin
If Not EAdmin() Then
    Response.Status = "403 Forbidden"
    Response.Write "Acesso negado"
    Response.End
End If

' Remover arquivo de manutenção
Dim fs
Set fs = Server.CreateObject("Scripting.FileSystemObject")

If fs.FileExists(Server.MapPath("manut.txt")) Then
    fs.DeleteFile Server.MapPath("manut.txt")
End If

Set fs = Nothing

' Redirecionar para a página inicial
Response.Redirect "manut_menu.asp"
%>