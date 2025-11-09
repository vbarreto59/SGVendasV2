<%@ Language="VBSCRIPT" %>
<!--#include file="manutencao_config.asp"-->
<%
' Verificar se é admin
If Not EAdmin() Then
    Response.Status = "403 Forbidden"
    Response.Write "Acesso negado"
    Response.End
End If

' Criar arquivo de manutenção
Dim fs, arquivo
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arquivo = fs.CreateTextFile(Server.MapPath("manut.txt"), True)
arquivo.Write "Manutenção ativada em " & Now()
arquivo.Close
Set arquivo = Nothing
Set fs = Nothing

' Redirecionar
if UCase(Session("Usuario")) = "BARRETO" then
    Response.Redirect("manut_menu.asp")
else    
   Response.Redirect "manutencao.asp"
end if   
%>