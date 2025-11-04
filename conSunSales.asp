<%
On Error Resume Next ' Habilita o tratamento de erros

' Constantes para strings de conexão
vDbLocalhost = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\SGVendas\db\SunSales.mdb"
vDbProduction = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\ClientHome\gabnetweb.com.br\bdados\SunSales.mdb;"


Const LOCAL_CONNECTION_STRINGx = ""

Const PRODUCTION_CONNECTION_STRINGx = ""

' Define a string de conexão com base no ambiente
Dim StrConnSales, connectionStringSales
Set StrConnSales = Server.CreateObject("ADODB.Connection")

If Request.ServerVariables("SERVER_NAME") = "localhost" Then
    connectionStringSales = vDbLocalhost
Else
    connectionStringSales = vDbProduction
End If

'response.Write connectionString
'response.end()

' Tenta abrir a conexão
StrConnSales.Open connectionStringSales

' Verifica se houve erro na conexão
If Err.Number <> 0 Then
    Response.Write("Erro ao conectar ao banco de dados (conexao.asp). Por favor, tente novamente mais tarde. Erro: " & Err.Number )
    StrConnSales.Close ' Fechar a conexão em caso de erro
    Set StrConnSales = Nothing ' Limpar o objeto de conexão
    Response.End()
End If

' Aqui pode ir o código que interage com o banco de dados...

' Fechamento da conexão
If Not StrConnSales Is Nothing Then
    StrConnSales.Close
    Set StrConnSales = Nothing ' Limpar o objeto de conexão
End If

On Error GoTo 0 ' Desabilita o tratamento de erros
StrConnSales = connectionStringSales
%>
