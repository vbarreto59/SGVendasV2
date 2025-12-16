<!-- ###################################### -->
<!-- SISTEMA: ToDoList                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: ELMVOFOOMN          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
      <title>ToDoList - Login</title>
      <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
      <meta http-equiv="X-UA-Compatible" content="IE=edge">  
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
      <link href="css/login.css" rel="stylesheet" type="text/css" />  
      <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
      <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>      
</head>
<body>
<%
	set objMail = server.createobject("CDONTS.NewMail")

	objMail.From = "gabnet@gabnetweb.com.br"
	objMail.To = "valterpb@hotmail.com"

	objMail.Subject = "SGVENDAS 3.0 - Backup ToccaMDB" 
	objMail.MailFormat = 0
	objMail.Attachfile "E:\ClientHome\gabnetweb.com.br\bdados\SunSales.mdb",   "SunSales.mdb"
	Response.Write "SunSales OK! <br> <br>"	
	objMail.Attachfile "E:\ClientHome\gabnetweb.com.br\bdados\ImobVendas.mdb", "ImobVendas.mdb"
	Response.Write "ImobVendas OK! <br> <br>"	
	objMail.Attachfile "E:\ClientHome\gabnetweb.com.br\bdados\SunnyLog.mdb",   "SunnyLog.mdb"
	Response.Write "SunnyLog OK! <br> <br>"	
	objMail.Send

	Response.Write "Backup Tocca. Mensagem Enviada! <br> <br>"
	set objMail = Nothing
%>
 <a href="menu.asp" class="btn btn-info">Home</a>
 
</body>