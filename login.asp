<%
Acao 	= Request.Form("acao")
If Acao = "login" Then

	Login = Request.Form("txtLogin")
	Senha = Request.Form("txtSenha")
		
	If Login = "" Or Senha = "" Then
		Response.Write "Os dois campos precisam ser preenchidos."
	Else
        Dim adoCon  
        Dim rsGuestbook
        Dim strSQL  

        Set adoCon = Server.CreateObject("ADODB.Connection")
        adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("dados\blog.mdb")     

        Set rsGuestbook = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT USUARIO, SENHA, ID_USUARIO FROM USUARIOS WHERE USUARIO = '"+ Login + "' AND SENHA = '"+ Senha +"' AND ATIVO = 'TRUE'"

        rsGuestbook.Open strSQL, adoCon,1,1
           
        If rsGuestbook.RecordCount = 1 Then
		    Response.Redirect "/admin/index.asp"
		Else
		    Response.Write "Usuário ou senha inválidos"
		End If 
		rsGuestbook.Close
	End If
End If 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd" 
>
<html>
<head>
    <title>Juan Saint' Clair</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="resource-type" content="document" />
<meta name="classification" content="Internet" />
<meta name="keywords" content="" />
<meta name="robots" content="index,no-follow" />
<meta name="distribution" content="Global" />
<meta name="language" content="pt-br" />
<meta name="title" content="" />
<meta name="author" content="Pysi" />

<!--<link rel="shortcut icon" href="/files/favicon.ico" type="image/x-icon" />-->
<link rel="stylesheet" type="text/css" href="css/estrutura.css" />
</head>
<body>
<div id="meio">
  <div id="meio_img"><img src="imagens/logo.JPG" alt="logo & cia" /></div>
<form id="frm_login" action="login.asp" method="post">
      <input type="hidden" name="acao" value="login" />
    <table cellspacing="1" cellpadding="1">
            <tr>
                <td align="center">
                    <b>Login:</b>
                </td>
                <td>
                    <input name="txtLogin" maxlength="15" />
                </td>
            </tr>
            <tr>
                <td>
                    <b>Senha:</b>
                </td>
                <td>
                    <input name="txtSenha" maxlength="15" type="password" />
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: right">
                  <input type="submit" name="Confirmar" value="Confirmar" class="botao_confirmar" />
                </td>
            </tr>
    </table>
            
    </form>
    </div>
</body>
</html>
