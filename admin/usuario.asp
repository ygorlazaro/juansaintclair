<!--#include file="../includes/usuarios.asp"-->
<%
set rsusuarios = Usuarios
Acao 	= Request.Form("acao")

Sub UsuariosCadastrados()
	for i = 0 to rsusuarios.recordcount -1
		response.write("<a href='alterarusuario.asp?id=" & rsusuarios("id_usuario") & "'>" & rsusuarios("usuario") & "</a>")
		response.write("<br/>")
	next
End Sub 

'If Acao = "submit" Then
'	set rstemp = UsuarioFLogin(txtLogin.Text)
'	if txtLogin.text = "" or txtsenha.text = "" then
'		response.write "Todos os campos são obrigatórios"
'	else if rstemp.RecordCount >0 Then
'		Response.Write "Este usuário já existe"
'	else
'		InserirUsuario txtLogin.Text, txtSenha.Text
'		Response.Write "Usuário cadastrado"
'	end if
'End If

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
  </head>
  <body>
    <form id="frm_login" action="usuario.asp" method="post">
      <input type="hidden" name="acao" value="submit" />
      <div id="usuarioscadastrados">
         Usuários cadastrados: 
        <br/>
        <% 

UsuariosCadastrados() 
        %>
<hr/>
      </div>
       Cadastrar novo usuário: 
      <br/>
      <div>
        <table border="1">
          <tbody>
            <tr>
              <td>
                 Usuário 
              </td>
              <td>
                <input name="txtLogin" maxlength="15" />
              </td>
            </tr>
            <tr>
              <td>
                 Senha 
              </td>
              <td>
                <input name="txtLogin" maxlength="15" type="password"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" style="text-align: right">
                <input  type="submit" name="Confirmar" value="Confirmar" class="botao_confirmar" >
	              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </form>
  </body>
</html>
