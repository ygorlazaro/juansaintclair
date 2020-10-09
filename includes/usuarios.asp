<!--#include file="../includes/dados.asp"-->
<%
Function UsuarioFLoginSenha(login, senha)
	
        Dim rs
        Dim sql  

        Set Conn = Server.CreateObject("ADODB.Connection")
        Conn.Open ConnString

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT USUARIO, SENHA, ID_USUARIO FROM USUARIOS WHERE USUARIO = '"+ Login + "' AND SENHA = '"+ Senha +"'"

        rs.Open sql, Conn,1,1

	set UsuarioFLoginSenha = rs
End Function

Function UsuarioFLogin(login)
	
        Dim rs
        Dim sql  

        Set Conn = Server.CreateObject("ADODB.Connection")
        Conn.Open ConnString

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT USUARIO, SENHA, ID_USUARIO FROM USUARIOS WHERE USUARIO = '"+ Login + "'"

        rs.Open sql, Conn,1,1

	set UsuarioFLogin = rs
End Function

Function Usuarios()
        Dim rs
        Dim sql  

	Set Conn = Server.CreateObject("ADODB.Connection")
        Conn.Open ConnString

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT USUARIO, SENHA, ID_USUARIO FROM USUARIOS"

        rs.Open sql, Conn,1,1

	set Usuarios = rs
End Function

Sub InserirUsuario(usuario, senha)
	Set Conn = Server.CreateObject("ADODB.Connection")
        Conn.Open ConnString
	
	sql = "INSERT INTO USUARIOS (USUARIO, SENHA) VALUES ('" + usuario + "', '" + senha + "')"
	Conn.Execute sql
End Sub
%>