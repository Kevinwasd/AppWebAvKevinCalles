<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' =====================================================================
'  PoliMusic - Punto de entrada
'  Redirige a la pagina principal si hay sesion o al login si no la hay.
' =====================================================================
If IsLoggedIn() Then
    Response.Redirect U("home.asp")
Else
    Response.Redirect U("login.asp")
End If
%>
