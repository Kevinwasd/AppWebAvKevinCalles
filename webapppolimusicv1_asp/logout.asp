<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' =====================================================================
'  PoliMusic - Cierre de sesion
' =====================================================================
LogoutUser
Response.Redirect U("login.asp")
%>
