<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/upload.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<!--#include file="includes/playlists_logic.asp"-->
<!--#include file="includes/views/router.asp"-->
<%
' =====================================================================
'  PoliMusic - Servicio de vistas parciales
'  Devuelve el HTML de una vista sin el marco de la pagina, para que el
'  cliente lo inserte sin recargar (y sin cortar el audio).
' =====================================================================
RequireLogin()

Response.CacheControl = "no-cache"
Response.Expires = -1
Response.ContentType = "text/html"

RenderView LCase(Param("view", "home")), "fragment"
%>
