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
<!--#include file="includes/actions.asp"-->
<!--#include file="includes/views/router.asp"-->
<%
' =====================================================================
'  PoliMusic - Pagina principal (shell de la aplicacion)
'  Renderiza la vista solicitada (?view=...) dentro del marco comun.
'  Las vistas se cargan tambien por AJAX desde api_view.asp, de modo que
'  la reproduccion de audio no se interrumpe al navegar.
' =====================================================================
Dim sView, sDenied

RequireLogin()
HandlePostActions()

sView = LCase(Param("view", "home"))
If Len(sView) = 0 Then sView = "home"
If sView = "playlist" And ParamInt("id", 0) <= 0 Then sView = "playlists"
If sView = "album" And ParamInt("id", 0) <= 0 Then sView = "albums"
If sView = "artist" And ParamInt("id", 0) <= 0 Then sView = "artists"

PAGE_TITLE = ViewLabel(sView)
VIEW_NAME  = sView
%>
<!--#include file="includes/layout_top.asp"-->
<%
sDenied = Param("e", "")
If sDenied = "denied" Then
    Response.Write "<div class=""flash flash-error""><span class=""flash-icon""></span>" _
        & "<span>No tiene permisos de administrador para acceder a esa seccion.</span></div>"
End If

RenderView sView, "full"
%>
<!--#include file="includes/layout_bottom.asp"-->
