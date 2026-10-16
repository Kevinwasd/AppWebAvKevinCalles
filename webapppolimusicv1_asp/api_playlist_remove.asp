<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/playlists_logic.asp"-->
<%
' =====================================================================
'  PoliMusic - Quitar una cancion de una playlist
'  POST playlist=<ID> & song=<ID>  ->  { "ok": true, "msg": "..." }
' =====================================================================
Dim nPl, nSong, sOk

Response.ContentType = "application/json"
Response.CacheControl = "no-cache"
Response.CharSet = "utf-8"

nPl   = ParamInt("playlist", 0)
nSong = ParamInt("song", 0)
sOk   = "false"

If IsLoggedIn() And nPl > 0 And nSong > 0 Then
    If PlaylistCanEdit(nPl) Then
        PlaylistRemoveSong nPl, nSong
        sOk = "true"
    End If
End If

If sOk = "true" Then
    Response.Write "{""ok"":true,""msg"":" & JsonEnc("La cancion fue quitada de la playlist.") & "}"
Else
    Response.Write "{""ok"":false,""msg"":" & JsonEnc("No se pudo quitar la cancion de la playlist.") & "}"
End If
%>
