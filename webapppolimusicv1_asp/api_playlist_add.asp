<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/playlists_logic.asp"-->
<%
' =====================================================================
'  PoliMusic - Agregar una cancion a una playlist
'  POST playlist=<ID> & song=<ID>  ->  { "ok": true, "msg": "..." }
' =====================================================================
Dim nPl, nSong, sMsg, sOk

Response.ContentType = "application/json"
Response.CacheControl = "no-cache"
Response.CharSet = "utf-8"

nPl   = ParamInt("playlist", 0)
nSong = ParamInt("song", 0)
sMsg  = "Debe iniciar sesion para modificar playlists."

If IsLoggedIn() Then
    sMsg = PlaylistAddSong(nPl, nSong)
    If Len(sMsg) = 0 Then
        sMsg = "La cancion fue agregada a la playlist."
    End If
End If

If Len(sMsg) = 0 Or sMsg = "La cancion fue agregada a la playlist." Then
    sOk = "true"
Else
    sOk = "false"
End If

Response.Write "{""ok"":" & sOk & ",""msg"":" & JsonEnc(sMsg) & "}"
%>
