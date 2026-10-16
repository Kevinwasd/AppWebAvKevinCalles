<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<%
' =====================================================================
'  PoliMusic - Selector de playlist (contenido del modal)
'  GET song=<ID_SONG>  ->  fragmento HTML con las playlists del usuario
' =====================================================================
Dim nSong, rs

RequireLogin()
Response.CacheControl = "no-cache"
Response.ContentType = "text/html"
Response.CharSet = "utf-8"

nSong = ParamInt("song", 0)

Set rs = DbQuery("SELECT pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, " _
    & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS, " _
    & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps2 WHERE ps2.ID_PLAYLIST = pl.ID_PLAYLIST " _
    & " AND ps2.ID_SONG = " & nSong & ") AS YA_ESTA " _
    & "FROM TBL_PLAYLIST pl WHERE pl.ID_USER = " & CurrentUserId() & " ORDER BY pl.TITLE")

If rs.EOF Then
    Response.Write "<div class=""empty"" style=""padding:28px 18px"">" & Icon("playlist") _
        & "<h3>Sin playlists</h3><p>Cree primero una playlist desde la seccion Playlists.</p></div>"
Else
    Response.Write "<div class=""pick-list"">"
    Do While Not rs.EOF
        Response.Write "<button type=""button"" class=""pick-item"" data-pick-playlist=""" _
            & CLng(rs("ID_PLAYLIST")) & """ data-song=""" & nSong & """>"
        Response.Write "<img src=""" & H(ImgUrl(rs("PHOTO"), "playlist", rs("ID_PLAYLIST"), rs("TITLE"))) _
            & """ alt=""" & H(rs("TITLE")) & """>"
        Response.Write "<span style=""flex:1""><b>" & H(rs("TITLE")) & "</b><br>" _
            & "<span class=""small muted"">" & FmtNum(rs("N_SONGS")) & " canciones"
        If CLng(NumOr(rs("YA_ESTA"), 0)) > 0 Then
            Response.Write " &middot; ya incluida"
        End If
        Response.Write "</span></span>"
        If CLng(NumOr(rs("YA_ESTA"), 0)) > 0 Then
            Response.Write "<span class=""badge gray"">Incluida</span>"
        Else
            Response.Write Icon("plus")
        End If
        Response.Write "</button>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End If

rs.Close
Set rs = Nothing
%>
