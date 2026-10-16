<%
' =====================================================================
'  Vista: Detalle de album (portada + listado de canciones)
' =====================================================================
Sub ViewAlbum()
    Dim nId, rs, sTitle, nYear, sImage, nSongs, sOrder
    nId = ParamInt("id", 0)
    sOrder = Param("order", "track")

    Set rs = DbQuery("SELECT TOP 1 al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, " _
        & "al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t " _
        & "WHERE t.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
        & "FROM TBL_ALBUM al WHERE al.ID_ALBUM = " & nId)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        EmptyState "Album no encontrado", "El album solicitado no existe o fue eliminado."
        Exit Sub
    End If

    sTitle = rs("TITLE_ALBUM") & ""
    nYear  = NumOr(rs("RELEASE_YEAR"), 0)
    sImage = rs("ALBUM_IMAGE_PATH") & ""
    nSongs = CLng(NumOr(rs("N_SONGS"), 0))
    rs.Close
    Set rs = Nothing
%>
    <div class="detail-head">
        <img class="detail-cover" src="<%=H(ImgUrl(sImage, "album", nId, sTitle))%>" alt="<%=H(sTitle)%>">
        <div class="detail-info">
            <span class="kind">Album</span>
            <h1><%=H(sTitle)%></h1>
            <div class="detail-meta">
                <span><%=FmtNum(nSongs)%> canciones</span>
                <% If nYear > 0 Then %><span class="sep">&middot;</span>
                <span>Publicado en <%=nYear%></span><% End If %>
            </div>
            <div class="detail-actions">
                <button type="button" class="btn btn-primary" data-play-all="#albumTracks">
                    <%=Icon("play")%><span>Reproducir album</span></button>
                <button type="button" class="btn btn-outline" data-play-all="#albumTracks" data-play-shuffle>
                    <%=Icon("shuffle")%><span>Aleatorio</span></button>
                <% If IsAdmin() Then %>
                <a class="btn btn-ghost" href="<%=U("admin_albums.asp?edit=" & nId)%>">
                    <%=Icon("edit")%><span>Editar album</span></a>
                <% End If %>
            </div>
        </div>
    </div>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Lista de canciones</h2>
            </div>
            <div class="flex gap-8">
                <a class="btn btn-sm <%=IIf2(sOrder = "track", "btn-primary", "btn-outline")%>"
                   href="<%=U("home.asp?view=album&id=" & nId & "&order=track")%>" data-nav>Orden del album</a>
                <a class="btn btn-sm <%=IIf2(sOrder = "plays", "btn-primary", "btn-outline")%>"
                   href="<%=U("home.asp?view=album&id=" & nId & "&order=plays")%>" data-nav>Mas escuchadas</a>
            </div>
        </div>
        <%
        If sOrder = "plays" Then
            Set rs = DbQuery(SongListSql("s.ID_SONG IN (SELECT t.ID_SONG FROM TBL_TRACKLIST t " _
                & "WHERE t.ID_ALBUM = " & nId & ")", "ISNULL(s.PLAYS,0) DESC, s.ID_SONG ASC", 0))
        Else
            Set rs = DbQuery(SongListSql("s.ID_SONG IN (SELECT t.ID_SONG FROM TBL_TRACKLIST t " _
                & "WHERE t.ID_ALBUM = " & nId & ")", "s.ID_SONG ASC", 0))
        End If
        SongTable rs, "albumTracks"
        rs.Close
        Set rs = Nothing
        %>
    </section>
<%
End Sub
%>
