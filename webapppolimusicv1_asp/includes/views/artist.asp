<%
' =====================================================================
'  Vista: Detalle de artista (canciones + albumes en los que participa)
' =====================================================================
Sub ViewArtist()
    Dim nId, rs, sName, sLast, sImage, nSongs, nAlbums, sFull
    nId = ParamInt("id", 0)

    Set rs = DbQuery("SELECT TOP 1 ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " _
        & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS N_SONGS " _
        & "FROM TBL_ARTIST ar WHERE ar.ID_ARTIST = " & nId)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        EmptyState "Artista no encontrado", "El artista solicitado no existe o fue eliminado."
        Exit Sub
    End If

    sName  = rs("ARTIST_NAME") & ""
    sLast  = rs("ARTIST_LASTNAME") & ""
    sImage = rs("ARTIST_IMAGE") & ""
    nSongs = CLng(NumOr(rs("N_SONGS"), 0))
    sFull  = ArtistName(sName, sLast)
    rs.Close
    Set rs = Nothing

    nAlbums = CLng(DbScalar("SELECT COUNT(DISTINCT t.ID_ALBUM) FROM TBL_TRACKLIST t " _
        & "WHERE t.ID_SONG IN (SELECT p.ID_SONG FROM TBL_PLAYER p WHERE p.ID_ARTIST = " & nId & ")", 0))
%>
    <div class="detail-head">
        <img class="detail-cover round" src="<%=H(ImgUrl(sImage, "artist", nId, sFull))%>" alt="<%=H(sFull)%>">
        <div class="detail-info">
            <span class="kind">Artista</span>
            <h1><%=H(sFull)%></h1>
            <div class="detail-meta">
                <span><%=FmtNum(nSongs)%> canciones</span>
                <span class="sep">&middot;</span>
                <span><%=FmtNum(nAlbums)%> albumes</span>
            </div>
            <div class="detail-actions">
                <button type="button" class="btn btn-primary" data-play-all="#artistTracks">
                    <%=Icon("play")%><span>Reproducir</span></button>
                <button type="button" class="btn btn-outline" data-play-all="#artistTracks" data-play-shuffle>
                    <%=Icon("shuffle")%><span>Aleatorio</span></button>
                <% If IsAdmin() Then %>
                <a class="btn btn-ghost" href="<%=U("admin_artists.asp?edit=" & nId)%>">
                    <%=Icon("edit")%><span>Editar artista</span></a>
                <% End If %>
            </div>
        </div>
    </div>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Canciones</h2>
                <p class="section-sub">Todas las canciones interpretadas por <%=H(sFull)%></p>
            </div>
        </div>
        <%
        Set rs = DbQuery(SongListSql("s.ID_SONG IN (SELECT p.ID_SONG FROM TBL_PLAYER p " _
            & "WHERE p.ID_ARTIST = " & nId & ")", "s.ID_SONG ASC", 0))
        SongTable rs, "artistTracks"
        rs.Close
        Set rs = Nothing
        %>
    </section>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Albumes</h2>
                <p class="section-sub">Albumes que incluyen canciones de este artista</p>
            </div>
        </div>
        <%
        Set rs = DbQuery("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " _
            & "(SELECT COUNT(*) FROM TBL_TRACKLIST t2 WHERE t2.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
            & "FROM TBL_ALBUM al WHERE al.ID_ALBUM IN (SELECT DISTINCT t.ID_ALBUM FROM TBL_TRACKLIST t " _
            & "WHERE t.ID_SONG IN (SELECT p.ID_SONG FROM TBL_PLAYER p WHERE p.ID_ARTIST = " & nId & ")) " _
            & "ORDER BY al.TITLE_ALBUM")
        If rs.EOF Then
            EmptyState "Sin albumes", "Este artista todavia no tiene albumes asociados."
        Else
        %>
        <div class="grid">
        <%
            Do While Not rs.EOF
                AlbumCard rs("ID_ALBUM"), rs("TITLE_ALBUM") & "", rs("RELEASE_YEAR"), _
                          rs("ALBUM_IMAGE_PATH") & "", rs("N_SONGS")
                rs.MoveNext
            Loop
        %>
        </div>
        <%
        End If
        rs.Close
        Set rs = Nothing
        %>
    </section>
<%
End Sub
%>
