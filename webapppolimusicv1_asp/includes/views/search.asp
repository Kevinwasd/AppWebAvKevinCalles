<%
' =====================================================================
'  Vista: Busqueda (canciones, artistas y albumes)
' =====================================================================
Sub ViewSearch()
    Dim sQ, sLike, rs, nSongs, nArtists, nAlbums
    sQ = ParamMax("q", 80)
    sLike = "%" & sQ & "%"

    If Len(sQ) = 0 Then
        EmptyState "Busque en PoliMusic", "Escriba el nombre de una cancion, artista o album en la barra superior."
        Exit Sub
    End If

    nSongs   = CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE SONG_NAME LIKE " & SqlStr(sLike), 0))
    nArtists = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST WHERE ARTIST_NAME LIKE " & SqlStr(sLike) _
        & " OR ISNULL(ARTIST_LASTNAME,'') LIKE " & SqlStr(sLike), 0))
    nAlbums  = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM WHERE TITLE_ALBUM LIKE " & SqlStr(sLike), 0))
%>
    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Resultados para &ldquo;<%=H(sQ)%>&rdquo;</h2>
                <p class="section-sub">
                    <%=FmtNum(nSongs)%> canciones &middot;
                    <%=FmtNum(nArtists)%> artistas &middot;
                    <%=FmtNum(nAlbums)%> albumes
                </p>
            </div>
        </div>

        <% If nSongs + nArtists + nAlbums = 0 Then %>
            <% EmptyState "Sin resultados", "No se encontro ninguna coincidencia. Intente con otro termino." %>
        <% End If %>
    </section>

    <% If nArtists > 0 Then %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Artistas</h2>
        </div></div>
        <div class="grid grid-artists">
        <%
        Set rs = DbQuery("SELECT TOP 8 ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS N_SONGS " _
            & "FROM TBL_ARTIST ar WHERE ar.ARTIST_NAME LIKE " & SqlStr(sLike) _
            & " OR ISNULL(ar.ARTIST_LASTNAME,'') LIKE " & SqlStr(sLike) & " ORDER BY ar.ARTIST_NAME")
        Do While Not rs.EOF
            ArtistCard rs("ID_ARTIST"), rs("ARTIST_NAME") & "", rs("ARTIST_LASTNAME") & "", _
                       rs("ARTIST_IMAGE") & "", rs("N_SONGS")
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
        </div>
    </section>
    <% End If %>

    <% If nAlbums > 0 Then %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Albumes</h2>
        </div></div>
        <div class="grid">
        <%
        Set rs = DbQuery("SELECT TOP 8 al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " _
            & "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
            & "FROM TBL_ALBUM al WHERE al.TITLE_ALBUM LIKE " & SqlStr(sLike) & " ORDER BY al.TITLE_ALBUM")
        Do While Not rs.EOF
            AlbumCard rs("ID_ALBUM"), rs("TITLE_ALBUM") & "", rs("RELEASE_YEAR"), _
                      rs("ALBUM_IMAGE_PATH") & "", rs("N_SONGS")
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
        </div>
    </section>
    <% End If %>

    <% If nSongs > 0 Then %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Canciones</h2>
        </div></div>
        <%
        Set rs = DbQuery(SongListSql("(s.SONG_NAME LIKE " & SqlStr(sLike) _
            & " OR s.ID_SONG IN (SELECT p.ID_SONG FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a " _
            & "ON a.ID_ARTIST = p.ID_ARTIST WHERE a.ARTIST_NAME LIKE " & SqlStr(sLike) _
            & " OR ISNULL(a.ARTIST_LASTNAME,'') LIKE " & SqlStr(sLike) & ") " _
            & "OR s.ID_SONG IN (SELECT t.ID_SONG FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al " _
            & "ON al.ID_ALBUM = t.ID_ALBUM WHERE al.TITLE_ALBUM LIKE " & SqlStr(sLike) & "))", _
            "s.SONG_NAME ASC", 50))
        SongTable rs, "searchSongs"
        rs.Close
        Set rs = Nothing
        %>
    </section>
    <% End If %>
<%
End Sub
%>
