<%
' =====================================================================
'  Vista: Inicio (panel principal del usuario)
' =====================================================================
Sub ViewHome()
    Dim rs, nSongs, nArtists, nAlbums, nPlaylists, sHello
%>
    <section class="hero">
        <span class="hero-kicker">Bienvenido a <%=H(APP_NAME)%></span>
        <h1>Hola, <%=H(CurrentUserName())%>.</h1>
        <p>Explore el catalogo completo, reproduzca sus canciones favoritas y
           administre sus playlists publicas y privadas.</p>
        <div class="hero-actions">
            <a class="btn btn-primary" href="<%=U("home.asp?view=songs")%>" data-nav>
                <%=Icon("music")%><span>Ir a las canciones</span>
            </a>
            <a class="btn btn-outline" href="<%=U("home.asp?view=playlists&mine=1")%>" data-nav>
                <%=Icon("playlist")%><span>Mis playlists</span>
            </a>
            <% If IsAdmin() Then %>
            <a class="btn btn-ghost" href="<%=U("admin.asp")%>">
                <%=Icon("shield")%><span>Panel de administracion</span>
            </a>
            <% End If %>
        </div>
    </section>

    <div class="stats">
        <div class="stat">
            <span class="ic"><%=Icon("music")%></span>
            <span><b><%=FmtNum(DbScalar("SELECT COUNT(*) FROM TBL_SONG", 0))%></b>
                  <span>Canciones</span></span>
        </div>
        <div class="stat">
            <span class="ic"><%=Icon("artist")%></span>
            <span><b><%=FmtNum(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", 0))%></b>
                  <span>Artistas</span></span>
        </div>
        <div class="stat">
            <span class="ic"><%=Icon("album")%></span>
            <span><b><%=FmtNum(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", 0))%></b>
                  <span>Albumes</span></span>
        </div>
        <div class="stat">
            <span class="ic"><%=Icon("playlist")%></span>
            <span><b><%=FmtNum(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", 0))%></b>
                  <span>Playlists</span></span>
        </div>
    </div>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Las mas escuchadas</h2>
                <p class="section-sub">Canciones ordenadas por numero de reproducciones</p>
            </div>
            <a class="link-more" href="<%=U("home.asp?view=songs")%>" data-nav>Ver todas &rarr;</a>
        </div>
        <%
        Set rs = DbQuery(SongListSql("1 = 1", "ISNULL(s.PLAYS,0) DESC, s.ID_SONG ASC", 5))
        SongTable rs, "topSongs"
        rs.Close
        Set rs = Nothing
        %>
    </section>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Albumes</h2>
                <p class="section-sub">Colecciones disponibles en el catalogo</p>
            </div>
            <a class="link-more" href="<%=U("home.asp?view=albums")%>" data-nav>Ver todos &rarr;</a>
        </div>
        <div class="grid">
        <%
        Set rs = DbQuery("SELECT TOP 6 al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, " _
            & "al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t " _
            & "WHERE t.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
            & "FROM TBL_ALBUM al ORDER BY al.ID_ALBUM")
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

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Artistas</h2>
                <p class="section-sub">Interpretes registrados</p>
            </div>
            <a class="link-more" href="<%=U("home.asp?view=artists")%>" data-nav>Ver todos &rarr;</a>
        </div>
        <div class="grid grid-artists">
        <%
        Set rs = DbQuery("SELECT TOP 6 ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, " _
            & "ar.ARTIST_IMAGE, (SELECT COUNT(*) FROM TBL_PLAYER p " _
            & "WHERE p.ID_ARTIST = ar.ID_ARTIST) AS N_SONGS " _
            & "FROM TBL_ARTIST ar ORDER BY ar.ID_ARTIST")
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

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Playlists publicas</h2>
                <p class="section-sub">Listas compartidas por la comunidad</p>
            </div>
            <a class="link-more" href="<%=U("home.asp?view=playlists")%>" data-nav>Ver todas &rarr;</a>
        </div>
        <div class="grid">
        <%
        Set rs = DbQuery("SELECT TOP 6 pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, " _
            & "u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps " _
            & "WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS " _
            & "FROM TBL_PLAYLIST pl INNER JOIN TBL_USER u ON u.ID_USER = pl.ID_USER " _
            & "WHERE pl.TYPE = " & PLAYLIST_PUBLIC & " ORDER BY pl.ID_PLAYLIST DESC")
        Do While Not rs.EOF
            PlaylistCard rs("ID_PLAYLIST"), rs("TITLE") & "", rs("PHOTO") & "", _
                         rs("TYPE"), rs("N_SONGS"), rs("USERNAME") & ""
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
        </div>
    </section>
<%
End Sub
%>
