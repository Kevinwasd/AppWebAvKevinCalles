<%
' =====================================================================
'  Vista: Canciones
'  Lista completa con buscador local, orden, paginacion y controles
'  de reproduccion (reproductor inferior + boton por fila).
' =====================================================================
Sub ViewSongs()
    Dim nPage, nTotal, nPages, rs, sOrder, sOrderKey, sBase, sWhere, sInner
    Dim nFav, nPlays

    nPage = ParamInt("page", 1)
    sOrderKey = LCase(Param("sort", "id"))

    Select Case sOrderKey
        Case "name"
            sOrder = "s.SONG_NAME ASC, s.ID_SONG ASC"
            sInner = "s2.SONG_NAME ASC, s2.ID_SONG ASC"
        Case "plays"
            sOrder = "ISNULL(s.PLAYS,0) DESC, s.ID_SONG ASC"
            sInner = "ISNULL(s2.PLAYS,0) DESC, s2.ID_SONG ASC"
        Case Else
            sOrderKey = "id"
            sOrder = "s.ID_SONG ASC"
            sInner = "s2.ID_SONG ASC"
    End Select

    nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG", 0))
    nPages = PageCount(nTotal, PAGE_SIZE)
    If nPages < 1 Then nPages = 1
    If nPage < 1 Then nPage = 1
    If nPage > nPages Then nPage = nPages

    nFav = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG", 0))
    nPlays = CLng(DbScalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", 0))

    sWhere = SongPagedWhere(nPage, PAGE_SIZE, sInner)
    sBase = "home.asp?view=songs&sort=" & sOrderKey
%>
    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Lista de canciones</h2>
                <p class="section-sub">
                    <%=FmtNum(nTotal)%> canciones disponibles &middot;
                    <%=FmtNum(nPlays)%> reproducciones acumuladas &middot;
                    <%=FmtNum(nFav)%> inclusiones en playlists
                </p>
            </div>
            <div class="flex gap-8 wrap">
                <a class="btn btn-sm <%=IIf2(sOrderKey = "id", "btn-primary", "btn-outline")%>"
                   href="<%=U("home.asp?view=songs&sort=id")%>" data-nav>Recientes</a>
                <a class="btn btn-sm <%=IIf2(sOrderKey = "name", "btn-primary", "btn-outline")%>"
                   href="<%=U("home.asp?view=songs&sort=name")%>" data-nav>A-Z</a>
                <a class="btn btn-sm <%=IIf2(sOrderKey = "plays", "btn-primary", "btn-outline")%>"
                   href="<%=U("home.asp?view=songs&sort=plays")%>" data-nav>Mas escuchadas</a>
            </div>
        </div>

        <div class="toolbar">
            <input class="filter" type="search" id="songFilter" placeholder="Filtrar por titulo, artista o album..."
                   data-filter-target="#songList .track" data-filter-count="#songCount" autocomplete="off">
            <button type="button" class="btn btn-primary btn-sm"
                    data-play-all="#songList"> <%=Icon("play")%><span>Reproducir todo</span></button>
            <button type="button" class="btn btn-outline btn-sm"
                    data-play-all="#songList" data-play-shuffle><%=Icon("shuffle")%><span>Aleatorio</span></button>
            <span class="count"><span id="songCount"><%=FmtNum(nTotal)%></span> canciones</span>
        </div>

        <%
        Set rs = DbQuery(SongListSql(sWhere, sOrder, 0))
        SongTable rs, "songList"
        rs.Close
        Set rs = Nothing
        %>

        <% Pager nPage, nPages, U(sBase) %>
    </section>

    <section class="section">
        <div class="panel">
            <h3 class="panel-title"><%=Icon("sparkle")%> Controles del reproductor</h3>
            <div class="form-grid">
                <div>
                    <p class="small muted mb-0"><b style="color:var(--pm-text)">Reproducir</b><br>
                    Pulse cualquier fila de la lista o el boton verde de la columna de acciones.</p>
                </div>
                <div>
                    <p class="small muted mb-0"><b style="color:var(--pm-text)">Cola de reproduccion</b><br>
                    El reproductor toma como cola todas las canciones visibles de la lista actual.</p>
                </div>
                <div>
                    <p class="small muted mb-0"><b style="color:var(--pm-text)">Atajos de teclado</b><br>
                    <code>Espacio</code> reproducir/pausar &middot;
                    <code>Shift + &rarr;</code> siguiente &middot;
                    <code>Shift + &larr;</code> anterior.</p>
                </div>
            </div>
        </div>
    </section>
<%
End Sub
%>
