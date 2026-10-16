<%
' =====================================================================
'  Vista: Playlists
'  mine=1 muestra unicamente las playlists del usuario conectado.
' =====================================================================
Sub ViewPlaylists()
    Dim rs, bMine, nTotal, sWhere, sBase
    bMine = (Param("mine", "") = "1")
    sBase = "home.asp?view=playlists"

    If bMine Then
        sWhere = "pl.ID_USER = " & CurrentUserId()
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = " & CurrentUserId(), 0))
    Else
        sWhere = "(pl.TYPE = " & PLAYLIST_PUBLIC & " OR pl.ID_USER = " & CurrentUserId() & ")"
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE " & sWhere, 0))
    End If
%>
    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>
                    <% If bMine Then %>Mis playlists<% Else %>Playlists<% End If %></h2>
                <p class="section-sub">
                    <%=FmtNum(nTotal)%> playlists
                    <% If bMine Then %>creadas por usted<% Else %>publicas y propias<% End If %>
                </p>
            </div>
            <div class="flex gap-8 wrap">
                <a class="btn btn-sm <%=IIf2(Not bMine, "btn-primary", "btn-outline")%>"
                   href="<%=U(sBase)%>" data-nav>Todas</a>
                <a class="btn btn-sm <%=IIf2(bMine, "btn-primary", "btn-outline")%>"
                   href="<%=U(sBase & "&mine=1")%>" data-nav>Solo mias</a>
                <% If IsAdmin() Then %>
                <a class="btn btn-sm btn-ghost" href="<%=U("admin_playlists.asp")%>">
                    <%=Icon("shield")%><span>Administrar</span></a>
                <% End If %>
            </div>
        </div>

        <details class="panel" style="padding:0">
            <summary style="padding:18px 22px;cursor:pointer;font-weight:650;list-style:none">
                <%=Icon("plus")%> Crear una nueva playlist
            </summary>
            <div style="padding:0 22px 22px">
                <form method="post" action="<%=U(sBase & IIf2(bMine, "&mine=1", ""))%>"
                      enctype="multipart/form-data">
                    <input type="hidden" name="action" value="playlist_save">
                    <input type="hidden" name="id" value="0">
                    <div class="form-grid">
                        <div class="field">
                            <label for="np_title">Titulo de la playlist</label>
                            <input class="input" type="text" id="np_title" name="title"
                                   maxlength="50" required placeholder="Ej. Favoritas 2026">
                        </div>
                        <div class="field">
                            <label for="np_type">Visibilidad</label>
                            <select class="select" id="np_type" name="type">
                                <%
                                Set rs = DbQuery("SELECT COD_CATALOG, [VALUE] FROM TBL_CATALOG " _
                                    & "WHERE COD_CATALOG_PARENT = '2' ORDER BY COD_CATALOG")
                                Do While Not rs.EOF
                                    Response.Write "<option value=""" & H(rs("COD_CATALOG")) & """>" _
                                        & H(rs("VALUE")) & "</option>"
                                    rs.MoveNext
                                Loop
                                rs.Close
                                Set rs = Nothing
                                %>
                            </select>
                        </div>
                        <div class="field">
                            <label for="np_photo">Portada <span class="hint">(opcional)</span></label>
                            <input class="input" type="file" id="np_photo" name="photo"
                                   accept="image/png,image/jpeg,image/gif,image/webp">
                        </div>
                    </div>
                    <button type="submit" class="btn btn-primary">
                        <%=Icon("check")%><span>Crear playlist</span></button>
                </form>
            </div>
        </details>

        <%
        Set rs = DbQuery("SELECT pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, u.USERNAME, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS " _
            & "FROM TBL_PLAYLIST pl INNER JOIN TBL_USER u ON u.ID_USER = pl.ID_USER " _
            & "WHERE " & sWhere & " ORDER BY pl.ID_PLAYLIST DESC")
        If rs.EOF Then
            EmptyState "Sin playlists", "Cree su primera playlist con el formulario de arriba."
        Else
        %>
        <div class="grid">
        <%
            Do While Not rs.EOF
                PlaylistCard rs("ID_PLAYLIST"), rs("TITLE") & "", rs("PHOTO") & "", _
                             rs("TYPE"), rs("N_SONGS"), rs("USERNAME") & ""
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
