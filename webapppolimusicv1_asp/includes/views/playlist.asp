<%
' =====================================================================
'  Vista: Detalle de playlist
'  El propietario (y el administrador) pueden editar, agregar y quitar
'  canciones. El resto de usuarios solo puede reproducirla.
' =====================================================================

' Fila de cancion con boton de quitar (solo si bCanEdit).
Sub PlaylistTrackRow(ByVal nIdx, ByVal rs, ByVal nPlaylist, ByVal bCanEdit)
    Dim sArtist, sAlbum, sSrc, sCover, sJson, nId
    nId = CLng(rs("ID_SONG"))
    sArtist = Trim(rs("ARTIST_NAME") & "")
    If Len(sArtist) = 0 Then sArtist = "Artista desconocido"
    sAlbum = Trim(rs("ALBUM_TITLE") & "")
    sSrc = AudioUrl(rs("SONG_PATH"))
    sCover = SongCover(rs)
    sJson = TrackJson(nId, rs("SONG_NAME") & "", sArtist, sSrc, sCover, sAlbum)

    Response.Write "<tr class=""track"" data-id=""" & nId & """ data-track=""" & H(sJson) & """>"
    Response.Write "<td class=""track-num""><span class=""n"">" & nIdx & "</span>" _
        & "<span class=""ico"">" & Icon("play") & "</span></td>"

    Response.Write "<td><div class=""track-cell"">" _
        & "<img class=""track-cover"" loading=""lazy"" src=""" & H(sCover) & """ alt=""" & H(rs("SONG_NAME")) & """>" _
        & "<div class=""track-meta"">" _
        & "<span class=""track-title"">" & H(rs("SONG_NAME")) & "</span>" _
        & "<span class=""track-sub"">" & H(sArtist) & "</span></div></div></td>"

    Response.Write "<td class=""track-artist hide-sm"">" & H(sArtist) & "</td>"

    If Len(sAlbum) = 0 Then
        Response.Write "<td class=""track-album hide-sm""><span class=""muted"">Sin album</span></td>"
    Else
        Response.Write "<td class=""track-album hide-sm""><a href=""" _
            & U("home.asp?view=album&id=" & rs("ID_ALBUM")) & """ data-nav>" & H(sAlbum) & "</a></td>"
    End If

    Response.Write "<td class=""track-plays hide-sm""><span data-plays-for=""" & nId & """>" _
        & FmtNum(rs("PLAYS")) & "</span></td>"

    Response.Write "<td class=""track-acts"">"
    If bCanEdit Then
        Response.Write "<form method=""post"" style=""display:inline"" action=""" _
            & H(U("home.asp?view=playlist&id=" & nPlaylist)) & """>" _
            & "<input type=""hidden"" name=""action"" value=""playlist_remove"">" _
            & "<input type=""hidden"" name=""playlist"" value=""" & nPlaylist & """>" _
            & "<input type=""hidden"" name=""song"" value=""" & nId & """>" _
            & "<button type=""submit"" class=""btn-icon danger"" data-confirm=""&iquest;Quitar esta cancion de la playlist?"" " _
            & "title=""Quitar de la playlist"">" & Icon("trash") & "</button></form>"
    End If
    Response.Write "<button type=""button"" class=""btn-icon solid"" data-play-id=""" & nId _
        & """ title=""Reproducir"">" & Icon("play") & "</button></td></tr>"
End Sub

Sub ViewPlaylist()
    Dim nId, rs, sTitle, sPhoto, nType, sOwner, nOwner, nSongs, bCanEdit, nAvail

    nId = ParamInt("id", 0)

    Set rs = DbQuery("SELECT TOP 1 pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, pl.ID_USER, " _
        & "pl.CREATION_DATE, u.USERNAME, " _
        & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS " _
        & "FROM TBL_PLAYLIST pl INNER JOIN TBL_USER u ON u.ID_USER = pl.ID_USER " _
        & "WHERE pl.ID_PLAYLIST = " & nId)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        EmptyState "Playlist no encontrada", "La playlist solicitada no existe o fue eliminada."
        Exit Sub
    End If

    sTitle = rs("TITLE") & ""
    sPhoto = rs("PHOTO") & ""
    nType  = CLng(NumOr(rs("TYPE"), PLAYLIST_PUBLIC))
    sOwner = rs("USERNAME") & ""
    nOwner = CLng(NumOr(rs("ID_USER"), 0))
    nSongs = CLng(NumOr(rs("N_SONGS"), 0))
    rs.Close
    Set rs = Nothing

    bCanEdit = (IsAdmin() Or (IsLoggedIn() And nOwner = CurrentUserId()))

    nAvail = CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG NOT IN " _
        & "(SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & nId & ")", 0))
%>
    <div class="detail-head">
        <img class="detail-cover" src="<%=H(ImgUrl(sPhoto, "playlist", nId, sTitle))%>" alt="<%=H(sTitle)%>">
        <div class="detail-info">
            <span class="kind">Playlist &middot;
                <% If nType = PLAYLIST_PRIVATE Then %>Privada<% Else %>Publica<% End If %>
            </span>
            <h1><%=H(sTitle)%></h1>
            <div class="detail-meta">
                <span>Creada por <b><%=H(sOwner)%></b></span>
                <span class="sep">&middot;</span>
                <span><%=FmtNum(nSongs)%> canciones</span>
            </div>
            <div class="detail-actions">
                <button type="button" class="btn btn-primary" data-play-all="#plTracks">
                    <%=Icon("play")%><span>Reproducir</span></button>
                <button type="button" class="btn btn-outline" data-play-all="#plTracks" data-play-shuffle>
                    <%=Icon("shuffle")%><span>Aleatorio</span></button>
                <% If IsAdmin() And nOwner <> CurrentUserId() Then %>
                <a class="btn btn-ghost" href="<%=U("admin_playlists.asp?edit=" & nId)%>">
                    <%=Icon("shield")%><span>Editar como admin</span></a>
                <% End If %>
            </div>
        </div>
    </div>

    <% If bCanEdit Then %>
    <details class="panel" style="padding:0">
        <summary style="padding:18px 22px;cursor:pointer;font-weight:650;list-style:none">
            <%=Icon("edit")%> Propiedades de la playlist
        </summary>
        <div style="padding:0 22px 22px">
            <form method="post" action="<%=H(U("home.asp?view=playlist&id=" & nId))%>"
                  enctype="multipart/form-data">
                <input type="hidden" name="action" value="playlist_save">
                <input type="hidden" name="id" value="<%=nId%>">
                <div class="form-grid">
                    <div class="field">
                        <label for="pl_title">Titulo</label>
                        <input class="input" type="text" id="pl_title" name="title"
                               maxlength="50" required value="<%=H(sTitle)%>">
                    </div>
                    <div class="field">
                        <label for="pl_type">Visibilidad</label>
                        <select class="select" id="pl_type" name="type">
                            <%
                            Set rs = DbQuery("SELECT COD_CATALOG, [VALUE] FROM TBL_CATALOG " _
                                & "WHERE COD_CATALOG_PARENT = '2' ORDER BY COD_CATALOG")
                            Do While Not rs.EOF
                                Response.Write "<option value=""" & H(rs("COD_CATALOG")) & """" _
                                    & IIf2(CLng(NumOr(rs("COD_CATALOG"), 0)) = nType, " selected", "") _
                                    & ">" & H(rs("VALUE")) & "</option>"
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </select>
                    </div>
                    <div class="field">
                        <label for="pl_photo">Cambiar portada <span class="hint">(opcional)</span></label>
                        <input class="input" type="file" id="pl_photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                    </div>
                </div>
                <div class="form-actions">
                    <button type="submit" class="btn btn-primary">
                        <%=Icon("check")%><span>Guardar cambios</span></button>
                    <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                            data-confirm="&iquest;Eliminar definitivamente esta playlist y todas sus canciones asociadas?">
                        <%=Icon("trash")%><span>Eliminar playlist</span></button>
                </div>
            </form>
        </div>
    </details>

    <details class="panel" style="padding:0">
        <summary style="padding:18px 22px;cursor:pointer;font-weight:650;list-style:none">
            <%=Icon("plus")%> Agregar canciones
            <span class="badge gray" style="margin-left:8px"><%=FmtNum(nAvail)%> disponibles</span>
        </summary>
        <div style="padding:0 22px 22px">
            <% If nAvail = 0 Then %>
            <p class="muted small mb-0">Todas las canciones del catalogo ya estan en esta playlist.</p>
            <% Else %>
            <form method="post" action="<%=H(U("home.asp?view=playlist&id=" & nId))%>">
                <input type="hidden" name="action" value="playlist_add">
                <input type="hidden" name="playlist" value="<%=nId%>">
                <div class="form-grid">
                    <div class="field">
                        <label for="add_song">Cancion</label>
                        <select class="select" id="add_song" name="song" required>
                            <%
                            Set rs = DbQuery(SongListSql("s.ID_SONG NOT IN (SELECT ps.ID_SONG " _
                                & "FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = " & nId & ")", _
                                "s.SONG_NAME ASC", 0))
                            Do While Not rs.EOF
                                Response.Write "<option value=""" & CLng(rs("ID_SONG")) & """>" _
                                    & H(rs("SONG_NAME")) & IIf2(Len(Trim(rs("ARTIST_NAME") & "")) > 0, _
                                    " - " & H(rs("ARTIST_NAME")), "") & "</option>"
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </select>
                    </div>
                </div>
                <button type="submit" class="btn btn-primary">
                    <%=Icon("plus")%><span>Agregar a la playlist</span></button>
            </form>
            <% End If %>
        </div>
    </details>
    <% End If %>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Canciones de la playlist</h2>
                <p class="section-sub"><%=FmtNum(nSongs)%> canciones en total</p>
            </div>
        </div>

        <%
        Set rs = DbQuery(SongListSql("s.ID_SONG IN (SELECT ps.ID_SONG FROM TBL_PLAYLIST_SONG ps " _
            & "WHERE ps.ID_PLAYLIST = " & nId & ")", "s.ID_SONG ASC", 0))
        If rs.EOF Then
            rs.Close
            Set rs = Nothing
            EmptyState "Playlist vacia", "Agregue canciones a esta playlist desde el formulario superior."
        Else
            Response.Write "<div class=""table-wrap""><table class=""tracklist"" data-tracklist id=""plTracks"">"
            TrackTableHead
            Response.Write "<tbody>"
            Dim nI
            nI = 0
            Do While Not rs.EOF
                nI = nI + 1
                PlaylistTrackRow nI, rs, nId, bCanEdit
                rs.MoveNext
            Loop
            Response.Write "</tbody></table></div>"
            rs.Close
            Set rs = Nothing
        End If
        %>
    </section>
<%
End Sub
%>
