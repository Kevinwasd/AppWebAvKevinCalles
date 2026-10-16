<%
' =====================================================================
'  Vista: Albumes (rejilla completa)
' =====================================================================
Sub ViewAlbums()
    Dim rs, nTotal
    nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", 0))
%>
    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Albumes</h2>
                <p class="section-sub"><%=FmtNum(nTotal)%> albumes en el catalogo</p>
            </div>
            <% If IsAdmin() Then %>
            <a class="btn btn-sm btn-primary" href="<%=U("admin_albums.asp")%>">
                <%=Icon("plus")%><span>Administrar albumes</span></a>
            <% End If %>
        </div>

        <%
        Set rs = DbQuery("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, " _
            & "al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t " _
            & "WHERE t.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
            & "FROM TBL_ALBUM al ORDER BY al.TITLE_ALBUM")
        If rs.EOF Then
            EmptyState "Sin albumes", "Aun no se han registrado albumes en el catalogo."
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
