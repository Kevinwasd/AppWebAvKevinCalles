<%
' =====================================================================
'  Vista: Artistas (rejilla completa)
' =====================================================================
Sub ViewArtists()
    Dim rs, nTotal
    nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", 0))
%>
    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Artistas</h2>
                <p class="section-sub"><%=FmtNum(nTotal)%> artistas registrados</p>
            </div>
            <% If IsAdmin() Then %>
            <a class="btn btn-sm btn-primary" href="<%=U("admin_artists.asp")%>">
                <%=Icon("plus")%><span>Administrar artistas</span></a>
            <% End If %>
        </div>

        <%
        Set rs = DbQuery("SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS N_SONGS " _
            & "FROM TBL_ARTIST ar ORDER BY ar.ARTIST_NAME")
        If rs.EOF Then
            EmptyState "Sin artistas", "Aun no se han registrado artistas en el catalogo."
        Else
        %>
        <div class="grid grid-artists">
        <%
            Do While Not rs.EOF
                ArtistCard rs("ID_ARTIST"), rs("ARTIST_NAME") & "", rs("ARTIST_LASTNAME") & "", _
                           rs("ARTIST_IMAGE") & "", rs("N_SONGS")
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
