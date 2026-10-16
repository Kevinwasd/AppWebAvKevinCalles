<%
' =====================================================================
'  PoliMusic - Componentes de interfaz reutilizables
' =====================================================================

' ---------------------------------------------------------------------
'  Consulta estandar de canciones (incluye artista, album y caratula).
' ---------------------------------------------------------------------
Function SongListSql(ByVal sWhere, ByVal sOrder, ByVal nTop)
    Dim sql, sLimit
    If nTop > 0 Then
        sLimit = "TOP " & CLng(nTop) & " "
    Else
        sLimit = ""
    End If
    sql = "SELECT " & sLimit & "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " _
        & "(SELECT TOP 1 a.ARTIST_NAME + ISNULL(' ' + NULLIF(LTRIM(RTRIM(a.ARTIST_LASTNAME)), ''), '') " _
        & "   FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " _
        & "  WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) AS ARTIST_NAME, " _
        & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_COUNT, " _
        & "(SELECT TOP 1 al.ID_ALBUM FROM TBL_TRACKLIST t " _
        & "   INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " _
        & "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ID_ALBUM, " _
        & "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t " _
        & "   INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " _
        & "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_TITLE, " _
        & "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t " _
        & "   INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " _
        & "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_IMAGE, " _
        & "(SELECT TOP 1 a.ARTIST_IMAGE FROM TBL_PLAYER p " _
        & "   INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " _
        & "  WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) AS ARTIST_IMAGE " _
        & "FROM TBL_SONG s "
    If Len(sWhere) > 0 Then sql = sql & "WHERE " & sWhere & " "
    If Len(sOrder) > 0 Then sql = sql & "ORDER BY " & sOrder
    SongListSql = sql
End Function

' ---------------------------------------------------------------------
'  Clausula WHERE para paginar canciones en SQL Server (ROW_NUMBER).
'  sInnerOrder usa el alias s2, por ejemplo: "ISNULL(s2.PLAYS,0) DESC, s2.ID_SONG ASC"
' ---------------------------------------------------------------------
Function SongPagedWhere(ByVal nPage, ByVal nPerPage, ByVal sInnerOrder)
    Dim nFrom, nTo
    nFrom = PageOffset(nPage, nPerPage) + 1
    nTo = nFrom + nPerPage - 1
    SongPagedWhere = "s.ID_SONG IN (SELECT x.ID_SONG FROM (" _
        & "SELECT s2.ID_SONG, ROW_NUMBER() OVER (ORDER BY " & sInnerOrder & ") AS RN " _
        & "FROM TBL_SONG s2) x WHERE x.RN BETWEEN " & nFrom & " AND " & nTo & ")"
End Function

' Caratula de una cancion: album -> artista -> marcador generado.
Function SongCover(ByVal rs)
    Dim p
    p = rs("ALBUM_IMAGE") & ""
    If Len(Trim(p)) = 0 Then p = rs("ARTIST_IMAGE") & ""
    If Len(Trim(p)) = 0 Then
        SongCover = PlaceholderUrl("song", rs("ID_SONG"), FirstLetter(rs("SONG_NAME")))
    Else
        SongCover = ImgUrl(p, "song", rs("ID_SONG"), rs("SONG_NAME"))
    End If
End Function

' Representacion JSON de una pista para el reproductor.
Function TrackJson(ByVal nId, ByVal sTitle, ByVal sArtist, ByVal sSrc, _
                   ByVal sCover, ByVal sAlbum)
    TrackJson = "{""id"":" & CLng(nId) _
        & ",""title"":" & JsonEnc(sTitle) _
        & ",""artist"":" & JsonEnc(sArtist) _
        & ",""album"":" & JsonEnc(sAlbum) _
        & ",""src"":" & JsonEnc(sSrc) _
        & ",""cover"":" & JsonEnc(sCover) & "}"
End Function

' ---------------------------------------------------------------------
'  Encabezado de la tabla de canciones
' ---------------------------------------------------------------------
Sub TrackTableHead()
    Response.Write "<thead><tr>" _
        & "<th class=""num"">#</th>" _
        & "<th>Titulo</th>" _
        & "<th class=""artist hide-sm"">Artista</th>" _
        & "<th class=""album hide-sm"">Album</th>" _
        & "<th class=""hide-sm"">Reproducciones</th>" _
        & "<th class=""acts"">Acciones</th>" _
        & "</tr></thead>"
End Sub

' ---------------------------------------------------------------------
'  Fila de cancion. rs debe venir de SongListSql().
' ---------------------------------------------------------------------
Sub TrackRow(ByVal nIdx, ByVal rs)
    Dim sArtist, sAlbum, sSrc, sCover, sJson, nId, sExtra
    nId = CLng(rs("ID_SONG"))
    sArtist = Trim(rs("ARTIST_NAME") & "")
    If Len(sArtist) = 0 Then sArtist = "Artista desconocido"
    If CLng(NumOr(rs("ARTIST_COUNT"), 1)) > 1 Then sArtist = sArtist & " y mas"
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
        & "<span class=""track-sub"">" & H(sArtist) & "</span>" _
        & "</div></div></td>"

    Response.Write "<td class=""track-artist hide-sm"">" & H(sArtist) & "</td>"

    If Len(sAlbum) = 0 Then
        Response.Write "<td class=""track-album hide-sm""><span class=""muted"">Sin album</span></td>"
    Else
        Response.Write "<td class=""track-album hide-sm""><a href=""" _
            & U("home.asp?view=album&id=" & rs("ID_ALBUM")) & """ data-nav>" & H(sAlbum) & "</a></td>"
    End If

    Response.Write "<td class=""track-plays hide-sm""><span data-plays-for=""" & nId & """>" _
        & FmtNum(rs("PLAYS")) & "</span></td>"

    sExtra = ""
    If IsLoggedIn() Then
        sExtra = "<button type=""button"" class=""btn-icon"" data-add-song=""" & nId _
            & """ title=""Anadir a una playlist"" aria-label=""Anadir a playlist"">" & Icon("plus") & "</button>"
    End If
    Response.Write "<td class=""track-acts"">" & sExtra _
        & "<button type=""button"" class=""btn-icon solid"" data-play-id=""" & nId _
        & """ title=""Reproducir"" aria-label=""Reproducir"">" & Icon("play") & "</button></td>"
    Response.Write "</tr>"
End Sub

' ---------------------------------------------------------------------
'  Bloque completo: barra de herramientas + tabla de canciones
' ---------------------------------------------------------------------
Sub SongTable(ByVal rs, ByVal sId)
    Dim n, nTotal
    Response.Write "<div class=""table-wrap""><table class=""tracklist"" data-tracklist id=""" & H(sId) & """>"
    TrackTableHead
    Response.Write "<tbody>"
    n = 0
    nTotal = 0
    Do While Not rs.EOF
        n = n + 1
        TrackRow n, rs
        nTotal = nTotal + 1
        rs.MoveNext
    Loop
    Response.Write "</tbody></table></div>"
    If nTotal = 0 Then
        EmptyState "Sin canciones", "Todavia no hay canciones que mostrar en esta lista."
    End If
End Sub

' ---------------------------------------------------------------------
'  Tarjetas
' ---------------------------------------------------------------------
Sub AlbumCard(ByVal nId, ByVal sTitle, ByVal nYear, ByVal sImage, ByVal nSongs)
    Dim img
    img = ImgUrl(sImage, "album", nId, sTitle)
    Response.Write "<a class=""card"" href=""" & U("home.asp?view=album&id=" & nId) & """ data-nav>" _
        & "<div class=""card-media""><img loading=""lazy"" src=""" & H(img) & """ alt=""" & H(sTitle) & """>" _
        & "<span class=""card-play"">" & Icon("play") & "</span></div>" _
        & "<p class=""card-title"">" & H(sTitle) & "</p>" _
        & "<p class=""card-sub"">" & FmtNum(nSongs) & " canciones" _
        & IIf2(Len(FmtYear(nYear)) > 0, " &middot; " & FmtYear(nYear), "") & "</p></a>"
End Sub

Sub ArtistCard(ByVal nId, ByVal sName, ByVal sLast, ByVal sImage, ByVal nSongs)
    Dim img, full
    full = ArtistName(sName, sLast)
    img = ImgUrl(sImage, "artist", nId, full)
    Response.Write "<a class=""card"" href=""" & U("home.asp?view=artist&id=" & nId) & """ data-nav>" _
        & "<div class=""card-media round""><img loading=""lazy"" src=""" & H(img) & """ alt=""" & H(full) & """>" _
        & "<span class=""card-play"">" & Icon("play") & "</span></div>" _
        & "<p class=""card-title"">" & H(full) & "</p>" _
        & "<p class=""card-sub"">" & FmtNum(nSongs) & " canciones</p></a>"
End Sub

Sub PlaylistCard(ByVal nId, ByVal sTitle, ByVal sPhoto, ByVal nType, ByVal nSongs, ByVal sOwner)
    Dim img, badge
    img = ImgUrl(sPhoto, "playlist", nId, sTitle)
    If CLng(nType) = PLAYLIST_PRIVATE Then
        badge = "<span class=""badge gray"">Privada</span>"
    Else
        badge = "<span class=""badge"">Publica</span>"
    End If
    Response.Write "<a class=""card"" href=""" & U("home.asp?view=playlist&id=" & nId) & """ data-nav>" _
        & "<div class=""card-media""><img loading=""lazy"" src=""" & H(img) & """ alt=""" & H(sTitle) & """>" _
        & "<span class=""card-play"">" & Icon("play") & "</span></div>" _
        & "<p class=""card-title"">" & H(sTitle) & "</p>" _
        & "<p class=""card-sub"">" & badge & " &middot; " & FmtNum(nSongs) & " canciones</p></a>"
End Sub

' ---------------------------------------------------------------------
'  Varios
' ---------------------------------------------------------------------
Function FmtYear(ByVal v)
    If IsNull(v) Or Not IsNumeric(v) Then
        FmtYear = ""
    Else
        If CLng(v) <= 0 Then
            FmtYear = ""
        Else
            FmtYear = CStr(CLng(v))
        End If
    End If
End Function

' Operador ternario simple: IIf2(condicion, valorSi, valorNo)
Function IIf2(ByVal bCond, ByVal vTrue, ByVal vFalse)
    If bCond Then
        IIf2 = vTrue
    Else
        IIf2 = vFalse
    End If
End Function

Sub EmptyState(ByVal sTitle, ByVal sText)
    Response.Write "<div class=""empty"">" & Icon("music") _
        & "<h3>" & H(sTitle) & "</h3><p>" & H(sText) & "</p></div>"
End Sub

Sub PageTitle(ByVal sKicker, ByVal sTitle, ByVal sSub)
    Response.Write "<div class=""section-head""><div>"
    If Len(sKicker) > 0 Then
        Response.Write "<div><span class=""kind"">" & H(sKicker) & "</span></div>"
    End If
    Response.Write "<h2 class=""section-title""><span class=""dot""></span>" & H(sTitle) & "</h2>"
    If Len(sSub) > 0 Then Response.Write "<p class=""section-sub"">" & H(sSub) & "</p>"
    Response.Write "</div></div>"
End Sub

' Paginacion: sBaseUrl es la URL (sin el parametro page) a la que se anade &page=N
Sub Pager(ByVal nPage, ByVal nTotalPages, ByVal sBaseUrl)
    Dim i, sep, cls
    If nTotalPages <= 1 Then Exit Sub
    sep = "?"
    If InStr(sBaseUrl, "?") > 0 Then sep = "&"
    Response.Write "<div class=""pagination"">"
    If nPage > 1 Then
        Response.Write "<a href=""" & H(sBaseUrl & sep & "page=" & (nPage - 1)) & """ data-nav>&laquo; Anterior</a>"
    Else
        Response.Write "<span class=""off"">&laquo; Anterior</span>"
    End If

    For i = 1 To nTotalPages
        If i = nPage Then
            Response.Write "<span class=""current"">" & i & "</span>"
        Else
            Response.Write "<a href=""" & H(sBaseUrl & sep & "page=" & i) & """ data-nav>" & i & "</a>"
        End If
    Next

    If nPage < nTotalPages Then
        Response.Write "<a href=""" & H(sBaseUrl & sep & "page=" & (nPage + 1)) & """ data-nav>Siguiente &raquo;</a>"
    Else
        Response.Write "<span class=""off"">Siguiente &raquo;</span>"
    End If
    Response.Write "</div>"
End Sub

' Boton "Reproducir todo" / "Aleatorio" para una lista de canciones.
Sub PlayAllButtons(ByVal sSelector, ByVal sLabel)
    Response.Write "<button type=""button"" class=""btn btn-primary"" data-play-all=""" & H(sSelector) & """>" _
        & Icon("play") & "<span>" & H(sLabel) & "</span></button>"
    Response.Write "<button type=""button"" class=""btn btn-outline"" data-play-all=""" & H(sSelector) _
        & """ data-play-shuffle>" & Icon("shuffle") & "<span>Aleatorio</span></button>"
End Sub

' ---------------------------------------------------------------------
'  Lista de casillas de verificacion (para relaciones N:M)
'  sSelected debe venir como ",1,4,7," para poder buscar por pertenencia.
' ---------------------------------------------------------------------
Sub CheckBoxList(ByVal sName, ByVal rs, ByVal sValueField, ByVal sLabelField, _
                 ByVal sSelected, ByVal sEmptyMsg)
    Dim n, sId
    Response.Write "<div class=""checklist"">"
    If rs.EOF Then
        Response.Write "<span class=""cl-empty"">" & H(sEmptyMsg) & "</span>"
    End If
    Do While Not rs.EOF
        n = CLng(NumOr(rs(sValueField), 0))
        sId = "," & n & ","
        Response.Write "<label><input type=""checkbox"" name=""" & H(sName) & """ value=""" & n & """" _
            & IIf2(InStr(sSelected, sId) > 0, " checked", "") & "><span>" _
            & H(rs(sLabelField)) & "</span></label>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End Sub

' Devuelve ",id1,id2," con los valores de una consulta.
Function SelectedIds(ByVal sql, ByVal sField)
    Dim rs, s
    s = ","
    Set rs = DbQuery(sql)
    Do While Not rs.EOF
        s = s & CLng(NumOr(rs(sField), 0)) & ","
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    SelectedIds = s
End Function

' Recorre una lista "1,2,3" (o "1, 2, 3") y devuelve un array de enteros.
Function IdList(ByVal s)
    Dim a, i, out, n, k
    out = ""
    a = Split(s & "", ",")
    For i = 0 To UBound(a)
        If IsNumeric(Trim(a(i))) Then
            n = CLng(Trim(a(i)))
            If n > 0 Then
                If Len(out) > 0 Then out = out & ","
                out = out & n
            End If
        End If
    Next
    If Len(out) = 0 Then
        IdList = Array()
    Else
        IdList = Split(out, ",")
    End If
End Function
%>
