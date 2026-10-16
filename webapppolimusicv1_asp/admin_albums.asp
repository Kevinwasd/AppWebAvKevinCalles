<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/upload.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<%
' =====================================================================
'  PoliMusic - Administracion de albumes (CRUD + lista de canciones)
'  La relacion album-cancion se guarda en TBL_TRACKLIST.
' =====================================================================
Dim sAction, nId, sMsg, bOk, sTitle, nYear, sImage, nNewId
Dim aIds, i, sSearch, nPage, nTotal, nPages, rs, bForm, nEdit
Dim sFTitle, nFYear, sFImage, sSelSongs, sWhere

RequireAdmin()

' --------------------------- Acciones POST ---------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sAction = LCase(UFField("action", ""))
    If UFField("do_delete", "") = "1" Then sAction = "delete"
    nId = UFFieldInt("id", 0)

    If sAction = "delete" Then
        If nId > 0 Then
            sImage = DbScalar("SELECT ISNULL(MIN(ALBUM_IMAGE_PATH),'') FROM TBL_ALBUM WHERE ID_ALBUM = " & nId, "")
            DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & nId
            DbExec "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = " & nId
            If Len(sImage & "") > 0 Then DeleteAsset sImage
            SetFlash "El album fue eliminado correctamente.", "success"
        End If
        Response.Redirect U("admin_albums.asp")
    End If

    If sAction = "save" Then
        sTitle = Left(UFField("title_album", ""), 50)
        nYear  = UFFieldInt("release_year", 0)
        bOk = True
        sMsg = ""

        If Len(sTitle) = 0 Then
            sMsg = "El titulo del album es obligatorio."
            bOk = False
        ElseIf nYear < 0 Or nYear > 2100 Then
            sMsg = "El anio de publicacion no es valido."
            bOk = False
        End If

        If bOk Then
            If nId > 0 Then
                If CLng(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM WHERE ID_ALBUM = " & nId, 0)) = 0 Then
                    bOk = False
                    sMsg = "El album indicado no existe."
                Else
                    DbExec "UPDATE TBL_ALBUM SET TITLE_ALBUM = " & SqlStr(sTitle) _
                        & ", RELEASE_YEAR = " & SqlIntNull(nYear) & " WHERE ID_ALBUM = " & nId
                End If
            Else
                nNewId = DbInsert("INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (" _
                    & SqlStr(sTitle) & ", " & SqlIntNull(nYear) & ", " & SqlStr("~/AlbumImages/album.png") & ")")
                If nNewId <= 0 Then
                    bOk = False
                    sMsg = "No fue posible crear el album."
                Else
                    nId = nNewId
                End If
            End If
        End If

        If bOk And HasUpload("photo") Then
            sImage = SaveUploaded("photo", DIR_ALBUMS, "album_" & nId)
            If Len(sImage) > 0 Then
                DbExec "UPDATE TBL_ALBUM SET ALBUM_IMAGE_PATH = " & SqlStr(sImage) & " WHERE ID_ALBUM = " & nId
            Else
                bOk = False
                sMsg = "No fue posible guardar la portada del album."
            End If
        End If

        If bOk Then
            DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & nId
            aIds = IdList(UFField("songs", ""))
            For i = 0 To UBound(aIds)
                DbExec "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" _
                    & nId & ", " & CLng(aIds(i)) & ")"
            Next
            SetFlash "El album fue guardado correctamente.", "success"
            Response.Redirect U("admin_albums.asp")
        Else
            SetFlash sMsg, "error"
        End If
    End If
End If

nEdit   = ParamInt("edit", 0)
bForm   = (nEdit > 0) Or (Param("new", "") = "1")
sSearch = ParamMax("q", 60)
nPage   = ParamInt("page", 1)

PAGE_TITLE   = "Administrar albumes"
ACTIVE_ADMIN = "albums"
%>
<!--#include file="includes/layout_top.asp"-->

<% If bForm Then
    sFTitle = ""
    nFYear = 0
    sFImage = ""
    sSelSongs = ","
    If nEdit > 0 Then
        Set rs = DbQuery("SELECT TOP 1 TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH " _
            & "FROM TBL_ALBUM WHERE ID_ALBUM = " & nEdit)
        If Not rs.EOF Then
            sFTitle = rs("TITLE_ALBUM") & ""
            nFYear  = CLng(NumOr(rs("RELEASE_YEAR"), 0))
            sFImage = rs("ALBUM_IMAGE_PATH") & ""
            sSelSongs = SelectedIds("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = " & nEdit, "ID_SONG")
        End If
        rs.Close
        Set rs = Nothing
    End If
%>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>
                <% If nEdit > 0 Then %>Editar album<% Else %>Nuevo album<% End If %></h2>
            <p class="section-sub">Seleccione las canciones que forman parte del album</p>
        </div>
        <a class="btn btn-ghost btn-sm" href="<%=U("admin_albums.asp")%>">
            <%=Icon("chevron-left")%><span>Volver al listado</span></a>
        </div>

        <form method="post" action="<%=U("admin_albums.asp")%>" enctype="multipart/form-data">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="id" value="<%=nEdit%>">

            <div class="panel">
                <h3 class="panel-title"><%=Icon("album")%> Datos del album</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="title_album">Titulo *</label>
                        <input class="input" type="text" id="title_album" name="title_album"
                               maxlength="50" required value="<%=H(sFTitle)%>" placeholder="Ej. Rock Album">
                    </div>
                    <div class="field">
                        <label for="release_year">Anio de publicacion</label>
                        <input class="input" type="number" id="release_year" name="release_year"
                               min="1900" max="2100" value="<%=IIf2(nFYear > 0, nFYear, "")%>">
                    </div>
                </div>
                <div class="form-grid">
                    <div class="field">
                        <label for="photo">Portada <span class="hint">(imagen)</span></label>
                        <input class="input" type="file" id="photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                        <span class="hint">Ruta actual: <%=H(IIf2(Len(sFImage) > 0, sFImage, "sin imagen"))%></span>
                    </div>
                    <div class="field">
                        <label>Vista previa</label>
                        <img style="width:110px;height:110px;border-radius:12px;object-fit:cover;border:1px solid var(--pm-border-2)"
                             alt="" src="<%=H(ImgUrl(sFImage, "album", IIf2(nEdit > 0, nEdit, 99), sFTitle))%>">
                    </div>
                </div>
            </div>

            <div class="panel">
                <h3 class="panel-title"><%=Icon("music")%> Lista de canciones del album</h3>
                <%
                Set rs = DbQuery("SELECT s.ID_SONG, s.SONG_NAME + ISNULL(' (' + " _
                    & "(SELECT TOP 1 a.ARTIST_NAME FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a " _
                    & "ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) + ')', '') AS NOMBRE " _
                    & "FROM TBL_SONG s ORDER BY s.SONG_NAME")
                CheckBoxList "songs", rs, "ID_SONG", "NOMBRE", sSelSongs, "No hay canciones registradas."
                rs.Close
                Set rs = Nothing
                %>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    <%=Icon("check")%><span>Guardar album</span></button>
                <a class="btn btn-ghost" href="<%=U("admin_albums.asp")%>">Cancelar</a>
                <% If nEdit > 0 Then %>
                <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                        data-confirm="&iquest;Eliminar este album y su lista de canciones?">
                    <%=Icon("trash")%><span>Eliminar</span></button>
                <% End If %>
            </div>
        </form>
    </section>
<% Else %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Albumes</h2>
            <p class="section-sub">Operaciones de creacion, edicion y eliminacion</p>
        </div>
        <a class="btn btn-primary btn-sm" href="<%=U("admin_albums.asp?new=1")%>">
            <%=Icon("plus")%><span>Nuevo album</span></a>
        </div>

        <form method="get" action="<%=U("admin_albums.asp")%>" class="toolbar">
            <input class="filter" type="search" name="q" value="<%=H(sSearch)%>"
                   placeholder="Buscar album..." autocomplete="off">
            <button type="submit" class="btn btn-outline btn-sm"><%=Icon("search")%><span>Buscar</span></button>
            <% If Len(sSearch) > 0 Then %>
            <a class="btn btn-ghost btn-sm" href="<%=U("admin_albums.asp")%>">Limpiar</a>
            <% End If %>
        </form>

        <%
        sWhere = ""
        If Len(sSearch) > 0 Then sWhere = "WHERE TITLE_ALBUM LIKE " & SqlStr("%" & sSearch & "%")
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM " & sWhere, 0))
        nPages = PageCount(nTotal, PAGE_SIZE)
        If nPages < 1 Then nPages = 1
        If nPage < 1 Then nPage = 1
        If nPage > nPages Then nPage = nPages

        Set rs = DbQuery("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " _
            & "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS N_SONGS " _
            & "FROM TBL_ALBUM al " & sWhere & " ORDER BY al.ID_ALBUM DESC " _
            & "OFFSET " & PageOffset(nPage, PAGE_SIZE) & " ROWS FETCH NEXT " & PAGE_SIZE & " ROWS ONLY")
        %>
        <div class="table-wrap">
            <table class="table">
                <thead><tr>
                    <th style="width:64px">ID</th>
                    <th>Album</th>
                    <th class="hide-sm">Anio</th>
                    <th>Canciones</th>
                    <th class="actions">Acciones</th>
                </tr></thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr><td colspan="5" class="muted">No se encontraron albumes.</td></tr>
                <% End If %>
                <%
                Do While Not rs.EOF
                %>
                    <tr>
                        <td class="muted">#<%=CLng(rs("ID_ALBUM"))%></td>
                        <td>
                            <div class="cell-flex">
                                <img class="table-cover" alt=""
                                     src="<%=H(ImgUrl(rs("ALBUM_IMAGE_PATH"), "album", rs("ID_ALBUM"), rs("TITLE_ALBUM")))%>">
                                <span class="cell-strong"><%=H(rs("TITLE_ALBUM"))%></span>
                            </div>
                        </td>
                        <td class="hide-sm muted"><%=H(IIf2(Len(FmtYear(rs("RELEASE_YEAR"))) > 0, FmtYear(rs("RELEASE_YEAR")), "-"))%></td>
                        <td><span class="badge gray"><%=FmtNum(rs("N_SONGS"))%></span></td>
                        <td class="actions">
                            <a class="btn-icon" title="Editar"
                               href="<%=U("admin_albums.asp?edit=" & rs("ID_ALBUM"))%>"><%=Icon("edit")%></a>
                            <form method="post" style="display:inline" action="<%=U("admin_albums.asp")%>">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=CLng(rs("ID_ALBUM"))%>">
                                <button type="submit" class="btn-icon danger" title="Eliminar"
                                        data-confirm="&iquest;Eliminar este album?"><%=Icon("trash")%></button>
                            </form>
                        </td>
                    </tr>
                <%
                    rs.MoveNext
                Loop
                %>
                </tbody>
            </table>
        </div>
        <%
        rs.Close
        Set rs = Nothing
        Pager nPage, nPages, U("admin_albums.asp" & IIf2(Len(sSearch) > 0, "?q=" & UrlEnc(sSearch), ""))
        %>
    </section>
<% End If %>

<!--#include file="includes/layout_bottom.asp"-->
