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
'  PoliMusic - Administracion de playlists de todo el sistema
'  La relacion playlist-cancion se guarda en TBL_PLAYLIST_SONG.
' =====================================================================
Dim sAction, nId, sMsg, bOk, sTitle, nUser, nType, sPhoto, nNewId
Dim aIds, i, sSearch, nPage, nTotal, nPages, rs, bForm, nEdit
Dim sFTitle, nFUser, nFType, sFPhoto, dFCreated, sSelSongs, sWhere, sIds

RequireAdmin()

' --------------------------- Acciones POST ---------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sAction = LCase(UFField("action", ""))
    If UFField("do_delete", "") = "1" Then sAction = "delete"
    nId = UFFieldInt("id", 0)

    If sAction = "delete" Then
        If nId > 0 Then
            sPhoto = DbScalar("SELECT ISNULL(MIN(PHOTO),'') FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & nId, "")
            DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & nId
            DbExec "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & nId
            If Len(sPhoto & "") > 0 Then DeleteAsset sPhoto
            SetFlash "La playlist fue eliminada correctamente.", "success"
        End If
        Response.Redirect U("admin_playlists.asp")
    End If

    If sAction = "save" Then
        sTitle = Left(UFField("title", ""), 50)
        nUser  = UFFieldInt("id_user", 0)
        nType  = UFFieldInt("type", PLAYLIST_PUBLIC)
        bOk = True
        sMsg = ""

        If nType <> PLAYLIST_PRIVATE Then nType = PLAYLIST_PUBLIC
        If Len(sTitle) = 0 Then
            sMsg = "El titulo de la playlist es obligatorio."
            bOk = False
        ElseIf CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE ID_USER = " & nUser, 0)) = 0 Then
            sMsg = "Debe seleccionar un usuario propietario valido."
            bOk = False
        End If

        If bOk Then
            If nId > 0 Then
                If CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & nId, 0)) = 0 Then
                    bOk = False
                    sMsg = "La playlist indicada no existe."
                Else
                    DbExec "UPDATE TBL_PLAYLIST SET TITLE = " & SqlStr(sTitle) _
                        & ", ID_USER = " & nUser & ", TYPE = " & nType & " WHERE ID_PLAYLIST = " & nId
                End If
            Else
                nNewId = DbInsert("INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (" _
                    & nUser & ", " & SqlStr(sTitle) & ", GETDATE(), " & nType & ", " _
                    & SqlStr("~/PlaylistImages/default.jpg") & ")")
                If nNewId <= 0 Then
                    bOk = False
                    sMsg = "No fue posible crear la playlist."
                Else
                    nId = nNewId
                End If
            End If
        End If

        If bOk And HasUpload("photo") Then
            sPhoto = SaveUploaded("photo", DIR_PLAYLISTS, "playlist_" & nId)
            If Len(sPhoto) > 0 Then
                DbExec "UPDATE TBL_PLAYLIST SET PHOTO = " & SqlStr(sPhoto) & " WHERE ID_PLAYLIST = " & nId
            Else
                bOk = False
                sMsg = "No fue posible guardar la portada de la playlist."
            End If
        End If

        If bOk Then
            DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & nId
            aIds = IdList(UFField("songs", ""))
            For i = 0 To UBound(aIds)
                DbExec "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (" _
                    & nId & ", " & CLng(aIds(i)) & ")"
            Next
            SetFlash "La playlist fue guardada correctamente.", "success"
            Response.Redirect U("admin_playlists.asp")
        Else
            SetFlash sMsg, "error"
        End If
    End If
End If

nEdit   = ParamInt("edit", 0)
bForm   = (nEdit > 0) Or (Param("new", "") = "1")
sSearch = ParamMax("q", 60)
nPage   = ParamInt("page", 1)

PAGE_TITLE   = "Administrar playlists"
ACTIVE_ADMIN = "playlists"
%>
<!--#include file="includes/layout_top.asp"-->

<% If bForm Then
    sFTitle = ""
    nFUser = CurrentUserId()
    nFType = PLAYLIST_PUBLIC
    sFPhoto = ""
    sSelSongs = ","
    If nEdit > 0 Then
        Set rs = DbQuery("SELECT TOP 1 TITLE, ID_USER, TYPE, PHOTO FROM TBL_PLAYLIST " _
            & "WHERE ID_PLAYLIST = " & nEdit)
        If Not rs.EOF Then
            sFTitle = rs("TITLE") & ""
            nFUser  = CLng(NumOr(rs("ID_USER"), CurrentUserId()))
            nFType  = CLng(NumOr(rs("TYPE"), PLAYLIST_PUBLIC))
            sFPhoto = rs("PHOTO") & ""
            sSelSongs = SelectedIds("SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & nEdit, "ID_SONG")
        End If
        rs.Close
        Set rs = Nothing
    End If
%>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>
                <% If nEdit > 0 Then %>Editar playlist<% Else %>Nueva playlist<% End If %></h2>
            <p class="section-sub">Como administrador puede asignar la playlist a cualquier usuario</p>
        </div>
        <a class="btn btn-ghost btn-sm" href="<%=U("admin_playlists.asp")%>">
            <%=Icon("chevron-left")%><span>Volver al listado</span></a>
        </div>

        <form method="post" action="<%=U("admin_playlists.asp")%>" enctype="multipart/form-data">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="id" value="<%=nEdit%>">

            <div class="panel">
                <h3 class="panel-title"><%=Icon("playlist")%> Datos de la playlist</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="title">Titulo *</label>
                        <input class="input" type="text" id="title" name="title"
                               maxlength="50" required value="<%=H(sFTitle)%>" placeholder="Ej. Favoritas 2026">
                    </div>
                    <div class="field">
                        <label for="id_user">Propietario *</label>
                        <select class="select" id="id_user" name="id_user" required>
                            <%
                            Set rs = DbQuery("SELECT ID_USER, USERNAME FROM TBL_USER ORDER BY USERNAME")
                            Do While Not rs.EOF
                                Response.Write "<option value=""" & CLng(rs("ID_USER")) & """" _
                                    & IIf2(CLng(rs("ID_USER")) = nFUser, " selected", "") & ">" _
                                    & H(rs("USERNAME")) & "</option>"
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </select>
                    </div>
                    <div class="field">
                        <label for="type">Visibilidad</label>
                        <select class="select" id="type" name="type">
                            <%
                            Set rs = DbQuery("SELECT COD_CATALOG, [VALUE] FROM TBL_CATALOG " _
                                & "WHERE COD_CATALOG_PARENT = '2' ORDER BY COD_CATALOG")
                            Do While Not rs.EOF
                                Response.Write "<option value=""" & H(rs("COD_CATALOG")) & """" _
                                    & IIf2(CLng(NumOr(rs("COD_CATALOG"), 0)) = nFType, " selected", "") & ">" _
                                    & H(rs("VALUE")) & "</option>"
                                rs.MoveNext
                            Loop
                            rs.Close
                            Set rs = Nothing
                            %>
                        </select>
                    </div>
                    <div class="field">
                        <label for="photo">Portada <span class="hint">(imagen)</span></label>
                        <input class="input" type="file" id="photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                        <span class="hint">Ruta actual: <%=H(IIf2(Len(sFPhoto) > 0, sFPhoto, "sin imagen"))%></span>
                    </div>
                </div>
            </div>

            <div class="panel">
                <h3 class="panel-title"><%=Icon("music")%> Canciones de la playlist</h3>
                <%
                Set rs = DbQuery(SongListSql("1 = 1", "s.SONG_NAME ASC", 0))
                ' Se reutiliza la consulta de canciones para mostrar titulo + artista
                sIds = sSelSongs
                Response.Write "<div class=""checklist"">"
                If rs.EOF Then
                    Response.Write "<span class=""cl-empty"">No hay canciones registradas.</span>"
                End If
                Do While Not rs.EOF
                    Response.Write "<label><input type=""checkbox"" name=""songs"" value=""" _
                        & CLng(rs("ID_SONG")) & """" _
                        & IIf2(InStr(sIds, "," & CLng(rs("ID_SONG")) & ",") > 0, " checked", "") _
                        & "><span>" & H(rs("SONG_NAME")) _
                        & IIf2(Len(Trim(rs("ARTIST_NAME") & "")) > 0, " - " & H(rs("ARTIST_NAME")), "") _
                        & "</span></label>"
                    rs.MoveNext
                Loop
                Response.Write "</div>"
                rs.Close
                Set rs = Nothing
                %>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    <%=Icon("check")%><span>Guardar playlist</span></button>
                <a class="btn btn-ghost" href="<%=U("admin_playlists.asp")%>">Cancelar</a>
                <% If nEdit > 0 Then %>
                <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                        data-confirm="&iquest;Eliminar esta playlist y su contenido?">
                    <%=Icon("trash")%><span>Eliminar</span></button>
                <% End If %>
            </div>
        </form>
    </section>
<% Else %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Playlists del sistema</h2>
            <p class="section-sub">Todas las playlists, de todos los usuarios</p>
        </div>
        <a class="btn btn-primary btn-sm" href="<%=U("admin_playlists.asp?new=1")%>">
            <%=Icon("plus")%><span>Nueva playlist</span></a>
        </div>

        <form method="get" action="<%=U("admin_playlists.asp")%>" class="toolbar">
            <input class="filter" type="search" name="q" value="<%=H(sSearch)%>"
                   placeholder="Buscar por titulo o usuario..." autocomplete="off">
            <button type="submit" class="btn btn-outline btn-sm"><%=Icon("search")%><span>Buscar</span></button>
            <% If Len(sSearch) > 0 Then %>
            <a class="btn btn-ghost btn-sm" href="<%=U("admin_playlists.asp")%>">Limpiar</a>
            <% End If %>
        </form>

        <%
        sWhere = ""
        If Len(sSearch) > 0 Then
            sWhere = "WHERE (pl.TITLE LIKE " & SqlStr("%" & sSearch & "%") _
                & " OR u.USERNAME LIKE " & SqlStr("%" & sSearch & "%") & ")"
        End If
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST pl INNER JOIN TBL_USER u " _
            & "ON u.ID_USER = pl.ID_USER " & sWhere, 0))
        nPages = PageCount(nTotal, PAGE_SIZE)
        If nPages < 1 Then nPages = 1
        If nPage < 1 Then nPage = 1
        If nPage > nPages Then nPage = nPages

        Set rs = DbQuery("SELECT pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, pl.CREATION_DATE, " _
            & "u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps " _
            & "WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS " _
            & "FROM TBL_PLAYLIST pl INNER JOIN TBL_USER u ON u.ID_USER = pl.ID_USER " & sWhere _
            & " ORDER BY pl.ID_PLAYLIST DESC " _
            & "OFFSET " & PageOffset(nPage, PAGE_SIZE) & " ROWS FETCH NEXT " & PAGE_SIZE & " ROWS ONLY")
        %>
        <div class="table-wrap">
            <table class="table">
                <thead><tr>
                    <th style="width:64px">ID</th>
                    <th>Playlist</th>
                    <th class="hide-sm">Propietario</th>
                    <th>Tipo</th>
                    <th class="hide-sm">Creada</th>
                    <th>Canciones</th>
                    <th class="actions">Acciones</th>
                </tr></thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr><td colspan="7" class="muted">No se encontraron playlists.</td></tr>
                <% End If %>
                <%
                Do While Not rs.EOF
                %>
                    <tr>
                        <td class="muted">#<%=CLng(rs("ID_PLAYLIST"))%></td>
                        <td>
                            <div class="cell-flex">
                                <img class="table-cover" alt=""
                                     src="<%=H(ImgUrl(rs("PHOTO"), "playlist", rs("ID_PLAYLIST"), rs("TITLE")))%>">
                                <span class="cell-strong"><%=H(rs("TITLE"))%></span>
                            </div>
                        </td>
                        <td class="hide-sm"><%=H(rs("USERNAME"))%></td>
                        <td>
                            <% If CLng(NumOr(rs("TYPE"), 0)) = PLAYLIST_PRIVATE Then %>
                            <span class="badge gray">Privada</span>
                            <% Else %>
                            <span class="badge">Publica</span>
                            <% End If %>
                        </td>
                        <td class="hide-sm muted small"><%=FmtDate(rs("CREATION_DATE"))%></td>
                        <td><span class="badge gray"><%=FmtNum(rs("N_SONGS"))%></span></td>
                        <td class="actions">
                            <a class="btn-icon" title="Ver" target="_blank"
                               href="<%=U("home.asp?view=playlist&id=" & rs("ID_PLAYLIST"))%>"><%=Icon("play")%></a>
                            <a class="btn-icon" title="Editar"
                               href="<%=U("admin_playlists.asp?edit=" & rs("ID_PLAYLIST"))%>"><%=Icon("edit")%></a>
                            <form method="post" style="display:inline" action="<%=U("admin_playlists.asp")%>">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=CLng(rs("ID_PLAYLIST"))%>">
                                <button type="submit" class="btn-icon danger" title="Eliminar"
                                        data-confirm="&iquest;Eliminar esta playlist?"><%=Icon("trash")%></button>
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
        Pager nPage, nPages, U("admin_playlists.asp" & IIf2(Len(sSearch) > 0, "?q=" & UrlEnc(sSearch), ""))
        %>
    </section>
<% End If %>

<!--#include file="includes/layout_bottom.asp"-->
