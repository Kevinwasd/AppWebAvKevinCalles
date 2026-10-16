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
'  PoliMusic - Administracion de canciones (CRUD completo)
'  Gestiona tambien las relaciones N:M con artistas (TBL_PLAYER) y
'  con albumes (TBL_TRACKLIST).
' =====================================================================
Dim sAction, nId, sMsg, bOk, sName, sPathIn, nPlays, sPath, nNewId
Dim aIds, i, sSearch, nPage, nTotal, nPages, rs, bForm, nEdit
Dim sFName, sFPath, nFPlays, sSelArtists, sSelAlbums, sWhere

' Normaliza la ruta del MP3 al formato usado en la base ("~/SongFiles/x.mp3").
Function NormalizeSongPath(ByVal s)
    Dim p
    p = Trim(s & "")
    If Len(p) = 0 Then
        NormalizeSongPath = ""
        Exit Function
    End If
    p = Replace(p, "\", "/")
    Do While InStr(p, "~/") > 0
        p = Replace(p, "~/", "")
    Loop
    Do While Left(p, 1) = "/"
        p = Mid(p, 2)
    Loop
    If InStr(p, "/") = 0 Then p = DIR_SONGS & "/" & p
    NormalizeSongPath = "~/" & p
End Function

RequireAdmin()

' --------------------------- Acciones POST ---------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sAction = LCase(UFField("action", ""))
    If UFField("do_delete", "") = "1" Then sAction = "delete"
    nId = UFFieldInt("id", 0)

    If sAction = "delete" Then
        If nId > 0 Then
            sPath = DbScalar("SELECT ISNULL(MIN(SONG_PATH),'') FROM TBL_SONG WHERE ID_SONG = " & nId, "")
            DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & nId
            DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & nId
            DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = " & nId
            DbExec "DELETE FROM TBL_SONG WHERE ID_SONG = " & nId
            If Len(sPath & "") > 0 Then DeleteAsset sPath
            SetFlash "La cancion fue eliminada junto con sus relaciones.", "success"
        End If
        Response.Redirect U("admin_songs.asp")
    End If

    If sAction = "save" Then
        sName    = Left(UFField("song_name", ""), 50)
        sPathIn  = Left(UFField("song_path", ""), 255)
        nPlays   = UFFieldInt("plays", 0)
        sMsg     = ""
        bOk      = True

        If Len(sName) = 0 Then
            sMsg = "El nombre de la cancion es obligatorio."
            bOk = False
        End If

        sPath = NormalizeSongPath(sPathIn)
        If bOk And Len(sPath) = 0 And Not HasUpload("mp3") Then
            sMsg = "Debe indicar la ruta del archivo MP3 o adjuntar un archivo de audio."
            bOk = False
        End If

        If bOk Then
            If nId > 0 Then
                If CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = " & nId, 0)) = 0 Then
                    bOk = False
                    sMsg = "La cancion indicada no existe."
                Else
                    If Len(sPath) = 0 Then
                        sPath = DbScalar("SELECT ISNULL(MIN(SONG_PATH),'') FROM TBL_SONG WHERE ID_SONG = " & nId, "")
                    End If
                    DbExec "UPDATE TBL_SONG SET SONG_NAME = " & SqlStr(sName) _
                        & ", SONG_PATH = " & SqlStr(sPath) _
                        & ", PLAYS = " & nPlays & " WHERE ID_SONG = " & nId
                End If
            Else
                If Len(sPath) = 0 Then sPath = "~/SongFiles/pendiente.mp3"
                nNewId = DbInsert("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (" _
                    & SqlStr(sName) & ", " & SqlStr(sPath) & ", " & nPlays & ")")
                If nNewId <= 0 Then
                    bOk = False
                    sMsg = "No fue posible crear la cancion."
                Else
                    nId = nNewId
                End If
            End If
        End If

        ' Archivo de audio adjunto
        If bOk And HasUpload("mp3") Then
            sPath = SaveUploaded("mp3", DIR_SONGS, "song_" & nId)
            If Len(sPath) > 0 Then
                DbExec "UPDATE TBL_SONG SET SONG_PATH = " & SqlStr(sPath) & " WHERE ID_SONG = " & nId
            Else
                bOk = False
                sMsg = "No fue posible guardar el archivo de audio (revise el tamano maximo permitido)."
            End If
        End If

        ' Relacion con artistas
        If bOk Then
            DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & nId
            aIds = IdList(UFField("artists", ""))
            For i = 0 To UBound(aIds)
                DbExec "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (" _
                    & nId & ", " & CLng(aIds(i)) & ")"
            Next

            DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & nId
            aIds = IdList(UFField("albums", ""))
            For i = 0 To UBound(aIds)
                DbExec "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" _
                    & CLng(aIds(i)) & ", " & nId & ")"
            Next
        End If

        If bOk Then
            SetFlash "La cancion fue guardada correctamente.", "success"
            Response.Redirect U("admin_songs.asp")
        Else
            SetFlash sMsg, "error"
        End If
    End If
End If

' --------------------------- Parametros -----------------------------
nEdit = ParamInt("edit", 0)
bForm = (nEdit > 0) Or (Param("new", "") = "1")
sSearch = ParamMax("q", 60)
nPage = ParamInt("page", 1)

PAGE_TITLE   = "Administrar canciones"
ACTIVE_ADMIN = "songs"
%>
<!--#include file="includes/layout_top.asp"-->

<% If bForm Then
    sFName = ""
    sFPath = ""
    nFPlays = 0
    sSelArtists = ","
    sSelAlbums = ","

    If nEdit > 0 Then
        Set rs = DbQuery("SELECT TOP 1 ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS " _
            & "FROM TBL_SONG WHERE ID_SONG = " & nEdit)
        If Not rs.EOF Then
            sFName  = rs("SONG_NAME") & ""
            sFPath  = rs("SONG_PATH") & ""
            nFPlays = CLng(NumOr(rs("PLAYS"), 0))
            sSelArtists = SelectedIds("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = " & nEdit, "ID_ARTIST")
            sSelAlbums = SelectedIds("SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = " & nEdit, "ID_ALBUM")
        End If
        rs.Close
        Set rs = Nothing
    End If
%>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>
                <% If nEdit > 0 Then %>Editar cancion<% Else %>Nueva cancion<% End If %></h2>
            <p class="section-sub">Los campos marcados con * son obligatorios</p>
        </div>
        <a class="btn btn-ghost btn-sm" href="<%=U("admin_songs.asp")%>">
            <%=Icon("chevron-left")%><span>Volver al listado</span></a>
        </div>

        <form method="post" action="<%=U("admin_songs.asp")%>" enctype="multipart/form-data">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="id" value="<%=nEdit%>">

            <div class="panel">
                <h3 class="panel-title"><%=Icon("music")%> Datos de la cancion</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="song_name">Nombre de la cancion *</label>
                        <input class="input" type="text" id="song_name" name="song_name"
                               maxlength="50" required value="<%=H(sFName)%>"
                               placeholder="Ej. Adventure">
                    </div>
                    <div class="field">
                        <label for="plays">Reproducciones</label>
                        <input class="input" type="number" id="plays" name="plays" min="0"
                               value="<%=nFPlays%>">
                    </div>
                </div>
                <div class="form-grid">
                    <div class="field">
                        <label for="song_path">Ruta del archivo (base de datos)</label>
                        <input class="input" type="text" id="song_path" name="song_path"
                               maxlength="255" value="<%=H(sFPath)%>"
                               placeholder="~/SongFiles/bensound-adventure.mp3">
                        <span class="hint">Puede escribir una ruta existente o adjuntar un archivo abajo.</span>
                    </div>
                    <div class="field">
                        <label for="mp3">Adjuntar MP3 <span class="hint">(opcional, reemplaza la ruta)</span></label>
                        <input class="input" type="file" id="mp3" name="mp3" accept="audio/mpeg,audio/mp3,audio/*">
                    </div>
                </div>
            </div>

            <div class="panel">
                <h3 class="panel-title"><%=Icon("artist")%> Artistas que interpretan la cancion</h3>
                <%
                Set rs = DbQuery("SELECT ID_ARTIST, ARTIST_NAME + ISNULL(' ' + ARTIST_LASTNAME,'') AS NOMBRE " _
                    & "FROM TBL_ARTIST ORDER BY ARTIST_NAME")
                CheckBoxList "artists", rs, "ID_ARTIST", "NOMBRE", sSelArtists, _
                    "No hay artistas registrados."
                rs.Close
                Set rs = Nothing
                %>
            </div>

            <div class="panel">
                <h3 class="panel-title"><%=Icon("album")%> Albumes que incluyen la cancion</h3>
                <%
                Set rs = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM AS NOMBRE FROM TBL_ALBUM ORDER BY TITLE_ALBUM")
                CheckBoxList "albums", rs, "ID_ALBUM", "NOMBRE", sSelAlbums, _
                    "No hay albumes registrados."
                rs.Close
                Set rs = Nothing
                %>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    <%=Icon("check")%><span>Guardar cancion</span></button>
                <a class="btn btn-ghost" href="<%=U("admin_songs.asp")%>">Cancelar</a>
                <% If nEdit > 0 Then %>
                <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                        data-confirm="&iquest;Eliminar esta cancion y todas sus relaciones?">
                    <%=Icon("trash")%><span>Eliminar</span></button>
                <% End If %>
            </div>
        </form>
    </section>
<% Else %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Canciones</h2>
            <p class="section-sub">Operaciones de creacion, edicion y eliminacion</p>
        </div>
        <a class="btn btn-primary btn-sm" href="<%=U("admin_songs.asp?new=1")%>">
            <%=Icon("plus")%><span>Nueva cancion</span></a>
        </div>

        <form method="get" action="<%=U("admin_songs.asp")%>" class="toolbar">
            <input class="filter" type="search" name="q" value="<%=H(sSearch)%>"
                   placeholder="Buscar por nombre o ruta..." autocomplete="off">
            <button type="submit" class="btn btn-outline btn-sm"><%=Icon("search")%><span>Buscar</span></button>
            <% If Len(sSearch) > 0 Then %>
            <a class="btn btn-ghost btn-sm" href="<%=U("admin_songs.asp")%>">Limpiar</a>
            <% End If %>
        </form>

        <%
        sWhere = ""
        If Len(sSearch) > 0 Then
            sWhere = "WHERE (SONG_NAME LIKE " & SqlStr("%" & sSearch & "%") _
                & " OR SONG_PATH LIKE " & SqlStr("%" & sSearch & "%") & ")"
        End If
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG " & sWhere, 0))
        nPages = PageCount(nTotal, PAGE_SIZE)
        If nPages < 1 Then nPages = 1
        If nPage < 1 Then nPage = 1
        If nPage > nPages Then nPage = nPages

        Set rs = DbQuery("SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG) AS N_ART, " _
            & "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_SONG = s.ID_SONG) AS N_ALB, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = s.ID_SONG) AS N_PL " _
            & "FROM TBL_SONG s " & sWhere _
            & " ORDER BY s.ID_SONG DESC " _
            & "OFFSET " & PageOffset(nPage, PAGE_SIZE) & " ROWS FETCH NEXT " & PAGE_SIZE & " ROWS ONLY")
        %>
        <div class="table-wrap">
            <table class="table">
                <thead><tr>
                    <th style="width:64px">ID</th>
                    <th>Cancion</th>
                    <th class="hide-sm">Ruta</th>
                    <th>Repro.</th>
                    <th class="hide-sm">Relaciones</th>
                    <th class="actions">Acciones</th>
                </tr></thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr><td colspan="6" class="muted">No se encontraron canciones.</td></tr>
                <% End If %>
                <%
                Do While Not rs.EOF
                %>
                    <tr>
                        <td class="muted">#<%=CLng(rs("ID_SONG"))%></td>
                        <td>
                            <div class="cell-flex">
                                <img class="table-cover" alt=""
                                     src="<%=H(ImgUrl("", "song", rs("ID_SONG"), FirstLetter(rs("SONG_NAME"))))%>">
                                <span class="cell-strong"><%=H(rs("SONG_NAME"))%></span>
                            </div>
                        </td>
                        <td class="hide-sm muted small"><%=H(Trunc(rs("SONG_PATH") & "", 38))%></td>
                        <td><%=FmtNum(rs("PLAYS"))%></td>
                        <td class="hide-sm muted small">
                            <%=FmtNum(rs("N_ART"))%> artistas &middot;
                            <%=FmtNum(rs("N_ALB"))%> albumes &middot;
                            <%=FmtNum(rs("N_PL"))%> playlists
                        </td>
                        <td class="actions">
                            <a class="btn-icon" title="Editar"
                               href="<%=U("admin_songs.asp?edit=" & rs("ID_SONG"))%>"><%=Icon("edit")%></a>
                            <form method="post" style="display:inline" action="<%=U("admin_songs.asp")%>">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=CLng(rs("ID_SONG"))%>">
                                <button type="submit" class="btn-icon danger" title="Eliminar"
                                        data-confirm="&iquest;Eliminar la cancion y todas sus relaciones?"><%=Icon("trash")%></button>
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
        Pager nPage, nPages, U("admin_songs.asp" & IIf2(Len(sSearch) > 0, "?q=" & UrlEnc(sSearch), ""))
        %>
    </section>
<% End If %>

<!--#include file="includes/layout_bottom.asp"-->
