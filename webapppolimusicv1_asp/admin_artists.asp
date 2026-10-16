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
'  PoliMusic - Administracion de artistas (CRUD completo)
' =====================================================================
Dim sAction, nId, sMsg, bOk, sName, sLast, sImage, nNewId
Dim sSearch, nPage, nTotal, nPages, rs, bForm, nEdit, sFName, sFLast, sFImage, sWhere

RequireAdmin()

' --------------------------- Acciones POST ---------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sAction = LCase(UFField("action", ""))
    If UFField("do_delete", "") = "1" Then sAction = "delete"
    nId = UFFieldInt("id", 0)

    If sAction = "delete" Then
        If nId > 0 Then
            sImage = DbScalar("SELECT ISNULL(MIN(ARTIST_IMAGE),'') FROM TBL_ARTIST WHERE ID_ARTIST = " & nId, "")
            DbExec "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = " & nId
            DbExec "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = " & nId
            If Len(sImage & "") > 0 Then DeleteAsset sImage
            SetFlash "El artista fue eliminado correctamente.", "success"
        End If
        Response.Redirect U("admin_artists.asp")
    End If

    If sAction = "save" Then
        sName = Left(UFField("artist_name", ""), 100)
        sLast = Left(UFField("artist_lastname", ""), 100)
        bOk = True
        sMsg = ""

        If Len(sName) = 0 Then
            sMsg = "El nombre del artista es obligatorio."
            bOk = False
        End If

        If bOk Then
            If nId > 0 Then
                If CLng(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST WHERE ID_ARTIST = " & nId, 0)) = 0 Then
                    bOk = False
                    sMsg = "El artista indicado no existe."
                Else
                    DbExec "UPDATE TBL_ARTIST SET ARTIST_NAME = " & SqlStr(sName) _
                        & ", ARTIST_LASTNAME = " & SqlStr(sLast) & " WHERE ID_ARTIST = " & nId
                End If
            Else
                nNewId = DbInsert("INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (" _
                    & SqlStr(sName) & ", " & SqlStr(sLast) & ", " & SqlStr("~/ArtistImages/artist.jpg") & ")")
                If nNewId <= 0 Then
                    bOk = False
                    sMsg = "No fue posible crear el artista."
                Else
                    nId = nNewId
                End If
            End If
        End If

        If bOk And HasUpload("photo") Then
            sImage = SaveUploaded("photo", DIR_ARTISTS, "artist_" & nId)
            If Len(sImage) > 0 Then
                DbExec "UPDATE TBL_ARTIST SET ARTIST_IMAGE = " & SqlStr(sImage) & " WHERE ID_ARTIST = " & nId
            Else
                bOk = False
                sMsg = "No fue posible guardar la imagen del artista."
            End If
        End If

        If bOk Then
            SetFlash "El artista fue guardado correctamente.", "success"
            Response.Redirect U("admin_artists.asp")
        Else
            SetFlash sMsg, "error"
        End If
    End If
End If

nEdit   = ParamInt("edit", 0)
bForm   = (nEdit > 0) Or (Param("new", "") = "1")
sSearch = ParamMax("q", 60)
nPage   = ParamInt("page", 1)

PAGE_TITLE   = "Administrar artistas"
ACTIVE_ADMIN = "artists"
%>
<!--#include file="includes/layout_top.asp"-->

<% If bForm Then
    sFName = ""
    sFLast = ""
    sFImage = ""
    If nEdit > 0 Then
        Set rs = DbQuery("SELECT TOP 1 ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE " _
            & "FROM TBL_ARTIST WHERE ID_ARTIST = " & nEdit)
        If Not rs.EOF Then
            sFName  = rs("ARTIST_NAME") & ""
            sFLast  = rs("ARTIST_LASTNAME") & ""
            sFImage = rs("ARTIST_IMAGE") & ""
        End If
        rs.Close
        Set rs = Nothing
    End If
%>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>
                <% If nEdit > 0 Then %>Editar artista<% Else %>Nuevo artista<% End If %></h2>
            <p class="section-sub">Los artistas se relacionan con las canciones a traves de TBL_PLAYER</p>
        </div>
        <a class="btn btn-ghost btn-sm" href="<%=U("admin_artists.asp")%>">
            <%=Icon("chevron-left")%><span>Volver al listado</span></a>
        </div>

        <form method="post" action="<%=U("admin_artists.asp")%>" enctype="multipart/form-data">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="id" value="<%=nEdit%>">

            <div class="panel">
                <h3 class="panel-title"><%=Icon("artist")%> Datos del artista</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="artist_name">Nombre *</label>
                        <input class="input" type="text" id="artist_name" name="artist_name"
                               maxlength="100" required value="<%=H(sFName)%>" placeholder="Ej. Juan">
                    </div>
                    <div class="field">
                        <label for="artist_lastname">Apellido</label>
                        <input class="input" type="text" id="artist_lastname" name="artist_lastname"
                               maxlength="100" value="<%=H(sFLast)%>" placeholder="Ej. Gabriel">
                    </div>
                </div>
                <div class="form-grid">
                    <div class="field">
                        <label for="photo">Fotografia <span class="hint">(imagen)</span></label>
                        <input class="input" type="file" id="photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                        <span class="hint">Ruta actual: <%=H(IIf2(Len(sFImage) > 0, sFImage, "sin imagen"))%></span>
                    </div>
                    <div class="field">
                        <label>Vista previa</label>
                        <img style="width:96px;height:96px;border-radius:50%;object-fit:cover;border:1px solid var(--pm-border-2)"
                             alt="" src="<%=H(ImgUrl(sFImage, "artist", IIf2(nEdit > 0, nEdit, 99), ArtistName(sFName, sFLast)))%>">
                    </div>
                </div>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    <%=Icon("check")%><span>Guardar artista</span></button>
                <a class="btn btn-ghost" href="<%=U("admin_artists.asp")%>">Cancelar</a>
                <% If nEdit > 0 Then %>
                <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                        data-confirm="&iquest;Eliminar este artista y sus relaciones con canciones?">
                    <%=Icon("trash")%><span>Eliminar</span></button>
                <% End If %>
            </div>
        </form>
    </section>
<% Else %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Artistas</h2>
            <p class="section-sub">Operaciones de creacion, edicion y eliminacion</p>
        </div>
        <a class="btn btn-primary btn-sm" href="<%=U("admin_artists.asp?new=1")%>">
            <%=Icon("plus")%><span>Nuevo artista</span></a>
        </div>

        <form method="get" action="<%=U("admin_artists.asp")%>" class="toolbar">
            <input class="filter" type="search" name="q" value="<%=H(sSearch)%>"
                   placeholder="Buscar artista..." autocomplete="off">
            <button type="submit" class="btn btn-outline btn-sm"><%=Icon("search")%><span>Buscar</span></button>
            <% If Len(sSearch) > 0 Then %>
            <a class="btn btn-ghost btn-sm" href="<%=U("admin_artists.asp")%>">Limpiar</a>
            <% End If %>
        </form>

        <%
        sWhere = ""
        If Len(sSearch) > 0 Then
            sWhere = "WHERE (ARTIST_NAME LIKE " & SqlStr("%" & sSearch & "%") _
                & " OR ISNULL(ARTIST_LASTNAME,'') LIKE " & SqlStr("%" & sSearch & "%") & ")"
        End If
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST " & sWhere, 0))
        nPages = PageCount(nTotal, PAGE_SIZE)
        If nPages < 1 Then nPages = 1
        If nPage < 1 Then nPage = 1
        If nPage > nPages Then nPage = nPages

        Set rs = DbQuery("SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS N_SONGS " _
            & "FROM TBL_ARTIST ar " & sWhere & " ORDER BY ar.ID_ARTIST DESC " _
            & "OFFSET " & PageOffset(nPage, PAGE_SIZE) & " ROWS FETCH NEXT " & PAGE_SIZE & " ROWS ONLY")
        %>
        <div class="table-wrap">
            <table class="table">
                <thead><tr>
                    <th style="width:64px">ID</th>
                    <th>Artista</th>
                    <th class="hide-sm">Imagen</th>
                    <th>Canciones</th>
                    <th class="actions">Acciones</th>
                </tr></thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr><td colspan="5" class="muted">No se encontraron artistas.</td></tr>
                <% End If %>
                <%
                Do While Not rs.EOF
                %>
                    <tr>
                        <td class="muted">#<%=CLng(rs("ID_ARTIST"))%></td>
                        <td>
                            <div class="cell-flex">
                                <img class="table-avatar" alt=""
                                     src="<%=H(ImgUrl(rs("ARTIST_IMAGE"), "artist", rs("ID_ARTIST"), ArtistName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME"))))%>">
                                <span class="cell-strong"><%=H(ArtistName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME")))%></span>
                            </div>
                        </td>
                        <td class="hide-sm muted small"><%=H(Trunc(rs("ARTIST_IMAGE") & "", 34))%></td>
                        <td><span class="badge gray"><%=FmtNum(rs("N_SONGS"))%></span></td>
                        <td class="actions">
                            <a class="btn-icon" title="Editar"
                               href="<%=U("admin_artists.asp?edit=" & rs("ID_ARTIST"))%>"><%=Icon("edit")%></a>
                            <form method="post" style="display:inline" action="<%=U("admin_artists.asp")%>">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=CLng(rs("ID_ARTIST"))%>">
                                <button type="submit" class="btn-icon danger" title="Eliminar"
                                        data-confirm="&iquest;Eliminar este artista?"><%=Icon("trash")%></button>
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
        Pager nPage, nPages, U("admin_artists.asp" & IIf2(Len(sSearch) > 0, "?q=" & UrlEnc(sSearch), ""))
        %>
    </section>
<% End If %>

<!--#include file="includes/layout_bottom.asp"-->
