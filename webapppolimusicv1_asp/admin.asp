<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<%
' =====================================================================
'  PoliMusic - Panel de administracion
' =====================================================================
Dim rs, nSongs, nArtists, nAlbums, nUsers, nPlaylists, nPlays, nAdmins
Dim fso, sRoot, aFolders, i, sFolderPath, nOk, f

RequireAdmin()

PAGE_TITLE   = "Panel de administracion"
ACTIVE_ADMIN = "dashboard"

nSongs     = CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG", 0))
nArtists   = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", 0))
nAlbums    = CLng(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", 0))
nUsers     = CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER", 0))
nAdmins    = CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = " & ROLE_ADMIN, 0))
nPlaylists = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", 0))
nPlays     = CLng(DbScalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", 0))
%>
<!--#include file="includes/layout_top.asp"-->

    <section class="hero">
        <span class="hero-kicker">Administracion</span>
        <h1>Panel de control</h1>
        <p>Desde aqui puede gestionar el catalogo completo de <%=H(APP_NAME)%>:
           canciones, artistas, albumes, playlists y usuarios.</p>
        <div class="hero-actions">
            <a class="btn btn-primary" href="<%=U("admin_songs.asp?new=1")%>">
                <%=Icon("plus")%><span>Nueva cancion</span></a>
            <a class="btn btn-outline" href="<%=U("admin_users.asp")%>">
                <%=Icon("users")%><span>Gestionar usuarios</span></a>
        </div>
    </section>

    <div class="stats">
        <div class="stat"><span class="ic"><%=Icon("music")%></span>
            <span><b><%=FmtNum(nSongs)%></b><span>Canciones</span></span></div>
        <div class="stat"><span class="ic"><%=Icon("artist")%></span>
            <span><b><%=FmtNum(nArtists)%></b><span>Artistas</span></span></div>
        <div class="stat"><span class="ic"><%=Icon("album")%></span>
            <span><b><%=FmtNum(nAlbums)%></b><span>Albumes</span></span></div>
        <div class="stat"><span class="ic"><%=Icon("playlist")%></span>
            <span><b><%=FmtNum(nPlaylists)%></b><span>Playlists</span></span></div>
        <div class="stat"><span class="ic"><%=Icon("users")%></span>
            <span><b><%=FmtNum(nUsers)%></b><span>Usuarios (<%=nAdmins%> admin)</span></span></div>
        <div class="stat"><span class="ic"><%=Icon("chart")%></span>
            <span><b><%=FmtNum(nPlays)%></b><span>Reproducciones</span></span></div>
    </div>

    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Accesos rapidos</h2>
            <p class="section-sub">Operaciones CRUD disponibles</p>
        </div></div>
        <div class="grid">
            <a class="card" href="<%=U("admin_songs.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("music")%></div>
                <p class="card-title">Canciones</p>
                <p class="card-sub">Crear, editar y eliminar</p></a>
            <a class="card" href="<%=U("admin_artists.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("artist")%></div>
                <p class="card-title">Artistas</p>
                <p class="card-sub">Crear, editar y eliminar</p></a>
            <a class="card" href="<%=U("admin_albums.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("album")%></div>
                <p class="card-title">Albumes</p>
                <p class="card-sub">Incluye lista de canciones</p></a>
            <a class="card" href="<%=U("admin_playlists.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("playlist")%></div>
                <p class="card-title">Playlists</p>
                <p class="card-sub">Todas las del sistema</p></a>
            <a class="card" href="<%=U("admin_users.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("users")%></div>
                <p class="card-title">Usuarios</p>
                <p class="card-sub">Roles y credenciales</p></a>
            <a class="card" href="<%=U("diagnostics.asp")%>">
                <div class="card-media" style="display:grid;place-items:center;color:var(--pm-primary-300)">
                    <%=Icon("shield")%></div>
                <p class="card-title">Diagnostico</p>
                <p class="card-sub">Estado del entorno</p></a>
        </div>
    </section>

    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Canciones mas escuchadas</h2>
        </div></div>
        <%
        Set rs = DbQuery(SongListSql("1 = 1", "ISNULL(s.PLAYS,0) DESC, s.ID_SONG ASC", 8))
        SongTable rs, "adminTop"
        rs.Close
        Set rs = Nothing
        %>
    </section>

    <section class="section">
        <div class="panel">
            <h3 class="panel-title"><%=Icon("folder")%> Estado de los recursos</h3>
            <div class="form-grid">
                <%
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                sRoot = AppPhysRoot()
                aFolders = Array(DIR_SONGS, DIR_USERS, DIR_PLAYLISTS, DIR_ALBUMS, DIR_ARTISTS, DIR_UPLOADS)
                For i = 0 To UBound(aFolders)
                    sFolderPath = sRoot & DIR_ASSETS & "\" & aFolders(i)
                    If fso.FolderExists(sFolderPath) Then
                        nOk = 0
                        For Each f In fso.GetFolder(sFolderPath).Files
                            nOk = nOk + 1
                        Next
                        Response.Write "<div><p class=""small mb-0""><b style=""color:var(--pm-text)"">assets/" _
                            & H(aFolders(i)) & "/</b><br><span class=""muted"">" & nOk _
                            & " archivo(s) &middot; <span class=""badge"">Disponible</span></span></p></div>"
                    Else
                        Response.Write "<div><p class=""small mb-0""><b style=""color:var(--pm-text)"">assets/" _
                            & H(aFolders(i)) & "/</b><br><span class=""muted""><span class=""badge danger"">No encontrada</span></span></p></div>"
                    End If
                Next
                Set fso = Nothing
                %>
            </div>
        </div>
    </section>

<!--#include file="includes/layout_bottom.asp"-->
